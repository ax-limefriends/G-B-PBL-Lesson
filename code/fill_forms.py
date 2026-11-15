"""fill_forms.py — 마스터 시트(xlsx) 한 장으로 인허가 서류(hwpx) 여러 개를 한 번에 채운다.

구조
  master.xlsx [프로젝트] 키·값  +  [서류별] 접수번호·제출일·수신처
        │  읽기 → 형식 맞추기(날짜 YYYY-MM-DD, kW·㎡ 숫자)
        ▼
  1) 사전 검증: 빈 값·날짜 형식·날짜 선후        → 문제는 [검증리포트]에 기록 (생성은 계속)
  2) 템플릿 {{키}} 치환 → out/*.hwpx 저장        → 빈 값은 [확인필요]
  3) 사후 검증: 생성 파일을 다시 열어 {{ }} 잔여·값 반영 여부 대조
  * 미리보기(PrvText/PrvImage)에 남은 옛 값을 지운다 — 파일 탐색기 미리보기로 옛 고객 정보가 새는 것 방지

사용: python3 fill_forms.py master.xlsx <템플릿폴더> <출력폴더> [--strip-lineseg]
  --strip-lineseg : 한글에서 글자가 겹쳐 보일 때만 사용(줄 배치 캐시 삭제, 한글이 다시 계산하도록 하는 용도 — 효과는 실물 확인 필요)
"""
import datetime as dt
import html
import io
import re
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook

MISSING = "[확인필요]"
DATE_KEYS = {"허가일", "공사시작일", "준공예정일", "공사완료일", "사용전검사확인일", "사업개시일", "제출일"}
PERMIT_DOCS = {"A2", "A3", "A4", "A5", "A6"}  # 전기사업 허가번호를 적는 서류


def fmt(key, v):
    """엑셀 값 → 서류에 들어갈 글자."""
    if v is None or str(v).strip() == "":
        return None
    if isinstance(v, (dt.datetime, dt.date)):  # 엑셀이 날짜로 바꿔 버린 경우도 받아 준다
        return v.strftime("%Y-%m-%d")
    if key == "설비용량":
        n = float(str(v).replace(",", "").replace("kW", "").strip())
        return f"{n:,.3f}".rstrip("0").rstrip(".")  # 998.4 / 1,000 / 99.75
    if key == "부지면적":
        n = float(str(v).replace(",", "").replace("㎡", "").strip())
        return f"{n:,.2f}".rstrip("0").rstrip(".")  # 12,500
    return str(v).strip()


def read_master(path):
    wb = load_workbook(path, data_only=True)
    proj, raw = {}, {}
    for k, v, *_ in wb["프로젝트"].iter_rows(min_row=2, values_only=True):
        if k:
            raw[k] = v
    docs = []
    for row in wb["서류별"].iter_rows(min_row=2, values_only=True):
        if row[0]:
            docs.append(dict(zip(["id", "name", "template", "접수번호", "제출일", "수신처"], row)))
    return raw, docs


def check(proj, docs):
    """사전 검증 — 사람이 자주 틀리는 것들."""
    msgs = []
    for k, v in proj.items():
        if v is None:
            msgs.append(f"[빈 값] 프로젝트.{k}")
        elif k in DATE_KEYS and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
            msgs.append(f"[날짜 형식] 프로젝트.{k} = {v} (YYYY-MM-DD 로)")

    def d(s):
        try:
            return dt.date.fromisoformat(s)
        except (TypeError, ValueError):
            return None

    order = [("공사시작일", "준공예정일"), ("공사시작일", "공사완료일"), ("공사완료일", "사용전검사확인일"),
             ("사용전검사확인일", "사업개시일")]
    for a, b in order:
        if d(proj.get(a)) and d(proj.get(b)) and d(proj[a]) > d(proj[b]):
            msgs.append(f"[날짜 순서] {a}({proj[a]}) 가 {b}({proj[b]}) 보다 늦음")
    for doc in docs:
        sub = d(doc["제출일"])
        if doc["id"] in PERMIT_DOCS and sub and d(proj.get("허가일")) and sub < d(proj["허가일"]):
            msgs.append(f"[날짜 순서] {doc['id']} {doc['name']} 제출일({doc['제출일']})이 전기사업 허가일({proj['허가일']})보다 빠른데 "
                        f"허가번호를 적게 되어 있음 → 확인필요")
    try:
        kw = float(proj["설비용량"].replace(",", ""))
        if kw <= 3000:
            msgs.append(f"[참고] 설비용량 {proj['설비용량']} kW ≤ 3,000 kW → 발전사업허가 신청처는 시·도지사 "
                        f"(전기사업법 시행규칙 제4조, 위임 조례는 지자체 확인) — A1 수신처 확인")
    except (KeyError, AttributeError, ValueError):
        msgs.append("[빈 값/형식] 설비용량은 kW 숫자만")
    return msgs


def fill_xml(xml, values, strip_lineseg):
    used = set()

    def put(m):
        k = m.group(1)
        used.add(k)
        return values.get(k) or MISSING  # 빈 값은 지어내지 않고 [확인필요]

    def sub_t(m):  # <hp:t> 안의 글자만 바꾼다 (태그 구조는 그대로)
        text = re.sub(r"\{\{(.+?)\}\}", put, html.unescape(m.group(1)))
        return "<hp:t>" + html.escape(text, quote=False) + "</hp:t>"  # & < > 이스케이프

    out = re.sub(r"<hp:t>(.*?)</hp:t>", sub_t, xml, flags=re.S)
    if strip_lineseg:
        out = re.sub(r"<hp:linesegarray>.*?</hp:linesegarray>", "", out, flags=re.S)
    return out, used


def plain_text(xml):
    return "\r\n".join(html.unescape(t) for t in re.findall(r"<hp:t>(.*?)</hp:t>", xml, flags=re.S))


def blank_png(png_bytes):
    """미리보기 그림을 같은 크기의 흰 이미지로 (옛 값이 그림으로 남는 것 방지)."""
    try:
        from PIL import Image
        w, h = Image.open(io.BytesIO(png_bytes)).size
        buf = io.BytesIO()
        Image.new("RGB", (w, h), "white").save(buf, "PNG")
        return buf.getvalue()
    except Exception:  # PIL 이 없으면 원본 유지 — 리포트에 남김
        return None


def write_hwpx(src, dst, new_xml, notes):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "Contents/section0.xml":
                data = new_xml.encode("utf-8")
            elif info.filename == "Preview/PrvText.txt":
                data = plain_text(new_xml).encode("utf-8")
            elif info.filename == "Preview/PrvImage.png":
                b = blank_png(data)
                if b is None:
                    notes.append(f"[주의] {dst.name}: 미리보기 그림을 지우지 못함(PIL 없음)")
                data = b or data
            comp = zipfile.ZIP_STORED if info.filename == "mimetype" else zipfile.ZIP_DEFLATED
            zout.writestr(info, data, compress_type=comp)  # mimetype 첫 항목·무압축 유지


def verify(dst, values, used):
    """사후 검증: 생성 파일을 다시 열어 대조."""
    with zipfile.ZipFile(dst) as z:
        text = plain_text(z.read("Contents/section0.xml").decode("utf-8"))
    errs = []
    if "{{" in text:
        errs.append(f"[치환 누락] {dst.name}: {set(re.findall(r'{{(.+?)}}', text))}")
    for k in sorted(used):
        v = values.get(k)
        if v and v not in text:
            errs.append(f"[반영 안 됨] {dst.name}: {k}={v}")
    blanks = sorted(k for k in used if not values.get(k))
    return errs, blanks


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    strip = "--strip-lineseg" in sys.argv
    master, tdir, odir = Path(args[0]), Path(args[1]), Path(args[2])
    odir.mkdir(parents=True, exist_ok=True)
    raw, docs = read_master(master)
    proj = {k: fmt(k, v) for k, v in raw.items()}
    for doc in docs:
        doc["제출일"] = fmt("제출일", doc["제출일"])
    report = ["# 검증 리포트", f"생성 시각: {dt.datetime.now():%Y-%m-%d %H:%M}", "", "## 1. 사전 검증"]
    report += check(proj, docs) or ["문제 없음"]
    report += ["", "## 2. 서류별 결과"]
    total_err = 0
    for doc in docs:
        src = tdir / doc["template"]
        if not src.exists():
            report.append(f"- {doc['id']} {doc['name']}: ⚠ 템플릿 없음 ({doc['template']})")
            total_err += 1
            continue
        values = dict(proj, 접수번호=fmt("", doc["접수번호"]), 제출일=doc["제출일"], 수신처=fmt("", doc["수신처"]))
        with zipfile.ZipFile(src) as z:
            xml = z.read("Contents/section0.xml").decode("utf-8")
        new_xml, used = fill_xml(xml, values, strip)
        dst = odir / f"{doc['id']}_{doc['name']}_{values['접수번호'] or '번호없음'}.hwpx"
        notes = []
        write_hwpx(src, dst, new_xml, notes)
        errs, blanks = verify(dst, values, used)
        total_err += len(errs)
        state = "✅ 정상" if not errs else "⚠ 확인"
        report.append(f"- {dst.name}: {state} · 채운 항목 {len(used)}개" + (f" · [확인필요] {blanks}" if blanks else ""))
        report += [f"  {m}" for m in errs + notes]
    report += ["", "## 3. 사람이 꼭 할 일", "- 한글에서 열어 표·줄 배치·서명란 확인", "- [참고]·[날짜 순서] 항목은 담당자가 판단"]
    (odir / "검증리포트.txt").write_text("\n".join(report), encoding="utf-8")
    print("\n".join(report))
    sys.exit(1 if total_err else 0)


if __name__ == "__main__":
    main()
