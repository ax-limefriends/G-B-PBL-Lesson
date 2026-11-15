"""make_templates.py — 값이 채워진 hwpx 서식을 {{키}} 자리표시 템플릿으로 바꾼다.

구조
  1) hwpx(ZIP) 안의 Contents/section0.xml 을 읽는다
  2) <hp:t> 글자 조각을 순서대로 보며
     - 표의 '항목명' 칸 다음 칸(값)은 항목명으로 키를 정한다 (같은 날짜라도 의미가 다르면 다른 키)
     - 머리부·서명란·본문 문장은 정해진 규칙으로 치환한다
  3) 원본 예시 값이 하나라도 남으면 멈춘다 (치환 누락 = 나중에 옛 값이 그대로 제출되는 사고)

사용: python3 make_templates.py <서식폴더> <템플릿출력폴더>
"""
import html
import re
import sys
import zipfile
from pathlib import Path

# ── 예시 서식에 들어 있던 값 (치환 대상 + 잔여 검사용) ──
SAMPLE = {
    "신청인주소": "충남 가상군 가상면 가상리 123-4",
    "사업위치": "충남 가상군 가상면 가상리 123-4",
    "발전소명": "가상솔라 1호 태양광발전소",
    "계통접속점": "가상변전소 22.9kV 배전선로",
    "허가번호": "가상전기-2025-0123",
    "사업자등록번호": "000-00-00001",
    "연락처": "000-0000-0001",
    "상호": "가상솔라(주)",
    "대표자": "김가상",
    "지번": "가상리 123-4",
    "부지면적": "12,500",
    "설비용량": "998.4",
    "발전방식": "태양광",
    "행위종류": "토지 형질 변경 (태양광발전시설 설치)",
    "주요설비": "태양광 모듈, 인버터, 접속반, 변압기, 계통연계 설비",
}
RECIPIENTS = ["가상전력(주) 가상지사", "가상군수"]  # 수신처 예시 값

# 표 항목명 → 키 (서식마다 같은 값을 다른 이름으로 부른다)
LABEL_KEY = {
    "상호(사업주)": "상호", "상호": "상호", "대표자": "대표자", "사업자등록번호": "사업자등록번호",
    "연락처": "연락처", "주소": "신청인주소", "사업명": "발전소명", "발전소명": "발전소명",
    "발전방식": "발전방식", "설비용량": "설비용량", "사업 위치": "사업위치", "위치": "사업위치",
    "발전소 소재지": "사업위치", "공사 위치": "사업위치", "소재지": "사업위치", "지번": "지번",
    "부지 면적": "부지면적", "면적": "부지면적", "전력계통 접속점": "계통접속점", "연계 계통": "계통접속점",
    "준공 예정일": "준공예정일", "상업운전 예정일": "준공예정일", "전기사업 허가번호": "허가번호",
    "허가일": "허가일", "공사 완료일": "공사완료일", "사용전검사 확인일": "사용전검사확인일",
    "사업 개시일": "사업개시일", "행위 종류": "행위종류", "주요 설비": "주요설비",
}
DATE = r"\d{4}-\d{2}-\d{2}"
# 일반 문장 치환 순서: 긴 값 먼저 (주소 안에 지번이 들어 있으므로)
GENERIC = ["사업위치", "발전소명", "계통접속점", "허가번호", "상호", "대표자", "지번", "부지면적", "설비용량"]


def templatize_value(label: str, text: str) -> str:
    """표의 값 칸: 항목명으로 키를 정해 치환."""
    if label == "공사 기간":  # 날짜 2개가 한 칸에 → 순서로 의미를 구분
        dates = re.findall(DATE, text)
        if len(dates) == 2:
            return text.replace(dates[0], "{{공사시작일}}", 1).replace(dates[1], "{{준공예정일}}", 1)
        if len(dates) == 1:
            return text.replace(dates[0], "{{준공예정일}}", 1)
        return text
    key = LABEL_KEY.get(label)
    if not key:
        return text
    if key in SAMPLE and SAMPLE[key] in text:
        return text.replace(SAMPLE[key], "{{%s}}" % key)
    if re.fullmatch(DATE, text.strip()):
        return "{{%s}}" % key
    return text


def templatize_text(text: str) -> str:
    """머리부·서명란·본문 문장."""
    text = re.sub(r"접수번호: \S+", "접수번호: {{접수번호}}", text)
    text = re.sub(r"(신청일|신고일): " + DATE, r"\1: {{제출일}}", text)
    if re.fullmatch(DATE, text.strip()):  # 서명란 위 날짜 한 줄
        return "{{제출일}}"
    for r in RECIPIENTS:
        text = text.replace(r, "{{수신처}}")
    for key in GENERIC:
        text = text.replace(SAMPLE[key], "{{%s}}" % key)
    return text


def templatize_xml(xml: str) -> str:
    labels = set(LABEL_KEY) | {"공사 기간"}
    prev = {"t": None}

    def sub(m):
        raw = html.unescape(m.group(1))
        new = templatize_value(prev["t"], raw) if prev["t"] in labels else raw
        new = templatize_text(new)
        prev["t"] = raw
        return "<hp:t>%s</hp:t>" % html.escape(new, quote=False)

    return re.sub(r"<hp:t>(.*?)</hp:t>", sub, xml, flags=re.S)


def leftovers(xml: str) -> list[str]:
    """치환되지 않고 남은 예시 값 목록 (비어 있어야 정상)."""
    texts = " ".join(html.unescape(t) for t in re.findall(r"<hp:t>(.*?)</hp:t>", xml, flags=re.S))
    found = [v for v in list(SAMPLE.values()) + RECIPIENTS if v in texts and v not in ("태양광",)]
    found += re.findall(DATE, texts)
    return found


def rewrite_zip(src: Path, dst: Path, new_section: str):
    """mimetype 을 첫 항목·무압축으로 유지해야 한글이 정상 인식한다."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "Contents/section0.xml":
                data = new_section.encode("utf-8")
            comp = zipfile.ZIP_STORED if info.filename == "mimetype" else zipfile.ZIP_DEFLATED
            zout.writestr(info, data, compress_type=comp)


def main(src_dir: str, out_dir: str):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    bad = 0
    for f in sorted(Path(src_dir).glob("*.hwp[xt]")):
        with zipfile.ZipFile(f) as z:
            xml = z.read("Contents/section0.xml").decode("utf-8")
        new = templatize_xml(xml)
        left = leftovers(new)
        keys = sorted(set(re.findall(r"\{\{(.+?)\}\}", new)))
        name = re.sub(r"_P\d{4}-\d{2}", "", f.stem) + "_템플릿.hwpx"
        rewrite_zip(f, out / name, new)
        print(f"{name}: 키 {len(keys)}개 {keys}")
        if left:
            bad += 1
            print(f"  ⚠ 남은 예시 값: {left}")
    if bad:
        sys.exit(f"치환 누락이 있는 서식 {bad}개 — 위 ⚠ 를 확인하세요")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
