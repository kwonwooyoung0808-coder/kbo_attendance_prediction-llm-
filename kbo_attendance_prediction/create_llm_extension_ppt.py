from __future__ import annotations

from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import html


OUT = Path("KBO_LLM_스마트티켓팅_확장발표.pptx")
W, H = 12192000, 6858000


def emu(x: float) -> int:
    return int(x * 914400)


def esc(text: str) -> str:
    return html.escape(str(text), quote=False)


def color(hex_color: str) -> str:
    return hex_color.replace("#", "").upper()


def fill_xml(hex_color: str | None) -> str:
    if not hex_color:
        return "<a:noFill/>"
    return f"<a:solidFill><a:srgbClr val=\"{color(hex_color)}\"/></a:solidFill>"


def line_xml(hex_color: str | None = None, width: int = 12700) -> str:
    if not hex_color:
        return "<a:ln><a:noFill/></a:ln>"
    return f"<a:ln w=\"{width}\"><a:solidFill><a:srgbClr val=\"{color(hex_color)}\"/></a:solidFill></a:ln>"


def run_xml(text: str, size: int = 24, bold: bool = False, hex_color: str = "#102033") -> str:
    b = ' b="1"' if bold else ""
    return (
        f"<a:r><a:rPr lang=\"ko-KR\" sz=\"{size * 100}\"{b}>"
        f"<a:solidFill><a:srgbClr val=\"{color(hex_color)}\"/></a:solidFill>"
        f"</a:rPr><a:t>{esc(text)}</a:t></a:r>"
    )


def textbox(shape_id: int, x: float, y: float, w: float, h: float, paragraphs: list[str],
            size: int = 22, bold: bool = False, font_color: str = "#102033",
            align: str = "l", bullet: bool = False, name: str = "Text") -> str:
    ps = []
    for text in paragraphs:
        if bullet:
            ps.append(
                f"<a:p><a:pPr marL=\"260000\" indent=\"-140000\" algn=\"{align}\">"
                f"<a:buChar char=\"•\"/></a:pPr>{run_xml(text, size, bold, font_color)}</a:p>"
            )
        else:
            ps.append(f"<a:p><a:pPr algn=\"{align}\"/>{run_xml(text, size, bold, font_color)}</a:p>")
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{name} {shape_id}\"/>"
        f"<p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr><a:xfrm><a:off x=\"{emu(x)}\" y=\"{emu(y)}\"/>"
        f"<a:ext cx=\"{emu(w)}\" cy=\"{emu(h)}\"/></a:xfrm>"
        f"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/>{line_xml()}</p:spPr>"
        f"<p:txBody><a:bodyPr wrap=\"square\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\"/>"
        f"<a:lstStyle/>{''.join(ps)}</p:txBody></p:sp>"
    )


def rect(shape_id: int, x: float, y: float, w: float, h: float, fill: str, line: str | None = None,
         radius: str = "roundRect", text: str | None = None, text_size: int = 20,
         text_color: str = "#102033", bold: bool = False) -> str:
    body = ""
    if text is not None:
        body = (
            f"<p:txBody><a:bodyPr anchor=\"ctr\" wrap=\"square\"/><a:lstStyle/>"
            f"<a:p><a:pPr algn=\"ctr\"/>{run_xml(text, text_size, bold, text_color)}</a:p></p:txBody>"
        )
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"Box {shape_id}\"/>"
        f"<p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr><a:xfrm><a:off x=\"{emu(x)}\" y=\"{emu(y)}\"/>"
        f"<a:ext cx=\"{emu(w)}\" cy=\"{emu(h)}\"/></a:xfrm>"
        f"<a:prstGeom prst=\"{radius}\"><a:avLst/></a:prstGeom>"
        f"{fill_xml(fill)}{line_xml(line)}</p:spPr>{body}</p:sp>"
    )


def line(shape_id: int, x1: float, y1: float, x2: float, y2: float, hex_color: str = "#9AA9B5") -> str:
    return (
        f"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id=\"{shape_id}\" name=\"Line {shape_id}\"/>"
        f"<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>"
        f"<a:xfrm><a:off x=\"{emu(min(x1, x2))}\" y=\"{emu(min(y1, y2))}\"/>"
        f"<a:ext cx=\"{emu(abs(x2 - x1))}\" cy=\"{emu(abs(y2 - y1))}\"/></a:xfrm>"
        f"<a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom>{line_xml(hex_color, 22000)}</p:spPr></p:cxnSp>"
    )


def slide_xml(title: str, page: int, total: int, elements: list[str], section: str = "LLM 확장 프로젝트") -> str:
    base = [
        rect(2, 0, 0, 13.333, 7.5, "#F6F8F6", None, "rect"),
        rect(3, 0, 0, 13.333, 0.16, "#1F6F58", None, "rect"),
        textbox(4, 0.55, 0.36, 8.2, 0.35, [section], 10, True, "#6B7A88"),
        textbox(5, 0.55, 0.78, 10.5, 0.55, [title], 24, True, "#102033"),
        line(6, 0.55, 1.42, 12.75, 1.42, "#D8E0DD"),
        textbox(7, 11.72, 6.93, 1.0, 0.22, [f"{page:02d} / {total:02d}"], 9, False, "#697887", "r"),
    ]
    return wrap_slide("".join(base + elements))


def wrap_slide(content: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/>
      <a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr>
    {content}
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def cover() -> str:
    elems = [
        rect(10, 0, 0, 13.333, 7.5, "#F4F7F4", None, "rect"),
        rect(11, 0.8, 0.78, 11.7, 5.85, "#FFFFFF", "#D8E0DD"),
        rect(12, 0.8, 0.78, 11.7, 0.18, "#1F6F58", None, "rect"),
        textbox(13, 1.25, 1.38, 10.4, 1.2, ["KBO 관중 수 예측 기반", "AI 스마트 티켓팅 가이드"], 34, True, "#102033"),
        textbox(14, 1.28, 2.82, 9.8, 0.52, ["딥러닝 관중 예측 모델에서 로컬 LLM/RAG 상담 서비스로 확장"], 18, False, "#4C5D6B"),
        rect(15, 1.28, 3.65, 2.2, 0.55, "#E7F1ED", "#C8DAD2", text="Dense 예측", text_size=16, text_color="#164C43", bold=True),
        rect(16, 3.75, 3.65, 2.2, 0.55, "#F3ECE2", "#E1D2BE", text="FAISS 검색", text_size=16, text_color="#704A2D", bold=True),
        rect(17, 6.22, 3.65, 2.2, 0.55, "#E9F1F5", "#CBDCE5", text="Ollama LLM", text_size=16, text_color="#285A70", bold=True),
        rect(18, 8.69, 3.65, 2.2, 0.55, "#F4E5DF", "#E2CBC0", text="티켓팅 상담", text_size=16, text_color="#813C2D", bold=True),
        textbox(19, 1.28, 5.35, 5.8, 0.42, ["2026 KBO 정규시즌 일정 기반 확장 프로젝트"], 15, True, "#1F6F58"),
        textbox(20, 1.28, 5.82, 5.8, 0.28, ["발표자: 권우영"], 12, False, "#687789"),
    ]
    return wrap_slide("".join(elems))


def build_slides() -> list[str]:
    total = 14
    slides = [cover()]
    slides.append(slide_xml("목차", 2, total, [
        textbox(10, 0.9, 1.75, 5.4, 4.6, [
            "1. 프로젝트 배경과 기획 의도",
            "2. 기존 딥러닝 프로젝트 요약",
            "3. 데이터 구성과 정제 기준",
            "4. 모델 성능 비교",
            "5. 서비스 확장 방향",
            "6. LLM/RAG 구조",
            "7. 화면 시연 및 한계점",
        ], 22, False, "#102033", bullet=True),
        rect(11, 7.1, 1.85, 4.7, 3.8, "#FFFFFF", "#D8E0DD"),
        textbox(12, 7.55, 2.28, 3.8, 0.45, ["발표 핵심"], 20, True, "#164C43", "c"),
        textbox(13, 7.55, 3.0, 3.8, 1.7, ["관중 수 예측 모델을 단순 분석에서 끝내지 않고, 실제 예매 의사결정을 돕는 LLM 상담 서비스로 확장"], 20, True, "#102033", "c"),
    ]))
    slides.append(slide_xml("프로젝트 배경과 문제 정의", 3, total, [
        textbox(10, 0.85, 1.85, 5.55, 3.8, [
            "KBO 경기는 날짜, 구장, 상대팀, 공휴일 여부에 따라 관중 수 차이가 큼",
            "사용자는 경기 선택 시 혼잡도와 예매 난이도를 미리 알고 싶어함",
            "기존 예측 모델 결과를 실제 서비스 기능으로 연결할 필요가 있음",
        ], 21, False, "#102033", bullet=True),
        rect(11, 7.05, 1.9, 4.9, 2.8, "#E7F1ED", "#C8DAD2"),
        textbox(12, 7.45, 2.34, 4.05, 0.5, ["기획 의도"], 22, True, "#164C43", "c"),
        textbox(13, 7.45, 3.07, 4.05, 0.9, ["관중 예측을 이용해 “어떤 경기를 언제 예매해야 하는지” 알려주는 스마트 티켓팅 가이드 구현"], 19, True, "#102033", "c"),
    ]))
    slides.append(slide_xml("기존 딥러닝 프로젝트 요약", 4, total, [
        rect(10, 0.75, 1.82, 3.55, 3.0, "#FFFFFF", "#D8E0DD", text="입력 특성\n홈팀, 원정팀, 구장\n월, 요일, 주말\n공휴일, 라이벌전, 시즌", text_size=18, text_color="#102033", bold=False),
        rect(11, 4.9, 1.82, 3.55, 3.0, "#FFFFFF", "#D8E0DD", text="예측 모델\nDense\nLSTM\nGRU", text_size=20, text_color="#102033", bold=True),
        rect(12, 9.05, 1.82, 3.55, 3.0, "#FFFFFF", "#D8E0DD", text="출력 결과\n예상 관중 수\n구장 점유율\n혼잡도 안내", text_size=18, text_color="#102033", bold=False),
        textbox(13, 1.05, 5.35, 10.9, 0.55, ["이번 확장에서는 예측 결과를 LLM이 해석하여 자연어 티켓팅 상담으로 연결"], 20, True, "#1F6F58", "c"),
    ]))
    slides.append(slide_xml("데이터 구성과 정제 기준", 5, total, [
        textbox(10, 0.85, 1.82, 5.7, 3.7, [
            "2024~2026 KBO 경기 일정 및 관중 데이터 사용",
            "2026년 4월 15일 이전 실제 경기에는 관중 수 포함",
            "우천 취소 경기는 CSV에서 제외하여 실제 열린 경기만 학습/표시",
            "2026년 9월 6일까지 남은 일정은 관중 수 없이 예측 대상으로 유지",
        ], 20, False, "#102033", bullet=True),
        rect(11, 7.1, 1.85, 4.9, 3.3, "#F3ECE2", "#E1D2BE"),
        textbox(12, 7.5, 2.22, 4.1, 0.45, ["데이터 처리 포인트"], 20, True, "#704A2D", "c"),
        textbox(13, 7.5, 3.0, 4.1, 1.2, ["취소 경기 제거\n실제 관중 데이터 반영\n미래 일정은 예측 서비스용으로 유지"], 20, True, "#102033", "c"),
    ]))
    slides.append(slide_xml("모델 성능 비교", 6, total, [
        textbox(10, 0.85, 1.7, 4.4, 0.45, ["MAE 기준 Dense 모델이 가장 안정적"], 20, True, "#164C43"),
        rect(11, 1.1, 2.65, 2.0, 2.15, "#1F6F58", None, "rect"),
        rect(12, 4.05, 2.15, 2.0, 2.65, "#9BB7AA", None, "rect"),
        rect(13, 7.0, 2.16, 2.0, 2.64, "#B7C8C1", None, "rect"),
        textbox(14, 0.95, 4.98, 2.3, 0.4, ["Dense\nMAE 3,863명"], 16, True, "#102033", "c"),
        textbox(15, 3.9, 4.98, 2.3, 0.4, ["LSTM\nMAE 4,763명"], 16, True, "#102033", "c"),
        textbox(16, 6.85, 4.98, 2.3, 0.4, ["GRU\nMAE 4,758명"], 16, True, "#102033", "c"),
        rect(17, 9.85, 1.95, 2.6, 3.2, "#FFFFFF", "#D8E0DD"),
        textbox(18, 10.15, 2.35, 2.0, 0.45, ["선택 이유"], 18, True, "#164C43", "c"),
        textbox(19, 10.12, 3.1, 2.05, 1.2, ["서비스에서는 설명 가능성과 안정성을 고려해 Dense 예측값을 기본으로 사용"], 15, False, "#102033", "c"),
    ]))
    slides.append(slide_xml("확장 방향: 예측 모델에서 상담 서비스로", 7, total, [
        rect(10, 0.8, 2.25, 2.35, 0.9, "#E7F1ED", "#C8DAD2", text="딥러닝 예측", text_size=18, text_color="#164C43", bold=True),
        rect(11, 3.85, 2.25, 2.35, 0.9, "#F3ECE2", "#E1D2BE", text="일정 검색", text_size=18, text_color="#704A2D", bold=True),
        rect(12, 6.9, 2.25, 2.35, 0.9, "#E9F1F5", "#CBDCE5", text="LLM 해석", text_size=18, text_color="#285A70", bold=True),
        rect(13, 9.95, 2.25, 2.35, 0.9, "#F4E5DF", "#E2CBC0", text="예매 전략", text_size=18, text_color="#813C2D", bold=True),
        line(14, 3.15, 2.7, 3.85, 2.7), line(15, 6.2, 2.7, 6.9, 2.7), line(16, 9.25, 2.7, 9.95, 2.7),
        textbox(17, 1.0, 4.0, 11.0, 0.8, ["기존 프로젝트의 산출물인 ‘관중 수 예측’을 사용자가 이해하기 쉬운 ‘티켓팅 의사결정’으로 변환"], 22, True, "#102033", "c"),
    ]))
    slides.append(slide_xml("스마트 티켓팅 가이드 화면", 8, total, [
        rect(10, 0.8, 1.75, 11.7, 4.55, "#FFFFFF", "#C8DAD2"),
        textbox(11, 1.1, 2.05, 11.1, 0.55, ["캡처 삽입 필요: Streamlit의 ‘스마트 티켓팅 가이드’ 화면"], 22, True, "#1F6F58", "c"),
        textbox(12, 1.4, 3.0, 10.5, 1.7, ["날짜 선택 → 해당일 경기 리스트 → 예측 점유율 → 혼잡도 배지\n예측값이 구장 정원을 넘으면 100%로 보정하고 ‘전좌석 매진 예상’으로 표시"], 20, False, "#102033", "c"),
    ]))
    slides.append(slide_xml("LLM/RAG 확장 구조", 9, total, [
        rect(10, 0.8, 1.85, 2.55, 0.78, "#FFFFFF", "#D8E0DD", text="사용자 질문", text_size=17, text_color="#102033", bold=True),
        rect(11, 3.85, 1.85, 2.55, 0.78, "#FFFFFF", "#D8E0DD", text="CSV 일정 문서화", text_size=17, text_color="#102033", bold=True),
        rect(12, 6.9, 1.85, 2.55, 0.78, "#FFFFFF", "#D8E0DD", text="FAISS 검색", text_size=17, text_color="#102033", bold=True),
        rect(13, 9.95, 1.85, 2.55, 0.78, "#FFFFFF", "#D8E0DD", text="Ollama 답변", text_size=17, text_color="#102033", bold=True),
        line(14, 3.35, 2.24, 3.85, 2.24), line(15, 6.4, 2.24, 6.9, 2.24), line(16, 9.45, 2.24, 9.95, 2.24),
        textbox(17, 1.0, 3.35, 5.3, 2.0, [
            "수업 실습의 CSV RAG 구조를 프로젝트 데이터에 적용",
            "경기 일정 텍스트를 벡터화하고 유사 경기 검색",
            "검색 결과와 Dense 예측값을 LLM 프롬프트에 제공",
        ], 18, False, "#102033", bullet=True),
        textbox(18, 7.0, 3.45, 4.8, 1.6, ["실시간 좌석 정보는 제공하지 않고, 예측 기반 예매 전략으로 한정"], 20, True, "#704A2D", "c"),
    ]))
    slides.append(slide_xml("AI 티켓팅 상담 화면", 10, total, [
        rect(10, 0.8, 1.75, 11.7, 4.55, "#FFFFFF", "#C8DAD2"),
        textbox(11, 1.1, 2.05, 11.1, 0.55, ["캡처 삽입 필요: Streamlit의 ‘AI 티켓팅 상담’ 화면"], 22, True, "#1F6F58", "c"),
        textbox(12, 1.4, 3.0, 10.5, 1.7, ["예시 질문 버튼, 자연어 입력창, AI 답변, 추천 근거 카드, 후보 경기 테이블이 함께 보이도록 캡처"], 20, False, "#102033", "c"),
    ]))
    slides.append(slide_xml("예시 질문과 추천 근거", 11, total, [
        textbox(10, 0.85, 1.82, 5.85, 3.4, [
            "5월 5일에 가족이랑 가기 좋은 경기 추천해줘",
            "롯데 경기 중 매진 가능성 높은 경기 알려줘",
            "잠실 경기 중 덜 붐비는 경기 추천해줘",
            "공휴일 경기 중 예매 빨리 해야 하는 경기 알려줘",
        ], 19, False, "#102033", bullet=True),
        rect(11, 7.15, 1.85, 4.8, 3.25, "#FFFFFF", "#D8E0DD"),
        textbox(12, 7.55, 2.25, 4.0, 0.45, ["추천 근거"], 20, True, "#164C43", "c"),
        textbox(13, 7.55, 2.95, 4.0, 1.45, ["예상 점유율\n공휴일 여부\n라이벌전 여부\n주말/평일 조건\n전좌석 매진 보정"], 18, False, "#102033", "c"),
    ]))
    slides.append(slide_xml("기존 딥러닝 대비 확장 포인트", 12, total, [
        rect(10, 0.8, 1.8, 5.4, 3.8, "#FFFFFF", "#D8E0DD"),
        textbox(11, 1.2, 2.18, 4.6, 0.45, ["기존 딥러닝 프로젝트"], 20, True, "#164C43", "c"),
        textbox(12, 1.2, 3.0, 4.6, 1.35, ["정형 데이터 기반 관중 수 예측\n모델 성능 비교 중심\n수치 결과 확인"], 19, False, "#102033", "c"),
        rect(13, 7.1, 1.8, 5.4, 3.8, "#FFFFFF", "#D8E0DD"),
        textbox(14, 7.5, 2.18, 4.6, 0.45, ["LLM 확장 버전"], 20, True, "#704A2D", "c"),
        textbox(15, 7.5, 3.0, 4.6, 1.35, ["자연어 질문 기반 상담\n일정 검색 + 예측값 해석\n사용자 예매 전략 제공"], 19, False, "#102033", "c"),
    ]))
    slides.append(slide_xml("한계점과 개선 방향", 13, total, [
        textbox(10, 0.85, 1.82, 5.6, 3.6, [
            "실시간 잔여 좌석과 실제 판매 현황은 반영하지 못함",
            "예측 모델 평균 오차가 약 3,863명 수준으로 존재",
            "날씨, 선발 투수, 팀 순위, 이벤트 정보는 추가 반영 가능",
            "향후 구단별 예매 안내 문서 RAG와 음성 상담 기능으로 확장 가능",
        ], 20, False, "#102033", bullet=True),
        rect(11, 7.1, 1.95, 4.8, 2.9, "#F3ECE2", "#E1D2BE"),
        textbox(12, 7.5, 2.35, 4.0, 0.45, ["발전 방향"], 20, True, "#704A2D", "c"),
        textbox(13, 7.5, 3.08, 4.0, 1.05, ["예측 정확도 향상 + 실시간 정보 연동 + 사용자 맞춤 추천"], 20, True, "#102033", "c"),
    ]))
    slides.append(slide_xml("결론", 14, total, [
        textbox(10, 1.15, 1.95, 10.9, 1.3, ["딥러닝 관중 예측 모델을 실제 사용자가 이해하고 활용할 수 있는 AI 티켓팅 상담 서비스로 확장했다."], 27, True, "#102033", "c"),
        rect(11, 1.35, 3.75, 3.1, 0.78, "#E7F1ED", "#C8DAD2", text="예측", text_size=20, text_color="#164C43", bold=True),
        rect(12, 5.1, 3.75, 3.1, 0.78, "#F3ECE2", "#E1D2BE", text="검색", text_size=20, text_color="#704A2D", bold=True),
        rect(13, 8.85, 3.75, 3.1, 0.78, "#E9F1F5", "#CBDCE5", text="상담", text_size=20, text_color="#285A70", bold=True),
        textbox(14, 1.2, 5.25, 10.8, 0.55, ["수업 실습의 딥러닝 모델링과 LLM/RAG 실습을 하나의 서비스 흐름으로 연결"], 20, False, "#4C5D6B", "c"),
    ]))
    return slides


def empty_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"""


def content_types(n: int) -> str:
    overrides = [
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    overrides += [
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, n + 1)
    ]
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
{''.join(overrides)}
</Types>"""


def presentation_xml(n: int) -> str:
    ids = "".join([f'<p:sldId id="{255+i}" r:id="rId{i+1}"/>' for i in range(1, n + 1)])
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>{ids}</p:sldIdLst>
<p:sldSz cx="{W}" cy="{H}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""


def presentation_rels(n: int) -> str:
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    rels += [
        f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, n + 1)
    ]
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{''.join(rels)}</Relationships>"""


def master_xml() -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/><a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>"""


def layout_xml() -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/><a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""


def theme_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="KBO Calm">
<a:themeElements><a:clrScheme name="KBO"><a:dk1><a:srgbClr val="102033"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="164C43"/></a:dk2><a:lt2><a:srgbClr val="F6F8F6"/></a:lt2><a:accent1><a:srgbClr val="1F6F58"/></a:accent1><a:accent2><a:srgbClr val="BF7A4F"/></a:accent2><a:accent3><a:srgbClr val="9BB7AA"/></a:accent3><a:accent4><a:srgbClr val="285A70"/></a:accent4><a:accent5><a:srgbClr val="704A2D"/></a:accent5><a:accent6><a:srgbClr val="D8E0DD"/></a:accent6><a:hlink><a:srgbClr val="285A70"/></a:hlink><a:folHlink><a:srgbClr val="704A2D"/></a:folHlink></a:clrScheme><a:fontScheme name="KBO"><a:majorFont><a:latin typeface="맑은 고딕"/><a:ea typeface="맑은 고딕"/></a:majorFont><a:minorFont><a:latin typeface="맑은 고딕"/><a:ea typeface="맑은 고딕"/></a:minorFont></a:fontScheme><a:fmtScheme name="KBO"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"""


def root_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def core_xml() -> str:
    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>KBO LLM 스마트 티켓팅 확장 발표</dc:title><dc:creator>Codex</dc:creator>
<cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>"""


def app_xml(n: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Codex</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>{n}</Slides></Properties>"""


def build() -> None:
    slides = build_slides()
    with ZipFile(OUT, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types(len(slides)))
        z.writestr("_rels/.rels", root_rels())
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", app_xml(len(slides)))
        z.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels(len(slides)))
        z.writestr("ppt/slideMasters/slideMaster1.xml", master_xml())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>""")
        z.writestr("ppt/slideLayouts/slideLayout1.xml", layout_xml())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>""")
        z.writestr("ppt/theme/theme1.xml", theme_xml())
        for i, slide in enumerate(slides, 1):
            z.writestr(f"ppt/slides/slide{i}.xml", slide)
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", empty_rels())
    print(OUT.resolve())


if __name__ == "__main__":
    build()
