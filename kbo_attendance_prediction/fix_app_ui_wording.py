from pathlib import Path


APP_PATH = Path(__file__).resolve().parent / "app.py"

REPLACEMENTS = {
    "AI 스마트 티켓팅 가이드": "AI 경기 추천 가이드",
    "스마트 티켓팅 가이드": "경기 추천 가이드",
    "AI 티켓팅 상담": "AI 경기 추천 상담",
    "티켓팅 상담": "경기 추천 상담",
    "원하는 조건에 맞는 경기와 예매 전략을 추천합니다.": "원하는 조건에 맞는 경기와 관람 판단을 추천합니다.",
    "자연어 예매 상담": "자연어 경기 추천 상담",
    "Ollama가 예매 전략을 정리하는 중입니다...": "Ollama가 추천 근거를 정리하는 중입니다...",
    "예매 안내": "관람 안내",
    "온라인 예매 권장": "사전 확인 권장",
    "현장 구매 검토 가능": "여유 관람 가능",
    "전좌석 매진 예상": "매우 혼잡 예상",
    "경기별 좌석난이도와 예매 우선순위를 한눈에 안내합니다.": "경기별 혼잡도와 관람 추천 우선순위를 한눈에 안내합니다.",
    "어떤 경기의 좌석난을 확인하고 싶으신가요?": "어떤 경기의 관람 혼잡도를 확인하고 싶으신가요?",
}


def main() -> None:
    text = APP_PATH.read_text(encoding="utf-8")
    for old, new in REPLACEMENTS.items():
        text = text.replace(old, new)
    APP_PATH.write_text(text, encoding="utf-8")

    remaining = [
        key
        for key in ["스마트 티켓팅", "스마트티켓팅", "AI 티켓팅", "티켓팅 상담"]
        if key in text
    ]
    print(f"updated={APP_PATH}")
    print(f"remaining={remaining}")


if __name__ == "__main__":
    main()
