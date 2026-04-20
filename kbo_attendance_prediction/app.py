from __future__ import annotations

from pathlib import Path
import json
import os
import pickle
import re

import altair as alt
import faiss
import numpy as np
import ollama
import pandas as pd
import streamlit as st
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import normalize
from tensorflow.keras.models import load_model

BASE_DIR = Path(__file__).resolve().parent
ARTIFACT_DIR = BASE_DIR / "artifacts"
MODEL_DIR = BASE_DIR / "models"
DATA_PATH = BASE_DIR / "data" / "kbo_attendance.csv"

DENSE_MODEL_PATH = MODEL_DIR / "attendance_dense_model.keras"
LSTM_MODEL_PATH = MODEL_DIR / "attendance_lstm_model.keras"
GRU_MODEL_PATH = MODEL_DIR / "attendance_gru_model.keras"
ENCODERS_FILE = ARTIFACT_DIR / "encoders.pkl"
SCALER_FILE = ARTIFACT_DIR / "scaler.pkl"
FEATURES_FILE = ARTIFACT_DIR / "feature_cols.json"
LSTM_SCALER_FILE = ARTIFACT_DIR / "lstm_target_scaler.pkl"
GRU_SCALER_FILE = ARTIFACT_DIR / "gru_target_scaler.pkl"
OLLAMA_MODEL = "gemma2:latest"
OLLAMA_HOST = os.getenv("OLLAMA_HOST", "").strip()
CURRENT_SEASON = 2026
CURRENT_DATE = pd.Timestamp("2026-04-16")
DENSE_MODEL_MAE = 3863
LLM_EXAMPLE_QUESTIONS = [
    "5월 5일에 가족이랑 가기 좋은 경기 추천해줘",
    "롯데 경기 중 매진 가능성 높은 경기 알려줘",
    "잠실 경기 중 덜 붐비는 경기 추천해줘",
    "공휴일 경기 중 예매 빨리 해야 하는 경기 알려줘",
]

TEAM_TO_STADIUM = {
    "LG": "잠실",
    "두산": "잠실",
    "삼성": "대구",
    "KIA": "광주",
    "KT": "수원",
    "SSG": "문학",
    "롯데": "사직",
    "한화": "대전",
    "NC": "창원",
    "키움": "고척",
}
STADIUM_CAPACITY = {
    "잠실": 23750,
    "대구": 24000,
    "광주": 20500,
    "문학": 23000,
    "창원": 18128,
    "수원": 18700,
    "고척": 16000,
    "대전": 17000,
    "사직": 22758,
    "한밭": 12000,
}
RIVAL_MATCHES = {tuple(sorted(["LG", "두산"]))}
WEEKDAY_TO_NUM = {"월": 0, "화": 1, "수": 2, "목": 3, "금": 4, "토": 5, "일": 6}
NUM_TO_WEEKDAY = {v: k for k, v in WEEKDAY_TO_NUM.items()}
TEAMS = ["LG", "두산", "삼성", "KIA", "KT", "SSG", "롯데", "한화", "NC", "키움"]

st.set_page_config(page_title="KBO 관중 수 예측", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Gowun+Dodum&family=IBM+Plex+Sans+KR:wght@400;500;600;700;800&display=swap');

    :root {
        --ink: #102033;
        --muted: #687789;
        --line: rgba(31, 57, 86, 0.12);
        --field: #f2f7f1;
        --grass: #1f6f58;
        --grass-dark: #164c43;
        --clay: #bf7a4f;
        --sand: #f7efe2;
        --sky: #f4f8fb;
        --card: rgba(255, 255, 255, 0.88);
    }

    .stApp {
        background:
            radial-gradient(circle at 12% 8%, rgba(191, 122, 79, 0.10), transparent 24%),
            radial-gradient(circle at 86% 2%, rgba(31, 111, 88, 0.10), transparent 26%),
            linear-gradient(180deg, #f6f8f6 0%, #eef5f2 48%, #fbfaf7 100%);
        color: var(--ink);
        font-family: 'IBM Plex Sans KR', 'Gowun Dodum', sans-serif;
    }

    .main .block-container {
        max-width: 1180px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #173f38 0%, #1f5f4f 100%);
        border-right: 1px solid rgba(255, 255, 255, 0.10);
    }

    section[data-testid="stSidebar"] * {
        color: rgba(255, 255, 255, 0.94);
    }

    div[role="radiogroup"] label {
        border-radius: 14px;
        padding: 8px 10px;
        transition: all 0.18s ease;
    }

    div[role="radiogroup"] label:hover {
        background: rgba(255, 255, 255, 0.10);
    }

    .hero {
        position: relative;
        overflow: hidden;
        padding: 30px 34px;
        border-radius: 24px;
        background:
            linear-gradient(135deg, rgba(22, 76, 67, 0.96), rgba(31, 111, 88, 0.92)),
            linear-gradient(180deg, rgba(255,255,255,0.10), transparent);
        color: white;
        box-shadow: 0 18px 38px rgba(31, 57, 86, 0.16);
        margin-bottom: 24px;
        isolation: isolate;
    }

    .hero:before {
        content: "";
        position: absolute;
        inset: auto -6% -50% -6%;
        height: 110px;
        background:
            repeating-linear-gradient(90deg, rgba(255,255,255,0.10) 0 1px, transparent 1px 56px),
            linear-gradient(180deg, rgba(255,255,255,0.08), transparent);
        transform: rotate(-1deg);
        z-index: -1;
    }

    .hero:after {
        content: "";
        position: absolute;
        right: 28px;
        bottom: 22px;
        width: 92px;
        height: 44px;
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 999px 999px 12px 12px;
        opacity: 0.42;
    }

    .hero h1 {
        margin: 0 0 10px 0;
        font-size: clamp(1.9rem, 3.2vw, 2.65rem);
        font-weight: 800;
        letter-spacing: -0.045em;
    }

    .hero p {
        margin: 0;
        max-width: 760px;
        font-size: 1.05rem;
        line-height: 1.6;
        opacity: 0.92;
    }

    .panel {
        background: var(--card);
        border: 1px solid var(--line);
        border-radius: 20px;
        padding: 20px 22px 16px 22px;
        box-shadow: 0 12px 28px rgba(31, 57, 86, 0.08);
        backdrop-filter: blur(10px);
        margin-bottom: 20px;
    }

    .mini-card {
        position: relative;
        overflow: hidden;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(248,251,248,0.94) 100%);
        border: 1px solid rgba(31, 111, 88, 0.12);
        border-radius: 18px;
        padding: 18px 18px 15px 18px;
        box-shadow: 0 10px 24px rgba(31, 57, 86, 0.07);
        min-height: 116px;
        transition: transform 0.18s ease, box-shadow 0.18s ease;
    }

    .mini-card:before {
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 4px;
        background: linear-gradient(90deg, var(--grass), #8fb8a7, var(--clay));
    }

    .mini-card:hover {
        transform: translateY(-1px);
        box-shadow: 0 14px 30px rgba(31, 57, 86, 0.10);
    }

    .mini-label {
        color: var(--muted);
        font-size: 0.9rem;
        font-weight: 600;
        margin-bottom: 8px;
    }

    .mini-value {
        color: var(--grass-dark);
        font-size: 1.65rem;
        font-weight: 800;
        line-height: 1.2;
        letter-spacing: -0.035em;
    }

    .mini-sub {
        color: var(--muted);
        font-size: 0.86rem;
        margin-top: 6px;
    }

    .section-title {
        display: flex;
        align-items: center;
        gap: 10px;
        font-size: 1.36rem;
        font-weight: 800;
        color: var(--grass-dark);
        margin: 4px 0 16px 0;
        letter-spacing: -0.035em;
    }

    .section-title:before {
        content: "";
        width: 10px;
        height: 10px;
        border-radius: 999px;
        background: var(--clay);
        box-shadow: 0 0 0 4px rgba(191, 122, 79, 0.12);
    }

    .game-card {
        position: relative;
        overflow: hidden;
        background: linear-gradient(135deg, rgba(255,255,255,0.98), rgba(247,250,248,0.94));
        border-radius: 20px;
        padding: 22px;
        border: 1px solid rgba(31, 111, 88, 0.12);
        margin-bottom: 16px;
        box-shadow: 0 12px 28px rgba(31, 57, 86, 0.08);
    }

    .game-card:before {
        content: "";
        position: absolute;
        top: 0;
        bottom: 0;
        left: 0;
        width: 5px;
        background: linear-gradient(180deg, var(--grass), #9bb7aa);
    }

    .ticket-badge {
        display: inline-block;
        padding: 8px 14px;
        border-radius: 999px;
        font-size: 0.85rem;
        font-weight: 700;
        margin-top: 8px;
        box-shadow: inset 0 0 0 1px rgba(255,255,255,0.55);
    }

    .badge-red { background: #f9e2de; color: #9a3027; }
    .badge-orange { background: #f8ead2; color: #8d5b22; }
    .badge-green { background: #e4f2e9; color: #1f6f58; }
    .badge-blue { background: #e8f1f7; color: #2f667f; }

    div.stButton > button {
        border: 0;
        border-radius: 14px;
        background: linear-gradient(135deg, #1f6f58, #2b8068);
        color: white;
        font-weight: 700;
        box-shadow: 0 8px 20px rgba(31, 111, 88, 0.16);
        transition: all 0.18s ease;
    }

    div.stButton > button:hover {
        color: white;
        transform: translateY(-1px);
        box-shadow: 0 12px 24px rgba(31, 111, 88, 0.22);
        filter: brightness(1.02);
    }

    div.stButton > button:active {
        transform: translateY(0);
    }

    [data-testid="stMetric"] {
        background: rgba(255,255,255,0.72);
        border: 1px solid rgba(31, 111, 88, 0.10);
        border-radius: 16px;
        padding: 14px 16px;
    }

    [data-testid="stMetricValue"] {
        color: var(--grass-dark);
        font-weight: 800;
    }

    [data-testid="stAlert"] {
        border-radius: 18px;
        border: 1px solid rgba(31, 111, 88, 0.10);
        box-shadow: 0 8px 22px rgba(31, 57, 86, 0.06);
    }

    [data-testid="stSelectbox"] label,
    [data-testid="stSelectbox"] label p {
        color: var(--ink) !important;
        font-weight: 700 !important;
    }

    [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        background: rgba(255, 255, 255, 0.96) !important;
        border-color: rgba(31, 111, 88, 0.20) !important;
        color: var(--ink) !important;
    }

    [data-testid="stSelectbox"] div[data-baseweb="select"] span,
    [data-testid="stSelectbox"] div[data-baseweb="select"] input,
    [data-testid="stSelectbox"] div[data-baseweb="select"] svg {
        color: var(--ink) !important;
        fill: var(--ink) !important;
    }

    textarea, input {
        border-radius: 16px !important;
    }

    [data-testid="stDataFrame"] {
        border-radius: 20px;
        overflow: hidden;
        box-shadow: 0 10px 24px rgba(31, 57, 86, 0.07);
    }

    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #2b8068, #c7a96b, #bf7a4f);
    }

    @media (max-width: 700px) {
        .hero {
            padding: 26px 22px;
            border-radius: 24px;
        }
        .panel, .game-card {
            padding: 18px;
            border-radius: 22px;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def load_data() -> pd.DataFrame:
    df = pd.read_csv(DATA_PATH)
    df["date"] = pd.to_datetime(df["date"])
    return df


@st.cache_resource
def load_dense_artifacts():
    if not (DENSE_MODEL_PATH.exists() and ENCODERS_FILE.exists() and SCALER_FILE.exists() and FEATURES_FILE.exists()):
        return None, None, None, None
    model = load_model(DENSE_MODEL_PATH, compile=False)
    with open(ENCODERS_FILE, "rb") as f:
        encoders = pickle.load(f)
    with open(SCALER_FILE, "rb") as f:
        scaler = pickle.load(f)
    feature_cols = json.loads(FEATURES_FILE.read_text(encoding="utf-8"))
    return model, encoders, scaler, feature_cols


@st.cache_resource
def load_sequence_artifacts():
    lstm_model = load_model(LSTM_MODEL_PATH, compile=False) if LSTM_MODEL_PATH.exists() else None
    gru_model = load_model(GRU_MODEL_PATH, compile=False) if GRU_MODEL_PATH.exists() else None
    lstm_scaler = None
    gru_scaler = None
    if LSTM_SCALER_FILE.exists():
        with open(LSTM_SCALER_FILE, "rb") as f:
            lstm_scaler = pickle.load(f)
    if GRU_SCALER_FILE.exists():
        with open(GRU_SCALER_FILE, "rb") as f:
            gru_scaler = pickle.load(f)
    return lstm_model, gru_model, lstm_scaler, gru_scaler


def render_hero(title: str, body: str):
    st.markdown(
        f"""
        <div class='hero'>
            <div style='font-size:0.82rem; font-weight:800; letter-spacing:0.16em; opacity:0.72; margin-bottom:10px;'>
                KBO DATA TICKETING GUIDE
            </div>
            <h1>{title}</h1>
            <p>{body}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_cards(items: list[tuple[str, str, str]]):
    cols = st.columns(len(items))
    for col, (label, value, sub) in zip(cols, items):
        with col:
            st.markdown(
                f"<div class='mini-card'><div class='mini-label'>{label}</div><div class='mini-value'>{value}</div><div class='mini-sub'>{sub}</div></div>",
                unsafe_allow_html=True,
            )


def predict_dense(home_team: str, away_team: str, month: int, weekday: str, is_holiday: bool, season: int) -> int | None:
    model, encoders, scaler, feature_cols = load_dense_artifacts()
    if model is None:
        return None
    stadium = TEAM_TO_STADIUM.get(home_team, "잠실")
    # 한화 구장 시즌별 예외 처리 (2024 한밭, 2025~ 대전)
    if home_team == "한화":
        stadium = "한밭" if season <= 2024 else "대전"

    weekday_num = WEEKDAY_TO_NUM[weekday]
    is_weekend = int(weekday_num >= 5)
    is_rival_match = int(tuple(sorted([home_team, away_team])) in RIVAL_MATCHES)
    try:
        row = {
            "home_team_enc": int(encoders["home_team"].transform([home_team])[0]),
            "away_team_enc": int(encoders["away_team"].transform([away_team])[0]),
            "stadium_enc": int(encoders["stadium"].transform([stadium])[0]),
            "month": month,
            "weekday_num": weekday_num,
            "is_weekend": is_weekend,
            "is_holiday": int(is_holiday),
            "is_rival_match": is_rival_match,
            "season": season,
        }
        X = pd.DataFrame([[row[col] for col in feature_cols]], columns=feature_cols)
        X_scaled = scaler.transform(X)
        pred = model.predict(X_scaled, verbose=0).reshape(-1)[0]
        return max(0, int(round(float(pred))))
    except Exception:
        return None


def get_recent_sequence(home_team: str, season: int, seq_len: int = 5):
    df = load_data()
    df = df[df["attendance"].notna()].copy()
    team_df = df[(df["home_team"] == home_team) & (df["season"] == season)].sort_values("date")
    if len(team_df) < seq_len:
        team_df = df[df["home_team"] == home_team].sort_values("date")
    if len(team_df) < seq_len:
        return None, None
    recent = team_df.tail(seq_len)[["date", "away_team", "attendance"]].copy()
    seq = recent["attendance"].astype(float).to_numpy().reshape(1, seq_len, 1)
    return recent, seq


def predict_sequence(model_type: str, home_team: str, season: int):
    lstm_model, gru_model, lstm_scaler, gru_scaler = load_sequence_artifacts()
    recent, seq = get_recent_sequence(home_team, season)
    if recent is None:
        return None, None
    if model_type == "LSTM" and lstm_model is not None and lstm_scaler is not None:
        pred_scaled = lstm_model.predict(seq, verbose=0).reshape(-1)[0]
        pred = lstm_scaler.inverse_transform(np.array([[pred_scaled]])).reshape(-1)[0]
        return max(0, int(round(float(pred)))), recent
    if model_type == "GRU" and gru_model is not None and gru_scaler is not None:
        pred_scaled = gru_model.predict(seq, verbose=0).reshape(-1)[0]
        pred = gru_scaler.inverse_transform(np.array([[pred_scaled]])).reshape(-1)[0]
        return max(0, int(round(float(pred)))), recent
    return None, None


def parse_ticketing_question(question: str) -> dict:
    filters: dict[str, object] = {}
    normalized = question.replace(".", "-").replace("/", "-")

    full_date = re.search(r"(20\d{2})-(\d{1,2})-(\d{1,2})", normalized)
    month_day = re.search(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일", question)
    month_only = re.search(r"(\d{1,2})\s*월", question)

    if full_date:
        year, month, day = map(int, full_date.groups())
        filters["date"] = pd.Timestamp(year=year, month=month, day=day)
    elif month_day:
        month, day = map(int, month_day.groups())
        filters["date"] = pd.Timestamp(year=CURRENT_SEASON, month=month, day=day)
    elif month_only:
        filters["month"] = int(month_only.group(1))

    filters["teams"] = [team for team in TEAMS if team in question]
    filters["stadiums"] = [stadium for stadium in STADIUM_CAPACITY if stadium in question]
    filters["weekend"] = "주말" in question
    filters["holiday"] = any(word in question for word in ["공휴일", "휴일", "어린이날", "현충일", "광복절"])
    filters["low_crowd"] = any(word in question for word in ["한산", "여유", "적은", "덜", "가족"])
    filters["high_crowd"] = any(word in question for word in ["매진", "혼잡", "많", "인기", "빨리"])
    return filters


def make_game_document(game: pd.Series) -> str:
    date = game["date"]
    weekday = NUM_TO_WEEKDAY[date.weekday()]
    home_team = game["home_team"]
    away_team = game["away_team"]
    stadium = game["stadium"]
    is_holiday = bool(game["is_holiday"])
    is_weekend = date.weekday() >= 5
    is_rival = tuple(sorted([home_team, away_team])) in RIVAL_MATCHES
    capacity = STADIUM_CAPACITY.get(stadium, 20000)

    tags = []
    if is_holiday:
        tags.append("공휴일")
    if is_weekend:
        tags.append("주말")
    if is_rival:
        tags.append("라이벌전")
    if not tags:
        tags.append("평일 일반 경기")

    return (
        f"{date.strftime('%Y-%m-%d')} {weekday}요일 {stadium}에서 "
        f"{away_team} 원정팀과 {home_team} 홈팀의 KBO 경기가 열린다. "
        f"구장 정원은 {capacity}명이다. 경기 특성은 {', '.join(tags)}이다. "
        f"팀명 키워드: {home_team}, {away_team}. 구장 키워드: {stadium}."
    )


@st.cache_resource(show_spinner=False)
def build_schedule_vectorstore():
    df = load_data()
    games = df[(df["season"] == CURRENT_SEASON) & (df["date"] >= CURRENT_DATE)].copy()
    games = games.sort_values("date").reset_index(drop=True)
    texts = [make_game_document(game) for _, game in games.iterrows()]
    vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), max_features=1024)
    vectors = vectorizer.fit_transform(texts)
    dense_vectors = normalize(vectors).toarray().astype("float32")
    index = faiss.IndexFlatIP(dense_vectors.shape[1])
    index.add(dense_vectors)
    return index, games, vectorizer


def retrieve_ticketing_games(question: str, max_games: int = 12) -> pd.DataFrame:
    index, games, vectorizer = build_schedule_vectorstore()
    query_vector = normalize(vectorizer.transform([question])).toarray().astype("float32")
    _, indexes = index.search(query_vector, max_games)
    valid_indexes = [int(index_value) for index_value in indexes[0] if index_value >= 0]
    if not valid_indexes:
        return pd.DataFrame()
    return games.loc[valid_indexes].copy()


def build_ticketing_candidates(question: str, max_games: int = 8) -> pd.DataFrame:
    filters = parse_ticketing_question(question)
    df = load_data()
    if "date" in filters or "month" in filters or filters.get("weekend") or filters.get("holiday"):
        games = df[(df["season"] == CURRENT_SEASON) & (df["date"] >= CURRENT_DATE)].copy()
    else:
        games = retrieve_ticketing_games(question, max_games=24)

    if "date" in filters:
        games = games[games["date"].dt.date == filters["date"].date()]
    else:
        games = games[games["date"] >= CURRENT_DATE]
        if "month" in filters:
            games = games[games["date"].dt.month == filters["month"]]
        if filters.get("weekend"):
            games = games[games["date"].dt.weekday >= 5]
        if filters.get("holiday"):
            games = games[games["is_holiday"].astype(int) == 1]

    teams = filters.get("teams", [])
    if teams:
        games = games[games["home_team"].isin(teams) | games["away_team"].isin(teams)]

    stadiums = filters.get("stadiums", [])
    if stadiums:
        games = games[games["stadium"].isin(stadiums)]

    if games.empty:
        return pd.DataFrame()

    rows = []
    for _, game in games.sort_values("date").iterrows():
        home_team = game["home_team"]
        away_team = game["away_team"]
        stadium = game["stadium"]
        date = game["date"]
        weekday = NUM_TO_WEEKDAY[date.weekday()]
        is_holiday = bool(game["is_holiday"])
        prediction = predict_dense(home_team, away_team, int(date.month), weekday, is_holiday, CURRENT_SEASON)
        capacity = STADIUM_CAPACITY.get(stadium, 20000)
        if prediction is None:
            continue
        occupancy = min((prediction / capacity) * 100, 100)
        rows.append({
            "date": date.strftime("%Y-%m-%d"),
            "weekday": weekday,
            "home_team": home_team,
            "away_team": away_team,
            "stadium": stadium,
            "capacity": capacity,
            "prediction": min(prediction, capacity),
            "raw_prediction": prediction,
            "occupancy": round(occupancy, 1),
            "ticket_status": "매우 혼잡 예상" if prediction >= capacity else "사전 확인 권장" if occupancy >= 70 else "여유 관람 가능",
            "holiday": "공휴일" if is_holiday else "일반일",
            "rival": "라이벌전" if tuple(sorted([home_team, away_team])) in RIVAL_MATCHES else "일반 경기",
        })

    result = pd.DataFrame(rows)
    if result.empty:
        return result

    if filters.get("low_crowd"):
        result = result.sort_values(["occupancy", "date"], ascending=[True, True])
    elif filters.get("high_crowd"):
        result = result.sort_values(["occupancy", "date"], ascending=[False, True])
    else:
        result = result.sort_values(["date", "occupancy"], ascending=[True, False])

    return result.head(max_games)


def make_ticketing_context(candidates: pd.DataFrame) -> str:
    lines = []
    for _, row in candidates.iterrows():
        lines.append(
            f"- {row['date']}({row['weekday']}) {row['away_team']} vs {row['home_team']} / "
            f"{row['stadium']} / 예상 {row['prediction']:,}명 / 점유율 {row['occupancy']}% / "
            f"{row['ticket_status']} / {row['holiday']} / {row['rival']}"
        )
    return "\n".join(lines)


def get_ticketing_reasons(row: pd.Series) -> list[str]:
    reasons = []
    if row["occupancy"] >= 95:
        reasons.append("예상 점유율이 매우 높아 조기 예매가 필요합니다.")
    elif row["occupancy"] >= 70:
        reasons.append("관중이 많은 편이라 온라인 예매를 권장합니다.")
    else:
        reasons.append("상대적으로 여유 있는 관람이 예상됩니다.")

    if row["holiday"] == "공휴일":
        reasons.append("공휴일 경기라 가족 관람 수요가 반영됩니다.")
    if row["rival"] == "라이벌전":
        reasons.append("라이벌전 특성상 관심도가 높게 반영됩니다.")
    if row["weekday"] in ["토", "일"]:
        reasons.append("주말 경기라 평일보다 관람 수요가 높을 수 있습니다.")
    if row["raw_prediction"] > row["capacity"]:
        reasons.append("모델 예측값이 구장 정원을 넘어 매우 혼잡 예상으로 보정했습니다.")
    return reasons


def build_rule_based_ticketing_answer(question: str, candidates: pd.DataFrame) -> str:
    top = candidates.head(3).reset_index(drop=True)
    best = top.iloc[0]
    lines = [
        f"조건에 가장 잘 맞는 경기는 {best['date']}({best['weekday']}) {best['away_team']} vs {best['home_team']} 경기입니다.",
        f"{best['stadium']} 기준 예상 관중은 {best['prediction']:,}명, 예상 점유율은 {best['occupancy']}%로 {best['ticket_status']}으로 분류됩니다.",
    ]

    if len(top) > 1:
        alternatives = ", ".join(
            f"{row['date']} {row['away_team']} vs {row['home_team']}({row['occupancy']}%)"
            for _, row in top.iloc[1:].iterrows()
        )
        lines.append(f"비교 후보로는 {alternatives}도 함께 볼 만합니다.")

    reason_text = " ".join(get_ticketing_reasons(best))
    lines.append(reason_text)
    lines.append(
        f"이 답변은 Dense 관중 예측과 일정 데이터 기반이며, 평균 절대 오차가 약 {DENSE_MODEL_MAE:,}명 있을 수 있습니다. 실제 예매 가능 좌석은 구단 예매처에서 다시 확인해 주세요."
    )
    return " ".join(lines)


def render_ticketing_reason_cards(candidates: pd.DataFrame):
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>추천 근거 요약</div>", unsafe_allow_html=True)
    top_candidates = candidates.head(3).reset_index(drop=True)
    cols = st.columns(len(top_candidates))
    for col, (_, row) in zip(cols, top_candidates.iterrows()):
        reasons = get_ticketing_reasons(row)
        reason_html = "".join(f"<li>{reason}</li>" for reason in reasons)
        with col:
            st.markdown(
                f"""
                <div class='mini-card'>
                    <div class='mini-label'>{row['date']}({row['weekday']}) · {row['stadium']}</div>
                    <div class='mini-value' style='font-size:1.2rem;'>{row['away_team']} vs {row['home_team']}</div>
                    <div class='mini-sub'>예상 점유율 {row['occupancy']}% · {row['ticket_status']}</div>
                    <ul style='margin:12px 0 0 18px; padding:0; color:#334155; font-size:0.9rem;'>{reason_html}</ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
    st.markdown("</div>", unsafe_allow_html=True)


def ask_ticketing_llm(question: str, context: str, candidates: pd.DataFrame) -> tuple[str, bool]:
    if not OLLAMA_HOST:
        return build_rule_based_ticketing_answer(question, candidates), False

    system_prompt = (
        "너는 KBO 경기 추천 상담 전문가다. "
        "제공된 경기 일정과 예측 관중, 점유율, 공휴일 여부, 라이벌전 여부만 근거로 답변한다. "
        f"관중 예측은 Dense 모델 기반이며 평균 절대 오차가 약 {DENSE_MODEL_MAE:,}명 수준이라는 한계를 함께 안내한다. "
        "실시간 잔여 좌석과 실제 예매 가능 여부는 알 수 없다고 밝혀라. "
        "점유율이 높으면 온라인 예매와 취소표 확인을 권장하고, 여유로운 관람을 원하면 점유율이 낮은 경기를 추천해라. "
        "한국어로 4~6문장 이내로 답변한다."
    )
    client = ollama.Client(host=OLLAMA_HOST, timeout=12)
    response = client.chat(
        model=OLLAMA_MODEL,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"질문: {question}\n\n경기 예측 데이터:\n{context}"},
        ],
    )
    return response["message"]["content"], True


def render_llm_ticketing_page():
    render_hero(
        "AI 경기 추천 상담",
        "2026 KBO 일정, 관중 예측 모델, 로컬 LLM을 연결해 원하는 조건에 맞는 경기와 관람 판단을 추천합니다.",
    )

    st.info(
        f"이 상담은 2026시즌 일정, Dense 관중 예측, FAISS 검색 결과를 함께 사용합니다. "
        f"관중 수 예측은 평균적으로 약 ±{DENSE_MODEL_MAE:,}명 오차가 있을 수 있고, 실제 잔여 좌석은 구단 예매처에서 확인해야 합니다."
    )

    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>자연어 경기 추천 상담</div>", unsafe_allow_html=True)
    st.caption("예시 질문을 눌러 바로 시연하거나, 원하는 날짜·팀·구장을 자연어로 입력해 보세요.")
    if "ticketing_question" not in st.session_state:
        st.session_state.ticketing_question = LLM_EXAMPLE_QUESTIONS[0]

    example_cols = st.columns(2)
    for index, example_question in enumerate(LLM_EXAMPLE_QUESTIONS):
        with example_cols[index % 2]:
            if st.button(example_question, key=f"llm_example_{index}", use_container_width=True):
                st.session_state.ticketing_question = example_question

    question = st.text_area(
        "궁금한 내용을 입력하세요",
        key="ticketing_question",
        height=100,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("AI 상담 받기", use_container_width=True):
        candidates = build_ticketing_candidates(question)
        if candidates.empty:
            st.warning("질문에 맞는 경기 일정을 찾지 못했습니다. 날짜나 팀명을 조금 더 구체적으로 입력해 주세요.")
            return

        context = make_ticketing_context(candidates)
        spinner_text = "Ollama가 추천 근거를 정리하는 중입니다..." if OLLAMA_HOST else "추천 근거를 정리하는 중입니다..."
        with st.spinner(spinner_text):
            try:
                answer, used_ollama = ask_ticketing_llm(question, context, candidates)
            except Exception as exc:
                answer = build_rule_based_ticketing_answer(question, candidates)
                used_ollama = False
                st.warning(f"Ollama 연결이 되지 않아 기본 추천 답변으로 표시합니다. ({exc})")

        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown("<div class='section-title'>AI 상담 답변</div>", unsafe_allow_html=True)
        if used_ollama:
            st.caption("Ollama와 검색된 일정, 예측 점유율을 근거로 생성한 답변입니다.")
        else:
            st.caption("배포 환경에서도 동작하도록 일정과 예측 점유율을 근거로 생성한 기본 답변입니다.")
        st.write(answer)
        st.markdown("</div>", unsafe_allow_html=True)

        render_ticketing_reason_cards(candidates)

        display_df = candidates[[
            "date", "weekday", "home_team", "away_team", "stadium",
            "prediction", "capacity", "occupancy", "ticket_status", "holiday", "rival"
        ]].copy()
        display_df.columns = ["날짜", "요일", "홈팀", "원정팀", "구장", "예상 관중", "정원", "예상 점유율(%)", "관람 안내", "공휴일", "라이벌"]
        st.dataframe(display_df, use_container_width=True, hide_index=True)


def render_smart_ticketing_page():
    render_hero(
        "AI 경기 추천 가이드",
        "올시즌 KBO 일정과 관중 예측 모델을 바탕으로 경기별 혼잡도와 예매 우선순위를 한눈에 안내합니다.",
    )

    df = load_data()
    
    # 2026년 데이터만 필터링하거나 전체 일정에서 선택 가능하게 함
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>관람 날짜 선택</div>", unsafe_allow_html=True)
    st.caption("2026년 3월 28일부터 9월 6일까지 등록된 정규시즌 일정에서 선택할 수 있습니다.")
    
    # 기본 날짜를 2026년 시즌 중 하루로 설정 (데이터가 있는 범위)
    min_date = pd.to_datetime("2026-03-28")
    max_date = pd.to_datetime("2026-09-06")
    default_date = pd.to_datetime("2026-05-05") # 어린이날 예시
    
    selected_date = st.date_input(
        "어떤 경기의 혼잡도를 확인하고 싶으신가요?",
        value=default_date,
        min_value=min_date,
        max_value=max_date
    )
    st.markdown("</div>", unsafe_allow_html=True)

    target_games = df[df["date"].dt.date == selected_date]

    if target_games.empty:
        st.warning(f"{selected_date}에는 예정된 경기가 없습니다.")
        return

    st.markdown(f"<div class='section-title'>{selected_date.strftime('%Y년 %m월 %d일')} 경기 리스트</div>", unsafe_allow_html=True)

    for idx, game in target_games.iterrows():
        home_team = game["home_team"]
        away_team = game["away_team"]
        stadium = game["stadium"]
        is_holiday = bool(game["is_holiday"])
        season = int(game["season"])
        weekday = NUM_TO_WEEKDAY[game["date"].weekday()]
        
        # 예측 수행
        with st.spinner(f"{home_team} vs {away_team} 경기 분석 중..."):
            pred = predict_dense(home_team, away_team, selected_date.month, weekday, is_holiday, season)
        
        # 경기장 정원 확인
        capacity = STADIUM_CAPACITY.get(stadium, 20000)
        
        # UI 구성
        with st.container():
            st.markdown("<div class='game-card'>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns([2, 2, 3])
            
            with c1:
                st.caption(f"{weekday}요일 · {'공휴일' if is_holiday else '일반일'}")
                st.subheader(f"{away_team} vs {home_team}")
                st.write(f"**{stadium} 구장**")
            
            with c2:
                if pred:
                    display_pred = min(pred, capacity)
                    occupancy = min((pred / capacity) * 100, 100)
                    delta_text = "매우 혼잡 예상" if pred >= capacity else f"{display_pred:,} / {capacity:,}명"
                    st.metric("예상 점유율", f"{occupancy:.1f}%", delta_text)
                else:
                    st.write("예측 모델을 불러올 수 없습니다.")
                    occupancy = 0
            
            with c3:
                if pred and pred >= capacity:
                    badge_class = "badge-red"
                    status_text = "매우 혼잡 예상 - 온라인 취소표 확인 권장"
                elif occupancy >= 90:
                    badge_class = "badge-red"
                    status_text = "매우 혼잡 - 매진 임박! 온라인 예매 필수"
                elif occupancy >= 70:
                    badge_class = "badge-orange"
                    status_text = "혼잡 - 빠른 온라인 예매를 권장합니다"
                elif occupancy >= 40:
                    badge_class = "badge-green"
                    status_text = "보통 - 여유 관람 가능"
                else:
                    badge_class = "badge-blue"
                    status_text = "여유 - 쾌적한 현장 관람이 가능합니다"
                
                st.markdown(f"<div class='ticket-badge {badge_class}'>{status_text}</div>", unsafe_allow_html=True)
                
                # 프로그레스 바 시각화
                st.progress(min(occupancy / 100, 1.0))
            
            st.markdown("</div>", unsafe_allow_html=True)


def render_dense_page():
    render_hero(
        "Dense 조건 기반 직접 매칭",
        "특정 팀과 경기 날짜 등 조건을 직접 설정하여 예측 모델의 시뮬레이션을 수행합니다.",
    )

    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>경기 조건 입력</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        home_team = st.selectbox("홈팀", TEAMS, key="dense_home")
        weekday = st.selectbox("요일", ["월", "화", "수", "목", "금", "토", "일"], index=5, key="dense_weekday")
        month = st.slider("경기 월 선택", min_value=3, max_value=10, value=7, key="dense_month")
    with c2:
        away_team = st.selectbox("원정팀", TEAMS, index=1, key="dense_away")
        is_holiday = st.toggle("공휴일 경기", value=False, key="dense_holiday")
        season = st.selectbox("시즌", [2024, 2025, 2026], index=2, key="dense_season")

    stadium = TEAM_TO_STADIUM.get(home_team, "미정")
    # UI 표시용 한화 구장 예외 처리
    if home_team == "한화":
        stadium = "한밭" if season <= 2024 else "대전"

    is_weekend = weekday in ["토", "일"]
    is_rival_match = int(tuple(sorted([home_team, away_team])) in RIVAL_MATCHES)
    st.markdown("</div>", unsafe_allow_html=True)

    render_cards([
        ("자동 매핑 구장", stadium, "홈팀 기준 구장 적용"),
        ("주말 여부", "주말" if is_weekend else "평일", "요일 자동 계산"),
        ("공휴일 여부", "공휴일" if is_holiday else "일반일", "직접 입력 반영"),
        ("라이벌전 여부", "라이벌전" if is_rival_match else "일반 경기", "LG-두산 자동 감지"),
    ])

    if st.button("Dense 예상 관중 확인", use_container_width=True):
        prediction = predict_dense(home_team, away_team, month, weekday, is_holiday, season)
        if prediction is None:
            st.info("Dense 모델 산출물이 없습니다. 먼저 학습을 실행해주세요.")
        else:
            st.markdown("<div class='panel'>", unsafe_allow_html=True)
            st.markdown("<div class='section-title'>예측 결과</div>", unsafe_allow_html=True)
            render_cards([
                ("예상 관중 수", f"{prediction:,}명", "Dense 회귀 모델 결과"),
                ("홈팀", home_team, f"원정팀 {away_team}"),
                ("경기 조건", f"{season} / {month}월 / {weekday}", "입력 조합 기준"),
            ])
            st.markdown("</div>", unsafe_allow_html=True)


def render_sequence_page(model_type: str):
    render_hero(
        f"{model_type} 시계열 예측",
        "최근 5경기 홈 관중 흐름을 시계열로 보고 다음 홈경기 관중 수를 예측합니다.",
    )

    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>시계열 입력 조건</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        home_team = st.selectbox("홈팀", TEAMS, key=f"{model_type}_home")
    with c2:
        season = st.selectbox("시즌", [2024, 2025, 2026], index=2, key=f"{model_type}_season")
    st.markdown("</div>", unsafe_allow_html=True)

    prediction, recent = predict_sequence(model_type, home_team, season)
    if prediction is None or recent is None:
        st.info("시계열 예측에 필요한 최근 5경기 데이터 또는 모델 산출물이 없습니다.")
        return

    render_cards([
        (f"{model_type} 예상 관중", f"{prediction:,}명", "다음 홈경기 예측값"),
        ("기준 홈팀", home_team, f"시즌 {season}"),
        ("시퀀스 길이", "최근 5경기", "홈 관중 흐름 사용"),
    ])

    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>최근 5경기 홈 관중 흐름</div>", unsafe_allow_html=True)
    display_df = recent.copy()
    display_df["date"] = display_df["date"].dt.strftime("%Y-%m-%d")
    display_df.columns = ["날짜", "원정팀", "관중 수"]
    left, right = st.columns([1.05, 1.25])
    with left:
        st.dataframe(display_df, use_container_width=True, hide_index=True)
    with right:
        chart_df = recent.copy()
        chart_df["date"] = chart_df["date"].dt.strftime("%m-%d")
        chart = (
            alt.Chart(chart_df)
            .mark_line(point=True, strokeWidth=4, color="#1976d2")
            .encode(
                x=alt.X("date:N", title="경기일"),
                y=alt.Y("attendance:Q", title="관중 수"),
                tooltip=["date", "away_team", "attendance"],
            )
            .properties(height=320)
        )
        st.altair_chart(chart, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)


st.sidebar.title("KBO 분석 센터")
page = st.sidebar.radio("서비스 메뉴", ["경기 추천 가이드", "AI 경기 추천 상담", "Dense 예측 시뮬레이션", "LSTM 흐름 분석", "GRU 흐름 분석"])

if page == "경기 추천 가이드":
    render_smart_ticketing_page()
elif page == "AI 경기 추천 상담":
    render_llm_ticketing_page()
elif page == "Dense 예측 시뮬레이션":
    render_dense_page()
elif page == "LSTM 흐름 분석":
    render_sequence_page("LSTM")
else:
    render_sequence_page("GRU")
