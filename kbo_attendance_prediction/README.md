# KBO Attendance Prediction

이번 주 딥러닝 프로젝트용 야구 관중 수 예측 폴더입니다.

구성:
- data/kbo_attendance.csv : 2024, 2025, 2026 시즌 초반 실제 경기 관중 데이터
- baseball_attendance_analysis.ipynb : 전처리, 시각화, Dense 모델 학습
- app.py : Streamlit 예측 화면
- models/ : 저장 모델 폴더
- artifacts/ : 그래프, 산출물 폴더

이번 프로젝트는 실제 경기 관중 데이터를 바탕으로 아래 특징을 사용합니다.
- 홈팀
- 원정팀
- 구장
- 월
- 요일
- 주말 여부
- 공휴일 여부
- 라이벌전 여부 (LG-두산)
- 시즌

모델은 수업 시간 실습 흐름에 맞춰 `Sequential + Dense` 회귀 모델을 사용합니다.
