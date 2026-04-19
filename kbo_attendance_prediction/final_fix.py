import pandas as pd

# 파일 경로
file_path = r'C:\Users\human-18\Desktop\kbo_attendance_prediction\data\kbo_attendance.csv'

# 데이터 로드 (UTF-8-SIG는 한글 지원 우수)
try:
    df = pd.read_csv(file_path, encoding='utf-8-sig')
except Exception as e:
    # 혹시 다른 인코딩일 경우 대비
    df = pd.read_csv(file_path, encoding='cp949')

# 날짜 변환
df['date_dt'] = pd.to_datetime(df['date'])

# 1. 2026년 실제 공휴일 목록 업데이트 (노동절, 제헌절 포함)
holidays_2026 = [
    '2026-05-01', # 노동절
    '2026-05-05', # 어린이날
    '2026-05-24', # 부처님오신날
    '2026-05-25', # 대체공휴일
    '2026-06-03', # 지방선거
    '2026-06-06', # 현충일
    '2026-07-17', # 제헌절
    '2026-08-15', # 광복절
    '2026-08-17'  # 대체공휴일
]

mask_2026 = df['date_dt'].dt.year == 2026
df.loc[mask_2026, 'is_holiday'] = 0 # 초기화
df.loc[df['date_dt'].dt.strftime('%Y-%m-%d').isin(holidays_2026), 'is_holiday'] = 1

# 2. SSG 홈구장을 포함한 모든 팀의 구장 이름 정상화 (2026년 데이터 대상)
team_to_stadium = {
    'LG': '잠실', '두산': '잠실', '삼성': '대구', 'KIA': '광주',
    'KT': '수원', 'SSG': '문학', '롯데': '사직', '한화': '대전',
    'NC': '창원', '키움': '고척'
}

for team, stadium in team_to_stadium.items():
    mask_team = (df['date_dt'].dt.year == 2026) & (df['home_team'] == team)
    df.loc[mask_team, 'stadium'] = stadium

# 3. 요일(weekday) 정보 정상화 (2026년 데이터 대상)
weekday_map = {0: '월', 1: '화', 2: '수', 3: '목', 4: '금', 5: '토', 6: '일'}
df.loc[mask_2026, 'weekday'] = df.loc[mask_2026, 'date_dt'].dt.weekday.map(weekday_map)

# 불필요한 날씨 정보가 다시 생기지 않도록 확인 (컬럼 유지)
cols = ['date', 'weekday', 'home_team', 'away_team', 'stadium', 'attendance', 'is_holiday', 'weather_main', 'temp_avg', 'rain_mm', 'season']
df = df[cols].copy()

# 저장
df.to_csv(file_path, index=False, encoding='utf-8-sig')

print("Successfully fixed 2026 data: Updated holidays (including 5/1, 7/17), normalized stadiums (SSG:문학), and fixed weekdays.")
