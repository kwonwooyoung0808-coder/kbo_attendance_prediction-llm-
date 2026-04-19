from pathlib import Path
import json
import pickle
import warnings

import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder, StandardScaler
from tensorflow.keras.callbacks import EarlyStopping
from tensorflow.keras.layers import Dense, Dropout, LSTM, GRU
from tensorflow.keras.models import Sequential

warnings.filterwarnings('ignore')

BASE_DIR = Path(r'C:\Users\human-18\Desktop\kbo_attendance_prediction')
DATA_PATH = BASE_DIR / 'data' / 'kbo_attendance.csv'
ARTIFACT_DIR = BASE_DIR / 'artifacts'
MODEL_DIR = BASE_DIR / 'models'
ARTIFACT_DIR.mkdir(exist_ok=True)
MODEL_DIR.mkdir(exist_ok=True)

RIVAL_MATCHES = {tuple(sorted(['LG', '두산']))}


def rmse_score(y_true, y_pred):
    return float(np.sqrt(mean_squared_error(y_true, y_pred)))


def train_dense(df: pd.DataFrame):
    df = df.copy()
    df['date'] = pd.to_datetime(df['date'])
    df['month'] = df['date'].dt.month
    df['weekday_num'] = df['date'].dt.weekday
    df['is_weekend'] = (df['weekday_num'] >= 5).astype(int)
    df['is_rival_match'] = df.apply(lambda row: int(tuple(sorted([row['home_team'], row['away_team']])) in RIVAL_MATCHES), axis=1)

    encoders = {}
    for col in ['home_team', 'away_team', 'stadium']:
        le = LabelEncoder()
        df[col + '_enc'] = le.fit_transform(df[col].astype(str))
        encoders[col] = le

    feature_cols = ['home_team_enc', 'away_team_enc', 'stadium_enc', 'month', 'weekday_num', 'is_weekend', 'is_holiday', 'is_rival_match', 'season']
    X = df[feature_cols].copy()
    y = df['attendance'].astype(float).copy()

    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    X_train, X_test, y_train, y_test = train_test_split(X_scaled, y, test_size=0.2, random_state=42)

    model = Sequential([
        Dense(64, activation='relu', input_shape=(X_train.shape[1],)),
        Dropout(0.2),
        Dense(32, activation='relu'),
        Dense(16, activation='relu'),
        Dense(1),
    ])
    model.compile(optimizer='adam', loss='mse', metrics=['mae'])
    history = model.fit(
        X_train, y_train,
        validation_split=0.2,
        epochs=80,
        batch_size=16,
        callbacks=[EarlyStopping(patience=8, restore_best_weights=True)],
        verbose=0,
    )

    pred = model.predict(X_test, verbose=0).reshape(-1).astype(float)
    y_true = y_test.to_numpy(dtype=float)
    baseline_pred = np.repeat(y_train.mean(), len(y_test))

    metrics = {
        'model': 'Dense',
        'mae': float(mean_absolute_error(y_true, pred)),
        'rmse': rmse_score(y_true, pred),
        'baseline_mae': float(mean_absolute_error(y_test, baseline_pred)),
        'baseline_rmse': rmse_score(y_test, baseline_pred),
        'sample_prediction': int(round(float(pred[0]))) if len(pred) else 0,
    }

    model.save(MODEL_DIR / 'attendance_dense_model.keras')
    with open(ARTIFACT_DIR / 'encoders.pkl', 'wb') as f:
        pickle.dump(encoders, f)
    with open(ARTIFACT_DIR / 'scaler.pkl', 'wb') as f:
        pickle.dump(scaler, f)
    with open(ARTIFACT_DIR / 'feature_cols.json', 'w', encoding='utf-8') as f:
        json.dump(feature_cols, f, ensure_ascii=False, indent=2)
    with open(ARTIFACT_DIR / 'dense_model_info.json', 'w', encoding='utf-8') as f:
        json.dump(metrics, f, ensure_ascii=False, indent=2)

    return metrics, history.history


def build_sequence_df(df: pd.DataFrame, seq_len: int = 5):
    df = df.copy()
    df['date'] = pd.to_datetime(df['date'])
    df.sort_values(['home_team', 'date'], inplace=True)
    sequences = []
    targets = []
    meta = []
    for home_team, group in df.groupby('home_team'):
        values = group['attendance'].astype(float).to_numpy()
        dates = group['date'].to_numpy()
        seasons = group['season'].to_numpy()
        for idx in range(seq_len, len(values)):
            sequences.append(values[idx-seq_len:idx].reshape(seq_len, 1))
            targets.append(values[idx])
            meta.append({
                'home_team': home_team,
                'target_date': str(pd.to_datetime(dates[idx]).date()),
                'season': int(seasons[idx]),
            })
    X = np.array(sequences, dtype=float)
    y = np.array(targets, dtype=float)
    meta_df = pd.DataFrame(meta)
    return X, y, meta_df


def train_sequence_model(X, y, model_type: str):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    scaler_y = StandardScaler()
    y_train_scaled = scaler_y.fit_transform(y_train.reshape(-1, 1)).reshape(-1)
    y_test_scaled = scaler_y.transform(y_test.reshape(-1, 1)).reshape(-1)

    recurrent_layer = LSTM(32) if model_type == 'LSTM' else GRU(32)
    model = Sequential([
        recurrent_layer,
        Dropout(0.2),
        Dense(16, activation='relu'),
        Dense(1),
    ])
    model.compile(optimizer='adam', loss='mse', metrics=['mae'])
    history = model.fit(
        X_train, y_train_scaled,
        validation_split=0.2,
        epochs=80,
        batch_size=8,
        callbacks=[EarlyStopping(patience=8, restore_best_weights=True)],
        verbose=0,
    )

    pred_scaled = model.predict(X_test, verbose=0).reshape(-1)
    pred = scaler_y.inverse_transform(pred_scaled.reshape(-1, 1)).reshape(-1)
    baseline_pred = np.repeat(y_train.mean(), len(y_test))
    metrics = {
        'model': model_type,
        'mae': float(mean_absolute_error(y_test, pred)),
        'rmse': rmse_score(y_test, pred),
        'baseline_mae': float(mean_absolute_error(y_test, baseline_pred)),
        'baseline_rmse': rmse_score(y_test, baseline_pred),
        'sample_prediction': int(round(float(pred[0]))) if len(pred) else 0,
    }

    model.save(MODEL_DIR / f'attendance_{model_type.lower()}_model.keras')
    with open(ARTIFACT_DIR / f'{model_type.lower()}_target_scaler.pkl', 'wb') as f:
        pickle.dump(scaler_y, f)
    with open(ARTIFACT_DIR / f'{model_type.lower()}_model_info.json', 'w', encoding='utf-8') as f:
        json.dump(metrics, f, ensure_ascii=False, indent=2)
    return metrics, history.history


def main():
    df = pd.read_csv(DATA_PATH)
    dense_metrics, dense_history = train_dense(df)
    X_seq, y_seq, meta_df = build_sequence_df(df, seq_len=5)
    lstm_metrics, lstm_history = train_sequence_model(X_seq, y_seq, 'LSTM')
    gru_metrics, gru_history = train_sequence_model(X_seq, y_seq, 'GRU')

    with open(ARTIFACT_DIR / 'sequence_meta.json', 'w', encoding='utf-8') as f:
        json.dump({'sequence_length': 5, 'rows': int(len(X_seq))}, f, ensure_ascii=False, indent=2)
    with open(ARTIFACT_DIR / 'sequence_recent_games.csv', 'w', encoding='utf-8-sig', newline='') as f:
        meta_df.to_csv(f, index=False)

    compare = [dense_metrics, lstm_metrics, gru_metrics]
    with open(ARTIFACT_DIR / 'model_compare.json', 'w', encoding='utf-8') as f:
        json.dump(compare, f, ensure_ascii=False, indent=2)
    pd.DataFrame(compare).to_csv(ARTIFACT_DIR / 'model_compare.csv', index=False, encoding='utf-8-sig')

    with open(ARTIFACT_DIR / 'training_history.json', 'w', encoding='utf-8') as f:
        json.dump({'dense': dense_history, 'lstm': lstm_history, 'gru': gru_history}, f, ensure_ascii=False)

    print(pd.DataFrame(compare))

if __name__ == '__main__':
    main()
