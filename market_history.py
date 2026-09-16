"""Dated observations only: never turn a cached quote into today's observation."""
from datetime import datetime, timezone
import math
import yfinance as yf


def fetch_history(symbol, days=90):
    try:
        frame = yf.Ticker(symbol).history(period='3mo', interval='1d', auto_adjust=False)
        rows = []
        for stamp, row in frame.iterrows():
            value = float(row['Close'])
            if math.isfinite(value) and value > 0:
                rows.append({'date': stamp.date().isoformat(), 'close': round(value, 4)})
        return rows[-days:]
    except Exception as exc:
        print(f'[Yahoo history {symbol}] {type(exc).__name__}')
        return []


def quote(rows, source='Yahoo Finance'):
    if not rows:
        return {'value': None, 'prev_value': None, 'date': None, 'chart': [], 'status': 'unavailable', 'source': source}
    last = rows[-1]
    age = (datetime.now(timezone.utc).date() - datetime.fromisoformat(last['date']).date()).days
    return {'value': last['close'], 'prev_value': rows[-2]['close'] if len(rows)>1 else None,
            'date': last['date'], 'chart': rows[-30:], 'source': source,
            'status': 'stale' if age > 4 else 'ok'}


def merge_observations(cache, fresh, limit=90):
    # Pre-fix cache contains synthetic dates; accept only caches explicitly migrated.
    points = {p['date']: p for p in cache if p.get('date') and p.get('close') is not None}
    points.update({p['date']: p for p in fresh if p.get('date') and p.get('close') is not None})
    return [points[d] for d in sorted(points)][-limit:]
