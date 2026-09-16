import json
from datetime import datetime, timezone
from pathlib import Path
j=json.loads(Path('latest_data.json').read_text());errors=[]
for key in ('vix','us10y'):
 q=j[key]
 if q.get('status')!='ok' or not q.get('date'):
  errors.append(key+': source unavailable/stale')
 elif (datetime.now(timezone.utc).date()-datetime.fromisoformat(q['date']).date()).days>4:
  errors.append(key+': observation stale')
 if len(q.get('chart',[]))<10 or len({p['close'] for p in q['chart']})<2:
  errors.append(key+': insufficient or frozen history')
Path('health.json').write_text(json.dumps({'checked_at':datetime.now(timezone.utc).isoformat(),'errors':errors,'ok':not errors},indent=2))
if errors:raise SystemExit('; '.join(errors))
print('Dated market history health: OK')
