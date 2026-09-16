import json
from datetime import datetime, timezone
from unittest.mock import patch
import generate
from market_history import quote


def test_failed_fetch_never_creates_today(tmp_path):
    old = {'date':'2026-06-01','close':4.336}
    file=tmp_path/'history.json'
    file.write_text(json.dumps({'us10y':[old]}))
    with patch.object(generate,'HISTORY_CACHE_PATH',str(file)):
        d=generate.merge_and_persist_history({'us10y':quote([old]),'vix':quote([])})
    assert d['us10y']['chart']==[old]
    assert json.loads(file.read_text())['us10y']==[old]


def test_history_is_backfilled_not_flattened(tmp_path):
    rows=[{'date':'2026-09-14','close':4.9},{'date':'2026-09-15','close':4.996}]
    with patch.object(generate,'HISTORY_CACHE_PATH',str(tmp_path/'history.json')):
        d=generate.merge_and_persist_history({'us10y':quote(rows),'vix':quote([])})
    assert d['us10y']['chart']==rows
    assert d['us10y']['date']=='2026-09-15'


def test_finmind_english_fields_keep_money_in_twd(monkeypatch):
    import data_fetcher as d
    rows=[{'date':'2026-09-16','name':name,'TodayBalance':today,'YesBalance':yesterday} for name,today,yesterday in [('MarginPurchase',9000,8000),('ShortSale',200,210),('MarginPurchaseMoney',586166660000,582240465000)]]
    monkeypatch.setattr(d,'_finmind_fetch',lambda *args:rows)
    value=d.fetch_margin_trading('20260916')
    assert value['margin_balance']==9000
    assert value['short_change']==-10
    assert value['margin_balance_amount']==586166660000
    assert value['date']=='2026-09-16'
