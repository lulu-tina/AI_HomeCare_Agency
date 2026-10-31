"""個案新增失智的資格重新檢核：受影響服務、三種處理方式的分法、接手人選與固定人選更新。"""
import pandas as pd

import qualification_recheck as qr


def _tasks():
    return pd.DataFrame({'任務ID': ['t1', 't2', 't3', 't4', 'x1'], '案家ID': ['C1'] * 4 + ['C2'],
                         '日期': ['2026-11-02', '2026-11-09', '2026-12-07', '2026-12-14', '2026-11-02'],
                         '時間窗_開始': ['09:00'] * 5, '時間窗_結束': ['10:00'] * 5})


def _cg():
    return pd.DataFrame({'居服員ID': ['A', 'B', 'C'], '取得資格': ['照顧服務員單一級技術士', '失智症照顧服務20小時訓練課程',
                                                              '失智症照顧服務20小時訓練課程']})


def test_affected_only_uncertified_from_effective_date():
    current = {'t1': 'A', 't2': 'A', 't3': 'B', 't4': 'A', 'x1': 'A'}
    out = qr.affected_tasks('C1', _tasks(), current, _cg(), since='2026-11-05')
    assert [a['任務ID'] for a in out] == ['t2', 't4']      # t1 早於生效日、t3 已是有證照的 B、x1 是別的案家


def test_split_never_keeps_uncertified_worker():
    # 法規：只能派給有訓練的人，三種方式都把全部受影響服務改派，沒有「寬限期內繼續服務」
    aff = [{'任務ID': t, '日期': d} for t, d in [('t1', '2026-11-02'), ('t2', '2026-11-09'), ('t3', '2026-12-07')]]
    for method in qr.METHODS:
        parts = qr.split(method, aff, '2026-11-30', handover=1)
        assert parts['move'] == ['t1', 't2', 't3'] and parts['keep'] == [] and parts['accompany'] == []
    assert qr.split('handover', aff, handover=2)['escort'] == ['t1', 't2']    # 原居服員陪同、不提供服務
    assert qr.split('grace', aff, '2026-11-30')['escort'] == []
    assert qr.DEFAULT_METHOD == 'now' and qr.METHODS[0] == 'now'


def test_covering_needs_every_task_and_ranks_by_average_position():
    avail = {'t1': [('B', 'b1'), ('C', 'c1')], 't2': [('C', 'c2'), ('B', 'b2')], 't3': [('C', 'c3')]}
    assert [c['cg'] for c in qr.covering(avail, ['t1', 't2'])] == ['B', 'C']
    assert [c['cg'] for c in qr.covering(avail, ['t1', 't2', 't3'])] == ['C']
    assert qr.covering(avail, []) == []


def test_team_after_removes_uncertified_owner():
    for method in qr.METHODS:
        assert qr.team_after(method, 'A', ['D', 'B'], 'B') == ('B', ['D'])


def test_foot_care_need_uses_ba08_training():
    cg = pd.DataFrame({'居服員ID': ['A', 'F'], '取得資格': ['失智症照顧服務20小時訓練課程', 'BA08足部照護']})
    assert qr.certified(cg, 'foot') == {'F'}
    out = qr.affected_tasks('C1', _tasks(), {'t1': 'A', 't2': 'F'}, cg, need='foot')
    assert [a['任務ID'] for a in out] == ['t1']
    row = qr.followup('grace', 'C1', 'A', 'F', qr.split('grace', out, '2026-12-31'), '2026-12-31', need='foot')
    assert '足部照護 10 小時' in row['待辦'] and 'BA08' in row['提醒'] and row['需求變更'].startswith('新增足部照護')
    assert '已改由 F 服務' in row['待辦'] and row['改由有訓練者服務'] == 1
    row = qr.followup('handover', 'C1', 'A', 'F', qr.split('handover', out, handover=1), need='foot')
    assert row['原居服員陪同'] == 1 and '不提供服務' in row['待辦']


def test_split_cover_by_weekday_when_nobody_takes_all():
    # t1、t3 是週一，t2 是週三；B 只能週一、C 只能週三
    avail = {'t1': [('B', 'b')], 't3': [('B', 'b')], 't2': [('C', 'c')]}
    days = {'t1': '2026-11-02', 't2': '2026-11-04', 't3': '2026-11-09'}
    assert qr.covering(avail, ['t1', 't2', 't3']) == []
    plan = qr.split_cover(avail, ['t1', 't2', 't3'], day_of=days)
    assert plan['assign'] == {'t1': 'B', 't3': 'B', 't2': 'C'} and plan['missing'] == []
    assert [(g['星期'], g['居服員'], g['筆數']) for g in plan['groups']] == [('週一', 'B', 2), ('週三', 'C', 1)]
    # 某一次完全沒人：列為缺人，不硬塞
    plan = qr.split_cover({'t1': [('B', 'b')], 't2': []}, ['t1', 't2'], day_of=days)
    assert plan['missing'] == ['t2'] and plan['assign'] == {'t1': 'B'}
