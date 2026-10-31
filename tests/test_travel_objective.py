"""派出數相同時偏好總轉場較少的派法（PipelineConfig.travel_objective_mode）。

合成案例：案家 X 的主責 O 當天請假，B、C 都能代班，派出數一樣；
B 年資較深、代班順序排第一，但 B 下一案（YB）離 X 約 11 公里，C 下一案（YC）只隔 0.5 公里。
off 維持舊行為選 B；weighted／two_stage 都要改選 C，且實際轉場（calculate_schedule_travel）較少。
"""
import pathlib
import sys

import pandas as pd
import pytest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / 'tools'))
import caregiver_engine as engine
import monthly_scheduler
import sandbox_travel
from agency_rules import payroll_tables
from monthly_scheduler import run_fixed_period_dispatch

DAY = '2026-11-02'
HOME = (25.03, 121.53)
SITES = {'X': (25.040, 121.540), 'YB': (25.140, 121.540), 'YC': (25.045, 121.540)}


@pytest.fixture(autouse=True)
def straight_line_travel(monkeypatch):
    # 交通改用直線距離概算（tools/sandbox_travel），不連網，且遠近真的會影響分鐘數。
    patches = sandbox_travel.patches()
    for p in patches:
        p.start()
    monkeypatch.setattr(monthly_scheduler, 'prefetch_osrm_travel_times', lambda *a, **k: None)
    yield
    for p in patches:
        p.stop()


def worker(identifier, **values):
    return {'居服員ID': identifier, '性別': '女', '常用交通工具': '機車',
            '取得資格': '失智症照顧服務20小時', '核心專長證照': '失智症照顧專長',
            '服務起點_緯度(家)': HOME[0], '服務起點_經度(家)': HOME[1],
            '每日工時上限(小時)': 12.0, '每日可服務時段_起': '00:00', '每日可服務時段_迄': '23:59',
            '可排班星期': '1,2,3,4,5,6,7', '可服務資料已確認': True,
            '當月累計服務時數(疲勞度)': 0.0, '歷史滿意度均值': 4.5,
            '特殊排斥條件': '', '請假紀錄': None, '到職日': '2024-01-01', **values}


def visit(identifier, client, start, end):
    duration = (pd.Timestamp(end) - pd.Timestamp(start)).total_seconds() / 60
    lat, lon = SITES[client]
    return {'任務ID': identifier, '案家ID': client, '日期': DAY, '星期': engine.get_weekday_name(DAY),
            '時間窗_開始': start, '時間窗_結束': end, '服務歷時(分鐘)': duration,
            'Service_Code_1': 'SA10', 'Units_1': duration / 60,
            '服務地點_緯度': lat, '服務地點_經度': lon, '任務優先級': '一般',
            '指定居服員性別': '', '特殊照護需求': '', '需重度移位協助(0/1)': 0,
            '案家環境特徵': '', '歷史首選居服員ID': '', '資料來源': '機構三檔'}


def dispatch(mode, **values):
    people = pd.DataFrame([
        worker('O', 請假紀錄=(f'{DAY} 00:00', f'{DAY} 23:59')),
        worker('B', 到職日='2010-01-01'),      # 年資深 → 舊行為的代班第一順位
        worker('C', 到職日='2024-01-01'),
    ])
    tasks = pd.DataFrame([visit('x1', 'X', '09:00', '10:00'),
                          visit('yb1', 'YB', '11:00', '12:00'),
                          visit('yc1', 'YC', '11:00', '12:00')])
    clients = pd.DataFrame([{'案家ID': c} for c in SITES])
    cfg = engine.PipelineConfig(travel_objective_mode=mode, enable_backup_caregiver=True,
                                max_backup_caregivers=2, enable_support_caregiver=True, **values)
    result = run_fixed_period_dispatch(tasks, people, clients, cfg, primary_map={'X': 'O', 'YB': 'B', 'YC': 'C'})
    assigned = result['df_result'].set_index('任務ID')['派單居服員'].to_dict()
    merged = engine.merge_task_client_tables(tasks, clients)
    _, daily, _, _ = payroll_tables(merged, result['df_result'], people, result['route_result'], cfg)
    return result, assigned, float(daily['轉場工時'].sum() * 60)


def test_off_keeps_previous_substitute_order():
    _, assigned, _ = dispatch('off')
    assert assigned == {'x1': 'B', 'yb1': 'B', 'yc1': 'C'}


@pytest.mark.parametrize('mode', ['weighted', 'two_stage'])
def test_travel_objective_prefers_fewer_transition_minutes(mode):
    base_result, base, base_minutes = dispatch('off')
    result, assigned, minutes = dispatch(mode)
    # 派出數不變、主責服務不動，只有請假代班改給轉場較少的 C。
    assert len(assigned) == len(base) == 3
    assert assigned['yb1'] == 'B' and assigned['yc1'] == 'C'
    assert assigned['x1'] == 'C'
    # 用最終班表的 calculate_schedule_travel 路段核對：真實轉場分鐘確實下降。
    assert minutes < base_minutes - 20
    routes = result['route_result'].set_index('任務ID')
    assert routes.loc['yc1', '預估車程(分)'] < base_result['route_result'].set_index('任務ID').loc['yb1', '預估車程(分)']


def test_two_stage_reports_its_moves_and_keeps_stage_one_on_zero_budget():
    result, assigned, _ = dispatch('two_stage')
    stats = result['dispatch_stats']['travel_stage2']
    assert stats['moves'] == 1 and stats['estimated_saved_minutes'] > 20
    assert '轉場最佳化改派' in result['df_result'].set_index('任務ID').loc['x1', '代班排序理由']
    # 第二階段沒有時間預算：立即停止，保留第一階段（舊行為）的解。
    result, assigned, _ = dispatch('two_stage', travel_stage2_seconds=0)
    assert result['dispatch_stats']['travel_stage2']['timed_out']
    assert assigned['x1'] == 'B'


def test_two_stage_skips_moves_below_paid_transition_threshold():
    # 改給 C 可省約 30 分付費轉場；門檻設 100 分時不得改派。
    result, assigned, _ = dispatch('two_stage', travel_stage2_min_gain_minutes=100)
    assert assigned['x1'] == 'B'
    assert result['dispatch_stats']['travel_stage2']['moves'] == 0
