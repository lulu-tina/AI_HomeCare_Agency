"""同一份 Excel、同一個排班期間跑一次自動排班，輸出固定格式的指標，供優化前後對照。

    python tools/benchmark.py 居服機構資料_三工作表_v9.xlsx
    python tools/benchmark.py 居服機構資料_三工作表_v9.xlsx --travel osrm --set time_flex_minutes=30 --set max_backup_caregivers=1
    python tools/benchmark.py --compare benchmarks/a.json benchmarks/b.json

排班流程比照 app.py「開始排班 · 全期間」：主責＋最多 2 位備援、臨時支援開啟、時段不挪動；
不讀 WorkflowStore（請假登錄、連續上班種子），確保每次輸入相同。
sandbox：交通改用 tools/sandbox_travel 的直線距離概算，無網路也能跑完。
osrm   ：機車走 OSRM；大眾運輸預設用引擎內建備援估算（不打 Google），加 --google 才改查 Google。
"""
import argparse, dataclasses, hashlib, json, math, os, pathlib, random, re, subprocess, sys, time
from datetime import datetime
ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / 'tools'))
import numpy as np
import pandas as pd
import caregiver_engine as e
from monthly_scheduler import run_fixed_period_dispatch
from source_fields import source_primary_map, normalize_clients
from agency_rules import payroll_tables
from dashboard_metrics import build_dashboard, CAREFLOW_BUILD

DEFAULT_EXCEL = ROOT / '居服機構資料_三工作表_v9.xlsx'
# (鍵, 中文名稱, 單位, 較佳方向 up/down/None)
METRICS = [
    ('services_assigned', '派出服務數', '筆', 'up'),
    ('services_unassigned', '未派服務數', '筆', 'down'),
    ('success_rate', '成功率', '%', 'up'),
    ('clients_unassigned', '未派案家數', '案', 'down'),
    ('transition_minutes_total', '總轉場分鐘', '分', 'down'),
    ('transition_minutes_per_worker_day', '平均每位居服員每日轉場分鐘', '分', 'down'),
    ('transition_wage_total', '轉場工資合計', '元', 'down'),
    ('overtime_hours_total', '加班時數合計', '時', 'down'),
    ('near_overtime_limit_people', '接近月加班上限(≥80%)人數', '人', 'down'),
    ('primary_share', '主責（固定人選）服務占比', '%', 'up'),
    ('workers_per_client', '每個案家平均搭配人數', '人', 'down'),
    ('worker_hours_mean', '居服員月工時平均', '時', None),
    ('worker_hours_std', '居服員月工時標準差', '時', 'down'),
    ('worker_hours_max', '居服員月工時最大', '時', None),
    ('worker_hours_min', '居服員月工時最小', '時', None),
    ('worker_hours_range', '居服員月工時最大減最小', '時', 'down'),
    ('solve_seconds', '核心求解秒數', '秒', 'down'),
]


def parse_value(field, text):
    """依 PipelineConfig 欄位預設值的型別轉換 --set 的字串。"""
    default = field.default if field.default is not dataclasses.MISSING else field.default_factory()
    low = text.strip().lower()
    if low in ('none', 'null'):
        return None
    if isinstance(default, bool):
        if low in ('1', 'true', 'yes', 'on'):
            return True
        if low in ('0', 'false', 'no', 'off'):
            return False
        raise ValueError(f'{field.name} 需為 true/false：{text}')
    if isinstance(default, int):
        return int(text)
    if isinstance(default, (float, type(None))):
        return float(text)
    if isinstance(default, (dict, list)):
        return json.loads(text)
    return text


def parse_overrides(pairs):
    fields = {f.name: f for f in dataclasses.fields(e.PipelineConfig)}
    out = {}
    for pair in pairs or []:
        if '=' not in pair:
            sys.exit(f'--set 格式應為 欄位=值：{pair}')
        name, text = (s.strip() for s in pair.split('=', 1))
        if name not in fields:
            sys.exit(f'PipelineConfig 沒有欄位「{name}」。可用欄位：\n  ' + '\n  '.join(sorted(fields)))
        out[name] = parse_value(fields[name], text)
    return out


def base_config(overrides):
    # 與 app.py 的機構資料排班一致的固定參數；--set 可再覆寫。
    values = dict(enable_backup_caregiver=True, max_backup_caregivers=2, enable_support_caregiver=True,
                  time_flex_minutes=0.0, plan_load_mode='hours', fatigue_reference_hours=40.0, fatigue_weight=10.0)
    values.update(overrides)
    return e.PipelineConfig(**values)


def _finite(value):
    try:
        value = float(value)
    except (TypeError, ValueError):
        return None
    return value if math.isfinite(value) else None


def osrm_only_patches():
    """機車照常查 OSRM；大眾運輸不打 Google，改用引擎內建的備援估算（步行／機車車程×係數＋候車）。"""
    from unittest.mock import patch
    e.set_transit_fallback_enabled(True)

    def travel(lat1, lon1, lat2, lon2, transport_mode='機車', travel_min_per_km=3.0, *a, **k):
        if transport_mode == '機車':
            return e.get_osrm_travel_time(lat1, lon1, lat2, lon2, travel_min_per_km, transport_mode='機車')
        return e.transit_fallback_minutes(lat1, lon1, lat2, lon2, travel_min_per_km)
    return [patch.object(e, 'get_google_travel_time', side_effect=travel),
            patch.object(e, 'prefetch_google_transit_pairs', lambda *a, **k: None)]


def run(excel, start, end, overrides):
    cfg = base_config(overrides)
    t0 = time.perf_counter()
    x = e.load_agency_workbook(str(excel), start, end, daily_start='08:00', daily_end='21:00',
                               daily_cap=cfg.max_daily_hours, availability_confirmed=True, biweekly_anchor_confirmed=True)
    cg = x['Caregiver_Profiles'].copy()
    cg['臨時不可排時段'] = '[]'
    ids = set(cg['居服員ID'].astype(str))
    cl = normalize_clients(x['Client_Profiles'].copy(), ids)
    tasks = x['Tasks'].copy()
    tasks['使用Service_code重算'] = False
    tasks = e.apply_service_duration(tasks, x['Service_Code'].copy())
    cfg.client_primary_map = source_primary_map(cl, ids, {})
    cfg.primary_profiles = {str(r['居服員ID']): r.to_dict() for _, r in cg.iterrows()}
    load_seconds = time.perf_counter() - t0

    t1 = time.perf_counter()
    r = run_fixed_period_dispatch(tasks, cg, cl, cfg, primary_map=cfg.client_primary_map, worker_workday_seed={})
    solve_seconds = time.perf_counter() - t1

    t2 = time.perf_counter()
    moved = r.get('time_shifts')
    if moved is not None and not moved.empty:
        new = moved.set_index('任務ID')
        hit = tasks['任務ID'].isin(new.index)
        tasks.loc[hit, '時間窗_開始'] = tasks.loc[hit, '任務ID'].map(new['新開始'])
        tasks.loc[hit, '時間窗_結束'] = tasks.loc[hit, '任務ID'].map(new['新結束'])
    cfg.client_primary_map = dict(r['client_primary_map'])
    cfg.client_backup_map = dict(r.get('client_backup_map') or {})
    assigned = r['df_result'].copy()
    if not assigned.empty:
        assigned = assigned.loc[assigned['派單居服員'].notna() &
                                ~assigned['派單居服員'].astype(str).str.strip().isin(['', '未指派', 'nan', 'None'])]
        assigned = assigned.drop_duplicates('任務ID', keep='last')
    travel_tasks = e.merge_task_client_tables(tasks, cl)
    routes = r.get('route_result')
    if (routes is None or routes.empty) and not assigned.empty:
        routes = e.calculate_schedule_travel(assigned, travel_tasks, cg, cfg)
    task_pay, daily_pay, monthly_pay, _ = payroll_tables(travel_tasks, assigned, cg, routes, cfg)
    diagnostics = r['df_matches'].attrs.get('matching_diagnostics', [])
    dash = build_dashboard(tasks, assigned, diagnostics, task_pay, daily_pay, monthly_pay, routes,
                           cfg.client_primary_map, caregivers=cg)
    metrics = collect_metrics(dash, tasks, assigned, task_pay, daily_pay, monthly_pay, cg, solve_seconds)
    extra = {'load_seconds': round(load_seconds, 2), 'metrics_seconds': round(time.perf_counter() - t2, 2),
             'caregivers': len(cg), 'clients': int(tasks['案家ID'].nunique()),
             'idle_caregivers': metrics.pop('_idle'), 'unknown_transition_days': metrics.pop('_unknown_days'),
             'dispatch_stats': r.get('dispatch_stats', {}), 'travel_legs': dash['travel_stats']}
    return cfg, metrics, extra


def collect_metrics(dash, tasks, assigned, task_pay, daily_pay, monthly_pay, cg, solve_seconds):
    s = dash['summary']
    t = tasks.drop_duplicates('任務ID')
    done = set(assigned['任務ID'].astype(str)) if not assigned.empty else set()
    pending = t.loc[~t['任務ID'].astype(str).isin(done)]
    num = lambda frame, col: pd.to_numeric(frame[col], errors='coerce') if not frame.empty and col in frame else pd.Series(dtype=float)
    transfer = num(daily_pay, '轉場工時')
    # 每日轉場工時：不同案家間的交通＋工作緩衝（首案通勤、末案返家不計），與薪資計算同一套。
    m = {
        'services_assigned': s['assigned'], 'services_unassigned': s['unassigned'], 'success_rate': s['rate'],
        'clients_unassigned': int(pending['案家ID'].nunique()),
        'transition_minutes_total': transfer.sum() * 60 if transfer.notna().all() else float('nan'),
        'transition_minutes_per_worker_day': transfer.mean() * 60,
        'transition_wage_total': num(task_pay, '轉場工資').sum(),
        'overtime_hours_total': num(monthly_pay, '加班時數').sum(min_count=1),
        'near_overtime_limit_people': s['near_limit_people'],
        'primary_share': s['primary_share'],
        '_unknown_days': int(transfer.isna().sum()),
    }
    if not assigned.empty:
        pairs = assigned[['任務ID', '派單居服員']].astype(str).merge(
            t[['任務ID', '案家ID']].astype(str), on='任務ID', how='left')
        m['workers_per_client'] = pairs.groupby('案家ID')['派單居服員'].nunique().mean()
    else:
        m['workers_per_client'] = float('nan')
    # 全部居服員都列入（整月沒排班者以 0 工時計），才看得出負荷是否集中在少數人。
    hours = dash['workload'].set_index('居服員')['總工時'] if not dash['workload'].empty else pd.Series(dtype=float)
    hours = pd.to_numeric(hours.reindex(cg['居服員ID'].astype(str).unique()), errors='coerce').fillna(0.0)
    m.update(worker_hours_mean=hours.mean(), worker_hours_std=hours.std(ddof=0), worker_hours_max=hours.max(),
             worker_hours_min=hours.min(), worker_hours_range=hours.max() - hours.min(),
             solve_seconds=solve_seconds, _idle=int((hours == 0).sum()))
    return {k: (_finite(v) if not isinstance(v, int) else v) for k, v in m.items()}


def git_commit():
    try:
        head = subprocess.run(['git', 'rev-parse', '--short', 'HEAD'], cwd=ROOT, capture_output=True, text=True).stdout.strip()
        dirty = subprocess.run(['git', 'status', '--porcelain', '--untracked-files=no'], cwd=ROOT, capture_output=True, text=True).stdout.strip()
        return head + ('-dirty' if dirty else '') if head else ''
    except OSError:
        return ''


def settings_tag(travel, overrides, seed):
    parts = [f'{k}={v}' for k, v in sorted(overrides.items())]
    tag = '_'.join(parts) or 'default'
    tag = re.sub(r'[^\w.=-]+', '-', tag)
    if len(tag) > 60:
        tag = tag[:40] + '-' + hashlib.sha1(tag.encode()).hexdigest()[:8]
    return f'{travel}_seed{seed}_{tag}'


def fmt(value):
    if value is None:
        return '—'
    if isinstance(value, int):
        return f'{value:,}'
    return f'{value:,.2f}' if abs(value) < 1000 else f'{value:,.1f}'


def to_markdown(report):
    meta, m, extra = report['meta'], report['metrics'], report['extra']
    lines = [f"# 排班基準測試 {meta['timestamp']}", '',
             f"- Excel：`{meta['excel']}`（sha256 {meta['excel_sha256'][:12]}）",
             f"- 期間：{meta['start']} ～ {meta['end']}；交通：{meta['travel']}；種子：{meta['seed']}",
             f"- 覆寫參數：{json.dumps(meta['overrides'], ensure_ascii=False) if meta['overrides'] else '無'}",
             f"- 版本：{meta['build']}；git {meta['git'] or '—'}",
             f"- 居服員 {extra['caregivers']} 人（整期未排 {extra['idle_caregivers']} 人）；案家 {extra['clients']} 案；"
             f"讀檔 {extra['load_seconds']} 秒、指標 {extra['metrics_seconds']} 秒",
             '', '| 指標 | 數值 | 單位 |', '|---|---:|---|']
    lines += [f'| {label} | {fmt(m.get(key))} | {unit} |' for key, label, unit, _ in METRICS]
    if extra['unknown_transition_days']:
        lines += ['', f"> 有 {extra['unknown_transition_days']} 個居服員工作日的轉場時間未知，總轉場分鐘留空。"]
    return '\n'.join(lines) + '\n'


def compare(path_a, path_b):
    a, b = (json.loads(pathlib.Path(p).read_text(encoding='utf-8')) for p in (path_a, path_b))
    lines = [f'# 基準測試比較', '', f'- A：`{path_a}`', f'- B：`{path_b}`']
    for key, label in [('excel_sha256', 'Excel 內容'), ('start', '起日'), ('end', '迄日'), ('travel', '交通模式')]:
        if a['meta'].get(key) != b['meta'].get(key):
            lines.append(f"- ⚠️ {label}不同：A={a['meta'].get(key)}　B={b['meta'].get(key)}（結果不可直接比較）")
    keys = sorted(set(a['meta']['overrides']) | set(b['meta']['overrides']))
    changed = [f"{k}: {a['meta']['overrides'].get(k, '預設')} → {b['meta']['overrides'].get(k, '預設')}"
               for k in keys if a['meta']['overrides'].get(k) != b['meta']['overrides'].get(k)]
    lines.append('- 參數差異：' + ('；'.join(changed) if changed else '無'))
    lines += ['', '| 指標 | A | B | 差異 | 變化% | 判讀 |', '|---|---:|---:|---:|---:|:---:|']
    for key, label, unit, better in METRICS:
        va, vb = a['metrics'].get(key), b['metrics'].get(key)
        if va is None or vb is None:
            lines.append(f'| {label}（{unit}） | {fmt(va)} | {fmt(vb)} | — | — | |')
            continue
        diff = vb - va
        pct = '0.0%' if abs(diff) < 1e-9 else f'{diff / abs(va) * 100:+.1f}%' if va else '—'
        verdict = '' if not better or abs(diff) < 1e-9 else ('較佳' if (diff > 0) == (better == 'up') else '較差')
        sign = '+' if diff > 0 else ''
        lines.append(f'| {label}（{unit}） | {fmt(va)} | {fmt(vb)} | {sign}{fmt(diff)} | {pct} | {verdict} |')
    lines += ['', '> 核心求解秒數受機器負載影響，請在同一台機器上比較。']
    return '\n'.join(lines) + '\n'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('excel', nargs='?', default=str(DEFAULT_EXCEL))
    ap.add_argument('--start', default='2026-10-01')
    ap.add_argument('--end', default='2026-10-31')
    ap.add_argument('--travel', choices=['sandbox', 'osrm'], default='sandbox')
    ap.add_argument('--set', action='append', default=[], metavar='欄位=值', help='覆寫 PipelineConfig 欄位，可重複')
    ap.add_argument('--google', action='store_true', help='osrm 模式下大眾運輸改查 Google（會產生費用，與網頁相同）')
    ap.add_argument('--seed', type=int, default=0)
    ap.add_argument('--out', default=str(ROOT / 'benchmarks'), help='輸出資料夾')
    ap.add_argument('--compare', nargs=2, metavar=('A.json', 'B.json'), help='比較兩次結果（不排班）')
    args = ap.parse_args()

    out_dir = pathlib.Path(args.out)
    if args.compare:
        text = compare(*args.compare)
        print(text)
        out_dir.mkdir(parents=True, exist_ok=True)
        name = f"compare_{pathlib.Path(args.compare[0]).stem}__vs__{pathlib.Path(args.compare[1]).stem}"[:150]
        (out_dir / f'{name}.md').write_text(text, encoding='utf-8')
        return

    excel = pathlib.Path(args.excel)
    if not excel.exists():
        sys.exit(f'找不到 Excel：{excel}')
    overrides = parse_overrides(args.set)
    random.seed(args.seed)
    np.random.seed(args.seed)

    patches = []
    if args.travel == 'sandbox':
        import sandbox_travel
        patches = sandbox_travel.patches()
    else:
        os.environ['CAREFLOW_TRAVEL_PROVIDER'] = 'osrm'
        e.begin_travel_query_run()
        e.set_google_queries_enabled(args.google)
        if not args.google:
            patches = osrm_only_patches()
    for p in patches:
        p.start()
    try:
        cfg, metrics, extra = run(excel, args.start, args.end, overrides)
    finally:
        for p in patches:
            p.stop()

    stamp = datetime.now()
    report = {
        'meta': {'timestamp': stamp.isoformat(timespec='seconds'), 'excel': str(excel),
                 'excel_sha256': hashlib.sha256(excel.read_bytes()).hexdigest(),
                 'start': args.start, 'end': args.end,
                 'travel': args.travel + ('+google' if args.travel == 'osrm' and args.google else ''), 'seed': args.seed,
                 'overrides': overrides, 'build': CAREFLOW_BUILD, 'git': git_commit(),
                 'config': {k: v for k, v in dataclasses.asdict(cfg).items()
                            if k not in ('client_primary_map', 'client_backup_map', 'primary_profiles', 'work_day_minutes',
                                         'dispatch_cache', 'worker_workdays', 'worker_workday_seed', 'task_support_map')}},
        'metrics': metrics,
        'labels': {key: label for key, label, _, _ in METRICS},
        'extra': extra,
    }
    out_dir.mkdir(parents=True, exist_ok=True)
    base = out_dir / f"{stamp:%Y%m%d-%H%M%S}_{args.start}_{args.end}_{settings_tag(report['meta']['travel'].replace('+', '-'), overrides, args.seed)}"
    pathlib.Path(f'{base}.json').write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    text = to_markdown(report)
    pathlib.Path(f'{base}.md').write_text(text, encoding='utf-8')
    print(text)
    print(f'已輸出：{base}.json / .md')


if __name__ == '__main__':
    main()
