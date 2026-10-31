"""Historical owners are reserved once; new clients are allocated atomically.

Every visit uses the central constraints. Only actual leave permits a backup.
This constructs a feasible draft, not a proven global optimum.
"""
CAREFLOW_BUILD = '2026.10.18-v19'
from dataclasses import replace
from collections import Counter
import re
import time
import pandas as pd
from ortools.linear_solver import pywraplp
from agency_rules import leave_overlap, clean, finite, monthly_overtime_error, consecutive_workdays_error
from source_fields import source_primary_map, seniority_score, resolve_worker
from care_policy import dementia_status, special_qualification_error, DEFAULT_CODE_CERTS
from recurring_schedule import visit_pattern, representative_days
from dispatch_workflow import blocked_reason
from caregiver_engine import (
    _evaluate_reassignment, _check_hard_constraints, _build_cg_busy_blocks,
    calculate_schedule_travel, caregiver_has_cert, calculate_task_revenue_and_salary,
    merge_task_client_tables, run_phase1_matching, validate_dispatch_inputs,
    prefetch_osrm_travel_times,
    warm_transit_routes, calc_travel_minutes, calc_distance_km,
    iter_service_codes, SPECIAL_CERT_REQUIREMENTS,
)

MATCH_COLUMNS = ['任務ID','案家ID','居服員ID','適配度分數','預估交通時間(分)',
    '地點緯度','地點經度','任務開始時間','任務結束時間','優先級',
    '是否歷史首選','具備失智症20小時認證(0/1)','常用交通工具']
RESULT_COLUMNS = ['任務ID','案家ID','派單居服員','適配分數','預估車程(分)',
    '服務時段','任務優先級','地點緯度','地點經度','預估服務營收','預估居服員薪資',
    '原首選替換原因','排班來源','排班日期','固定居服員ID','是否請假代班','代班原因','代班排序理由']


BACKUP_STAGES = ('歷史首選核對','新案固定人選核對','新案整案配對')


def failure_kind(detail):
    for needles,label,action in [
        (['缺特殊訓練','專長認證'],'必要證照','首選需具備所有必要訓練；核對取得資格'),
        (['到期','到職'],'資格日期','核對證明到期日與到職日'),
        (['重疊','不能同時'],'時段重疊','調整服務時段或尚無歷史首選的新案整案人選'),
        (['休息不足'],'休息不足','保留不含交通的休息空檔'),
        (['空檔','交通加緩衝'],'轉場空檔不足','延長轉場空檔或調整新案整案人選'),
        (['不排班','可排班星期'],'可服務星期','核對休息日與固定服務星期'),
        (['連續上班'],'連續工作天數','不得連續上班逾上限天數；核對休息日安排或改由人工插單並確認提醒'),
        (['工時超標'],'工時上限','調整服務時段與工作量'),
        (['交通','座標'],'交通資料','核對座標與路網連線'),
    ]:
        if any(word in detail for word in needles):
            return label,action
    return '資料或其他限制','查看逐人原因並核對服務計畫'


def run_fixed_period_dispatch(tasks,caregivers,clients,config,primary_map=None,progress=None,routine_templates=None,
        worker_workday_seed=None):
    """Choose new owners on the recurring template; extend fixed routes.

    Regular motorcycle worker/day routes are copied after date eligibility and
    month-ledger checks. Leave, changed plans and transit are checked separately.

    worker_workday_seed：{居服員ID:{日期字串,...}}，本期間開始日之前（跨月）已確定的上班日，
    供「連續上班天數」硬性規則回溯種子；只套用在真實日期序列的那一次 _run_expanded_dispatch，
    代表性週型（representative_days）那一次日期是抽樣、非連續日曆，不套用此規則。
    """
    tasks = merge_task_client_tables(tasks,clients).copy()
    tasks['日期'] = pd.to_datetime(tasks['日期'],errors='coerce').dt.strftime('%Y-%m-%d')
    validate_dispatch_inputs(tasks,caregivers)
    original = source_primary_map(clients,set(caregivers['居服員ID'].astype(str)),primary_map)
    unknown = set(tasks['案家ID'].astype(str)) - set(original)
    if not unknown:
        result = _run_expanded_dispatch(tasks,caregivers,clients,config,original,progress,day_templates=routine_templates,
            worker_workday_seed=worker_workday_seed)
        result['dispatch_stats'].update(template_service_count=len(representative_days(tasks,caregivers)),
                                        period_service_count=len(tasks))
        return result
    basis = representative_days(tasks,caregivers)
    notify = (lambda v,label:progress(v*.55,'週期模板：'+label)) if progress else None
    regular = _run_expanded_dispatch(basis,caregivers,clients,config,original,notify,
                                    compute_routes=False,day_templates=routine_templates,allow_backup=False,
                                    enforce_consecutive_days=False)
    notify = (lambda v,label:progress(.55+v*.45,'延展班表：'+label)) if progress else None
    # Pass the chosen owners directly. Candidate search does not repeat on each
    # later week. Unresolved clients retain their template diagnostic.
    result = _run_expanded_dispatch(tasks,caregivers,clients,config,
        regular['client_primary_map'],notify,day_templates=regular['_day_templates'],
        historical_labels=original,skip_owner_search=True,
        missing_diagnostics=regular['diagnostics'],worker_workday_seed=worker_workday_seed)
    result['dispatch_stats'].update(template_service_count=len(basis),period_service_count=len(tasks),
        template_day_count=basis['日期'].nunique(),ownership_search_clients=len(unknown),
        historical_clients=len(original),new_clients=len(unknown))
    return result


def _run_expanded_dispatch(tasks,caregivers,clients,config,primary_map=None,progress=None,
        compute_routes=True,day_templates=None,historical_labels=None,
        skip_owner_search=False,missing_diagnostics=None,allow_backup=True,
        enforce_consecutive_days=True,worker_workday_seed=None):
    tasks = merge_task_client_tables(tasks,clients).copy()
    tasks['日期'] = pd.to_datetime(tasks['日期'],errors='coerce').dt.strftime('%Y-%m-%d')
    if tasks['日期'].isna().any():
        raise ValueError('任務含無效日期，請先修正。')
    validate_dispatch_inputs(tasks,caregivers)
    motor = caregivers[caregivers['常用交通工具'].astype(str).eq('機車')]
    if not motor.empty:
        prefetch_osrm_travel_times(motor[['服務起點_緯度(家)','服務起點_經度(家)']].drop_duplicates().itertuples(index=False,name=None),
            tasks[['服務地點_緯度','服務地點_經度']].drop_duplicates().itertuples(index=False,name=None),
            config.travel_min_per_km,transport_mode='機車')
        # 案家↔案家（轉場檢查會用到）也一次批次查完，避免排班中逐筆 HTTP。
        places = list(tasks[['服務地點_緯度','服務地點_經度']].drop_duplicates().itertuples(index=False,name=None))
        prefetch_osrm_travel_times(places,places,config.travel_min_per_km,transport_mode='機車')
    tasks = tasks.sort_values(['日期','時間窗_開始','任務ID']).reset_index(drop=True)
    profiles = {str(r['居服員ID']):r.to_dict() for _,r in caregivers.iterrows()}
    profile_rows = {worker:pd.Series(profile) for worker,profile in profiles.items()}
    owners = source_primary_map(clients,set(profiles),primary_map)
    fixed = dict(owners)
    historical = dict(owners) if historical_labels is None else dict(historical_labels)
    config.client_primary_map,config.primary_profiles,config.work_day_minutes = owners,profiles,{}
    # 連續上班天數的硬性規則狀態：只有真實日期序列的那一次套用種子並累計；
    # 代表性週型（representative_days）那一次日期是抽樣、非連續日曆，一律從空白起算且不強制。
    config.worker_workdays = ({str(k):set(v) for k,v in (worker_workday_seed or {}).items()}
                               if enforce_consecutive_days else {})
    backups = {}
    config.client_backup_map = backups
    supports = {}
    config.task_support_map = supports
    days = {day:group.copy() for day,group in tasks.groupby('日期')}
    config.dispatch_cache = {'caregivers':caregivers,'people':profiles,'busy':_build_cg_busy_blocks(caregivers),
        'days':{id(frame):{'frame':frame,'ids':set(frame['任務ID']),
            'lookup':frame.set_index('任務ID').to_dict('index')} for frame in days.values()}}
    financial = calculate_task_revenue_and_salary(tasks,config).set_index('任務ID').to_dict('index')
    assigned = {day:pd.DataFrame(columns=['任務ID','派單居服員']) for day in days}
    rows,matches,diagnostics,rankings,leave_tasks = [],[],[],{},[]
    served = Counter()   # (案家ID, 居服員ID) → 本次班表已服務次數；代班挑選的「熟悉」依據
    checked_patterns = {}
    hard_results,occupied_slots = {},{}
    task_records = {t['任務ID']:t for t in tasks.to_dict('records')}
    task_patterns = {tid:visit_pattern(t) for tid,t in task_records.items()}
    # 大眾運輸的固定人選：整個計畫的路段先批次預抓 Google（同星期同時刻只查一次），排班中就不必逐筆等候。
    _transit_owned = [(tid,fixed[str(t['案家ID'])]) for tid,t in task_records.items()
                      if fixed.get(str(t['案家ID'])) and profiles.get(fixed[str(t['案家ID'])],{}).get('常用交通工具')=='大眾運輸']
    if _transit_owned:
        try:
            warm_transit_routes(pd.DataFrame(_transit_owned,columns=['任務ID','派單居服員']),task_records,profiles,config)
        except Exception:
            pass  # 預抓只是加速；真正需要時的逐筆查詢仍會依原本規則檢查並回報錯誤。
    pattern_of = lambda t:task_patterns[t['任務ID']]
    calendar = {}
    for worker,profile in profiles.items():
        joined = pd.to_datetime(profile.get('到職日'),errors='coerce')
        expiry = pd.to_datetime(profile.get('長照服務人員到期日'),errors='coerce')
        calendar[worker] = (joined.date() if pd.notna(joined) else None,
                            expiry.date() if pd.notna(expiry) else None,
                            set(clean(profile.get('請假或不排班日期')).split(',')))
    templates = dict(day_templates or {})
    profile_signatures = {worker:repr(profile) for worker,profile in profiles.items()}
    rule_signature = (config.buffer_mins,config.travel_min_per_km,config.continuous_work_limit_mins,
                      config.mandatory_break_mins,config.max_daily_hours,config.normal_daily_hours)
    copied_days, checked_days = 0, 0
    normal_checks = {}
    month_overtime = {}

    def set_work_minutes(worker,day,minutes):
        old = config.work_day_minutes.get((worker,day),0.0)
        key = worker,day[:7]
        month_overtime[key] = month_overtime.get(key,0.0) + max(0.0,minutes/60-config.normal_daily_hours) - max(0.0,old/60-config.normal_daily_hours)
        config.work_day_minutes[(worker,day)] = minutes

    def template_month_fits(worker,day,minutes):
        baseline = max(0.0,finite(profiles.get(worker,{}).get('排班期外本月加班時數'),0.0))
        old = config.work_day_minutes.get((worker,day),0.0)
        total = baseline + month_overtime.get((worker,day[:7]),0.0) - max(0.0,old/60-config.normal_daily_hours) + max(0.0,minutes/60-config.normal_daily_hours)
        return total <= config.max_monthly_overtime_hours + 1e-7

    def template_key(records,worker):
        patterns = [pattern_of(t) for t in records]
        if len(set(patterns)) != len(patterns) or profiles.get(worker,{}).get('常用交通工具') != '機車':
            return None
        states = tuple(sorted((p,hard_check(t,worker,config) or '',leave_overlap(t,profiles.get(worker)))
                              for p,t in zip(patterns,records)))
        return worker,profile_signatures[worker],rule_signature,states

    def hard_check(task,worker,cfg):
        owner = cfg.client_primary_map.get(str(task['案家ID']))
        joined,expiry,excluded = calendar[worker]
        date = pd.Timestamp(task['日期']).date()
        profile = profiles[worker]
        dates = (bool(joined and date<joined),bool(expiry and date>expiry),
                 date.isoformat() in excluded,leave_overlap(task,profile),blocked_reason(task,profile) or '')
        foreign = (owner,leave_overlap(task,profiles.get(owner)),
                   worker in cfg.client_backup_map.get(str(task['案家ID']),()),
                   cfg.task_support_map.get(str(task['任務ID']))==worker) if owner and owner!=worker else ()
        key = pattern_of(task),worker,foreign,dates
        if key not in hard_results:
            hard_results[key] = _check_hard_constraints(pd.Series(task) if isinstance(task,dict) else task,profile_rows[worker],cfg)
        return hard_results[key]

    def check_visit(task,worker,current,cfg):
        profile = profiles.get(worker)
        if profile is None:
            return _evaluate_reassignment(task['任務ID'],worker,days[task['日期']],current,caregivers,cfg)
        hard = hard_check(task,worker,cfg)
        if hard:
            return {'available':False,'detail':hard,'reason':'hard_constraint'}
        if enforce_consecutive_days:
            # 絕對硬性規則：AI 不得把居服員排到連續上班第 7 天（見 agency_rules.consecutive_workdays_error）。
            # 這是隨排班進度變動的狀態性檢查，不能跟上面 hard_check 共用記憶化快取，因此獨立判斷、每次都查。
            consecutive = consecutive_workdays_error(worker,task['日期'],cfg)
            if consecutive:
                return {'available':False,'detail':consecutive,'reason':'consecutive_days'}
        lookup = cfg.dispatch_cache['days'][id(days[task['日期']])]['lookup']
        others = current['任務ID'].to_numpy()[current['派單居服員'].to_numpy()==worker] if len(current) else []
        neighbors = []
        for tid in others:
            row = lookup[tid]
            if task['時間窗_開始'] < row['時間窗_結束'] and row['時間窗_開始'] < task['時間窗_結束']:
                return {'available':False,'detail':'服務時間重疊，不能同時接兩案。','reason':'time_or_rest'}
            neighbors.append((row['案家ID'],row['時間窗_開始'],row['時間窗_結束'],
                              row.get('服務地點_緯度'),row.get('服務地點_經度')))
        signature = (worker,pd.Timestamp(task['日期']).isoweekday(),task['案家ID'],
            task['時間窗_開始'],task['時間窗_結束'],task.get('服務地點_緯度'),task.get('服務地點_經度'),
            tuple(sorted(neighbors)))
        # OSRM is a static road model. Google TRANSIT always checks the actual
        # date/time. Eligibility and the monthly ledger are checked every visit.
        cached = checked_patterns.get(signature) if profile.get('常用交通工具')=='機車' else None
        if cached:
            error = monthly_overtime_error(worker,task['日期'],cached['working_minutes'],cfg,profile)
            if error:
                return {'available':False,'detail':error,'reason':'monthly_overtime'}
            return dict(cached)
        result = _evaluate_reassignment(task['任務ID'],worker,days[task['日期']],current,caregivers,cfg)
        if result['available'] and profile.get('常用交通工具')=='機車':
            checked_patterns[signature] = dict(result)
        return result

    def note(task,stage,detail,reasons=None,count=0,main_reason=None):
        category,action = failure_kind(main_reason or detail)
        if stage=='請假代班':
            category,action='請假代班缺人','調整代班服務時段或安排合格支援人員'
        diagnostics.append({'任務ID':task['任務ID'],'案家ID':task['案家ID'],'來源列':task.get('來源列'),
            '失敗階段':stage,'問題分類':category,'建議處理':action,'候選數':count,
            '剔除原因':detail,'逐人排班原因':reasons or {}})

    def commit(task,worker,checked,absence='',batch=False,source='',rank_reason=''):
        day = task['日期']
        served[(str(task['案家ID']),str(worker))] += 1
        if not batch:
            assigned[day] = pd.concat([assigned[day],pd.DataFrame([{'任務ID':task['任務ID'],'派單居服員':worker}])],ignore_index=True)
        config.worker_workdays.setdefault(str(worker),set()).add(str(day)[:10])
        set_work_minutes(worker,day,checked['working_minutes'])
        occupied_slots.setdefault((worker,day),[]).append(task)
        if not absence and not source:
            normal_checks[task['任務ID']] = dict(checked)
        match = _match_row(task,worker,profiles,rankings,config,historical)
        matches.append(match)
        fin = financial[task['任務ID']]
        rows.append({'任務ID':task['任務ID'],'案家ID':task['案家ID'],'派單居服員':worker,
            '適配分數':match['適配度分數'],'預估車程(分)':float('nan'),
            '服務時段':f"{task['時間窗_開始']}-{task['時間窗_結束']}",
            '任務優先級':task.get('任務優先級','一般'),'地點緯度':task.get('服務地點_緯度'),
            '地點經度':task.get('服務地點_經度'),'預估服務營收':fin['預估服務營收'],
            '預估居服員薪資':fin['預估居服員薪資'],'原首選替換原因':absence,'代班原因':absence,
            '排班來源':source or ('請假代班' if absence else '歷史首選' if str(task['案家ID']) in historical else '新案固定人選'),
            '排班日期':day,'固定居服員ID':owners[str(task['案家ID'])],'是否請假代班':bool(absence),
            '代班排序理由':rank_reason})

    # ---- 轉場成本：同一居服員同一天依開始時間排序後「實際相鄰」的不同案家才計（含工作緩衝），
    # 與 calculate_schedule_travel／payroll_tables 的轉場定義一致；首案住家通勤以 travel_home_weight 計。
    # 大眾運輸居服員改用直線距離估算，比較派法時不額外觸發 Google 查詢；可不可派仍由 check_visit 決定。
    travel_mode = str(getattr(config,'travel_objective_mode','off') or 'off')
    travel_weight = max(0.0,finite(getattr(config,'travel_objective_weight',1.0),1.0))
    home_weight = max(0.0,finite(getattr(config,'travel_home_weight',0.0),0.0))
    leg_cache = {}

    def leg_minutes(worker,a,b):
        transit = profiles[worker].get('常用交通工具')=='大眾運輸'
        key = (*a,*b,transit)
        if key not in leg_cache:
            if not all(v==v for v in (*a,*b)):
                leg_cache[key] = 0.0
            else:
                try:
                    leg_cache[key] = (calc_distance_km(*a,*b)*float(config.travel_min_per_km) if transit
                                      else float(calc_travel_minutes(*a,*b,config)))
                except Exception:
                    leg_cache[key] = calc_distance_km(*a,*b)*float(config.travel_min_per_km)
        return leg_cache[key]

    def route_cost(worker,day,tids,paid_only=False):
        # paid_only：只算付費轉場（不同案家間車程＋緩衝），不含首案住家通勤。
        lookup = config.dispatch_cache['days'][id(days[day])]['lookup']
        profile = profiles[worker]
        previous,previous_client,cost = None,None,0.0
        for tid in sorted(tids,key=lambda t:(lookup[t]['時間窗_開始'],t)):
            row = lookup[tid]
            here = finite(row.get('服務地點_緯度')),finite(row.get('服務地點_經度'))
            if previous is None:
                if paid_only:
                    previous,previous_client = here,str(row['案家ID'])
                    continue
                home = finite(profile.get('服務起點_緯度(家)')),finite(profile.get('服務起點_經度(家)'))
                cost += home_weight*leg_minutes(worker,home,here)
            elif str(row['案家ID']) != previous_client:
                cost += leg_minutes(worker,previous,here)+config.buffer_mins
            previous,previous_client = here,str(row['案家ID'])
        return cost

    def worker_ids(frame,worker):
        return list(frame['任務ID'].to_numpy()[frame['派單居服員'].astype(str).to_numpy()==str(worker)]) if len(frame) else []

    def added_cost(worker,frames):
        """frames：{日期: 加入後的當日班表}；回傳相對 assigned 的新增轉場成本。"""
        return sum(route_cost(worker,day,worker_ids(frame,worker))-route_cost(worker,day,worker_ids(assigned[day],worker))
                   for day,frame in frames.items())

    def objective(served_count,cost):
        # 派出一筆＝1,000,000；轉場成本再大也不會超過一筆服務的價值，派出數永遠優先。
        return served_count*1_000_000-travel_weight*cost

    # Reserve each fixed worker's weekly route once, then copy identical days.
    for i,(day,frame) in enumerate(days.items()):
        if progress:
            progress(.35*i/max(len(days),1),f'直接核對歷史首選 {i+1}/{len(days)} 日 · {day}')
        worker_groups = {}
        for tid in frame['任務ID']:
            record = task_records[tid]
            worker = fixed.get(str(record['案家ID']))
            if worker:
                worker_groups.setdefault(worker,[]).append(record)
        for worker,records in sorted(worker_groups.items()):
            records.sort(key=lambda t:(str(t['案家ID']) not in historical,t['時間窗_開始'],t['任務ID']))
            key = template_key(records,worker)
            cached = templates.get(key) if key else None
            # 連續上班天數規則是逐日曆日生效的，「相同週型直接複製」這條快速路徑假設各天可互換，
            # 在這點上不成立：同一人、同一週型，連續天數到了第7天就必須擋，不能照抄前面天的核可結果。
            # 不滿足就退回下面逐任務 check_visit 的慢路徑，讓硬性規則照常生效、原因照常記錄。
            template_blocked_by_streak = enforce_consecutive_days and bool(consecutive_workdays_error(worker,day,config))
            if cached and template_month_fits(worker,day,cached['working_minutes']) and not template_blocked_by_streak:
                lookup = {pattern_of(t):t for t in records}
                substitutions = {old:lookup[pattern]['任務ID'] for pattern,old in cached['old_ids'].items()}
                token = re.compile('|'.join(re.escape(k) for k in sorted(substitutions,key=len,reverse=True)))
                translate = lambda value:token.sub(lambda m:str(substitutions[m.group()]),str(value))
                batch_rows = []
                for pattern,checked in cached['accepted']:
                    task = lookup[pattern]
                    commit(task,worker,checked,batch=True)
                    batch_rows.append({'任務ID':task['任務ID'],'派單居服員':worker})
                if batch_rows:
                    assigned[day] = pd.concat([assigned[day],pd.DataFrame(batch_rows)],ignore_index=True)
                    set_work_minutes(worker,day,cached['working_minutes'])
                for pattern,detail in cached['failed']:
                    task = lookup[pattern]
                    detail = translate(detail)
                    stage = '歷史首選核對' if str(task['案家ID']) in historical else '新案固定人選核對'
                    note(task,stage,f'固定人選 {worker}：'+detail,{worker:detail})
                for task in lookup.values():
                    if leave_overlap(task,profiles.get(worker)):
                        leave_tasks.append(pd.Series(task))
                copied_days += 1
                continue
            checked_days += 1
            accepted,failed = [],[]
            for task_record in records:
                task = pd.Series(task_record)
                if leave_overlap(task,profiles.get(worker)):
                    leave_tasks.append(task)
                    continue
                checked = check_visit(task,worker,assigned[day],config)
                if checked['available']:
                    commit(task,worker,checked)
                    accepted.append((pattern_of(task),dict(checked)))
                else:
                    stage = '歷史首選核對' if str(task['案家ID']) in historical else '新案固定人選核對'
                    note(task,stage,f'固定人選 {worker}：'+checked['detail'],{worker:checked['detail']})
                    failed.append((pattern_of(task),checked['detail']))
            if key and not any('每月加班' in detail or '同月既有工時' in detail for _,detail in failed):
                templates[key] = {'accepted':accepted,'failed':failed,
                    'working_minutes':config.work_day_minutes.get((worker,day),0),
                    'old_ids':{pattern_of(t):t['任務ID'] for t in records}}

    remaining,static_reasons = {},{}
    for client,visits in tasks.groupby('案家ID',sort=True):
        client = str(client)
        if client in fixed:
            continue
        if skip_owner_search:
            source = missing_diagnostics
            details = source.loc[source['案家ID'].astype(str).eq(client)].to_dict('records') if source is not None and not source.empty else []
            detail = details[0].get('剔除原因','週期模板沒有合格且可承接的固定人選。') if details else '週期模板沒有合格且可承接的固定人選。'
            for _,task in visits.iterrows():
                note(task,'新案整案配對',detail,count=details[0].get('候選數',0) if details else 0)
            continue
        columns = ['星期','時間窗_開始','時間窗_結束','服務地點_緯度','服務地點_經度']
        columns += [c for c in visits if c.startswith(('Service_Code_','Units_'))]
        # The last date also checks future expiry/join dates before selection.
        samples = pd.concat([visits.drop_duplicates(columns),visits.tail(1)]).drop_duplicates('任務ID')
        candidates,errors = [],{}
        for worker,profile in profiles.items():
            qualified = True
            for _,task in samples.iterrows():
                needed = set(clean(task.get('服務碼必要證照')).split('、')) - {''}
                needed.update(DEFAULT_CODE_CERTS[code] for code,units in iter_service_codes(task) if units>0 and code in DEFAULT_CODE_CERTS)
                reason = special_qualification_error(task,profile,needed)
                special = SPECIAL_CERT_REQUIREMENTS.get(task.get('特殊照護需求'))
                if special and not any(caregiver_has_cert(profile,cert) for cert in special):
                    reason = '缺乏該需求類別之法定專長認證'
                if reason:
                    errors[worker] = reason
                    qualified = False
                    break
            if qualified:
                candidates.append(worker)
        static_reasons[client] = errors
        remaining[client] = visits,candidates,samples

    def qualification_reason(task,profile):
        needed = set(clean(task.get('服務碼必要證照')).split('、')) - {''}
        needed.update(DEFAULT_CODE_CERTS[code] for code,units in iter_service_codes(task) if units>0 and code in DEFAULT_CODE_CERTS)
        reason = special_qualification_error(task,profile,needed)
        special = SPECIAL_CERT_REQUIREMENTS.get(task.get('特殊照護需求'))
        if special and not any(caregiver_has_cert(profile,cert) for cert in special):
            reason = '缺乏該需求類別之法定專長認證'
        return reason

    def no_overlap(worker,samples):
        for task in samples:
            if leave_overlap(task,profiles[worker]):
                return True
            if hard_check(task,worker,config):
                continue
            day = task['日期']
            for other in occupied_slots.get((worker,day),[]):
                if task['時間窗_開始'] < other['時間窗_結束'] and other['時間窗_開始'] < task['時間窗_結束']:
                    break
            else:
                return True
        return False

    total_new = len(remaining)
    visit_records = {c:v[0].to_dict('records') for c,v in remaining.items()}
    overlap_cache = {}
    def available_workers(client,candidates):
        # 結果只取決於該居服員已佔用時段；某人新增班表後，只作廢他的快取。
        out = []
        for w in candidates:
            key = client,w
            if key not in overlap_cache:
                overlap_cache[key] = no_overlap(w,visit_records[client])
            if overlap_cache[key]:
                out.append(w)
        return out
    def coverage(checked_rows):
        return sum(float(task.get('_template_occurrences',1)) for task,_ in checked_rows)
    while remaining:
        choices = []
        for client,(visits,candidates,samples) in remaining.items():
            available = available_workers(client,candidates)
            choices.append((len(available),-float((visits['服務歷時(分鐘)']*visits.get('_template_occurrences',1)).sum()),client,available))
        _,_,client,candidates = min(choices,key=lambda item:item[:3])
        visits,original,samples = remaining.pop(client)
        if progress:
            progress(.35+.5*(total_new-len(remaining))/max(total_new,1),f'新案整案配對 {total_new-len(remaining)}/{total_new}')
        errors = dict(static_reasons[client])
        errors.update({w:'服務時間重疊，不能同時接兩案。' for w in original if w not in candidates})
        if candidates:
            exemplars = {}
            for worker in candidates:
                for _,task in visits.iterrows():
                    if leave_overlap(task,profiles[worker]) or not hard_check(task,worker,config):
                        exemplars.setdefault(task['任務ID'],[]).append(worker)
                        break
            frames=[]
            for tid,workers in exemplars.items():
                shortlist = caregivers[caregivers['居服員ID'].astype(str).isin(workers)].copy()
                shortlist['請假紀錄'],shortlist['請假日期'] = None,''
                frames.append(run_phase1_matching(visits[visits['任務ID'].eq(tid)],shortlist,config))
            scored = pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()
            rankings[client] = scored.set_index('居服員ID').to_dict('index') if not scored.empty else {}
            if scored.empty:
                errors['候選交通與資格'] = str(scored.attrs.get('matching_diagnostics',[]))
            candidates = scored.sort_values(['適配度分數','居服員ID'],ascending=[False,True])['居服員ID'].astype(str).tolist() if not scored.empty else []
        best = None
        for worker in candidates:
            # 同一候選人要接下該案整期間的所有服務（多天），worker_workdays 同樣要給獨立複本並逐筆
            # 累計，否則連續上班天數會被低估（理由見上方 try_backup 的註解）。
            trial = replace(config,client_primary_map={**owners,client:worker},work_day_minutes=dict(config.work_day_minutes),
                            worker_workdays={k:set(v) for k,v in config.worker_workdays.items()})
            tentative,checked_rows,task_errors = {},[],{}
            potentially_available = []
            for _,task in visits.iterrows():
                if leave_overlap(task,profiles[worker]):
                    continue
                reason = hard_check(task,worker,trial)
                if reason:
                    task_errors[task['任務ID']] = reason
                else:
                    potentially_available.append(task)
            potential = sum(float(t.get('_template_occurrences',1)) for t in potentially_available)
            if best and (potential < coverage(best[1]) if travel_mode=='weighted' else potential <= coverage(best[1])):
                continue
            for _,task in visits.iterrows():
                if leave_overlap(task,profiles[worker]):
                    continue
                if task['任務ID'] in task_errors:
                    continue
                day = task['日期']
                current = tentative.get(day,assigned[day])
                checked = check_visit(task,worker,current,trial)
                if not checked['available']:
                    errors[worker] = checked['detail']
                    task_errors[task['任務ID']] = checked['detail']
                    continue
                checked_rows.append((task,checked))
                tentative[day] = pd.concat([current,pd.DataFrame([{'任務ID':task['任務ID'],'派單居服員':worker}])],ignore_index=True)
                trial.work_day_minutes[(worker,day)] = checked['working_minutes']
                trial.worker_workdays.setdefault(worker,set()).add(str(day)[:10])
            if travel_mode=='weighted':
                # 不在第一位排滿的人停下：同樣排滿時比較整期新增轉場。
                score = objective(coverage(checked_rows),added_cost(worker,tentative)) if checked_rows else None
                if checked_rows and (best is None or score>best[3]):
                    best = worker,checked_rows,task_errors,score
                continue
            if checked_rows and (best is None or coverage(checked_rows)>coverage(best[1])):
                best = worker,checked_rows,task_errors,None
            if len(checked_rows)+sum(bool(leave_overlap(task,profiles[worker])) for _,task in visits.iterrows())==len(visits):
                break
        if best:
            worker,checked_rows,task_errors,_ = best
            owners[client] = worker
            for key in [k for k in overlap_cache if k[1]==worker]:
                del overlap_cache[key]
            for task,checked in checked_rows:
                commit(task,worker,checked)
            for _,task in visits.iterrows():
                if leave_overlap(task,profiles[worker]):
                    leave_tasks.append(task)
                elif task['任務ID'] in task_errors:
                    note(task,'新案固定人選核對',f'固定人選 {worker}：'+task_errors[task['任務ID']],
                         {worker:task_errors[task['任務ID']]},len(original))
        else:
            summary = '無法由同一人承接全期間服務。'+'；'.join(f'{w}：{r}' for w,r in errors.items())
            relevant = [errors[w] for w in original if w in errors] or list(errors.values())
            categories = Counter(failure_kind(r)[0] for r in relevant)
            common = categories.most_common(1)[0][0] if categories else ''
            main_reason = next((r for r in relevant if failure_kind(r)[0]==common),summary)
            for _,task in visits.iterrows():
                note(task,'新案整案配對',summary,errors,len(original),main_reason)

    # Save complete normal routes after new owners have been added. Backups are
    # deliberately excluded: they belong only to the actual leave date.
    diag = {d['任務ID']:d for d in diagnostics}
    if total_new:
        for day,frame in days.items():
            worker_groups = {}
            for tid in frame['任務ID']:
                record = task_records[tid]
                worker = owners.get(str(record['案家ID']))
                if worker:
                    worker_groups.setdefault(worker,[]).append(record)
            for worker,records in worker_groups.items():
                key = template_key(records,worker)
                failed = [(pattern_of(t),diag[t['任務ID']]['剔除原因'].split('：',1)[-1])
                          for t in records if t['任務ID'] in diag]
                if not key or any('每月加班' in detail or '同月既有工時' in detail for _,detail in failed):
                    continue
                templates[key] = {'accepted':[(pattern_of(t),normal_checks[t['任務ID']])
                    for t in records if t['任務ID'] in normal_checks],
                    'failed':failed,'working_minutes':config.work_day_minutes.get((worker,day),0),
                    'old_ids':{pattern_of(t):t['任務ID'] for t in records}}

    seed_backups = {}
    if '歷史備援' in clients:
        known_workers = set(profiles)
        for _,row in clients.iterrows():
            ids = [resolve_worker(v,known_workers) for v in re.split(r'[、,，;；\s]+',clean(row.get('歷史備援'))) if v]
            ids = tuple(dict.fromkeys(i for i in ids if i in known_workers))
            if ids:
                seed_backups[clean(row.get('案家ID'))] = ids

    # ---- 代班挑選順序（與畫面上的臨時異動共用 substitute_rank）：熟悉 → 前一案交通區段 → 本月工時 → 年資 ----
    from substitute_rank import build_profile as _sub_profile, order_workers as _sub_order, month_summary as _sub_month
    _estimate = lambda la1,lo1,la2,lo2: calc_distance_km(float(la1),float(lo1),float(la2),float(lo2))*float(config.travel_min_per_km)

    def _same_day_visits(worker,day,current=None):
        frame = assigned[day] if current is None else current
        if not len(frame):
            return []
        lookup = config.dispatch_cache['days'][id(days[day])]['lookup']
        ids = frame['任務ID'].to_numpy()[frame['派單居服員'].astype(str).to_numpy()==str(worker)]
        return [{'start':str(lookup[t]['時間窗_開始'])[:5],'end':str(lookup[t]['時間窗_結束'])[:5],
                 'lat':lookup[t].get('服務地點_緯度'),'lon':lookup[t].get('服務地點_經度'),'client':str(lookup[t]['案家ID'])}
                for t in ids if t in lookup]

    def substitute_order(task,workers,current=None):
        """回傳 (依代班順序排好的居服員, {居服員: 排序資料})；只排先後，可不可派仍由 check_visit 決定。"""
        client = str(task['案家ID'])
        summary = _sub_month(config.work_day_minutes,str(task['日期'])[:7],config)
        profiles_ = []
        for worker in workers:
            worker = str(worker)
            if worker not in profiles:
                continue
            familiar = served[(client,worker)] or int(worker in seed_backups.get(client,()) or worker in backups.get(client,()))
            profiles_.append(_sub_profile(worker,task,profiles[worker],config,others=_same_day_visits(worker,task['日期'],current),
                familiar=familiar,ledger=config.work_day_minutes,summary=summary,travel_fn=calc_travel_minutes,estimate_fn=_estimate))
        ordered = _sub_order(profiles_)
        return [p['worker'] for p in ordered],{p['worker']:p for p in ordered}

    def sub_tier(info):
        # 代班順序的前兩層（月加班近上限者最後、服務過此案家者優先）視為連續性政策，轉場只在同一層內比較。
        return bool(info['skipped']),not (bool(info['familiar']) and not info['demoted'])

    def pick_substitute(task,candidates,rank_info,reasons,before=None,after=None):
        """off：代班順序第一位可行者。weighted：同一層可行者中新增轉場最少者（平手維持原順序）。"""
        day = task['日期']
        best,tier = None,None
        for worker in candidates:
            worker = str(worker)
            if tier is not None and sub_tier(rank_info[worker])!=tier:
                break
            if before:
                before(worker)
            checked = check_visit(task,worker,assigned[day],config)
            if after:
                after(worker)
            if not checked['available']:
                reasons[worker] = checked['detail']
                continue
            if travel_mode!='weighted':
                return worker,checked
            frame = pd.concat([assigned[day],pd.DataFrame([{'任務ID':task['任務ID'],'派單居服員':worker}])],ignore_index=True)
            score = objective(1,added_cost(worker,{day:frame}))
            if best is None or score>best[2]:
                best = worker,checked,score
            tier = sub_tier(rank_info[worker])
        return best[:2] if best else None

    for i,task in enumerate(leave_tasks):
        if progress:
            progress(.85+.1*i/max(len(leave_tasks),1),f'核對請假代班 {i+1}/{len(leave_tasks)}')
        scored = run_phase1_matching(pd.DataFrame([task]),caregivers,config)
        reasons = {}
        # 請假代班不再依適配度分數（含滿意度與住家距離）排序，改用代班挑選順序。
        candidates,rank_info = substitute_order(task,scored['居服員ID'].astype(str).tolist() if not scored.empty else [])
        picked = pick_substitute(task,candidates,rank_info,reasons)
        if picked:
            worker,checked = picked
            rankings.setdefault(str(task['案家ID']),{}).update(scored.set_index('居服員ID').to_dict('index'))
            commit(task,worker,checked,leave_overlap(task,profiles.get(owners[str(task['案家ID'])])),
                   rank_reason=rank_info[worker]['reason'])
        else:
            note(task,'請假代班','固定人選請假，沒有通過資格、交通與工時檢查的代班人選。',reasons,len(scored))

    # ---- 彈性派單：固定備援 ＋ 時段彈性 ---------------------------------------------
    # 順序：①請假代班（上方已完成）→ ②同案固定備援承接 → ③挪動開始時間（服務分鐘不變）。
    # 備援與挪時段都只處理「前面步驟仍排不進」的服務，不會搶走已排好的班。
    backup_services,time_shifts = 0,[]
    shift_fail = Counter()
    rescued = set()
    extras_on = allow_backup and getattr(config,'enable_backup_caregiver',False)
    max_extra = max(1,int(getattr(config,'max_backup_caregivers',1))) if extras_on else 0
    done_ids = {tid for frame in assigned.values() for tid in frame['任務ID']}
    leave_ids = {str(t['任務ID']) for t in leave_tasks}

    def open_visits():
        found = {}
        for d in diagnostics:
            tid = d['任務ID']
            if d.get('問題分類')=='必要證照' and not getattr(config,'rescue_missing_cert_with_backup',False):
                continue  # 缺特殊訓練（如失智症照顧專長）需人工重新派單，不由備援自動接手
            if d['失敗階段'] in BACKUP_STAGES and tid not in done_ids and tid not in rescued and str(tid) not in leave_ids:
                found.setdefault(str(d['案家ID']),set()).add(tid)
        return found

    def ranked_workers(client,owner,site,limit,task=None):
        if task is not None:
            # 固定備援、時段調整與臨時支援：同一套代班挑選順序（以第一筆待補服務為代表）。
            ordered,_ = substitute_order(task,[w for w in profiles if w!=owner])
            seeded = [w for w in ordered if w in seed_backups.get(client,())]
            others = [w for w in ordered if w not in seeded]
            keep = set(seeded + (others[:limit] if limit else others))
            return [w for w in ordered if w in keep]
        def distance(worker):
            home = profiles[worker].get('服務起點_緯度(家)'),profiles[worker].get('服務起點_經度(家)')
            try:
                return (float(home[0])-float(site[0]))**2+((float(home[1])-float(site[1]))*0.91)**2
            except (TypeError,ValueError):
                return float('inf')
        seeded = [w for w in seed_backups.get(client,()) if w!=owner]
        # 大眾運輸居服員每次檢查都要查 Google，排在機車之後；機車能接滿就不必查。
        others = sorted((w for w in profiles if w!=owner and w not in seeded),
                        key=lambda w:(profiles[w].get('常用交通工具')=='大眾運輸',distance(w),w))
        return seeded + (others[:limit] if limit else others)

    def try_backup(client,visits,worker):
        # worker_workdays 必須給 trial 一份獨立複本（而非與主 config 共用參照）並隨著本次試算逐筆
        # 累計：visits 按日期排序試算，同一次試算裡前面幾天「暫定」會接受，必須讓後面天數的連續
        # 上班檢查看得到，否則同一個候選人一次試算涵蓋多天時，連續天數會被低估而漏放第7天。
        trial = replace(config,client_backup_map={**backups,client:backups.get(client,())+(worker,)},
                        work_day_minutes=dict(config.work_day_minutes),
                        worker_workdays={k:set(v) for k,v in config.worker_workdays.items()})
        tentative,accepted = {},[]
        for task in visits:
            if leave_overlap(task,profiles[worker]) or hard_check(task,worker,trial):
                continue
            day = task['日期']
            current = tentative.get(day,assigned[day])
            checked = check_visit(task,worker,current,trial)
            if not checked['available']:
                continue
            accepted.append((task,checked))
            tentative[day] = pd.concat([current,pd.DataFrame([{'任務ID':task['任務ID'],'派單居服員':worker}])],ignore_index=True)
            trial.work_day_minutes[(worker,day)] = checked['working_minutes']
            trial.worker_workdays.setdefault(worker,set()).add(str(day)[:10])
        return accepted

    if extras_on:
        for client,tids in sorted(open_visits().items(),key=lambda kv:(-len(kv[1]),kv[0])):
            owner = owners.get(client)
            if not owner:
                continue
            visits = sorted((task_records[t] for t in tids),key=lambda t:(t['日期'],t['時間窗_開始'],t['任務ID']))
            site = visits[0].get('服務地點_緯度'),visits[0].get('服務地點_經度')
            ranked = ranked_workers(client,owner,site,int(getattr(config,'backup_candidate_limit',0)) or None,task=visits[0])
            while len(backups.get(client,())) < max_extra:
                visits = [t for t in visits if t['任務ID'] not in rescued]
                if not visits:
                    break
                best = None
                for worker in (w for w in ranked if w not in backups.get(client,())):
                    accepted = try_backup(client,visits,worker)
                    if travel_mode=='weighted':
                        # 承接筆數優先，其次歷史備援，再比整期新增轉場；不在第一位排滿的人停下。
                        if accepted:
                            frames = {}
                            for task,_ in accepted:
                                day = task['日期']
                                frames[day] = pd.concat([frames.get(day,assigned[day]),
                                    pd.DataFrame([{'任務ID':task['任務ID'],'派單居服員':worker}])],ignore_index=True)
                            score = (objective(len(accepted),0.0),worker in seed_backups.get(client,()),
                                     -travel_weight*added_cost(worker,frames))
                            if best is None or score>best[2]:
                                best = worker,accepted,score
                        continue
                    if accepted and (best is None or len(accepted)>len(best[1])):
                        best = worker,accepted,None
                    if len(accepted)==len(visits):
                        break
                if not best:
                    break
                worker,accepted,_ = best
                backups[client] = backups.get(client,())+(worker,)
                for task,checked in accepted:
                    _,_info = substitute_order(task,[worker])
                    commit(task,worker,checked,source='備援居服員',rank_reason=_info.get(worker,{}).get('reason',''))
                    rescued.add(task['任務ID'])
                backup_services += len(accepted)

    # ---- 時段彈性：只挪動開始時間（服務分鐘不變），同案同一模式每週套用相同位移 ----
    flex_default = max(0.0,finite(getattr(config,'time_flex_minutes',0.0),0.0))
    flex_step = max(5,int(getattr(config,'time_flex_step_minutes',15)))
    task_flex = lambda rec:(lambda v:v if v==v and v>=0 else flex_default)(finite(rec.get('時間彈性(分)')))
    if allow_backup and any(task_flex(r)>=flex_step for r in task_records.values()):
        to_min = lambda text:int(text[:2])*60+int(text[3:5])
        to_hhmm = lambda minutes:f'{minutes//60:02d}:{minutes%60:02d}'

        def set_times(tid,start,end):
            rec = task_records[tid]
            rec['時間窗_開始'],rec['時間窗_結束'] = start,end
            task_patterns[tid] = visit_pattern(rec)
            frame = days[rec['日期']]
            frame.loc[frame['任務ID'].eq(tid),['時間窗_開始','時間窗_結束']] = [start,end]
            entry = config.dispatch_cache['days'][id(frame)]['lookup'].get(tid)
            if entry is not None:
                entry['時間窗_開始'],entry['時間窗_結束'] = start,end

        def clear_of_overlap(worker,day,start,end):
            return not any(start < other['時間窗_結束'] and other['時間窗_開始'] < end
                           for other in occupied_slots.get((worker,day),()))

        def try_shift(rec,start,end,worker,trial):
            tid = rec['任務ID']
            old = rec['時間窗_開始'],rec['時間窗_結束']
            set_times(tid,start,end)
            try:
                if leave_overlap(rec,profiles[worker]):
                    shift_fail['請假']+=1
                    return None
                hard = hard_check(rec,worker,trial)
                if hard:
                    shift_fail['硬性條件：'+failure_kind(hard)[0]]+=1
                    return None
                checked = check_visit(rec,worker,assigned[rec['日期']],trial)
                if not checked['available']:
                    shift_fail[failure_kind(checked['detail'])[0]]+=1
                return checked if checked['available'] else None
            finally:
                set_times(tid,*old)

        shift_limit = int(getattr(config,'time_flex_candidate_limit',8)) or None
        time_dependent = ('時段','請假','服務時間')
        static_block = {}

        def worker_static_ok(rec,worker,trial):
            # 休息日、資格日期、證照等與「幾點開始」無關的限制：原時段不通過就不必逐一試挪動。
            key = pattern_of(rec),rec['日期'],worker
            if key not in static_block:
                reason = hard_check(rec,worker,trial)
                static_block[key] = bool(reason) and not any(word in reason for word in time_dependent)
            return not static_block[key]

        def run_groups(groups,allow_new_team=True):
            for _,recs in sorted(groups.items(),key=lambda kv:(-len(kv[1]),str(kv[0]))):
                recs = sorted(recs,key=lambda t:(t['日期'],t['任務ID']))
                client = str(recs[0]['案家ID'])
                owner = owners.get(client)
                flex = task_flex(recs[0])
                if not owner or flex < flex_step:
                    continue
                shifts = [sign*k for k in range(flex_step,int(flex)+1,flex_step) for sign in (1,-1)]
                site = recs[0].get('服務地點_緯度'),recs[0].get('服務地點_經度')
                for _ in range(3):
                    recs = [r for r in recs if r['任務ID'] not in rescued]
                    if not recs:
                        break
                    current_team = [owner]+list(backups.get(client,()))
                    room = max_extra-len(backups.get(client,())) if extras_on else 0
                    pool = current_team + ([w for w in ranked_workers(client,owner,site,shift_limit,task=recs[0]) if w not in current_team] if room>0 else [])
                    best = None
                    for shift in shifts:
                        for worker in pool:
                            # 同一位候選人要接下 recs 裡該案的多天時段，worker_workdays 同樣要給獨立複本並
                            # 逐筆累計（理由見上方 try_backup 的註解），否則連續上班天數會被低估。
                            trial = replace(config,client_backup_map={**backups,client:backups.get(client,())+(() if worker in current_team else (worker,))},
                                            work_day_minutes=dict(config.work_day_minutes),
                                            worker_workdays={k:set(v) for k,v in config.worker_workdays.items()})
                            accepted = []
                            for rec in recs:
                                begin,finish = to_min(rec['時間窗_開始'])+shift,to_min(rec['時間窗_結束'])+shift
                                if begin<0 or finish>1439:
                                    continue
                                start,end = to_hhmm(begin),to_hhmm(finish)
                                if not clear_of_overlap(worker,rec['日期'],start,end):
                                    shift_fail['時段重疊（預篩）']+=1
                                    continue
                                if not worker_static_ok(rec,worker,trial):
                                    shift_fail['休息日／資格（預篩）']+=1
                                    continue
                                checked = try_shift(rec,start,end,worker,trial)
                                if checked:
                                    accepted.append((rec,start,end,checked))
                                    trial.work_day_minutes[(worker,rec['日期'])] = checked['working_minutes']
                                    trial.worker_workdays.setdefault(worker,set()).add(str(rec['日期'])[:10])
                            if accepted and (best is None or len(accepted)>len(best[2])):
                                best = shift,worker,accepted
                            if len(accepted)==len(recs):
                                break
                        if best and len(best[2])==len(recs):
                            break
                    if not best:
                        break
                    shift,worker,accepted = best
                    if worker not in current_team:
                        backups[client] = backups.get(client,())+(worker,)
                    for rec,start,end,checked in accepted:
                        before = f"{rec['時間窗_開始']}-{rec['時間窗_結束']}"
                        set_times(rec['任務ID'],start,end)
                        commit(rec,worker,checked,source='時段調整')
                        rescued.add(rec['任務ID'])
                        time_shifts.append({'任務ID':rec['任務ID'],'案家ID':client,'日期':rec['日期'],'星期':rec.get('星期',''),
                            '來源列':rec.get('來源列'),'原時段':before,'新時段':f'{start}-{end}','位移(分)':shift,
                            '派單居服員':worker,'調整類型':'每週固定位移' if len(groups_kind)==1 else '單次調整','新開始':start,'新結束':end})

        groups = {}
        for client,tids in open_visits().items():
            for tid in tids:
                groups.setdefault(pattern_of(task_records[tid]),[]).append(task_records[tid])
        groups_kind = {'pattern'}
        run_groups(groups)
        # 同模式無法用同一位移／同一人承接時，才退而求其次逐次挪動；結果標示「單次調整」供居督確認。
        if getattr(config,'time_flex_allow_single_visit',True):
            groups = {}
            for client,tids in open_visits().items():
                for tid in tids:
                    groups[(pattern_of(task_records[tid]),task_records[tid]['日期'])] = [task_records[tid]]
            groups_kind = {'pattern','single'}
            run_groups(groups)

    # ---- 臨時支援：主責＋固定備援（含挪時段）都排不進的服務，逐次另選合格居服員 ----
    # 只記在該次服務（task_support_map），不加入固定備援，所以每案固定人選仍是「1 主責＋最多 N 備援」。
    # 同案已支援過的人優先，其次依距離；一樣要通過資格、請假、交通、工時與月加班檢查。
    support_services = 0
    if allow_backup and getattr(config,'enable_support_caregiver',False):
        limit = int(getattr(config,'backup_candidate_limit',0)) or None
        for client,tids in sorted(open_visits().items(),key=lambda kv:(-len(kv[1]),kv[0])):
            owner = owners.get(client)
            if not owner:
                continue
            team = {owner,*backups.get(client,())}
            visits = sorted((task_records[t] for t in tids),key=lambda t:(t['日期'],t['時間窗_開始'],t['任務ID']))
            site = visits[0].get('服務地點_緯度'),visits[0].get('服務地點_經度')
            pool = [w for w in ranked_workers(client,owner,site,limit,task=visits[0]) if w not in team]
            helpers = []
            for task in visits:
                tid = str(task['任務ID'])
                # 每一筆依當天實際的前一案重新排序；已支援過本案的人因「服務過」自然排在前面。
                ordered,rank_info = substitute_order(task,helpers+[w for w in pool if w not in helpers])
                picked = pick_substitute(task,ordered,rank_info,{},before=lambda w:supports.__setitem__(tid,w),
                                         after=lambda w:supports.pop(tid,None))
                if picked:
                    worker,checked = picked
                    supports[tid] = worker
                    commit(task,worker,checked,source='臨時支援',rank_reason=rank_info[worker]['reason'])
                    rescued.add(task['任務ID'])
                    support_services += 1
                    if worker not in helpers:
                        helpers.append(worker)

    # ---- 第二階段（two_stage）：降低轉場的局部搬移與交換 ----
    # 只動非主責服務（請假代班、固定備援、臨時支援）；主責服務一筆都不動，派出數也不變。
    # 每一步都重新通過 check_visit（資格、請假、時段、交通空檔、休息、日／月工時、連續上班），
    # 且精確的「居服員×日」付費轉場（不含住家通勤）至少少 travel_stage2_min_gain_minutes 分才接受，所以任何時刻的班表都是可行解；
    # 逾時就停在目前已驗證的解（不會比第一階段差）。固定備援服務只在該案既有備援之間調動，
    # 不新增備援人數；臨時支援可換成主責＋備援以外的人；請假代班可換成主責以外的人。
    stage2 = {'moves':0,'swaps':0,'estimated_saved_minutes':0.0,'seconds':0.0,'timed_out':False}
    if allow_backup and travel_mode=='two_stage' and rows:
        started = time.perf_counter()
        deadline = started+max(0.0,finite(getattr(config,'travel_stage2_seconds',0.0),0.0))
        row_at = {r['任務ID']:i for i,r in enumerate(rows)}
        match_at = {m['任務ID']:i for i,m in enumerate(matches)}
        movable_sources = ('請假代班','備援居服員','臨時支援')
        # 只看付費轉場、且至少省 travel_stage2_min_gain_minutes 分才換人，避免為了一兩分鐘通勤就改派。
        gain_floor = max(0.01,finite(getattr(config,'travel_stage2_min_gain_minutes',5.0),5.0))
        paid = lambda worker,day,tids:route_cost(worker,day,tids,paid_only=True)
        first_worker,saved_by_task = {},{}   # 任務ID → 第一階段原派居服員／累計省下的估計分鐘

        def movable():
            return [r for r in rows if r['排班來源'] in movable_sources]

        def allowed(row):
            client,now = str(row['案家ID']),str(row['派單居服員'])
            owner = owners.get(client)
            if row['排班來源']=='備援居服員':
                pool = backups.get(client,())
            elif row['排班來源']=='臨時支援':
                pool = [w for w in profiles if w!=owner and w not in backups.get(client,())]
            else:
                pool = [w for w in profiles if w!=owner]
            return [w for w in pool if w!=now]

        def check_as(tid,worker,frame):
            support = rows[row_at[tid]]['排班來源']=='臨時支援'
            previous = supports.get(str(tid))
            if support:
                supports[str(tid)] = worker
            try:
                return check_visit(task_records[tid],worker,frame,config)
            finally:
                if support:
                    supports[str(tid)] = previous

        def day_minutes(worker,day,frame):
            ids = worker_ids(frame,worker)
            if not ids:
                return 0.0
            checked = check_as(ids[0],worker,frame[frame['任務ID'].ne(ids[0])])
            return checked['working_minutes'] if checked['available'] else None

        def keeps_team(changes):
            # 不增加案家的搭配人數：接手者已服務過該案，或原服務者只剩這一筆（換人後人數不變）。
            for tid,old,new in changes:
                client = str(task_records[tid]['案家ID'])
                if not served[(client,new)] and served[(client,old)] > 1:
                    return False
            return True

        def overtime(day,minutes):
            return sum(max(0.0,m/60-config.normal_daily_hours) for m in minutes.values())

        def no_extra_overtime(day,minutes):
            # 轉場省下的分鐘若換來加班（1.34 倍工資）就不划算：相關居服員當天加班不得增加。
            before = {w:config.work_day_minutes.get((w,day),0.0) for w in minutes}
            return overtime(day,minutes) <= overtime(day,before)+1e-9

        def with_rows(frame,pairs):
            return pd.concat([frame,pd.DataFrame([{'任務ID':t,'派單居服員':w} for t,w in pairs])],ignore_index=True)

        def apply(changes,day,frame,minutes,gain):
            assigned[day] = frame
            for worker,value in minutes.items():
                set_work_minutes(worker,day,value)
                if not value:
                    config.work_day_minutes.pop((worker,day),None)
            for tid,old,new in changes:
                task = task_records[tid]
                if not worker_ids(frame,old):
                    config.worker_workdays.get(old,set()).discard(str(day)[:10])
                config.worker_workdays.setdefault(new,set()).add(str(day)[:10])
                occupied_slots[(old,day)] = [t for t in occupied_slots.get((old,day),[]) if t['任務ID']!=tid]
                occupied_slots.setdefault((new,day),[]).append(task)
                client = str(task['案家ID'])
                served[(client,old)] -= 1
                served[(client,new)] += 1
                row = rows[row_at[tid]]
                row['派單居服員'] = new
                # 原本的排序理由描述的是第一階段的人選，改派後不再適用；改寫成改派說明（多次改派只留一句）。
                first_worker.setdefault(tid,old)
                saved_by_task[tid] = saved_by_task.get(tid,0.0)+gain
                row['代班排序理由'] = (f'轉場最佳化改派：原派 {first_worker[tid]}，改派後估計轉場少 {saved_by_task[tid]:.0f} 分'
                                   if first_worker[tid]!=new else '轉場最佳化後維持原派')
                matches[match_at[tid]] = _match_row(task,new,profiles,rankings,config,historical)
                if row['排班來源']=='臨時支援':
                    supports[str(tid)] = new

        def relocate(row):
            tid,a,day = row['任務ID'],str(row['派單居服員']),row['排班日期']
            frame = assigned[day]
            without = frame[frame['任務ID'].ne(tid)]
            ids_a = worker_ids(frame,a)
            left = paid(a,day,[t for t in ids_a if t!=tid])-paid(a,day,ids_a)
            for b in allowed(row):
                ids_b = worker_ids(frame,b)
                gain = -(left+paid(b,day,ids_b+[tid])-paid(b,day,ids_b))
                if gain < gain_floor or not keeps_team([(tid,a,b)]):
                    continue
                checked = check_as(tid,b,without)
                if not checked['available']:
                    continue
                final = with_rows(without,[(tid,b)])
                minutes_a = day_minutes(a,day,final)
                if minutes_a is None:
                    continue
                minutes = {b:checked['working_minutes'],a:minutes_a}
                if not no_extra_overtime(day,minutes):
                    continue
                apply([(tid,a,b)],day,final,minutes,gain)
                return gain
            return 0.0

        def swap(r1,r2):
            t1,t2,day = r1['任務ID'],r2['任務ID'],r1['排班日期']
            a,b = str(r1['派單居服員']),str(r2['派單居服員'])
            if a==b or b not in allowed(r1) or a not in allowed(r2):
                return 0.0
            frame = assigned[day]
            ids_a,ids_b = worker_ids(frame,a),worker_ids(frame,b)
            gain = (paid(a,day,ids_a)+paid(b,day,ids_b)
                    -paid(a,day,[t for t in ids_a if t!=t1]+[t2])-paid(b,day,[t for t in ids_b if t!=t2]+[t1]))
            if gain < gain_floor or not keeps_team([(t1,a,b),(t2,b,a)]):
                return 0.0
            base = frame[~frame['任務ID'].isin([t1,t2])]
            first = check_as(t1,b,base)
            if not first['available']:
                return 0.0
            middle = with_rows(base,[(t1,b)])
            second = check_as(t2,a,middle)
            if not second['available']:
                return 0.0
            minutes = {b:first['working_minutes'],a:second['working_minutes']}
            if not no_extra_overtime(day,minutes):
                return 0.0
            apply([(t1,a,b),(t2,b,a)],day,with_rows(middle,[(t2,a)]),minutes,gain)
            return gain

        def out_of_time():
            stage2['timed_out'] = time.perf_counter() > deadline
            return stage2['timed_out']

        if progress:
            progress(.95,'轉場最佳化（第二階段）')
        improved = True
        while improved and not out_of_time():
            improved = False
            for row in movable():
                if out_of_time():
                    break
                gain = relocate(row)
                if gain:
                    stage2['moves'] += 1
                    stage2['estimated_saved_minutes'] += gain
                    improved = True
            by_day = {}
            for row in movable():
                by_day.setdefault(row['排班日期'],[]).append(row)
            for day,group in sorted(by_day.items()):
                for i in range(len(group)):
                    for j in range(i+1,len(group)):
                        if out_of_time():
                            break
                        gain = swap(group[i],group[j])
                        if gain:
                            stage2['swaps'] += 1
                            stage2['estimated_saved_minutes'] += gain
                            improved = True
        stage2['seconds'] = round(time.perf_counter()-started,2)
        stage2['estimated_saved_minutes'] = round(stage2['estimated_saved_minutes'],1)
    diagnostics[:] = [d for d in diagnostics if d['任務ID'] not in rescued]
    if time_shifts:
        moved = {t['任務ID']:t for t in time_shifts}
        mask = tasks['任務ID'].isin(moved)
        tasks.loc[mask,'時間窗_開始'] = tasks.loc[mask,'任務ID'].map(lambda t:moved[t]['新開始'])
        tasks.loc[mask,'時間窗_結束'] = tasks.loc[mask,'任務ID'].map(lambda t:moved[t]['新結束'])

    result = pd.DataFrame(rows,columns=RESULT_COLUMNS).sort_values(['排班日期','服務時段','任務ID']).reset_index(drop=True)
    match_frame = pd.DataFrame(matches)
    for column in MATCH_COLUMNS:
        if column not in match_frame:
            match_frame[column] = pd.Series(dtype=object)
    match_frame.attrs['matching_diagnostics'] = diagnostics
    config.dispatch_cache = {}
    if progress:
        progress(.96,'計算最終交通路段')
    routes = calculate_schedule_travel(result,tasks,caregivers,config) if compute_routes and not result.empty else result.copy()
    if progress:
        progress(1.0,'全期間排班完成')
    return {'df_result':result,'df_matches':match_frame,'route_result':routes,
        'client_primary_map':dict(owners),'client_backup_map':dict(backups),'task_support_map':dict(supports),
        'time_shifts':pd.DataFrame(time_shifts),'tasks_adjusted':tasks,'status':pywraplp.Solver.FEASIBLE,'assigned_count':len(result),
        'df_valid':pd.DataFrame(),'pending_details':{d['任務ID']:d['逐人排班原因'] for d in diagnostics},
        'daily_counts':{day:len(frame) for day,frame in assigned.items()},'diagnostics':pd.DataFrame(diagnostics),
        '_day_templates':templates,
        'dispatch_stats':{'historical_clients':len(historical),'new_clients':total_new,'ownership_search_clients':total_new,
                          'copied_worker_days':copied_days,'checked_worker_days':checked_days,
                          'backup_clients':len(backups),'backup_services':backup_services,'support_services':support_services,
                          'time_shifted_services':len(time_shifts),'time_shift_fail_reasons':dict(shift_fail),
                          'travel_objective_mode':travel_mode,'travel_stage2':stage2}}


def _match_row(task,worker,profiles,rankings,config,historical):
    ranked = rankings.get(str(task['案家ID']),{}).get(worker)
    profile = profiles[worker]
    years,bonus = seniority_score(profile,task['日期'],config)
    preferred = historical.get(str(task['案家ID'])) == worker
    continuity = config.preferred_caregiver_bonus if preferred else 0.0
    skill = config.cert_bonus_dementia if dementia_status(task) is True and caregiver_has_cert(profile,'失智症照顧專長') else 0.0
    score = (ranked['適配度分數']-ranked.get('年資加分',0)+bonus) if ranked else config.base_score+continuity+bonus+skill
    return {**(ranked or {}),'任務ID':task['任務ID'],'案家ID':task['案家ID'],'居服員ID':worker,
        '適配度分數':round(score,2),'歷史首選加分':continuity,'年資年數':years,'年資加分':bonus,
        '預估交通時間(分)':float('nan'),'地點緯度':task.get('服務地點_緯度'),'地點經度':task.get('服務地點_經度'),
        '任務開始時間':task['時間窗_開始'],'任務結束時間':task['時間窗_結束'],'優先級':task.get('任務優先級'),
        '是否歷史首選':preferred,'具備失智症20小時認證(0/1)':int(caregiver_has_cert(profile,'失智症照顧專長')),
        '常用交通工具':profile.get('常用交通工具')}
