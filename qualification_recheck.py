"""個案照顧需求變更的資格重新檢核：案家新增「失智」或新增「BA08 足部照護」。

原本由 A 長期服務的案家，照專複評後新增需求（確認失智，或核定 BA08 足部照護），
A 沒有對應訓練（失智症 20 小時／足部照護 10 小時）時：

法規原則（v20 起）：服務只能派給「具備該案所需訓練資格」的居服員；派案當下不符資格的人，
系統不派、不留「寬限期內繼續服務」。因此三種處理方式的「服務提供者」一律是有訓練的居服員，
差別只在怎麼降低換人衝擊、以及原居服員之後能不能回來：
  now      立即改派（預設）：生效日起全部改由有訓練的居服員；一位接不完就依星期分給固定幾位。
  handover 熟面孔陪同交接：同樣從第一次起由有訓練的居服員服務；前幾次原居服員以「陪同」身分到場
           介紹、協助長輩認識新面孔（不是服務提供者、不申報服務、不計服務時數；需機構同意）。
  grace    補訓後回任：同樣立即改派；另追蹤原居服員在期限內完成補訓，取得資格並登錄後，
           居督可評估把主責改回原居服員（恢復照顧連續性）。補訓期間原居服員不得服務此案。
BA08 另有一種常見作法：原居服員繼續原本的 BA02／BA20 等服務，BA08 另排一位有訓練者單獨服務
（週服務計畫另加一列），這屬於「新增服務」流程，不在本模組。

這裡只放純計算（可單獨測試）；畫面與寫入班表在 app.py。
"""
CAREFLOW_BUILD = '2026.10.18-v19'
from datetime import date, timedelta

import pandas as pd

import caregiver_engine as engine

CERT = '失智症照顧專長'
# 需求種類：(畫面名稱, 需要的資格, 補訓課程, 降低換人衝擊的說明)
NEEDS = {
    'dementia': {'label': '新增失智', 'cert': '失智症照顧專長', 'course': '失智症照顧服務 20 小時訓練',
                 'why': '失智長輩對熟悉的照顧者比較穩定，換生面孔容易拒絕服務或出現情緒行為，交接時可請原居服員陪同。',
                 'excel': '把機構 Excel「個案資料」的是否失智改為「是」'},
    'foot': {'label': '新增足部照護（BA08）', 'cert': 'BA08足部照護', 'course': 'BA08 足部照護 10 小時訓練',
             'why': '長期服務的居服員熟悉長輩與家屬；也可以原居服員續做原本服務、BA08 另排有訓練的人單獨服務。',
             'excel': '在機構 Excel「週服務計畫」這個案家的服務項目加上 BA08'},
}
METHODS = ('now', 'handover', 'grace')
DEFAULT_METHOD = 'now'
RULE_TEXT = '法規規定：只能派給具備此個案所需訓練資格的居服員；派案當下不符資格者，系統不會派給他。'
METHOD_SHORT = {'now': '立即改派', 'handover': '熟面孔陪同交接', 'grace': '補訓後回任'}


def method_label(method, need='dementia'):
    course = NEEDS[need]['course']
    return {'now': '① 立即改派（法規規定，預設：生效日起全部由有訓練的居服員服務）',
            'handover': '② 熟面孔陪同交接（有訓練的居服員服務，前幾次原居服員陪同、不提供服務）',
            'grace': f'③ 補訓後回任（立即改派，原居服員期限內完成{course}後可評估改回主責）'}[method]


def method_help(method, need='dementia'):
    info = NEEDS[need]
    return {'now': '單位與縣市都要求服務人員須具備資格，生效日起不符資格的人不能再服務此案。'
                   '一位接不完時，依星期分給固定幾位，讓長輩每週同一天看到同一個人。',
            'handover': info['why'] + '陪同不另計一筆服務、不申報，原居服員只做介紹與情緒安撫，'
                        '照顧服務（含失智照顧、足部照護）由有訓練的居服員執行；陪同的交通與工時請機構另行認定。',
            'grace': '服務仍立即改由有訓練的人；另把原居服員的補訓期限記入追蹤清單。完成訓練並在員工資料登錄「取得資格」後，'
                     '可到固定人選把主責改回原居服員。補訓期間原居服員不得服務此案，也不列為此案備援。'}[method]


# 舊版相容（只有失智時的文字）
METHOD_LABEL = {m: method_label(m) for m in METHODS}
METHOD_HELP = {m: method_help(m) for m in METHODS}
GRACE_DAYS = 60          # 原居服員補訓期限預設天數（補訓後回任用；不是可繼續服務的寬限期）
HANDOVER_VISITS = 2
EMPTY = ('', '未指派', 'None', 'nan', 'NaT')


def _text(value):
    if value is None:
        return ''
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _day(value):
    stamp = pd.to_datetime(value, errors='coerce')
    return '' if pd.isna(stamp) else stamp.strftime('%Y-%m-%d')


def certified(df_cg, need='dementia'):
    """有這項需求所需訓練的居服員ID。"""
    cert = NEEDS[need]['cert']
    return {_text(r.get('居服員ID')) for r in df_cg.to_dict('records') if engine.caregiver_has_cert(r, cert)}


def affected_tasks(client, tasks, current_map, df_cg, since=None, need='dementia'):
    """這個案家從 since 起、目前由「沒有這項訓練」的人服務的服務，依日期排序。"""
    ok = certified(df_cg, need)
    since = _day(since) if since else ''
    rows = tasks.loc[tasks['案家ID'].astype(str) == str(client)]
    out = []
    for r in rows.to_dict('records'):
        tid, day = _text(r.get('任務ID')), _day(r.get('日期'))
        cg = _text(current_map.get(tid))
        if not day or (since and day < since) or cg in EMPTY or cg in ok:
            continue
        out.append({'任務ID': tid, '日期': day, '開始': _text(r.get('時間窗_開始'))[:5],
                    '結束': _text(r.get('時間窗_結束'))[:5], '居服員': cg})
    return sorted(out, key=lambda x: (x['日期'], x['開始']))


def split(method, affected, deadline=None, handover=HANDOVER_VISITS):
    """依處理方式把受影響服務分成：
    move（改由有訓練的居服員服務）：三種方式都是全部，法規不允許不符資格者繼續服務；
    escort（原居服員陪同、不提供服務）：只有熟面孔陪同交接的前 handover 次；
    keep／accompany：舊版欄位，固定為空（保留給舊程式與存檔相容）。"""
    ids = [a['任務ID'] for a in affected]
    escort = ids[:max(0, int(handover))] if method == 'handover' else []
    return {'keep': [], 'accompany': [], 'escort': escort, 'move': ids}


def availability(task_ids, candidate_ids, tasks, effective, df_cg, config, df_cl=None, task_locations=None):
    """每筆服務、每位有證照的候選人能不能接（與改派同一套檢查）；固定人選限制暫時開放，
    因為原居服員資格不符本來就要找人。回傳 {任務ID: [(居服員ID, 理由), ...]}（依代班順序）。"""
    out = {}
    for tid in task_ids:
        with engine._OpenSupport(config, tid):
            try:
                ranked = engine.rank_candidates_by_availability(
                    tid, list(candidate_ids), tasks, effective, df_cg, config,
                    task_locations=task_locations, date_column='日期', df_cl=df_cl)
            except Exception:  # noqa: BLE001
                ranked = []
        out[tid] = [(_text(r['cg_id']), _text(r.get('rank_reason'))) for r in ranked if r.get('available')]
    return out


def covering(avail, need_ids, accompany_ids=()):
    """能接下全部 need_ids（及陪同 accompany_ids）的人，依各筆代班順位的平均名次排序。
    回傳 [{'cg', 'reason', 'avg_rank'}]；need_ids 與 accompany_ids 都空時回傳 []。"""
    need = list(dict.fromkeys([*need_ids, *accompany_ids]))
    if not need:
        return []
    sets = [{cg for cg, _ in avail.get(t, [])} for t in need]
    common = set.intersection(*sets) if sets else set()
    ranked = []
    for cg in common:
        positions = [next(i for i, (c, _) in enumerate(avail[t]) if c == cg) for t in need]
        reason = next((r for c, r in avail[need[0]] if c == cg), '')
        ranked.append({'cg': cg, 'reason': reason, 'avg_rank': sum(positions) / len(positions)})
    return sorted(ranked, key=lambda x: (x['avg_rank'], x['cg']))


def team_after(method, owner, backups, newcomer):
    """處理後的固定人選（主責, 備援）。三種方式都由有訓練的新居服員主責，原居服員移出此案團隊
    （不符資格不能當備援）；補訓後回任要等原居服員取得資格，再由居督手動改回。"""
    backups = [b for b in (backups or []) if b and b not in (owner, newcomer)]
    return newcomer, backups[:2]


def default_deadline(since):
    start = pd.to_datetime(since, errors='coerce')
    start = date.today() if pd.isna(start) else start.date()
    return start + timedelta(days=GRACE_DAYS)


def followup(method, client, owner, newcomer, parts, deadline=None, created='', need='dementia'):
    """追蹤清單的一筆（04 與臨時異動頁顯示，自動存檔一起保存）。"""
    info = NEEDS[need]
    if method == 'grace':
        todo = (f'已改由 {newcomer} 服務；{owner} 需在 {_day(deadline)} 前完成{info["course"]}，'
                '完成後更新員工資料「取得資格」，再評估是否改回主責')
    elif method == 'handover':
        todo = f'已改由 {newcomer} 服務；前 {len(parts.get("escort", []))} 次請 {owner} 陪同交接（不提供服務、不申報）'
    else:
        todo = f'已全部改由 {newcomer} 服務'
    return {'案家': client, '需求變更': info['label'], '原居服員': owner, '接手居服員': newcomer, '處理方式': METHOD_SHORT[method],
            '期限': _day(deadline) if method == 'grace' else '', '改由有訓練者服務': len(parts['move']),
            '原居服員陪同': len(parts.get('escort', [])), '待辦': todo,
            '提醒': '請立即同步' + info['excel'] + '，之後重新排班時系統才會持續擋下沒有訓練的人',
            '建立時間': created, '狀態': '追蹤中'}


def split_cover(avail, need_ids, accompany_ids=(), day_of=None):
    """沒有一位能全包時，依「星期」分給多位有訓練的人（同一個星期固定同一位，案家比較好適應）。
    先挑能包下最多個星期的人，剩下的星期再找下一位；某個星期找不到一位包下時，那個星期逐次找人。
    回傳 {'assign': {任務ID: 居服員ID}, 'groups': [{'星期','居服員','筆數','理由'}], 'missing': [任務ID]}。"""
    need = list(dict.fromkeys([*need_ids, *accompany_ids]))
    day_of = day_of or {}
    week = '一二三四五六日'

    def weekday(tid):
        stamp = pd.to_datetime(day_of.get(tid), errors='coerce')
        return '未定' if pd.isna(stamp) else '週' + week[stamp.weekday()]
    groups = {}
    for tid in need:
        groups.setdefault(weekday(tid), []).append(tid)
    able = {g: {c['cg']: c for c in covering(avail, ids)} for g, ids in groups.items()}
    assign, rows, left = {}, [], [g for g in groups if able[g]]
    while left:
        counts = {}
        for g in left:
            for cg, info in able[g].items():
                counts.setdefault(cg, []).append(info['avg_rank'])
        best = min(counts, key=lambda cg: (-len(counts[cg]), sum(counts[cg]) / len(counts[cg]), cg))
        for g in [g for g in left if best in able[g]]:
            for tid in groups[g]:
                assign[tid] = best
            rows.append({'星期': g, '居服員': best, '筆數': len(groups[g]), '理由': able[g][best]['reason']})
            left.remove(g)
    missing = []
    for g, ids in groups.items():
        if able[g]:
            continue
        for tid in ids:                 # 這個星期沒人能全包：逐次找
            options = avail.get(tid, [])
            if options:
                assign[tid] = options[0][0]
                rows.append({'星期': g + '（' + str(day_of.get(tid, ''))[:10] + '）', '居服員': options[0][0], '筆數': 1, '理由': options[0][1]})
            else:
                missing.append(tid)
    order = {w: i for i, w in enumerate(['週' + d for d in week])}
    rows.sort(key=lambda r: (order.get(r['星期'][:2], 9), r['星期']))
    return {'assign': assign, 'groups': rows, 'missing': missing}
