"""Shared service-hour and rest checks; durations come from occupied time slots."""

def workload_error(slots, cap_hours=8, occupied_minutes=0, continuous_limit=240, rest_minutes=30, transfers=None):
    slots=sorted(slots)
    durations=[(b-a).total_seconds()/60 for a,b in slots]
    if any(d<=0 for d in durations):return '服務結束時間必須晚於開始時間。'
    total=sum(durations)+max(0,float(occupied_minutes))
    cap=min(8.0,float(cap_hours))
    if total>cap*60+1e-7:return f'每日服務工時超標：加入後 {total/60:.2f} 小時，上限 {cap:g} 小時。'
    continuous=0
    for i,((start,end),duration) in enumerate(zip(slots,durations)):
        if i:
            gap=(start-slots[i-1][1]).total_seconds()/60
            if gap<0:return '服務時間重疊，不能同時接兩案。'
            transfer=(transfers or {}).get(i,0)
            if gap-transfer>=rest_minutes:continuous=0
        continuous+=duration
        if continuous>continuous_limit+1e-7:
            return f'休息不足：連續服務將達 {continuous:g} 分鐘，上限 {continuous_limit:g} 分鐘；需保留至少 {rest_minutes:g} 分鐘休息（不含轉場）。'
    return None


def incremental_travel(result, tasks, caregivers, config, cache, calculate):
    """Cache each caregiver/day route. Reuse travel fields, never stale assignment data."""
    import hashlib
    import pandas as pd
    if result.empty:
        cache.clear()
        return result.copy()
    task_map=tasks.drop_duplicates('任務ID').set_index('任務ID').to_dict('index')
    cg_map=caregivers.set_index('居服員ID').to_dict('index')
    groups={}
    for index,row in result.reset_index(drop=True).iterrows():
        t=task_map.get(row['任務ID'],{})
        day=str(t.get('日期',row.get('排班日期','單日')))
        groups.setdefault((str(row['派單居服員']),day),[]).append(index)
    frames=[]; active={}
    base=result.reset_index(drop=True)
    for key,indices in groups.items():
        group=base.iloc[indices].copy()
        ids=group['任務ID'].tolist()
        subset=tasks[tasks['任務ID'].isin(ids)].copy()
        payload=(group[['任務ID','派單居服員']].sort_values('任務ID').to_dict('records'),
                 subset.sort_values('任務ID').to_dict('records'),
                 cg_map.get(group.iloc[0]['派單居服員'],{}),vars(config))
        signature=hashlib.sha256(repr(payload).encode()).hexdigest()
        old=cache.get(key)
        if old and old[0]==signature:
            travel=old[1]
        else:
            computed=calculate(group,subset,caregivers,config)
            fields=[c for c in computed if c not in group or c in ('預估車程(分)','交通估算來源') or c.startswith(('交通','路段','含緩衝','轉場','配對起點'))]
            travel=computed[['任務ID']+[c for c in fields if c!='任務ID']].copy()
        active[key]=(signature,travel)
        group=group.drop(columns=[c for c in travel if c!='任務ID' and c in group])
        frames.append(group.merge(travel,on='任務ID',how='left',validate='one_to_one'))
    cache.clear();cache.update(active)
    return pd.concat(frames,ignore_index=True)
