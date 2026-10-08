"""Unified local dispatch follow-up and availability. No external messages are sent."""
import hashlib
import json
import sqlite3
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd

STATES = ['建議', '督導確認', '待回覆', '已接單', '已通知案家']

def stamp():
    return datetime.now(ZoneInfo('Asia/Taipei')).isoformat(timespec='seconds')

def task_signature(task, cg):
    values = [str(task.get(k, '')) for k in ['任務ID','案家ID','日期','時間窗_開始','時間窗_結束']]
    return hashlib.sha256(json.dumps(values + [str(cg or '')], ensure_ascii=False).encode()).hexdigest()

def overlap(task, interval):
    try:
        day = pd.Timestamp(task.get('日期')).date().isoformat()
        start = pd.Timestamp(day + ' ' + str(task['時間窗_開始'])[:5])
        end = pd.Timestamp(day + ' ' + str(task['時間窗_結束'])[:5])
        return start < pd.Timestamp(interval['end']) and end > pd.Timestamp(interval['start'])
    except (ValueError, TypeError, KeyError):
        return False

def blocked_reason(task, cg):
    blocks = cg.get('臨時不可排時段', [])
    if isinstance(blocks, str):
        try: blocks = json.loads(blocks)
        except ValueError: return '臨時不可排時段資料格式錯誤'
    if not isinstance(blocks, list): return None
    for item in blocks:
        if overlap(task, item):
            return f"臨時不可排：{item['start']}～{item['end']}（{item.get('reason', '請假')}）"
    return None

class WorkflowStore:
    def __init__(self, scope, path='output_results/dispatch_workflow.sqlite3'):
        self.scope = hashlib.sha256(str(scope).encode()).hexdigest()
        self.path = str(path)
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as db:
            db.execute('CREATE TABLE IF NOT EXISTS records(scope TEXT, task TEXT, body TEXT, PRIMARY KEY(scope,task))')
            db.execute('CREATE TABLE IF NOT EXISTS blocks(id INTEGER PRIMARY KEY, scope TEXT, cg TEXT, start TEXT, end TEXT, reason TEXT)')
            db.execute('CREATE TABLE IF NOT EXISTS pending(scope TEXT, task TEXT, body TEXT, PRIMARY KEY(scope,task))')
            db.execute('CREATE TABLE IF NOT EXISTS history(id INTEGER PRIMARY KEY, scope TEXT, task TEXT, at TEXT, body TEXT)')
    def connect(self): return sqlite3.connect(self.path, timeout=15)
    def get(self, tid):
        with self.connect() as db:
            row=db.execute('SELECT body FROM records WHERE scope=? AND task=?',(self.scope,str(tid))).fetchone()
        return json.loads(row[0]) if row else {}
    def save(self, tid, record):
        record = {**record, 'updated':stamp()}
        body=json.dumps(record,ensure_ascii=False,default=str)
        with self.connect() as db:
            db.execute('INSERT OR REPLACE INTO records VALUES(?,?,?)',(self.scope,str(tid),body))
            db.execute('INSERT INTO history(scope,task,at,body) VALUES(?,?,?,?)',(self.scope,str(tid),stamp(),body))
    def record(self, task, cg, confirmed=False):
        old=self.get(task['任務ID']);sig=task_signature(task,cg)
        if old.get('signature')==sig:return old
        # Replies apply only to the exact person/date/time, never to a changed assignment.
        record = {'signature':sig,'cg':cg or '', 'state':('督導確認' if confirmed else '建議') if cg else '待安排',
                'owner':old.get('owner',''),'due':old.get('due',''),'note':old.get('note',''), 'previous_state':old.get('state','')}
        if old:self.save(task['任務ID'],record)
        return record
    def blocks(self):
        with self.connect() as db:
            rows=db.execute('SELECT id,cg,start,end,reason FROM blocks WHERE scope=? ORDER BY start',(self.scope,)).fetchall()
        return [dict(zip(['id','cg','start','end','reason'],r)) for r in rows]
    def add_blocks(self, cg, day, spans, reason):
        if not reason.strip():raise ValueError('請填寫不可排原因。')
        items=[]
        for start,end in spans:
            a=pd.Timestamp(f'{day} {start}');b=pd.Timestamp(f'{day} {end}')
            if b<=a:raise ValueError('每段結束時間必須晚於開始時間；跨日請分日期登錄。')
            items.append((a,b))
        if not items:raise ValueError('請至少填入一段時段。')
        merged=[]
        for a,b in sorted(items):
            if merged and a<=merged[-1][1]:merged[-1]=(merged[-1][0],max(b,merged[-1][1]))
            else:merged.append((a,b))
        with self.connect() as db:
            for a,b in merged:
                data=(self.scope,str(cg),a.isoformat(),b.isoformat(),reason.strip())
                if not db.execute('SELECT 1 FROM blocks WHERE scope=? AND cg=? AND start=? AND end=? AND reason=?',data).fetchone():
                    db.execute('INSERT INTO blocks(scope,cg,start,end,reason) VALUES(?,?,?,?,?)',data)
            db.execute('INSERT INTO history(scope,task,at,body) VALUES(?,?,?,?)',(self.scope,'availability',stamp(),json.dumps({'cg':cg,'spans':[(str(a),str(b)) for a,b in merged],'reason':reason},ensure_ascii=False)))
    def remove_block(self, block_id):
        with self.connect() as db:
            row=db.execute('SELECT cg,start,end,reason FROM blocks WHERE scope=? AND id=?',(self.scope,block_id)).fetchone()
            db.execute('DELETE FROM blocks WHERE scope=? AND id=?',(self.scope,block_id))
            db.execute('INSERT INTO history(scope,task,at,body) VALUES(?,?,?,?)',(self.scope,'availability',stamp(),json.dumps({'removed':row},ensure_ascii=False)))
    def pending(self):
        with self.connect() as db:
            return [json.loads(r[0]) for r in db.execute('SELECT body FROM pending WHERE scope=?',(self.scope,))]
    def save_pending(self, task):
        with self.connect() as db:
            db.execute('INSERT OR REPLACE INTO pending VALUES(?,?,?)',(self.scope,str(task['任務ID']),json.dumps(task,ensure_ascii=False,default=str)))
    def remove_pending(self, tid):
        with self.connect() as db:
            db.execute('DELETE FROM pending WHERE scope=? AND task=?',(self.scope,str(tid)))
    def history(self, tid):
        with self.connect() as db:
            return db.execute('SELECT at,body FROM history WHERE scope=? AND task=? ORDER BY id DESC LIMIT 20',(self.scope,str(tid))).fetchall()

def attach_blocks(frame, store):
    result=frame.copy();items=store.blocks()
    result['臨時不可排時段']=[json.dumps([b for b in items if b['cg']==str(cg)],ensure_ascii=False) for cg in result.get('居服員ID',[])]
    return result

def render_leave(store, caregivers, tasks, result, overrides):
    import streamlit as st
    st.caption('整天或多段時段共用同一流程。登錄後，重疊服務會轉為待安排；其他時段保留。')
    labels={str(r['居服員ID']):str(r['居服員ID']) for _,r in caregivers.iterrows()}
    names=st.session_state.get('_workflow_cg_labels',{})
    if not labels:
        st.info('請先載入居服員資料。');return
    with st.form('availability_form'):
        cg=st.selectbox('居服員',list(labels),format_func=lambda x:names.get(x,x))
        day=st.date_input('日期',value=pd.to_datetime(tasks['日期']).min().date())
        mode=st.radio('不可排範圍',['整天','指定時段（可多段）'],horizontal=True)
        raw=st.text_area('指定時段：每行一段',value='09:00-11:00\n14:00-15:00',help='選整天時忽略此欄。時間格式 HH:MM-HH:MM。')
        reason=st.selectbox('原因',['請假','看診','教育訓練','家庭事務','其他不可排'])
        submitted=st.form_submit_button('登錄不可排時間',type='primary')
    if submitted:
        try:
            spans=[('00:00','23:59')] if mode=='整天' else [tuple(line.strip().split('-')) for line in raw.splitlines() if line.strip()]
            if any(len(x)!=2 for x in spans):raise ValueError('請使用 HH:MM-HH:MM，每行一段。')
            store.add_blocks(cg,day.isoformat(),spans,reason)
            assignment=dict(zip(result.get('任務ID',[]),result.get('派單居服員',[])))
            affected=sum(1 for _,t in tasks.iterrows() if str(assignment.get(t['任務ID'],''))==str(cg) and any(b['cg']==str(cg) and overlap(t,b) for b in store.blocks()))
            st.session_state['workflow_go_queue']=True
            st.session_state['workflow_notice']=f'已登錄不可排時間，本次有 {affected} 筆已派服務受影響，已轉至接單與待辦。'
            st.rerun()
        except ValueError as exc:st.error(str(exc))
    blocks=store.blocks()
    if blocks:
        with st.expander(f'已登錄不可排時段 · {len(blocks)}'):
            st.dataframe(pd.DataFrame([{'編號':b['id'],'居服員':names.get(b['cg'],b['cg']),'開始':b['start'],'結束':b['end'],'原因':b['reason']} for b in blocks]),hide_index=True,width='stretch')
            bid=st.selectbox('撤銷時段', [b['id'] for b in blocks])
            if st.button('撤銷選取時段'):
                store.remove_block(bid)
                st.session_state['workflow_notice']='已撤銷不可排時段；已改派或待安排服務不會自動還原，請到接單與待辦確認。'
                st.rerun()

def render_queue(store, tasks, result, caregivers, overrides, on_assign, on_rank, on_move, find_slots):
    import streamlit as st
    mapping=dict(zip(result.get('任務ID',[]).astype(str),result.get('派單居服員',[]))) if not result.empty else {}
    records={}; table=[];names=st.session_state.get('_workflow_cg_labels',{}); clients=st.session_state.get('_workflow_cl_labels',{})
    for _,task in tasks.drop_duplicates('任務ID').iterrows():
        tid=str(task['任務ID']);cg=str(mapping.get(tid,'') or '')
        if cg in ['nan','未指派','None']:cg=''
        record=store.record(task,cg,tid in overrides)
        if not cg and not record.get('note'):record['note']=overrides.get(tid,{}).get('reason','')
        records[tid]=(task,cg,record)
        due=pd.to_datetime(record.get('due'),errors='coerce')
        overdue=pd.notna(due) and due.date()<datetime.now(ZoneInfo('Asia/Taipei')).date() and record['state']!='已通知案家'
        table.append({'任務ID':tid,'日期':str(task.get('日期',''))[:10],'時段':str(task.get('時間窗_開始',''))+'–'+str(task.get('時間窗_結束','')),'案家':clients.get(str(task.get('案家ID')),str(task.get('案家ID',''))),'居服員':names.get(cg,cg or '未安排'),'狀態':record['state'],'追蹤人':record.get('owner',''),'期限':record.get('due',''),'逾期':'逾期' if overdue else ''})
    df=pd.DataFrame(table)
    if df.empty:st.info('目前沒有服務任務。');return
    st.caption('追蹤人為人工填寫，非登入驗證。此處統一登錄電話／訊息回覆，不會自動發送通知。待回覆的人選仍保留時段，避免重複安排；拒接後釋出。')
    a,b,c=st.columns(3);a.metric('待安排',int((df['狀態']=='待安排').sum()));b.metric('待接單確認',int(df['狀態'].isin(STATES[:3]).sum()));c.metric('已接單／已通知',int(df['狀態'].isin(STATES[3:]).sum()))
    view=st.radio('查看',['待處理','全部'],horizontal=True,key='workflow_filter')
    search=st.text_input('搜尋案家、居服員或任務',key='workflow_search')
    shown=df if view=='全部' else df[df['狀態']!='已通知案家']
    if search:shown=shown[shown.astype(str).apply(lambda r:r.str.contains(search,regex=False).any(),axis=1)]
    st.dataframe(shown,hide_index=True,width='stretch')
    if shown.empty:return
    index=shown.set_index('任務ID')
    tid=st.selectbox('選取一筆處理',shown['任務ID'].tolist(),format_func=lambda t:f"{index.loc[t,'日期']} {index.loc[t,'時段']} · {index.loc[t,'案家']} · {t}")
    task,cg,record=records[tid];key=tid+record['signature'][:8]
    st.write(f"目前：{record['state']} · {names.get(cg,cg or '未安排')}")
    if not cg and record.get('note'):st.warning(record['note'])
    if record.get('previous_state'):st.caption('人選或時段已變更，前次接單／通知狀態不再適用。')
    if cg:
        with st.form('reply_'+key):
            target=st.selectbox('更新狀態',STATES,index=STATES.index(record['state']) if record['state'] in STATES else 0)
            note=st.text_input('聯繫／確認紀錄',value=record.get('note',''))
            checked=st.checkbox('已向居服員確認接單；如選已通知案家，也已完成案家通知')
            if st.form_submit_button('儲存狀態',type='primary'):
                if target in STATES[3:] and (not checked or not note.strip()):st.error('請勾選實際確認結果並填寫聯繫紀錄。')
                else:store.save(tid,{**record,'state':target,'note':note});st.rerun()
        with st.expander('拒接／撤回人選'):
            reason=st.text_input('拒接或撤回原因',key='reject_'+key)
            if st.button('釋出人選並轉待安排',key='release_'+key):
                if not reason.strip():st.error('請填寫原因。')
                else:
                    error=on_assign(tid,None,reason)
                    if error:st.error(error)
                    else:
                        store.save(tid,{**store.record(task,''),'note':reason,'rejected_cg':cg,'rejected_signature':task_signature(task,cg)});st.rerun()
    with st.expander('找人／無人可派處理',expanded=not bool(cg)):
        if st.button('重新檢查所有候選人',key='rank_'+key):
            with st.spinner('檢查資格、可上班時段、工時與交通…'):
                try:ranked=on_rank(tid,list(caregivers['居服員ID'].astype(str)))
                except Exception as exc:
                    st.error('候選檢查未完成：'+str(exc));ranked=[]
            st.session_state['workflow_rank']={'key':key,'rows':ranked}
        saved=st.session_state.get('workflow_rank',{})
        ranked=saved.get('rows',[]) if saved.get('key')==key else []
        if ranked:
            st.dataframe(pd.DataFrame([{'居服員':names.get(str(r['cg_id']),str(r['cg_id'])),'可派':bool(r.get('available')),'原因':r.get('detail') or r.get('reason','')} for r in ranked]),hide_index=True,width='stretch')
            available=[str(r['cg_id']) for r in ranked if r.get('available')]
            if available:
                choice=st.selectbox('備援／接手人員',available,format_func=lambda x:names.get(x,x),key='choice_'+key)
                if st.button('督導確認此人選',key='assign_'+key):
                    error=on_assign(tid,choice,'統一待辦選擇接手人員')
                    if error:st.error(error)
                    else:
                        store.save(tid,{**store.record(task,choice,True),'state':'督導確認'});st.session_state.pop('workflow_rank',None);st.rerun()
            else:st.warning('目前無可派人選。請登錄追蹤期限，協商其他時間或聯絡備援人力。')
        if st.button('找可協商時段（前後 30／60 分鐘）',key='slots_'+key):
            with st.spinner('重新檢查替代時段…'):
                try:options=find_slots(task)
                except Exception as exc:
                    st.error('協商時段檢查未完成：'+str(exc));options=[]
            st.session_state['workflow_slots']={'key':key,'rows':options}
        found=st.session_state.get('workflow_slots',{})
        slots=found.get('rows',[]) if found.get('key')==key else []
        if found.get('key')==key and not slots:st.info('前後 30／60 分鐘內未找到可行方案；可回「案家臨時改期」協商其他日期。')
        if slots:
            selected=st.selectbox('可協商方案',range(len(slots)),format_func=lambda i:f"{slots[i]['start']}–{slots[i]['end']} · {names.get(slots[i]['cg_id'],slots[i]['cg_id'])}",key='slot_'+key)
            agreed=st.checkbox('已與案家協商同意新時段',key='agree_'+key)
            if st.button('確認改期人選（仍待接單）',disabled=not agreed,key='applyslot_'+key):
                error=on_move(slots[selected],'案家已同意協商時段','僅本次')
                if error:st.error(error)
                else:st.session_state.pop('workflow_slots',None);st.rerun()
    with st.expander('追蹤人、期限與處理紀錄'):
        with st.form('follow_'+key):
            owner=st.text_input('負責追蹤人',value=record.get('owner',''))
            due=st.date_input('追蹤期限',value=pd.Timestamp(record['due']).date() if record.get('due') else None)
            note=st.text_area('處理紀錄／待協商事項',value=record.get('note',''))
            if st.form_submit_button('儲存追蹤'):
                store.save(tid,{**record,'owner':owner,'due':due.isoformat() if due else '', 'note':note});st.rerun()
        history=store.history(tid)
        if history:st.dataframe(pd.DataFrame([{'時間':at, '狀態':json.loads(body).get('state',''),'紀錄':json.loads(body).get('note','')} for at,body in history]),hide_index=True,width='stretch')
