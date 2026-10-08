"""CareFlow 班表工作台：單日全員、個人整週與異動清單。"""
from datetime import date, datetime, timedelta
import re
import pandas as pd
import streamlit as st


def apply_overrides_to_result(frame, overrides):
    result = frame.copy() if frame is not None else pd.DataFrame()
    for task_id, override in (overrides or {}).items():
        cg_id = override.get("cg_id")
        mask = result["任務ID"].astype(str).eq(str(task_id)) if "任務ID" in result else pd.Series(False, index=result.index)
        if mask.any():
            if cg_id is None:
                result = result.loc[~mask].copy()
            else:
                result.loc[mask, "派單居服員"] = cg_id
        elif cg_id is not None:
            result = pd.concat([result, pd.DataFrame([{"任務ID": task_id, "派單居服員": cg_id}])], ignore_index=True)
    return result


def build_schedule_rows(result, tasks, caregivers, overrides=None, client_names=None, caregiver_names=None):
    """ID 保持唯一鍵；遮罩姓名僅用於畫面。"""
    task_rows = tasks.drop_duplicates("任務ID") if "任務ID" in tasks else pd.DataFrame()
    assigned = apply_overrides_to_result(result, overrides)
    assigned_map = dict(zip(assigned.get("任務ID", []), assigned.get("派單居服員", [])))
    cg_ids = [str(x) for x in caregivers.get("居服員ID", pd.Series(dtype=object)).dropna().unique()]
    name_map_clients = {str(k): str(v) for k, v in (client_names or {}).items()}
    rows = []
    for _, row in task_rows.iterrows():
        task_id = str(row.get("任務ID", ""))
        day = pd.to_datetime(row.get("日期"), errors="coerce")
        start, end = str(row.get("時間窗_開始", ""))[:5], str(row.get("時間窗_結束", ""))[:5]
        if pd.isna(day) or not re.fullmatch(r"\d{2}:\d{2}", start) or not re.fullmatch(r"\d{2}:\d{2}", end):
            continue
        cg = str(assigned_map.get(row.get("任務ID"), ""))
        if cg in ("", "未指派", "nan"):
            cg = ""
        client_id = str(row.get("案家ID", ""))
        recurrence = str(row.get("頻率", ""))
        tag = "隔週" if "隔週" in recurrence or "隔周" in recurrence else "每週" if "每週" in recurrence else "單次"
        override = (overrides or {}).get(task_id)
        if override and override.get("cg_id"):
            tag = "代班" if "請假" in str(override.get("reason", "")) else "已調整"
        codes = [str(row[key]).strip() for key in row.index if re.fullmatch(r"Service_Code_\d+", str(key)) and pd.notna(row[key]) and str(row[key]).strip()]
        code = ' / '.join(codes)
        rows.append({"id": task_id, "date": day.date().isoformat(), "start": start, "end": end,
                     "cg": cg, "client": client_id, "client_label": (name_map_clients[client_id] + '｜' + client_id) if client_id in name_map_clients else '姓名未對照｜' + client_id,
                     "cg_label": str((caregiver_names or {}).get(cg, '')) + '｜' + cg,
                     "code": code if code != "nan" else "", "tag": tag, "reason": str((override or {}).get("reason", ""))})
    name_map = {str(k): str(v) for k, v in (caregiver_names or {}).items()}
    labels = {cg: f"{name_map.get(cg, '姓名未對照')}｜{cg}" for cg in cg_ids}
    return rows, labels


_GRID_CSS = """
.cf-workspace {font-family:var(--st-font),sans-serif; color:#183d3d; overflow:auto; max-height:780px; border:1px solid #dae8e2;border-radius:16px;background:#fff}
.cf-head {display:grid;position:sticky;top:0;z-index:5;background:#f1f7f4;border-bottom:1px solid #dce9e3}
.cf-head div {min-width:155px;padding:13px 10px;font-size:13px;font-weight:750;border-left:1px solid #e3eee8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cf-head div:first-child {min-width:64px;border-left:0}
.cf-board {position:relative;display:grid;min-width:max-content}
.cf-time {position:absolute;left:0;width:64px;text-align:center;color:#6b807b;font-size:12px;padding-top:2px}
.cf-lane {position:relative;min-width:155px;border-left:1px solid #e3eee8;background:repeating-linear-gradient(to bottom,#fff 0,#fff 83px,#edf3ef 84px)}
.cf-card {position:absolute;left:5px;right:5px;overflow:hidden;border:1px solid #a9d9c7;border-left:4px solid #16816c;border-radius:8px;background:#dff4eb;padding:5px 7px;box-sizing:border-box;cursor:pointer;font-size:12px;line-height:1.4;z-index:2}
.cf-card strong {display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:13px;padding-right:22px}
.cf-card span {display:block}
.cf-card[data-tag="隔週"] {background:#eeebff;border-color:#c7bcf5;border-left-color:#8067ce}
.cf-card[data-tag="代班"],.cf-card[data-tag="已調整"] {background:#fff1d9;border-color:#f4d394;border-left-color:#e49a19}
.cf-card:focus {outline:2px solid #127b6b;outline-offset:1px}
.cf-card[draggable="true"]:active {cursor:grabbing}
.cf-lane.cf-over {background-color:#f0f9f5}
.cf-hint {color:#667e76;font-size:12px;margin:9px 0}
.cf-handle {position:absolute;right:2px;top:2px;border:0;border-radius:5px;background:#ffffffaa;color:#155b4d;padding:2px 5px;cursor:grab;touch-action:none;user-select:none;font-size:18px;line-height:22px}
.cf-details {position:absolute;bottom:3px;right:6px;border:0;background:transparent;color:#175b4d;text-decoration:underline;cursor:pointer;padding:0;margin-top:2px;font-size:12px}
.cf-live {padding:9px 12px;margin-bottom:8px;background:#eaf5ee;color:#235b4b;border-radius:8px;font-size:13px;min-height:22px}
.cf-drag-ghost {position:fixed;pointer-events:none;z-index:10000;background:#135e50;color:white;white-space:pre-line;padding:10px 14px;border-radius:9px;font-size:13px;box-shadow:0 4px 18px #0003}
.cf-lane.cf-over {box-shadow:inset 0 0 0 2px #248974}
.cf-card {user-select:none}
.cf-event-person {padding-right:28px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cf-head div {min-width:180px;white-space:normal}

"""

_GRID_JS = r"""
export default function({parentElement,data,setStateValue}) {
  const root=parentElement.querySelector('#cf-calendar'); root.replaceChildren();
  const days=data.days||[], people=data.people||[], rows=data.rows||[];
  const first=480,last=1260,step=30,px=84;
  const minute=v=>{const p=(v||'').split(':').map(Number);return p.length===2&&p.every(Number.isFinite)?p[0]*60+p[1]:null};
  const clock=n=>String(Math.floor(n/60)).padStart(2,'0')+':'+String(n%60).padStart(2,'0');
  const el=(tag,cls,text)=>{const n=document.createElement(tag);if(cls)n.className=cls;if(text!==undefined)n.textContent=text;return n};
  const emit=(kind,payload)=>{
    status.textContent=kind==='select'?'正在開啟服務詳情…':'正在開啟改班確認視窗…';
    setStateValue('action',{kind,...payload,nonce:globalThis.crypto?.randomUUID?.()||String(Date.now())+Math.random()});
  };
  const lanes=data.mode==='week'?days.map(d=>({id:data.person,day:d,label:d})):people.map(p=>({id:p.id,day:days[0],label:p.label}));
  const status=el('div','cf-live','點卡片或「詳情」查看；按住右上角 ⠿ 移動。');status.setAttribute('aria-live','polite');root.appendChild(status);
  const board=el('div','cf-workspace');
  const columns='64px repeat('+lanes.length+', minmax(180px, 1fr))';
  const head=el('div','cf-head');head.style.gridTemplateColumns=columns;head.appendChild(el('div','','時間'));
  lanes.forEach(l=>head.appendChild(el('div','',l.label)));board.appendChild(head);
  const body=el('div','cf-board');body.style.gridTemplateColumns=columns;body.style.height=((last-first)/step*px)+'px';
  const gutter=el('div');gutter.style.position='relative';
  for(let t=first;t<last;t+=60){const n=el('div','cf-time',clock(t));n.style.top=((t-first)/step*px)+'px';gutter.appendChild(n)}body.appendChild(gutter);
  const laneElements=[];let active=null,ghost=null;
  function clean(){if(active?.card)active.card.style.opacity='';active=null;ghost?.remove();ghost=null;laneElements.forEach(x=>x.node.classList.remove('cf-over'))}
  function targetAt(x,y){return laneElements.find(l=>{const r=l.node.getBoundingClientRect();return x>=r.left&&x<=r.right&&y>=r.top&&y<=r.bottom})}
  function position(e){
    const target=targetAt(e.clientX,e.clientY);if(!target)return null;
    const r=target.node.getBoundingClientRect();let start=first+Math.round(((e.clientY-r.top-active.offset)/px*step)/5)*5;
    if(Math.abs(e.clientY-active.y)<8)start=minute(active.row.start);
    const duration=minute(active.row.end)-minute(active.row.start);start=Math.max(first,Math.min(last-duration,start));
    return {target,start,end:start+duration};
  }
  for(const lane of lanes){
    const col=el('div','cf-lane');laneElements.push({node:col,...lane});
    for(const row of rows.filter(x=>x.cg===lane.id&&x.date===lane.day)){
      const begin=minute(row.start),end=minute(row.end);if(begin===null||end===null||end<=first||begin>=last)continue;
      const card=el('div','cf-card');card.dataset.tag=row.tag;card.tabIndex=0;card.setAttribute('role','button');
      card.style.top=(Math.max(0,(begin-first)/step*px)+2)+'px';card.style.height=Math.max(30,(Math.min(last,end)-Math.max(first,begin))/step*px-4)+'px';
      card.title=[row.client_label,row.start+'–'+row.end,row.code,row.tag,row.cg_label].join('｜');card.setAttribute('aria-label',card.title);
      const title=el('strong','',row.client_label);card.appendChild(title);
      card.appendChild(el('span','cf-event-time',row.start+'–'+row.end+' · '+(row.workflow_state||row.tag)));
      card.appendChild(el('span','cf-event-code',row.code));
      card.appendChild(el('span','cf-event-person',row.cg_label));
      const details=el('button','cf-details','詳情');details.type='button';details.onclick=e=>{e.stopPropagation();emit('select',{task_id:row.id})};card.appendChild(details);
      const handle=el('button','cf-handle','⠿');handle.type='button';handle.title='按住並移動：換人／改時間';handle.setAttribute('aria-label','拖曳調整班表');card.appendChild(handle);
      card.onclick=e=>{if(e.target!==handle)emit('select',{task_id:row.id})};
      card.onkeydown=e=>{if(e.target===card&&(e.key==='Enter'||e.key===' ')){e.preventDefault();emit('select',{task_id:row.id})}};
      handle.onclick=e=>e.stopPropagation();
      handle.onpointerdown=e=>{if(e.button!==0)return;e.preventDefault();e.stopPropagation();clean();active={row,card,x:e.clientX,y:e.clientY,offset:e.clientY-card.getBoundingClientRect().top,moved:false};handle.setPointerCapture(e.pointerId)};
      handle.onpointermove=e=>{
        if(!active)return;
        if(!active.moved&&Math.hypot(e.clientX-active.x,e.clientY-active.y)<6)return;
        active.moved=true;active.card.style.opacity='.45';
        if(!ghost){ghost=el('div','cf-drag-ghost');root.appendChild(ghost)}
        ghost.style.left=(e.clientX+14)+'px';ghost.style.top=(e.clientY+12)+'px';
        laneElements.forEach(l=>l.node.classList.remove('cf-over'));const p=position(e);
        if(p){p.target.node.classList.add('cf-over');ghost.textContent=p.target.label+'\n'+clock(p.start)+'–'+clock(p.end);status.textContent='放開後預覽：'+p.target.label+' '+clock(p.start)+'–'+clock(p.end)}
        else {ghost.textContent='請移到班表欄位內';status.textContent='移至班表內才可放開調整。'}
      };
      handle.onpointerup=e=>{
        if(!active)return;const source=active.row;const p=position(e);const moved=active.moved;clean();
        if(handle.hasPointerCapture(e.pointerId))handle.releasePointerCapture(e.pointerId);
        if(moved&&p&&(p.target.id!==source.cg||p.target.day!==source.date||p.start!==minute(source.start)))emit('move',{task_id:source.id,cg_id:p.target.id,date:p.target.day,start:clock(p.start),end:clock(p.end)});
        else status.textContent='未更動班表。';
      };
      handle.onpointercancel=()=>{clean();status.textContent='已取消拖曳。'};
      col.appendChild(card);
    }body.appendChild(col);
  }
  board.appendChild(body);root.appendChild(board);
  return ()=>clean();
}
"""


def _render_grid(rows, people, days, mode, person=""):
    if not hasattr(st.components, "v2"):
        st.warning("請更新 Streamlit 至 1.52 以上，才能顯示互動班表。")
        return None
    component = st.components.v2.component("careflow_time_grid", html='<div id="cf-calendar"></div>', css=_GRID_CSS, js=_GRID_JS)
    return component(data={"rows": rows, "people": people, "days": days, "mode": mode, "person": person},
                     key="careflow_grid_v2", default={"action": None}, on_action_change=lambda: None)


def _close_assignment_dialog():
    st.session_state.pop("cf_selected_task", None)
    st.session_state.pop("cf_move_preview", None)
    st.session_state.pop("cf_dialog_error", None)
    st.session_state.pop("cf_dialog_candidates", None)


def _open_assignment_dialog(task_id):
    # Fresh widget values when reopening a task after a prior move or cancellation.
    for prefix in ('cf_dialog_cg_', 'cf_dialog_scope_', 'cf_edit_date_', 'cf_edit_time_', 'cf_dialog_reason_'):
        st.session_state.pop(prefix + str(task_id), None)
    st.session_state["cf_selected_task"] = str(task_id)
    st.session_state.pop("cf_move_preview", None)
    st.session_state.pop("cf_dialog_error", None)
    st.session_state.pop("cf_dialog_candidates", None)


def _process_assignment_request(on_reassign, on_move=None):
    # Run once during the full app rerun, before the scope widget is rendered.
    request = st.session_state.pop("cf_assignment_request", None)
    if not request:
        return
    st.session_state["cf_selected_task"] = request["task_id"]
    callback = on_move if request.get('kind') == 'move' else on_reassign
    if not callable(callback):
        st.session_state["cf_dialog_error"] = "目前未接上改派功能，請確認 app.py 與 calendar_view.py 版本一致。"
        return
    st.session_state["override_apply_scope"] = request["scope"]
    try:
        with st.spinner("正在檢查資格、時段、工時與交通，請稍候…"):
            error = callback(request['draft'], request['reason'], request['scope']) if request.get('kind') == 'move' else callback(request["task_id"], request["cg_id"], request["reason"])
    except Exception as exc:
        error = "安排未完成：" + str(exc)
    if error:
        st.session_state["cf_dialog_error"] = str(error)
        st.session_state.pop("cf_dialog_candidates", None)
    else:
        _close_assignment_dialog()
        st.session_state["cf_assignment_notice"] = "已安排居服員，班表與待安排清單已更新。"
        st.rerun()


@st.dialog("安排居服員", width="medium", on_dismiss=_close_assignment_dialog)
def _assignment_dialog(row, labels, on_list_candidates, can_reassign, can_move=False):
    task_id = row["id"]
    st.markdown(f"**{row['client_label']}** · {row['code']} · {row['tag']}")
    st.caption(f"{row['date']} {row['start']}–{row['end']} · 目前：{labels.get(row['cg'], '未安排')} · 任務 {task_id}")
    error = st.session_state.get("cf_dialog_error")
    if error:
        st.error(error)
    if not labels:
        st.warning("目前沒有可選擇的居服員，請先確認員工資料。")
        if st.button("關閉", key="cf_empty_dialog_close"):
            _close_assignment_dialog()
            st.rerun()
        return
    if st.button("檢查並排序可派人選", key="cf_check_" + task_id, disabled=not callable(on_list_candidates)):
        try:
            with st.spinner("正在檢查候選人，包含交通時間，請稍候…"):
                ranked = on_list_candidates(task_id, list(labels))
            st.session_state["cf_dialog_candidates"] = {"task_id": task_id, "rows": ranked or []}
        except Exception as exc:
            st.error("候選人檢查未完成：" + str(exc))
    saved = st.session_state.get("cf_dialog_candidates", {})
    ranked = saved.get("rows", []) if saved.get("task_id") == task_id else []
    status = {str(r.get("cg_id") or r.get("居服員ID")): r for r in ranked}
    choices = sorted(labels, key=lambda cg: (not status.get(cg, {}).get("available", False), cg))
    def option_label(cg):
        info = status.get(cg)
        if info is None:
            return labels[cg] + " · 尚未檢查"
        return labels[cg] + (" · " + str(info.get("detail") or ("可派" if info.get("available") else "不符合條件")))
    choice = st.selectbox("指定居服員", choices, index=choices.index(row["cg"]) if row["cg"] in choices else None, placeholder="搜尋或選擇居服員", format_func=option_label, key="cf_dialog_cg_" + task_id)
    st.caption("可直接選人；按『確認安排』時仍會重新檢查是否可派。")
    scope = st.radio("套用範圍", ["僅本次", "後續同週期服務"], horizontal=True, key="cf_dialog_scope_" + task_id)
    with st.expander("需要更改日期或開始時間？"):
        changed_date = st.date_input("新日期", value=pd.Timestamp(row['date']).date(), key='cf_edit_date_' + task_id)
        changed_time = st.time_input("新開始時間", value=datetime.strptime(row['start'], '%H:%M').time(), step=300, key='cf_edit_time_' + task_id)
        st.caption('服務分鐘維持原長度；只更換居服員時不用修改。')
    reason = st.text_input("安排／改派原因（必填）", placeholder="例如：補排未安排服務、臨時代班", key="cf_dialog_reason_" + task_id)
    c1, c2 = st.columns(2)
    if c1.button("確認安排", type="primary", width="stretch", key="cf_dialog_confirm_" + task_id, disabled=not can_reassign):
        if choice is None:
            st.warning("請先選擇居服員。")
        elif not reason.strip():
            st.warning("請填寫安排原因。")
        else:
            if changed_date.isoformat() != row['date'] or changed_time.strftime('%H:%M') != row['start']:
                if not can_move:
                    st.error('主程式未接上時段調整功能，請一起更新 app.py。')
                    return
                if scope != '僅本次':
                    st.warning('更改日期或時段請選「僅本次」。')
                    return
                duration = datetime.strptime(row['end'], '%H:%M') - datetime.strptime(row['start'], '%H:%M')
                end = datetime.combine(changed_date, changed_time) + duration
                if end.date() != changed_date:
                    st.warning('服務不能跨日。')
                    return
                draft = {'task_id':task_id, 'cg_id':choice, 'date':changed_date.isoformat(), 'start':changed_time.strftime('%H:%M'), 'end':end.strftime('%H:%M')}
                st.session_state['cf_assignment_request'] = {'kind':'move','task_id':task_id,'draft':draft,'scope':scope,'reason':reason.strip()}
            else:
                st.session_state["cf_assignment_request"] = {"task_id": task_id, "cg_id": choice, "scope": scope, "reason": reason.strip()}
            st.rerun()
    if c2.button("取消", width="stretch", key="cf_dialog_cancel_" + task_id):
        _close_assignment_dialog()
        st.rerun()


@st.dialog("確認拖曳調班", width="medium", on_dismiss=_close_assignment_dialog)
def _move_dialog(original, draft, labels, can_move):
    st.markdown(f"**{original['client_label']}** · {original['code']}")
    st.write(f"原班：{original['date']} {original['start']}–{original['end']} · {labels.get(original['cg'], '未安排')}")
    st.write(f"新班：{draft['date']} {draft['start']}–{draft['end']} · {labels.get(draft['cg_id'], draft['cg_id'])}")
    if st.session_state.get('cf_dialog_error'):
        st.error(st.session_state['cf_dialog_error'])
    scope = st.radio('套用範圍', ['僅本次', '後續同週期服務'], horizontal=True, key='cf_move_scope')
    reason = st.text_input('調整原因（必填）', key='cf_move_reason')
    c1,c2=st.columns(2)
    if c1.button('確認調整', type='primary', key='cf_confirm_move', disabled=not can_move):
        if not reason.strip():
            st.warning('請填寫調整原因。')
        elif scope != '僅本次' and (draft['date'], draft['start'], draft['end']) != (original['date'], original['start'], original['end']):
            st.warning('更改日期或時段時請選「僅本次」。')
        else:
            st.session_state['cf_assignment_request'] = {'kind':'move', 'task_id':original['id'], 'draft':draft, 'scope':scope, 'reason':reason.strip()}
            st.rerun()
    if c2.button('取消', key='cf_cancel_move'):
        _close_assignment_dialog()
        st.rerun()


def render_calendar_overview(result, tasks, caregivers, overrides=None, on_reassign=None,
                             on_list_candidates=None, on_move=None, client_names=None, caregiver_names=None, **kwargs):
    _process_assignment_request(on_reassign, on_move)
    rows, labels = build_schedule_rows(result, tasks, caregivers, overrides, client_names, caregiver_names)
    workflow = kwargs.get('workflow_store')
    if workflow:
        task_map = {str(t['任務ID']): t for _,t in tasks.iterrows()}
        for row in rows:
            row['workflow_state'] = workflow.record(task_map[row['id']],row['cg'],row['id'] in (overrides or {}))['state']
    st.markdown("### 班表工作台")
    if any(r["client_label"].startswith("姓名未對照｜") for r in rows) or any(v.startswith("姓名未對照｜") for v in labels.values()):
        st.info("部分姓名尚未顯示。請在『個資顯示設定』開啟姓名，並確認 Excel 的『案家姓名對照』與『居服員姓名對照』包含對應 ID。")
    notice = st.session_state.pop("cf_assignment_notice", None)
    if notice:
        st.success(notice)
    if not rows:
        st.info("目前沒有帶日期和時段的服務任務。")
        return
    dates = sorted({r["date"] for r in rows})
    st.caption("下方日期只切換查看範圍，不會重新執行排班。全員日班一次看一天；個人週班一次看連續 7 天。")
    if st.session_state.get('_cf_week_view_version') != 1:
        st.session_state['cf_calendar_mode'] = '個人週班'
        st.session_state.pop('cf_week', None)
        st.session_state['_cf_week_view_version'] = 1
    mode = st.segmented_control("班表視角", ["全員日班", "個人週班", "月異動"], key="cf_calendar_mode")
    if mode == "月異動":
        month = st.date_input("查看月份", value=pd.Timestamp(dates[0]).date(), key="cf_month")
        prefix = month.strftime("%Y-%m")
        changes = [r for r in rows if r["date"].startswith(prefix) and
                   (r["tag"] in ("代班", "已調整") or not r["cg"])]
        display_rows = [{"日期":r["date"], "時間":r["start"]+"–"+r["end"], "案家":r["client_label"],
                         "狀態":"待安排" if not r["cg"] else r["tag"], "居服員":labels.get(r["cg"],"")} for r in changes]
        leave_count = 0
        for _, cg_row in caregivers.iterrows():
            leave = cg_row.get("請假紀錄")
            if isinstance(leave, (list, tuple)) and len(leave) == 2:
                parts = leave
            elif isinstance(leave, str) and " ~ " in leave:
                parts = leave.split(" ~ ", 1)
            else:
                continue
            start_leave = pd.to_datetime(parts[0], errors="coerce")
            end_leave = pd.to_datetime(parts[1], errors="coerce")
            if pd.isna(start_leave) or pd.isna(end_leave) or end_leave < start_leave:
                continue
            if start_leave.strftime("%Y-%m") == prefix or end_leave.strftime("%Y-%m") == prefix:
                leave_count += 1
                display_rows.append({"日期":start_leave.date().isoformat(), "時間":start_leave.strftime("%H:%M")+"–"+end_leave.strftime("%H:%M"),
                                     "案家":"—", "狀態":"居服員請假", "居服員":labels.get(str(cg_row.get("居服員ID", "")), "")})
        c1,c2,c3=st.columns(3)
        c1.metric("待安排",sum(not r["cg"] for r in changes))
        c2.metric("已調整／代班",sum(bool(r["cg"]) for r in changes))
        c3.metric("請假紀錄",leave_count)
        if display_rows:
            st.dataframe(pd.DataFrame(display_rows).sort_values(["日期","時間"]), hide_index=True, width="stretch")
        else:
            st.success("這個月沒有待安排或已調整的服務。")
        return
    left, right = st.columns([3.3, 1.1], gap="medium")
    with left:
        if mode == "全員日班":
            selected_day = st.date_input("日期", value=pd.Timestamp(dates[0]).date(), key="cf_day")
            selected = st.multiselect("顯示居服員（最多 6 位）", list(labels), default=list(labels)[:5],
                                      format_func=lambda x: labels[x], key="cf_people")
            if len(selected)>6:
                st.warning("一次最多顯示 6 位，請縮小範圍。")
                selected=selected[:6]
            if not selected:
                st.info("請選擇至少一位居服員。")
                return
            people=[{"id":x,"label":labels[x]} for x in selected]
            grid=_render_grid(rows,people,[selected_day.isoformat()],"day")
        else:
            active_people = {r['cg'] for r in rows if r['cg']}
            ordered_people = sorted(labels, key=lambda x: (x not in active_people, x))
            if not ordered_people:
                st.info("目前没有可顯示的居服員。")
                return
            person = st.selectbox("居服員", ordered_people, format_func=lambda x: labels[x], key="cf_person")
            view_start = pd.Timestamp(kwargs.get('period_start') or dates[0]).date()
            chosen = st.date_input("顯示起日（連續 7 天）",value=view_start,key="cf_week")
            days=[(chosen+timedelta(days=i)).isoformat() for i in range(7)]
            st.caption(f"查看 {days[0]} ～ {days[-1]}；切換居服員即可查看其他人的整週班表。")
            grid=_render_grid(rows,[],days,"week",person)
        st.caption("點卡片或「詳情」開啟視窗；按住卡片右上角 ⠿ 拖曳，放開後確認。可精確至 5 分鐘，原服務分鐘不變；也可在詳情內直接輸入日期和時間。")
    pending_clicked = False
    with left:
        visible_rows = [r for r in rows if r['date'] == selected_day.isoformat()] if mode == '全員日班' else [r for r in rows if r['date'] in days and r['cg'] == person]
        with st.expander('搜尋服務／直接開啟詳情'):
            by_id = {r['id']: r for r in visible_rows}
            def service_label(task_id):
                r = by_id[task_id]
                return f"{r['date']} {r['start']} · {r['client_label']} · {labels.get(r['cg'], '未安排')} · {r['code']}"
            picked = st.selectbox('服務', list(by_id), index=None, format_func=service_label, placeholder='搜尋案家、居服員或服務代碼', key='cf_direct_service')
            if st.button('開啟服務詳情', disabled=picked is None, key='cf_direct_open'):
                pending_clicked = True
                _open_assignment_dialog(picked)
    with right:
        visible_day = selected_day.isoformat() if mode == "全員日班" else None
        pending=[r for r in rows if not r["cg"] and (visible_day is None and r["date"] in days or r["date"]==visible_day)]
        st.markdown(f"**待安排 · {len(pending)}**")
        for row in pending[:12]:
            with st.container(border=True):
                st.markdown(f"**{row['client_label']}** · {row['code']}")
                st.caption(f"{row['date']} {row['start']}–{row['end']}")
                if st.button("安排居服員", key="cf_pending_"+row["id"]):
                    pending_clicked = True
                    _open_assignment_dialog(row["id"])
        if not pending:
            st.caption("這段期間沒有未安排服務。")
        if len(pending)>12:
            st.caption(f"還有 {len(pending)-12} 筆，請縮小日期範圍。")
    action = getattr(grid, 'action', None) if grid is not None else None
    if action and action.get('nonce') != st.session_state.get('cf_last_grid_action'):
        st.session_state['cf_last_grid_action'] = action.get('nonce')
        if not pending_clicked:
            if action.get('kind') == 'select':
                _open_assignment_dialog(str(action.get('task_id', '')))
            elif action.get('kind') == 'move':
                source = next((r for r in rows if r['id'] == action.get('task_id')), None)
                if source and action.get('cg_id') in labels:
                    st.session_state['cf_move_preview'] = action
                    st.session_state.pop('cf_selected_task', None)
                    st.session_state.pop('cf_dialog_error', None)
    draft = st.session_state.get('cf_move_preview')
    if draft:
        original = next((r for r in rows if r['id'] == draft.get('task_id')), None)
        if original:
            _move_dialog(original, draft, labels, callable(on_move))
        else:
            st.session_state.pop('cf_move_preview', None)
    else:
        selected_id = st.session_state.get('cf_selected_task')
        if selected_id:
            selected_row = next((r for r in rows if r['id'] == selected_id), None)
            if selected_row:
                _assignment_dialog(selected_row, labels, on_list_candidates, callable(on_reassign), callable(on_move))
            else:
                _close_assignment_dialog()
                st.warning('選取的任務已不在目前班表，請重新選擇。')


def render_task_override_picker(tasks,result,caregivers,overrides,on_reassign=None,on_clear_override=None,on_list_candidates=None,**kwargs):
    st.subheader("人工搜尋改派")
    ids=[str(x) for x in tasks.get("任務ID",pd.Series(dtype=object)).dropna()]
    if not ids:
        st.info("目前沒有可調整任務。")
        return
    task_id=st.selectbox("選擇任務",ids,key="calendar_override_task")
    caregiver_ids=[str(x) for x in caregivers.get("居服員ID",pd.Series(dtype=object)).dropna()]
    ranked=on_list_candidates(task_id,caregiver_ids) if callable(on_list_candidates) else []
    labels={str(r.get("cg_id") or r.get("居服員ID")): ("✅ 可派" if r.get("available") else "⚠️ "+str(r.get("detail") or "需確認")) for r in ranked or []}
    choice=st.selectbox("改派居服員",["未指派"]+caregiver_ids,format_func=lambda x:x if x=="未指派" else x+"｜"+labels.get(x,""),key="calendar_override_caregiver")
    reason=st.text_input("改派原因",key="calendar_override_reason")
    c1,c2=st.columns(2)
    if c1.button("套用改派",key="calendar_apply_override",width="stretch"):
        if not reason.strip():st.warning("請填寫原因。")
        elif callable(on_reassign):
            err=on_reassign(task_id,None if choice=="未指派" else choice,reason.strip())
            if err:st.error(str(err))
            else:st.success("已套用改派。");st.rerun()
    if c2.button("清除此任務覆寫",key="calendar_clear_override",width="stretch") and callable(on_clear_override):
        on_clear_override(task_id);st.rerun()
