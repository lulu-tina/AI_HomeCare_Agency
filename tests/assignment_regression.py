import sys,ast
from pathlib import Path
from datetime import datetime,timedelta
from types import SimpleNamespace
import pandas as pd
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1].resolve()))
from schedule_checks import workload_error,incremental_travel

def dt(s):return datetime.strptime(s,'%H:%M')
def slots(*pairs):return [(dt(a),dt(b)) for a,b in pairs]
assert workload_error(slots(('08:00','12:00'),('13:00','17:00'))) is None
assert '工時超標' in workload_error(slots(('08:00','12:00'),('13:00','17:01')))
assert '工時超標' in workload_error(slots(('08:00','12:00'),('13:00','15:00')),5)
assert '休息不足' in workload_error(slots(('08:00','12:00'),('12:40','13:00')),transfers={1:15})
assert workload_error(slots(('08:00','12:00'),('12:45','13:00')),transfers={1:15}) is None
source=(Path(__file__).resolve().parents[1] / 'caregiver_engine.py').read_text();tree=ast.parse(source)
ns=dict(pd=pd,np=np,Optional=object,datetime=datetime,timedelta=timedelta)
ns['Optional']=__import__('typing').Optional
ns.update(_text=lambda v:str(v) if pd.notna(v) else '',_number=lambda v,default=np.nan:float(v) if pd.notna(v) else default,
          _check_hard_constraints=lambda *a:None,_build_cg_busy_blocks=lambda d:{},
          parse_time=lambda s:tuple(map(dt,s.split('-'))),_interval_travel=lambda *a:0,
          calc_travel_minutes=lambda *a,**k:0,_task_clock=lambda t,v:v)
node=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='_evaluate_reassignment')
exec(compile(ast.Module(body=[node],type_ignores=[]),'engine','exec'),ns)
config=SimpleNamespace(continuous_work_limit_mins=240,mandatory_break_mins=30)
tasks=pd.DataFrame([{'任務ID':f'T{i}','日期':day,'案家ID':f'A{i}','時間窗_開始':a,'時間窗_結束':b,'服務地點_緯度':25,'服務地點_經度':121} for i,(a,b,day) in enumerate([('08:00','12:00','2026-10-05'),('13:00','17:00','2026-10-05'),('18:00','18:01','2026-10-05'),('08:00','12:00','2026-10-06')])])
cgs=pd.DataFrame([{'居服員ID':f'C{i}','每日工時上限(小時)':8,'常用交通工具':'機車','服務起點_緯度(家)':25,'服務起點_經度(家)':121} for i in range(6)])
assigned=pd.DataFrame([{'任務ID':'T0','派單居服員':'C0'},{'任務ID':'T1','派單居服員':'C0'}])
eval=ns['_evaluate_reassignment']
assert not eval('T2','C0',tasks,assigned,cgs,config)['available']
assert eval('T1','C0',tasks,assigned,cgs,config)['available'] # exclude itself
assert eval('T3','C0',tasks,assigned,cgs,config)['available'] # different day
# Test the actual repair block: first 5 candidates fail, sixth succeeds; no unsafe fallback.
phase=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='run_phase2_optimization')
start=next(i for i,n in enumerate(phase.body) if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='initial' for t in n.targets))
repair=ast.FunctionDef(name='repair',args=ast.arguments(posonlyargs=[],args=[ast.arg(arg=x) for x in ['result','df_matches','tasks','df_cg','config','extra_busy_blocks','task_revenue_map','status']],kwonlyargs=[],kw_defaults=[],defaults=[]),body=phase.body[start:],decorator_list=[])
ns.update(pywraplp=SimpleNamespace(Solver=SimpleNamespace(FEASIBLE=1)),calculate_schedule_travel=lambda r,*a:r)
ns['_evaluate_reassignment']=lambda tid,cid,*a,**kw:{'available':cid=='C5','detail':'服務時間重疊'}
exec(compile(ast.fix_missing_locations(ast.Module(body=[repair],type_ignores=[])),'repair','exec'),ns)
matches=pd.DataFrame([{'任務ID':'T0','居服員ID':f'C{i}','適配度分數':100-i,'地點緯度':25,'地點經度':121} for i in range(6)])
r=ns['repair']({'df_result':pd.DataFrame(),'df_matches':matches.copy()},matches,tasks,cgs,config,None,{},0)
assert r['df_result'].iloc[0]['派單居服員']=='C5' and r['repair_count']==1
r=ns['repair']({'df_result':pd.DataFrame(),'df_matches':matches.copy()},matches.iloc[:5],tasks,cgs,config,None,{},0)
assert r['assigned_count']==0 and len(r['pending_details']['T0'])==5
# Cache correctness: unchanged = zero calls, moving person recalculates both affected routes only.
cache={};calls=[]
def calculate(r,t,c,config):
 calls.append(tuple(r['任務ID']));return r.assign(**{'交通路段':'calculated','預估車程(分)':10})
r=pd.DataFrame([{'任務ID':'T0','派單居服員':'C0'},{'任務ID':'T1','派單居服員':'C1'},{'任務ID':'T2','派單居服員':'C2'},{'任務ID':'T3','派單居服員':'C0'}])
incremental_travel(r,tasks,cgs,config,cache,calculate);assert len(calls)==4
calls.clear();incremental_travel(r,tasks,cgs,config,cache,calculate);assert calls==[]
r.loc[0,'派單居服員']='C1'
incremental_travel(r,tasks,cgs,config,cache,calculate);assert calls==[('T0','T1')]
calls.clear();t2=tasks.copy();t2.loc[2,'時間窗_結束']='18:02'
incremental_travel(r,t2,cgs,config,cache,calculate);assert calls==[('T2',)]
print('PASS: 8-hour boundary, lower employee cap, travel-aware rest, self-exclusion, date isolation, sixth-candidate repair, no unsafe fallback, incremental route caching and time-change invalidation.')
