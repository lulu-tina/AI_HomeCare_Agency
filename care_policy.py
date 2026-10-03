"""機構可覆核的排班參考規則；不是臨牀量表或完整法規資格審核。"""
import re
import math

CERT_ALIASES = {
 'BA08足部照護': ('BA08足部照護','足部照護'),
 'BA17A抽吸訓練': ('BA17A','人工氣道管內分泌物抽吸'),
 'BA17B抽吸訓練': ('BA17B','口腔內分泌物抽吸'),
 '失智症照顧專長': ('失智症照顧專長','失智症照顧服務20小時'),
 '精神疾病照顧專長': ('精神疾病照顧專長','精神疾病'),
 '身心障礙支持服務核心課程': ('身心障礙支持服務核心課程','身心障礙支持服務20小時'),
 '單一級照服證照': ('單一級照服證照','照顧服務員單一級'),
}
DEFAULT_CODE_CERTS = {'BA08':'BA08足部照護','BA17A':'BA17A抽吸訓練','BA17B':'BA17B抽吸訓練'}
LEVEL_WEIGHT = {2:1.0,3:1.1,4:1.2,5:1.3,6:1.4,7:1.5,8:1.6}
SEVERITY_WEIGHT = {'無':1.0,'輕':1.1,'輕度':1.1,'中':1.3,'中度':1.3,'重':1.5,'重度':1.5,'極重':1.7,'極重度':1.7}

def text(value):
 if value is None:return ''
 try:
  if math.isnan(value):return ''
 except (TypeError,ValueError):pass
 return str(value).strip()

def certificate_names(cg):
 raw_parts=re.split(r'[、,，;；\n]',text(cg.get('取得資格'))+'、'+text(cg.get('核心專長證照')))
 raw='、'.join(p for p in raw_parts if not any(n in p for n in ['未完成','未取得','未受訓','尚未','無此'])).upper()
 return {name for name,aliases in CERT_ALIASES.items() if any(a.upper() in raw for a in aliases)}

def difficulty_reference(row):
 weights=[];reasons=[];hints=[]
 value=text(row.get('核定等級'))
 try:level=float(value)
 except ValueError:level=0
 if level in LEVEL_WEIGHT:
  weights.append(LEVEL_WEIGHT[level]);reasons.append(f'核定等級{level:g}→{LEVEL_WEIGHT[level]:g}')
 for field in ['失能程度','障礙程度']:
  value=text(row.get(field))
  if value in SEVERITY_WEIGHT:
   weights.append(SEVERITY_WEIGHT[value]);reasons.append(f'{field}{value}→{SEVERITY_WEIGHT[value]:g}')
 proof=text(row.get('身障證明'));book=text(row.get('身障手冊'))
 for needles,label in [(['第01類','新制一類','精神、心智'],'認知／精神功能需確認'),(['第02類','新制二類','感官'],'溝通／感官協助需確認'),(['第07類','新制七類','移動相關'],'移動／移位需求需確認')]:
  if any(n in proof for n in needles):hints.append(label)
 if '多重障礙' in proof or len(set(re.findall(r'第0?([1-8])類',proof)))>1:hints.append('多重障礙照護需求需確認')
 if book.startswith('有') or proof:
  hints.append('可優先配對身心障礙支持訓練者')
 if book and book!='無':reasons.append('身障手冊：'+book+'（不單獨加重）')
 if proof:reasons.append('身障證明：'+proof+'（類別提示，不診斷）')
 return {'照護難度參考係數':max(weights,default=1.0),'照護難度判斷依據':'；'.join(reasons) or '未提供；暫用中性係數1',
         '照護提示':'；'.join(hints),'照護難度規則版本':'機構參考v1：各程度取最大值，不重複累加'}

def certificate_bonus(task,cg):
 # 身障類別本身不判定失智，不把年齡/障礙等級當作體力。
 hint=text(task.get('照護提示'))
 return 5.0 if '身心障礙支持訓練' in hint and '身心障礙支持服務核心課程' in certificate_names(cg) else 0.0

REGISTRATION_FIELDS = {
 '失智症照顧專長':'失智症訓練登錄確認(0/1)',
 'BA08足部照護':'BA08訓練登錄確認(0/1)',
 'BA17A抽吸訓練':'BA17A訓練登錄確認(0/1)',
 'BA17B抽吸訓練':'BA17B訓練登錄確認(0/1)',
}

def is_confirmed(value):
 return text(value).lower() in {'1','1.0','是','已確認','已登錄','true'}

def dementia_status(task):
 value=text(task.get('失智症個案(0/1)'))
 if is_confirmed(value):return True
 if value.lower() in {'0','0.0','否','false'}:return False
 # 只接受已由操作者明確指定的照護需求，不從身障第1類推論。
 if text(task.get('特殊照護需求')) in {'失智引導與精神陪伴','失智症照顧','失智症照護'}:return True
 return None

def special_qualification_error(task,cg,code_requirements):
 needed=set(code_requirements)
 if dementia_status(task) is True:needed.add('失智症照顧專長')
 certs=certificate_names(cg)
 for cert in sorted(needed):
  if cert not in certs and cert not in text(cg.get('核心專長證照')).split('、'):
   return '缺特殊訓練：'+cert
 return None
