"""在你的活頁簿加入「示範未派案例」（全部是虛構案家，名稱為 DEMO…），用來練習居督的處理流程。
    python tools/make_cases_workbook.py 你的檔.xlsx 輸出.xlsx

加入的東西（其他資料完全不動，樣式沿用原檔）：
  • 機動居服員 M04（休週日）、M05（休週六）— 沒有掛成任何案家的固定人選，所以不改變自動排班結果。
  • DEMO0003 吳○蘭：每週一 12:30～13:20 足部照護（BA08），固定居服員 E6279。
        前一個服務 12:00～12:15 結束，只剩 15 分鐘，趕過來要 18 分鐘 → 排不進；往後移 30 分鐘就可以。
  • DEMO0004 周○雄：每週二 13:00～14:00 基本日常照顧＋陪伴（BA02＋BA20），固定居服員 03816，正常排得進。
        用來示範「臨時異動 → 🩺 個案照顧需求變更 → 新增足部照護（BA08）」：照專複評後新增 BA08，
        03816 沒有足部照護訓練 → 依法規改由有訓練的人（例如 F5005），範例預設「補訓後回任」。
  • DEMO0005 陳○妹：每週二、五 08:00～09:00 基本日常照顧＋陪伴（BA02＋BA20），固定居服員 M04（沒有失智症訓練）。
        用來示範「🩺 個案照顧需求變更 → 新增失智」：照專複評後確認失智 → 依法規改由有失智症訓練的人，範例預設「熟面孔陪同交接」。
再加上你原本就有的 N1112090025（週三 5 小時連續服務）：必須拆成兩段，示範由 M04＋M05 各做一段。
"""
import sys
from copy import copy
import openpyxl

FLOATERS = [('M04', '黃○芳', '週日', 24.9990, 121.5180), ('M05', '林○慧', '週六', 25.0010, 121.5230)]
CASES = [
    # 案號, 姓名, 歷史首選, 星期, 開始, 結束, 分鐘, 服務項目, 說明
    ('DEMO0003', '吳○蘭', 'E6279', '星期一', '12:30:00', '13:20:00', 50, 'BA08(1)', '空檔不夠 → 後移 30 分鐘'),
    ('DEMO0004', '周○雄', '03816', '星期二', '13:00:00', '14:00:00', 60, 'BA02(1) BA20(1)', '03816 長期服務；示範新增足部照護需求'),
    ('DEMO0005', '陳○妹', 'M04', '星期二、星期五', '08:00:00', '09:00:00', 60, 'BA02(1) BA20(1)', 'M04 長期服務；示範新增失智'),
]
LON, LAT = 121.5175, 24.9912


def _headers(ws):
    return {str(c.value).strip(): i for i, c in enumerate(ws[1], start=1) if c.value is not None}


def _append(ws, values, template_row):
    row = ws.max_row + 1
    for col in range(1, ws.max_column + 1):
        src, dst = ws.cell(template_row, col), ws.cell(row, col)
        if src.has_style:
            dst._style = copy(src._style)
    for key, value in values.items():
        ws.cell(row, key).value = value
    return row


def build(src, out):
    wb = openpyxl.load_workbook(src)
    staff, clients, weekly, names_c, names_s = (wb[n] for n in ('員工資料', '個案資料', '週服務計畫', '案家姓名對照', '居服員姓名對照'))
    hs, hc, hw = _headers(staff), _headers(clients), _headers(weekly)
    ids = {str(r[0].value) for r in staff.iter_rows(min_row=2)}
    tpl = next(r for r in range(2, staff.max_row + 1) if str(staff.cell(r, hs['編號']).value) == 'M01') if 'M01' in ids else 2
    for cid, name, rest, lat, lon in FLOATERS:
        if cid in ids:
            continue
        vals = {hs[k]: staff.cell(tpl, hs[k]).value for k in hs}
        vals.update({hs['編號']: cid, hs['休息日']: rest, hs['交通工具']: '機車',
                     hs['經度 (WGS84_Lon)']: lon, hs['緯度 (WGS84_Lat)']: lat})
        for col in ('類別', '開始', '結束', '時間', '時數', '說明'):
            if col in hs:
                vals[hs[col]] = None
        _append(staff, vals, tpl)
        _append(names_s, {1: cid, 2: name}, names_s.max_row)
    # 非失智、一般條件的案家當範本
    wk_rows = {}
    for r in range(2, weekly.max_row + 1):
        wk_rows.setdefault(str(weekly.cell(r, hw['案號']).value), r)
    ctpl = next(r for r in range(2, clients.max_row + 1)
                if str(clients.cell(r, hc['是否失智']).value) == '否' and str(clients.cell(r, hc['內部案號']).value) in wk_rows)
    wtpl = wk_rows[str(clients.cell(ctpl, hc['內部案號']).value)]
    have = {str(clients.cell(r, hc['內部案號']).value) for r in range(2, clients.max_row + 1)}
    for case, name, first, day, start, end, minutes, items, _note in CASES:
        if case in have:
            continue
        vals = {hc[k]: clients.cell(ctpl, hc[k]).value for k in hc}
        vals.update({hc['內部案號']: case, hc['性別']: '女', hc['年齡']: 78, hc['是否失智']: '否', hc['歷史首選']: first,
                     hc['經度 (WGS84_Lon)']: LON, hc['緯度 (WGS84_Lat)']: LAT})
        for col, val in (('核定等級', 2), ('失能程度', '輕'), ('障礙程度', '無'), ('身障證明', None), ('身障手冊', None), ('收入別', '第三類')):
            if col in hc:
                vals[hc[col]] = val
        if '歷史備援' in hc:
            vals[hc['歷史備援']] = None
        _append(clients, vals, ctpl)
        for one_day in day.split('、'):          # 一週多天：每天一列
            w = {hw[k]: weekly.cell(wtpl, hw[k]).value for k in hw}
            w.update({hw['案號']: case, hw['星期']: one_day, hw['開始時間']: start, hw['結束時間']: end, hw['服務時數']: minutes / 60,
                      hw['服務分鐘']: minutes, hw['費用類型']: '基本', hw['頻率']: '每週', hw['生效日']: '2026-04-24', hw['到期日']: None,
                      hw['服務項目']: items, hw['經度 (WGS84_Lon)']: LON, hw['緯度 (WGS84_Lat)']: LAT})
            if '狀態' in hw:
                w[hw['狀態']] = None
            _append(weekly, w, wtpl)
        _append(names_c, {1: case, 2: name}, names_c.max_row)
    wb.save(out)
    return out


if __name__ == '__main__':
    print(build(sys.argv[1], sys.argv[2]))
