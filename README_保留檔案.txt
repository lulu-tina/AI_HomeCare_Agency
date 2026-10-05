CareFlow 最終精簡版

專案根目錄只需要保留：
1. app.py
2. caregiver_engine.py
3. care_policy.py
4. calendar_view.py
5. requirements.txt
6. packages.txt
7. .gitignore
8. .streamlit/secrets.toml   ← 你自己的正式金鑰檔，不要上傳 GitHub

可選：
- 00_DB/你的 Excel：只有你想讓程式預設讀本機資料時才留；平常用上傳 Excel 可不留。
- 測試檔：如果你還要跑 pytest 才留。

可以刪：
- apply_*.py
- fix_*.py
- repair_*.py
- app.py.bak_*
- caregiver_engine.py.bak_*
- 舊版 app(數字).py / caregiver_engine(數字).py
- requirements_agency.txt（已整合為 requirements.txt）

啟動：
py -m pip install -r requirements.txt
py -m py_compile app.py caregiver_engine.py care_policy.py calendar_view.py
py -m streamlit run app.py

速度版重點：
- 機車：OSRM /table 批次快取
- 大眾運輸：Google ComputeRouteMatrix 批次 + 平行請求 + 同次排班 memo cache
- Phase 1：大眾運輸只查硬限制通過後最近 5 位候選
- Phase 2：每個任務最多保留 10 位最高分候選再做 OR-Tools
- 第一週完整 AI；確認後固定週期直接延展；新週期首次才 AI

v2 排班速度修正：
- 不再於 Phase 2 對所有大眾運輸候選任務做 O(n²) Google 查詢。
- 先求解，再只對真正已選中的大眾運輸相鄰轉場用 Google Transit 驗證。
- 若 Google 驗證發現轉場不足，自動加限制後重新求解，最多 4 輪。
- 大眾運輸候選 shortlist 4 人；Phase 2 每任務最多 6 名候選。

UltraFast v3：
- Phase 1：Google Transit 用 15 分鐘時槽批次化，降低 HTTP request 數；仍以 Google 回傳時間評分。
- 每個任務的大眾運輸只保留最近 3 位硬限制合格候選。
- Phase 2：不再對所有候選任務兩兩查 Google；先求解，只驗證真正已派的相鄰轉場。
- 驗證衝突時自動加限制重解，最多 3 輪。
- 快速模式 OR-Tools 3 秒；極速 1 秒；完整 15 秒。

UltraFast v3.1 修正：
- 修正 Phase 2 大眾運輸 busy_after / busy_before 檢核 tuple 解包錯誤。
- 不改變 v3 的速度優化邏輯。
