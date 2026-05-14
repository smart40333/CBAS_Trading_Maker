# 2026-04-17 Update Summary

## 1. VIP名單：新增「80元手續費」欄位

### 影響檔案 / 位置
- `main.py:428` — `vip_list_columns` 加入 `80元手續費`
- `main.py:1118-1137` `loading_vips()` — 載入時若該欄=Y → 儲存格上淺藍色 `QColor(204,229,255)`
- `file_reader.py:67-86` `read_vip_list()` — 欄位向後相容（舊 CSV 自動補空欄）
- `file_reader.py:173-208` `load_vip_data()` — 同上
- `bargaining.py:298-316` `calculate_new_trade_batch()` — 手續費判斷新增 `80元手續費=Y → 80`

### 手續費優先序（由高到低）
1. 特殊報價（VIP_Quote）
2. **80元手續費=Y → 80**（新增）
3. 不限張數低手續費=Y → 100
4. 特殊ID（H122699830 等）→ 60
5. 基礎規則（成交張數≥10 或 STORQTY≥200 → 100；SRC=E → 110；其他 → 150）

---

## 2. 讀取暫存檔時自動套用檢查上色

### 影響檔案 / 位置
- `main.py` `temp_load_all` — 資料填回表格後，依序呼叫三個 check 函式（silent）：
  - `check_buy_table_with_quote(silent=True)` → 買進百元價綠/紅
  - `check_buy_table_with_qty(silent=True)` → 買進 CB 不符紅
  - `check_sell_table_with_qty(silent=True)` → 賣出重複橙、CB 不符紅
- `check_buy_table_with_quote` / `check_buy_table_with_qty` / `check_sell_table_with_qty` 新增 `silent` 參數：跳過 QMessageBox，但保留全部上色

---

## 3. CB自營系統不符：顯示明細 + 紅色標記

### 影響檔案 / 位置
- `main.py` `check_buy_table_with_qty` — 不符時依 `(客戶ID, CB代號)` vs `(客戶代碼, 商品代碼)` outer merge 找差異
- `main.py` `check_sell_table_with_qty` — 同上（sell 側用 `履約張數` + `賣出金額/股數`）

### 行為
- 不符時：彈窗列出每一筆不符的 `客戶ID / CB代號: 中台=X張,Y元  CB自營=A張,B元  (差 Δ張, Δ元)`（前 30 筆）
  - 「中台」= AP 側（本系統 table_buy / table_sell 的盤面交易列）
  - 「CB自營」= CB自營軟體資料庫（`RPT_Monitor_Fill`）
  - outer merge：CB 有、中台沒有的 key 也會列出（中台側顯示為 0）；差為負數代表 CB 比較多
- 盤面交易列上紅色 `QColor(255,180,180)`；紅色優先於賣出的重複橙色
  - 注意：「CB 有、中台沒有」的 key 因為中台表格無對應列，不會被紅色標記，只會出現在彈窗
- silent 模式只上色不彈窗

---

## 4. 特殊報價：新增「張數」「備註」+ 拆單與自動扣減

### 影響檔案 / 位置
- `main.py:468` — `vip_quote_columns` 加入 `張數`, `備註`
- `main.py:283` — `buy_columns` 加入 `VIP報價張數`（追蹤每列消耗了幾張 VIP）
- `main.py` `show_buy_table` — `required_columns_order` 加入 `VIP報價張數`
- `main.py` `generate_buy_upload_file` — 產檔成功後呼叫扣減
- `main.py` `_deduct_vip_quote_after_upload` — 新增，處理扣減/備註/刪列/寫回CSV/刷UI
- `file_reader.py` `read_vip_quote` / `load_vip_data` — 向後相容補欄
- `bargaining.py:280-343` — 重寫「1. 檢查特殊報價」，支援張數上限與拆單

### 計算邏輯

| VIP張數 | 成交張數 | 行為 |
|--------|---------|------|
| 空/NaN | 任意 | 全量套 VIP（舊行為） |
| 0 或非數值 | 任意 | 不套 VIP |
| 正數 N | ≤ N | 全量套 VIP，`VIP報價張數=成交張數` |
| 正數 N | > N | **拆兩列**：N張套VIP；(成交-N)張走原邏輯 |

### 扣減流程（按下「產生上傳檔」成功後）
1. 依 `(客戶ID, CB代號)` 加總 `table_buy.VIP報價張數`
2. 讀取當下 VIP_Quote.csv（fresh）
3. 對每組 key：扣 `min(消耗, 原張數)`
4. 備註**追加** `於{YYYYMMDD}扣除了{X}張`（以 `; ` 分隔舊備註）
5. 新張數 ≤ 0 → 整列刪除
6. 寫回 CSV → 刷新 `table_vip_quote`
7. 扣減失敗不影響產檔流程

---

## 檔案變更清單
- `main.py`
- `bargaining.py`
- `file_reader.py`

## 影響的資料檔（欄位升級，向後相容）
- `\\10.72.228.112\...\CBAS_Trading_Maker\VIP_List.csv` — 新增欄位 `80元手續費`
- `\\10.72.228.112\...\CBAS_Trading_Maker\VIP_Quote.csv` — 新增欄位 `張數`, `備註`
