# Board1 & Board2 Pin Matching Tool (ECL_Mapping4.py)

一個以 **Python + Tkinter + Pandas + OpenPyXL** 打造的桌面工具，用於將 **兩張板子（Board1 / Board2）** 的
ECL 走線報表（由 `ECL_Tool15.py` 等工具產出的多分頁 Excel）依照一份 **Mapping Table（元件/Pin 對應表）**
進行 **跨板 Pin 對接匹配**，自動找出兩板之間實際相連的訊號路徑，並合併計算 **總長度（Total Length）** 與
**總插入損耗（Total Loss at 8G / 16G）**，最終輸出成一份含 Summary、Config 與各連接分頁的整合 Excel 報表。

適合用於 **Board-to-Board（B2B）/ Connector 對接**、**跨板訊號完整性分析**、**跨板 Loss Budget 加總** 等場景。

---

## 目錄

- [功能特色](#功能特色)
- [系統需求](#系統需求)
- [安裝方式](#安裝方式)
- [使用方式](#使用方式)
- [輸入檔案格式說明](#輸入檔案格式說明)
- [Pin 對接匹配邏輯](#pin-對接匹配邏輯)
- [輸出 Excel 內容說明](#輸出-excel-內容說明)
- [程式架構 / 模組說明](#程式架構--模組說明)
- [常見問題](#常見問題)
- [授權](#授權)

---

## 功能特色

- **跨板 Pin 自動配對**：依據 Mapping Table，將 Board1 的 Start/End Pin 對應到 Board2 對應元件/Pin，自動找出所有可能連接。
- **支援兩種對應層級**：可指定「整顆元件對應」（如 `U1 -> U2`）或「特定 Pin 對應」（如 `U1.A1 -> U2.B3`），特定 Pin 對應優先套用。
- **四種連接方向全覆蓋**：自動比對 Start-to-Start、Start-to-End、End-to-Start、End-to-End 四種可能的接續關係，避免遺漏。
- **自動去重**：同一筆 Board2 資料若透過多種路徑重複比對到，僅保留一筆，避免報表重複計算。
- **自動加總長度與損耗**：自動偵測 `Total Length`、`Loss at 8G`、`Loss at 16G` 相關欄位，計算兩板加總後的 `Total Length`、`Total Loss at 8G`、`Total Loss at 16G`。
- **依元件對分頁**：每組 `Board1 元件 -> Board2 元件` 的對應關係會產生獨立分頁，並以顏色區分 Board1（綠色）、Board2（藍色）、Total（紅色粗體）欄位資料，方便閱讀。
- **Summary 總覽頁（含公式）**：自動生成總覽頁，包含每個分頁的最短/最長 Net、最小/最大 Loss Net 超連結、以及依 Interface Spec 判定的 **Risk Level**（High Risk / Low Risk）。
- **Config 介面規格表**：內建 PCIe4 / PCIe5 / PCIe6 / xGMI / UPI 的 Nyquist frequency、Spec、Overhead 設定，並透過下拉選單套用到 Summary 頁。
- **背景執行緒處理**：實際運算於獨立 Thread 執行，避免處理大檔案時 GUI 視窗凍結，並即時顯示 Log 訊息。
- **自動化 Excel 排版**：自動調整欄寬、加上格線框線、超連結返回 Summary。

---

## 系統需求

| 項目 | 說明 |
|---|---|
| Python 版本 | Python 3.8 以上（建議 3.9+） |
| 作業系統 | Windows / macOS / Linux（需支援 Tkinter GUI） |
| 必要套件 | `pandas`、`openpyxl` |
| 內建套件 | `tkinter`、`os`、`threading`、`collections`（Python 標準庫，無需另外安裝） |

> ⚠️ 讀取 `.xlsx` 檔案需要 `openpyxl` 作為 pandas 的引擎，安裝 `openpyxl` 即可同時滿足讀取與寫入需求。

---

## 安裝方式

1. 確認已安裝 Python 3.8 以上版本：

   ```bash
   python --version
   ```

2. 安裝所需套件：

   ```bash
   pip install pandas openpyxl
   ```

3. 下載 `ECL_Mapping4.py` 至本機任意資料夾。

---

## 使用方式

### 啟動程式

```bash
python ECL_Mapping4.py
```

程式會開啟標題為「Board1 & Board2 Pin Matching Tool」的視窗（尺寸 600x550）。

### 操作步驟

1. **選擇 Board1 File**：點選對應「Browse」按鈕，選取第一張板子的 ECL 走線 Excel 報表（`.xlsx`）。
2. **選擇 Board2 File**：同上，選取第二張板子的 ECL 走線 Excel 報表。
3. **選擇 Mapping Table**：選取用於定義兩板元件/Pin 對應關係的 Excel 檔（格式詳見下方說明）。
4. **點選「Start」**：程式會跳出儲存對話框，指定輸出檔名（預設 `Board_Merged_Report.xlsx`）。
5. 選擇儲存路徑後，程式會於背景執行緒開始比對，並於下方 Log 區即時顯示進度訊息（如載入筆數、找到的連接數量等）。
6. 完成後會跳出「Success」訊息框並顯示輸出摘要（Summary / Config / N 個資料分頁）；若失敗則顯示錯誤訊息。

---

## 輸入檔案格式說明

### Board1 File / Board2 File（`.xlsx`，必要）

- 需為多分頁 Excel 檔（例如 `ECL_Tool15.py` 產出的走線報表）。
- 程式會自動忽略名為 `Summary`、`Config`、`loss data` 的分頁，其餘分頁皆視為資料分頁讀取。
- 每個資料分頁的表頭需位於**第 2 列**（`header=1`），且必須包含 `Start Pin` 欄位，否則該分頁會被跳過。
- 需包含 `End Pin` 欄位以支援雙向比對；若欄位名稱含前後空白，程式會自動去除。

### Mapping Table（`.xlsx`，必要）

- 至少需包含 **2 欄**：第 1 欄為來源（Board1 端）、第 2 欄為目標（Board2 端）。
- 對應規則依內容格式分為兩種：

| 內容格式 | 對應層級 | 範例 |
|---|---|---|
| 不含 `.`（純元件名稱） | 元件層級對應（Component-level） | `U1` → `U2` |
| 含 `.`（元件.Pin 名稱） | 特定 Pin 對應（Pin-level，優先套用） | `U1.A1` → `U2.B3` |

- 同一元件可對應多個目標元件（程式以 list 儲存，會逐一嘗試比對）。

---

## Pin 對接匹配邏輯

1. 針對 Board1 每一筆資料的 `Start Pin` 與 `End Pin`，先查詢是否有 **特定 Pin 對應**（`pin_map`），若無則查詢其 **元件是否有對應規則**（`comp_map`），產生一組或多組「目標 Pin 字串」。
2. 將每個目標 Pin 字串，分別與 Board2 依 `Start Pin`、`End Pin` 建立的索引表進行比對，找出所有可能相符的 Board2 資料列。
3. 比對結果會標記四種連接類型之一：`Board1_Start_to_Board2_Start`、`Board1_Start_to_Board2_End`、`Board1_End_to_Board2_Start`、`Board1_End_to_Board2_End`。
4. 同一筆 Board2 資料若被多條路徑重複比對到，僅保留一次（以資料列的物件 id 去重）。
5. 依「來源元件 -> 目標元件」（如 `U1->U2`）分組，將 Board1 與 Board2 對應資料列以欄位前綴 `Board1_` / `Board2_` 合併成一筆完整記錄。

---

## 輸出 Excel 內容說明

輸出的 `.xlsx` 檔案包含以下工作表：

| 工作表 | 內容 |
|---|---|
| **Summary** | 總覽頁，列出每組元件對應分頁的 Interface（可下拉選擇）、Net 數量、最短/最長 Net 與長度、最小/最大 Loss Net、以及自動判定的 Risk Level |
| **Config** | PCIe4 / PCIe5 / PCIe6 / xGMI / UPI 的 Nyquist frequency（GHz）、Spec、Overhead 設定表，供 Summary 頁公式查表使用 |
| **各元件對分頁**（如 `U1->U2`） | 該組元件對應下所有匹配到的連接記錄，包含 Board1 與 Board2 原始欄位、連接類型（Connection_Type）、以及自動計算的 `Total Length`、`Total Loss at 8G`、`Total Loss at 16G` |

### 資料分頁欄位配色

- **綠色文字**：來自 Board1 的欄位（欄位名前綴 `Board1_`）。
- **藍色文字**：來自 Board2 的欄位（欄位名前綴 `Board2_`）。
- **紅色粗體文字**：加總計算欄位（`Total Length`、`Total Loss at 8G`、`Total Loss at 16G`）。

### Summary 頁公式與 Risk Level

- 使用 `HYPERLINK` + `INDEX/MATCH` 公式自動連結到該分頁中長度最短/最長、Loss 最小/最大的 Net。
- 依所選 Interface 的 Nyquist frequency（≤8GHz 或 >8GHz）自動切換參照 `Loss at 8G` 或 `Loss at 16G` 欄位。
- **Risk Level** 判斷邏輯：若最大 Loss 超過 `(Spec - Overhead)` 門檻，標示為 `High Risk`，否則為 `Low Risk`；若資料無法判讀則顯示 `Data Error` 或 `No Data`。

---

## 程式架構 / 模組說明

以下依處理流程列出主要函式（Function）／類別用途：

- `auto_adjust_column_width(ws)` — 自動依內容長度調整 Excel 欄寬（中文字元以雙倍寬度計算，並跳過合併儲存格）。
- `add_borders_to_sheet(ws)` — 為所有有內容的儲存格加上細框線。
- `load_mapping_table(file_path, log_func)` — 讀取 Mapping Table，拆分成元件層級對應表（`comp_map`）與特定 Pin 對應表（`pin_map`）。
- `parse_pin(pin_str)` — 將 `元件.Pin` 格式字串拆解為 `(元件, Pin)` tuple。
- `load_all_sheets(file_path, log_func)` — 讀取 Excel 檔案所有資料分頁（排除 Summary/Config/loss data），合併為單一 DataFrame。
- `create_config_sheet(wb)` — 建立內建 Interface 規格對照表（Config 分頁）。
- `create_summary_sheet(wb, summary_info_list)` — 建立含公式、超連結與 Risk Level 判定的 Summary 總覽頁。
- `run_processing(file_board1, file_board2, file_mapping, output_path, log_func, finish_callback)` — 主流程函式：讀檔、建立索引、執行 Pin 對接比對、計算 Total Length/Loss、寫出各分頁與 Summary/Config，最後儲存 Excel。
- `App`（class，繼承 `tk.Tk`） — 主視窗類別，負責建立 GUI 版面（檔案選擇欄、Start 按鈕、Log 顯示區）與事件綁定。
  - `setup_ui()` — 建立視窗內所有元件版面。
  - `add_file_row(parent, label, var)` — 建立單一「標籤 + 輸入框 + Browse 按鈕」的檔案選擇列。
  - `browse(var)` — 開啟檔案選擇對話框並更新對應路徑變數。
  - `log(s)` — 於畫面 Log 區域附加訊息並自動捲動到底部。
  - `run()` — 檢查輸入完整性、選擇輸出路徑，並啟動背景執行緒呼叫 `run_processing`。
  - `done(success, msg)` — 背景執行緒完成後的回呼，依結果顯示成功或錯誤訊息框。

---

## 常見問題

**Q: 執行後顯示「No valid data sheets found」？**
A: 表示 Board1 或 Board2 檔案中沒有任何分頁包含 `Start Pin` 欄位，或表頭未位於第 2 列，請確認輸入檔案格式（例如是否為 `ECL_Tool15.py` 產生的標準格式報表）。

**Q: Mapping Table 需要包含表頭嗎？**
A: 需要，程式讀取時會使用第一列作為欄位名稱（`header=0`），第 1、2 欄的內容才是實際對應資料。

**Q: 為什麼某些連接沒有出現在輸出檔案中？**
A: 請確認 Mapping Table 是否正確設定該元件（或該 Pin）的對應關係；若 Board1 與 Board2 皆無法透過 `comp_map` 或 `pin_map` 找到目標 Pin，該筆資料將不會被匹配。

**Q: 「No matching data found」是什麼意思？**
A: 表示依據目前的 Mapping Table，未能在 Board1 與 Board2 之間找到任何符合的 Pin 對接關係，請確認對應表內容與板子的 Pin 命名是否一致。

**Q: Total Loss 欄位為什麼沒有出現？**
A: 程式會自動搜尋欄位名稱中含有「Board1/Board2」與「Loss at 8G/16G」的欄位；若原始資料分頁中沒有對應的 Loss 欄位，則不會產生該項加總欄位。

**Q: GUI 處理大檔案時會不會卡住？**
A: 實際運算已放在獨立 Thread（`threading.Thread`）中執行，主視窗與 Log 區域仍可即時更新，不會完全凍結。

---

## 授權

本 README 為根據所提供程式碼 `ECL_Mapping4.py` 反推生成之技術文件，實際授權條款請依專案內部規範或使用者自行補充。
