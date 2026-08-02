# ECL Report Combiner (ECL_Tool15.py)

一個以 **Python + Tkinter** 打造的桌面小工具，用於解析 Signal Integrity (SI) 領域常見的
**ECL (Electrical Constraint / Length) 報告** (`.txt`)，並可選擇性地整合 **走線 HTML 報告**
(Trace Length by Width Report)，最終輸出成一份格式化、含條件式格式與計算公式的 **Excel (.xlsx) 彈性報表**。

適合用於 PCIe / xGMI / UPI 等高速介面的 **Insertion Loss 計算**、**Via 數量統計**、**Neck-down（線寬縮頸）長度統計** 等 Signal Integrity 前處理工作。

---

## 目錄

- [功能特色](#功能特色)
- [系統需求](#系統需求)
- [安裝方式](#安裝方式)
- [使用方式](#使用方式)
- [輸入檔案格式說明](#輸入檔案格式說明)
- [輸出 Excel 內容說明](#輸出-excel-內容說明)
- [Neckdown（線寬縮頸）功能說明](#neckdown線寬縮頸功能說明)
- [程式架構 / 模組說明](#程式架構--模組說明)
- [常見問題](#常見問題)
- [授權](#授權)

---

## 功能特色

- **多檔案批次匯入**：可同時載入多個 ECL `.txt` 報告，程式會自動合併同一組 Net。
- **自動解析 Pin / RefDes / Via / 分段長度**：從原始 ECL 文字報告中提取每段走線長度、層別 (Layer) 與 Via 數量。
- **被動元件（R/L/C）自動合併**：偵測經過同一顆電阻/電感/電容（如 `_C_`、`_R_`）的兩段 Net，自動合併為單一完整路徑，還原真實訊號路徑。
- **依 RefDes pair 分組建立工作表**：例如 `U1_U2`，每組元件對會產生獨立分頁，方便逐一檢視。
- **PCB Loss 對照表**：內建 Dell's Material Category（Mid loss / Low loss 系列）在 8GHz / 16GHz 下的典型 Loss 數值，供插入損耗估算參考。
- **可自訂 Interface Config**：內建 PCIe4 / PCIe5 / PCIe6 / xGMI / UPI 的預設頻率、Spec、Overhead 設定表，並支援下拉選單驗證。
- **Insertion Loss 自動計算欄位**：針對 PE / XGMI / UPI 開頭的 Net，自動加入 Loss 計算相關欄位與公式。
- **Neck-down（縮頸長度）統計（選用）**：可載入額外的走線寬度 HTML 報告，統計低於指定線寬門檻的走線長度，按層別加總。
- **自動化 Excel 排版**：自動調整欄寬、加上格線框線、Summary 頁超連結回跳等，輸出即可直接使用，不需再手動排版。
- **單位自動轉換**：輸入資料若為 mm，會自動換算為 mil（1 mm = 39.3700787 mil）。
- **圖形化介面（GUI）**：不需寫程式或下指令，透過視窗介面即可完成檔案選取、參數輸入與報表匯出。

---

## 系統需求

| 項目 | 說明 |
|---|---|
| Python 版本 | Python 3.8 以上（建議 3.9+） |
| 作業系統 | Windows / macOS / Linux（需支援 Tkinter GUI） |
| 必要套件 | [`openpyxl`](https://openpyxl.readthedocs.io/)（用於產生 Excel） |
| 內建套件 | `tkinter`、`re`、`os`、`typing`、`html.parser`（Python 標準庫，無需另外安裝） |

> ⚠️ 若未安裝 `openpyxl`，程式啟動 GUI 後會跳出提示視窗，告知需先安裝該套件才能使用 Excel 匯出功能。

---

## 安裝方式

1. 確認已安裝 Python 3.8 以上版本：

   ```bash
   python --version
   ```

2. 安裝所需套件：

   ```bash
   pip install openpyxl
   ```

3. 下載 `ECL_Tool15.py` 至本機任意資料夾。

---

## 使用方式

### 啟動程式

在終端機 / 命令提示字元中執行：

```bash
python ECL_Tool15.py
```

程式會開啟一個標題為「ECL Report Combiner」的視窗（尺寸約 820x520）。

### 操作步驟

1. **加入 ECL 報告檔案**
   - 點選「Add files...」按鈕，選擇一個或多個 `.txt` 格式的 ECL 報告。
   - 可用「Remove selected」移除選取檔案，或「Clear」清空整份清單。

2. **（選用）加入 HTML 走線寬度報告**
   - 點選「Browse...」選擇 `.htm` / `.html` 檔案（用於 Neck-down 統計）。
   - 若不需要此功能，保持空白即可，或按「Clear」清除已選路徑。

3. **（選用）設定 Neckdown 門檻**
   - 於「Neckdown threshold」欄位輸入門檻值，支援格式：`4`、`4mil`、`4mils`、`0.1mm`。
   - 若只設定門檻但未選擇 HTML 檔案，程式會跳出提示並自動略過 Neck-down 計算。

4. **設定輸出檔案路徑**
   - 預設輸出檔名為 `ecl_reports_combined.xlsx`，可點選「Browse...」自訂儲存路徑與檔名。

5. **執行**
   - 點選右下角「Run」按鈕開始處理。
   - 處理過程中按鈕會暫時停用並顯示「Running... please wait.」。
   - 完成後會跳出「Completed」訊息框，並顯示輸出檔案路徑；若失敗則顯示錯誤訊息。

---

## 輸入檔案格式說明

### ECL 報告（`.txt`，必要）

程式解析的每一行資料格式大致為：

```
<Pin/Via 名稱>  <X座標>  <Y座標>  [L/B/D/V]  <累計長度>  [Layer/其他資訊]
```

- 檔案中會忽略頁首、頁尾、分隔線（如以 `|`、`ECL `、`Page`、`C:/`、`dimensions`、`refdes` 等開頭的行）。
- 每個 Net 區塊以 Net 名稱行開始，並以 `TOTAL <n> VIA(S) <length> mils/millimeters` 結尾行結束，作為該 Net 的分段標記。
- 支援長度單位為 `mils` 或 `millimeters`（`mm`），程式會自動判斷並統一換算為 mil。

### HTML 走線寬度報告（`.htm` / `.html`，選用）

- 需包含表頭含有「Net Name」與「Layer Name」欄位的表格（僅解析檔案中第一個符合條件的表格）。
- 欄位需包含：Net 名稱、層別（Layer）、線寬（Line Width）、對應線寬下的走線長度（Length at Width）。
- 支援單位為 mil 或 mm（程式會依表頭文字自動偵測並換算）。

---

## 輸出 Excel 內容說明

輸出的 `.xlsx` 檔案主要包含以下工作表：

| 工作表 | 內容 |
|---|---|
| **Summary** | 總覽頁，列出各 RefDes pair 分頁的統計摘要與跳轉連結 |
| **PCB loss** | Dell's Material Category（Mid loss / Low loss 1 / Low loss 2...）於 8GHz、16GHz 下的典型 Loss 對照表 |
| **Config** | PCIe4 / PCIe5 / PCIe6 / xGMI / UPI 等介面預設頻率（freq）、規格門檻（spec）、Overhead 設定表，並提供下拉選單資料驗證 |
| **各 RefDes Pair 分頁**（如 `U1_U2`） | 每組元件對之間的走線分段明細，含層別、長度、Via 數量；若符合 PE/XGMI/UPI 條件則額外附上 Loss 計算欄位與公式 |

輸出格式特色：

- 每個分頁 A1 儲存格提供「Return to Summary」超連結，方便快速跳回總覽頁。
- 所有含資料的儲存格自動加上細框線（thin border）。
- 所有欄位依內容長度自動調整寬度（中日文字元以雙倍寬度計算）。
- 分頁名稱超過 31 字元（Excel 限制）會自動截斷，並移除非法字元（`[ ] : * ? / \`）。

---

## Neckdown（線寬縮頸）功能說明

Neck-down 是指走線在特定區段因設計限制（如連接器 Pin 間距、BGA Escape）而**局部縮小線寬**的現象，通常會增加插入損耗。此工具可協助統計：

1. 每條 Net 在走線寬度報告中，**線寬小於或等於設定門檻**的走線長度。
2. 依 **層別（Layer）** 加總這些縮頸長度，回傳如 `{層別: 縮頸總長(mil)}` 的統計結果。

啟用條件：

- 必須同時提供 **HTML 走線寬度報告** 與 **Neckdown 門檻值**，兩者缺一則此功能自動略過（不會產生 Neck_down 相關欄位）。
- 門檻輸入格式範例：`4`（預設 mil）、`4mil`、`4mils`、`0.1mm`。

---

## 程式架構 / 模組說明

以下依處理流程列出主要函式（Function）用途：

- `get_refdes(pin_name)` — 從 Pin 名稱解析出對應的元件 RefDes（排除 Via 與 T 特殊標記）。
- `is_component_pin(name)` — 判斷名稱是否為一般元件 Pin（非 Via、非 T）。
- `parse_file(path)` — 解析單一 ECL `.txt` 檔，切分出每個 Net 及其分段資料。
- `convert_net_to_segments(net)` — 將 Net 的累計長度資料轉換為各段獨立長度（segment），並統計 Via 數量。
- `merge_passive_nets(rows)` — 自動偵測並合併經過同一被動元件（R/L/C）的兩段 Net，還原完整訊號路徑。
- `collect_and_sort_layers(grouped_data)` / `layer_sort_key(...)` — 收集所有出現過的層別並依規則排序，用於建立分頁欄位。
- `parse_trace_html_to_map(html_path)` — 解析 HTML 走線寬度報告，回傳依 Net 分類的線寬/長度資料。
- `parse_neckdown_input(s)` — 解析使用者輸入的 Neckdown 門檻字串（支援 mil / mm 格式）。
- `calc_neckdown_layer_map(html_rows, threshold)` — 依門檻值計算各層別的縮頸總長度。
- `create_pcb_loss_sheet(wb)` — 建立內建 PCB Loss 對照表工作表。
- `should_have_loss_calc(net_name)` — 判斷 Net 名稱是否需要附加 Insertion Loss 計算欄位（PE / XGMI / UPI 開頭）。
- `combine_to_excel(...)` — 主流程函式，整合以上所有步驟並輸出最終 Excel 檔案。
- `auto_adjust_column_width(ws)` / `add_borders_to_sheet(ws)` — Excel 排版輔助函式，自動調整欄寬與加上框線。
- `SimpleTableParser`（class） — 繼承 `HTMLParser`，用於解析 HTML 報告中的第一個表格。
- `run_gui()` — 建立並啟動整個 Tkinter GUI 主視窗與事件邏輯。

---

## 常見問題

**Q: 執行後直接顯示「Missing Library」錯誤？**
A: 尚未安裝 `openpyxl`，請執行 `pip install openpyxl` 後重新啟動程式。

**Q: 為什麼設定了 Neckdown 門檻卻沒有出現 Neck_down 欄位？**
A: Neck-down 功能需同時提供 HTML 走線寬度報告與門檻值。若只設定門檻但未選擇 HTML 檔案，程式會提示並自動略過該功能。

**Q: 分頁名稱看起來被截斷或跟原本 RefDes 不完全一致？**
A: Excel 分頁名稱上限為 31 字元，且不可包含 `[ ] : * ? / \` 等字元，程式會自動截斷與替換以符合限制。

**Q: 支援哪些介面類型的 Insertion Loss 自動計算？**
A: 目前內建 PCIe4、PCIe5、PCIe6、xGMI、UPI 五種介面的預設頻率、規格與 Overhead 參數，Net 名稱以 `PE`、`XGMI`、`UPI` 開頭者會自動套用計算欄位。

**Q: 可以在無 GUI 環境（如遠端伺服器）執行嗎？**
A: 目前程式僅提供 GUI 操作模式（`run_gui()`），若當前環境無法建立 Tkinter 視窗，程式會捕捉 `TclError` 並印出「GUI not available.」，此工具目前不支援純命令列（CLI）模式。

---

## 授權

本 README 為根據所提供程式碼 `ECL_Tool15.py` 反推生成之技術文件，實際授權條款請依專案內部規範或使用者自行補充。
