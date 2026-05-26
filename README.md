# 票價比對工具

票價比對系統，支援安源 `.xls` 與廠商 `.xlsx` 的票價核對，提供兩種模式：固定格式版與彈性格式版。

---

## 包含工具

| 工具 | 主程式 | 說明 |
|------|--------|------|
| **固定格式比對** | `app.py` | 廠商 Excel 格式固定（標題列固定在第 3 列），需登入 |
| **彈性格式比對** | `flexible_app.py` | 廠商 Excel 格式不限，自動偵測欄位 + 使用者確認 |

---

## 目錄結構

```
ticket-compare/
├── README.md               # 本文件
├── requirements.txt        # Python 套件清單
├── app.py                  # 固定格式比對（原版）
├── flexible_app.py         # 彈性格式比對（新版）
├── compare.py              # 共用：比對邏輯 + Excel 報告產生
└── ticket_parser.py        # 彈性版：Excel 解析器（自動偵測欄位）
```

---

## 功能說明

### 固定格式比對（`app.py`）

> 適用廠商 Excel 格式固定、標題列在第 3 列的情境。

**流程：**
1. 登入（帳號密碼）
2. 上傳安源 `.xls` + 廠商 `.xlsx`
3. 選擇場次（安源日期 × 廠商票價頁籤）
4. 執行比對，查看結果，下載 Excel 報告

**廠商 Excel 格式要求：**
- 頁籤名稱需包含「票價」
- 第 3 列（row index 2）為票種標題列
- 第 2 欄（col index 1）為區域名稱

---

### 彈性格式比對（`flexible_app.py`）

> 適用廠商 Excel 格式不固定的情境，無需登入。

**流程：**
1. 上傳安源 `.xls` + 廠商 `.xlsx`
2. 選擇廠商頁籤 → 查看自動偵測結果
3. 確認或調整欄位指定（區域欄、票種欄、票價欄）
4. 執行比對，查看結果，下載 Excel 報告

**自動偵測機制：**
- 掃描前 20 列，找最多關鍵字（區域、票種、票價）的列作為標題列
- 依關鍵字比對自動指定欄位，使用者可手動修正
- 若區域有合併儲存格，自動向下填補

---

## 比對結果說明

| 類別 | 說明 |
|------|------|
| ✅ 票價完全相符 | 區域 + 票種 + 票價三者完全一致 |
| ❌ 票價不符 | 區域 + 票種相同但票價不同 |
| ⚠️ 廠商有、安源沒有 | 廠商有此票種，安源系統無對應資料 |
| ⚠️ 安源有、廠商沒有 | 安源有此票種，廠商 Excel 無對應資料 |
| ℹ️ 安源票種無對應 | 安源票種名稱不在票種對照表（`TICKET_MAP`）中 |

---

## 票種對照表（TICKET_MAP）

安源票種名稱與廠商票種的標準化對應，定義於 `compare.py`：

| 安源票種 | 標準化名稱 |
|----------|------------|
| 全票 | 全票 |
| 貴賓券 | 貴賓券 |
| 公關票 | 公關票 |
| 信用卡優惠票 | 信用卡優惠票 |
| 啦啦隊票 | 啦啦隊票 |
| 眷屬票 | 眷屬票 |
| 內野優惠票 | 內野優惠票 |
| 內野身心優惠票 | 身心優惠票 |
| 外野身心優惠票 | 身心優惠票 |
| 內野半票 | 半票 |
| 外野半票 | 半票 |
| 統一獅會員優惠票 | 會員優惠票(折100) |

> 如需新增票種對應，修改 `compare.py` 中的 `TICKET_MAP` 字典。

---

## 本地運行

### 環境需求

- Python 3.10+

### 安裝套件

```bash
pip install -r requirements.txt
```

### 啟動固定格式版

```bash
streamlit run app.py
```

需建立 `.streamlit/secrets.toml`（不推 GitHub）：

```toml
[auth]
username = "your_username"
password = "your_password"
```

### 啟動彈性格式版

```bash
streamlit run flexible_app.py
```

---

## Streamlit Cloud 部署

### 固定格式版（已部署）

- **Repository：** `lard23chen/ticket-compare`
- **Branch：** `main`
- **Main file：** `app.py`
- **Secrets：** 需在 Streamlit Cloud 設定 `[auth]` 區塊

### 彈性格式版

- **Repository：** `lard23chen/ticket-compare`
- **Branch：** `main`
- **Main file：** `flexible_app.py`
- **Secrets：** 不需要

---

## 技術棧

| 類別 | 套件 |
|------|------|
| UI 框架 | [Streamlit](https://streamlit.io/) |
| 資料處理 | [pandas](https://pandas.pydata.org/) |
| Excel 讀取（.xlsx） | [openpyxl](https://openpyxl.readthedocs.io/) |
| Excel 讀取（.xls） | [xlrd](https://xlrd.readthedocs.io/) |

---

## 注意事項

- 安源 `.xls` 格式固定，兩個工具均使用相同的 `parse_ansource()` 解析
- 彈性版的**票種欄為選填**，若廠商 Excel 無票種欄位，可不指定，系統預設以「全票」代入
- 比對以「區域名稱 + 標準化票種」為鍵值進行 outer join
- Excel 報告包含摘要頁與各類別詳細頁，以色彩標示結果（綠/紅/橘/藍）
