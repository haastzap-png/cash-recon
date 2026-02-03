# 現金對帳表（Web App）

目的：使用者上傳報表後，自動產出「現金對帳表」(xlsx)，並列出漏結帳清單。

特性：App **不保存上傳檔**（處理在記憶體中完成，不會把原始報表寫入伺服器磁碟），也不提供任何「瀏覽歷史上傳檔」的功能。

## 支援的輸入
- Hotcake：帳單紀錄（含 `服務`、`儲值金` 分頁）
- Hotcake：訂單/預約報表（含 `訂單報表` 分頁）
- 收銀機：歷史訂單（選配；上傳後會核對現金是否一致）

系統會嘗試自動辨識檔案類型，並對欄位標題做容錯（例如去空白/同義欄名）。

## 輸出（現金對帳表 xlsx）
- `Summary`：區間、Hotcake 現金合計、（若有）收銀機現金合計、差額、漏結帳筆數與可否視為正確對帳
- `MissingBills`：已報到但帳單金額空/無帳單編號的明細（店、日期、師傅、訂單）
- `HotcakeBills_Service`：區間內（依訂單日期時間）對應到的帳單現金明細
- `HotcakeBills_Topup`：區間內（依結帳操作時間）儲值金現金明細

## 本機啟動（macOS）
```bash
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

## CLI（可用於排程/批次）
```bash
python3 -m cash_recon.cli \
  --store '中壢三光店' \
  --start '2026-01-01 00:00' \
  --end '2026-01-31 23:59' \
  --hotcake-bills '/path/to/帳單紀錄.xlsx' \
  --hotcake-orders '/path/to/訂單報表.xlsx' \
  --pos-orders '/path/to/歷史訂單.xlsx' \
  --out 'output/spreadsheet/cash_recon.xlsx'
```

## 雲端部署（手機 iOS 可用）
這個專案是純 Web App（使用者用瀏覽器上傳/下載），部署到第三方主機後 iPhone/iPad 可直接使用。

### 最省事：Streamlit Community Cloud（免費）
1. 把此資料夾推到 GitHub（公有 repo 最方便分享連結）。
2. 到 Streamlit Community Cloud 建立新 App，選 repo 與 `app.py` 作為入口。
3. Deploy（會自動用 `requirements.txt` 安裝套件）。

### 其他平台
- Cloud Run / Render 等也能跑（可用本專案的 `Dockerfile`），啟動命令：
  - `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
