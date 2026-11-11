# 東海實習職缺一覽（每日自動更新）

每天台灣時間早上 8 點，GitHub 會自動執行 `thu_internships.py`，抓取企管、國貿、財金、經濟系、實習中心與 GIFT 的實習職缺，更新 `index.html`。網頁網址固定，網站上的按鈕連到這個網址即可。

## 檔案說明
| 檔案 | 用途 |
|---|---|
| `thu_internships.py` | 抓取程式 |
| `manual_overrides.csv` | 手動補充表（網站沒寫截止日時，在這裡補） |
| `eligibility_notes.csv` | 報名限制表（其他系職缺限本系學生時在這裡填，網頁會標 ❌／⚠️；表裡沒有的國貿、財金、經濟系職缺，程式會讀內文自動判斷） |
| `index.html` | 自動產生的網頁（不要手動改） |
| `internships.csv` | 自動產生的資料表，可用 Excel 開 |
| `.github/workflows/update.yml` | 每日排程設定 |

## 第一次設定（約 10 分鐘）
1. 註冊 GitHub 帳號。建議用畢專或系辦共用的信箱，或建立一個 Organization，方便之後交接。
2. 右上角「＋」→「New repository」，名稱例如 `thu-internships`，選 **Public**，按 Create。
3. 在新 repository 頁面按「uploading an existing file」，把這個資料夾裡的**所有檔案和 `.github` 資料夾**一起拖進去，按「Commit changes」。
   - Windows 檔案總管預設看不到 `.github` 資料夾：在檔案總管上方「檢視」→ 勾選「隱藏的項目」。
4. 「Settings」→「Pages」→ Source 選「Deploy from a branch」，Branch 選 `main`、資料夾 `/ (root)`，按 Save。
5. 「Actions」分頁 → 左邊點「每日更新實習職缺」→ 右邊「Run workflow」，手動跑第一次（約 2～5 分鐘）。
6. 跑完後回到「Settings」→「Pages」，上方會顯示網址，例如 `https://帳號.github.io/thu-internships/`。把 WordPress 按鈕的連結改成這個網址。

## 平常維護
- **補截止日**：在 GitHub 網頁上點 `manual_overrides.csv` → 右上鉛筆圖示 → 修改 → Commit。隔天早上 8 點自動套用，想馬上更新就去 Actions 按「Run workflow」。
- **看有沒有跑成功**：Actions 分頁，綠色勾勾是成功，紅色叉叉是失敗（點進去看錯誤訊息）。失敗時 GitHub 會寄信通知帳號擁有者。
- **某個系網站改版**抓不到資料時，程式會略過那個系，其他系照常更新，這時需要修改 `thu_internships.py`。

## 交接給下一屆
- 方法 1：「Settings」→ 最下方「Transfer ownership」，把 repository 轉給新負責人的帳號。網址會變成新帳號的，記得改 WordPress 按鈕連結。
- 方法 2（推薦）：一開始就建在 Organization 底下，交接時只要在 Organization 的「People」加入新成員、移除舊成員，網址不變。
