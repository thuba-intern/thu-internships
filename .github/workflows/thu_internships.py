# -*- coding: utf-8 -*-
"""
擷取東海大學各系的實習資訊，整理成一個新的網頁 internships.html

資料來源：
  1. 企管系「實習與徵才」            https://ba.thu.edu.tw/
  2. 國貿系「實習與徵才」            https://inttrade.thu.edu.tw/
  3. 財金系「實習職缺總覽」          https://fin.thu.edu.tw/
  4. 經濟系「實習資訊」              https://economic.thu.edu.tw/intern-news/
  5. 實習與學生成就中心「實習最速報」  https://isac-net.thu.edu.tw/
  6. GIFT 全球專業實習聯盟平臺「學期間實習」 https://www.gift.org.tw/

頁面列出：實習職缺、職缺來源（點了直接開啟原公告）、截止報名日期、類別（hashtag）。
可以用下拉選單篩選來源與類別；報名已截止的實習會標示「已截止」。另外輸出一份 CSV 給 Excel 使用。

使用方式（在命令提示字元 / 終端機）：
    pip install requests beautifulsoup4
    python thu_internships.py

執行完會在同一個資料夾產生 internships.html，用瀏覽器打開即可。
"""
import copy
import csv
import os
import html
import re
import time
from datetime import date, datetime, timedelta, timezone
from urllib.parse import urljoin, urlsplit, unquote

import requests
from bs4 import BeautifulSoup

# ───────────── 可以自己調整的設定 ─────────────
OUTPUT = "internships.html"        # 輸出的新頁面檔名
CSV_OUTPUT = "internships.csv"     # 同時輸出一份 CSV，可用 Excel 開啟或匯入 WordPress
MANUAL_CSV = "manual_overrides.csv"   # 手動補充表：程式抓不到或抓錯的截止日、實習期間，填在這裡（跟程式放同一個資料夾）
ELIGIBILITY_CSV = "eligibility_notes.csv"   # 報名限制表：人工確認過的「限本系」等資格，會標在職缺名稱下面（跟程式放同一個資料夾）
DELAY = 1.0                        # 每抓一頁停幾秒（避免對學校網站造成負擔）

# 來源 1：企管系
BA_BASE = "https://ba.thu.edu.tw/"
BA_CATEGORY = "實習與徵才"
BA_CATEGORY_URL = f"{BA_BASE}Education/cate/{BA_CATEGORY}/"
BA_MAX_PAGES = 5

# 來源 2：國貿系（WordPress，用它提供的資料介面 wp-json 讀取）
INTTRADE_BASE = "https://inttrade.thu.edu.tw/"
INTTRADE_CATEGORY_SLUG = "cat-announcement-internship"
INTTRADE_PAGE_URL = INTTRADE_BASE + "news/?tx_category=" + INTTRADE_CATEGORY_SLUG
INTTRADE_MAX_PAGES = 3             # 一頁 20 則

# 來源 3：財金系（清單頁是用 JavaScript 產生的，程式讀不到；改讀首頁置頂的「實習職缺總覽」表格）
FIN_BASE = "https://fin.thu.edu.tw/"
FIN_LIST_URL = FIN_BASE + "web/news/list.php?cid=15"
FIN_MAX_ITEMS = 60

# 來源 4：經濟系（所有公告都寫在同一頁，每則以標題分隔）
ECON_URL = "https://economic.thu.edu.tw/intern-news/"
ECON_SKIP_PREFIXES = ("【徵才】", "【活動】")      # 這些開頭的不是實習

# 來源 5：實習與學生成就中心（cid=2 是「實習最速報」分類）
ISAC_BASE = "https://isac-net.thu.edu.tw/"
ISAC_LIST_URL = ISAC_BASE + "web/research/list.php?page={page}&cid=2&action=search"
ISAC_MAX_PAGES = 3                 # 一頁 20 則

# 來源 6：GIFT 全球專業實習聯盟平臺（s_1=2 是「學期間實習」；1 寒暑假實習、4 海外實習）
GIFT_BASE = "https://www.gift.org.tw/"
GIFT_LIST_URL = GIFT_BASE + "index.php/index.php/welcome/jobs/?page={page}&s_1=2"
GIFT_MAX_PAGES = 5                 # 一頁 10 則
GIFT_DETAIL_URL = GIFT_BASE + "index.php/welcome/company_job_dialog"   # 「詳細」視窗的內容（POST job_id）
GIFT_JOB_URL = GIFT_BASE + "index.php/welcome/jobs?s_1=2#{id}"         # 打開就會直接跳出那則的詳細視窗

# 篩選規則
MAX_AGE_DAYS = 120                 # 發布超過幾天的舊公告就不列（還沒截止的除外）
NO_DEADLINE_KEEP_DAYS = 60         # 找不到截止日的公告：發布超過幾天就不列
EXPIRED_KEEP_DAYS = None           # 已截止超過幾天就不列；None 表示已截止的全部列出
EXCLUDE_TITLE_WORDS = ["錄取名單", "【活動", "總覽", "Google Taiwan 實習說明會",
                       "助教", "企業實習申請"]     # 標題含這些字的公告不是招募資訊，直接略過
OVERVIEW_WORDS = ["總覽"]          # 標題含這些字的是「彙整表」：一律列出，不判斷截止日

# 詳細內容
MAX_CONTENT_CHARS = 3000           # 每則公告最多帶過來幾個字，超過就截斷

TODAY = datetime.now(timezone(timedelta(hours=8))).date()     # 台灣時間的今天

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
})


# ═════════════ 1. 下載網頁 ═════════════
def get_with_retry(url, tries=4):
    """連線被中斷或逾時時，等一下再試（等 5、15、30 秒），最多試 4 次。
    有些學校網站偶爾會直接斷線（例如 Remote end closed connection），多試幾次通常就能抓到。"""
    for k in range(tries):
        try:
            return SESSION.get(url, timeout=30)
        except (requests.ConnectionError, requests.Timeout) as e:
            if k == tries - 1:
                raise
            wait = (5, 15, 30)[min(k, 2)]
            print(f"  連線失敗，{wait} 秒後重試（第 {k + 2} 次）：{e.__class__.__name__}")
            time.sleep(wait)


def fetch(url):
    """下載一個網頁，回傳 BeautifulSoup 物件。"""
    resp = get_with_retry(url)
    if resp.status_code != 404:      # 404 先不當成錯誤：有些網站會回 404 卻仍附上內容
        resp.raise_for_status()
    time.sleep(DELAY)
    return BeautifulSoup(resp.content, "html.parser")


def fetch_json(url):
    """下載 JSON 資料（國貿系的 WordPress 介面用）。"""
    resp = get_with_retry(url)
    resp.raise_for_status()
    time.sleep(DELAY)
    return resp.json()


def same_site(url, base):
    return urlsplit(url).netloc == urlsplit(base).netloc


def path_of(url):
    """取出網址的路徑並把 %E5%AF... 轉回中文，方便比對。"""
    return unquote(urlsplit(url).path)


# ═════════════ 2. 把網頁轉成文字 ═════════════
# 內文裡的連結和圖片先換成「記號」夾在文字中，最後產生頁面時再還原成 <a>、<img>。
# 記號用的是一般文章不會出現的特殊字元：
#   連結   網址  連結文字         圖片   圖片網址 
LINK_MARK = re.compile("([^]*)([^]*)")
IMG_MARK = re.compile("([^]*)")
SPLIT_MARK = ""                # 經濟系：標示每一則公告的開頭


def page_text(soup, base_url):
    """把網頁（或其中一塊）轉成文字：區塊（段落、表格列…）之間換行，行內文字保持在同一行；
    連結與圖片以記號保留。注意：會直接改動傳進來的 soup。"""
    for tag in soup(["script", "style", "noscript", "nav", "footer"]):
        tag.decompose()
    for img in soup.find_all("img"):
        src = urljoin(base_url, img.get("src") or "")
        ok = src.startswith("http") and "emoji" not in src
        img.replace_with(f"\n{src}\n" if ok else "")
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        label = a.get_text(" ", strip=True)
        url = href if href.startswith("mailto:") else urljoin(base_url, href)
        if "" in label or not url.startswith(("http", "mailto:")):
            a.replace_with(label)                    # 包著圖片的連結、javascript: 等：只留文字
        else:
            a.replace_with(f"{url}{label or url}")
    for br in soup.find_all("br"):
        br.replace_with("\n")
    for cell in soup.find_all(["td", "th"]):         # 表格：同一列的儲存格用「｜」隔開
        cell.append(" ｜ ")
    for tag in soup.find_all(["p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
                              "section", "article", "table", "ul", "ol", "figure", "hr"]):
        tag.insert_before("\n")
        tag.append("\n")
    text = (soup.body or soup).get_text("")
    text = re.sub(r"[ \t\xa0　]+", " ", text)
    return re.sub(r"\s*\n\s*", "\n", text).strip()


def strip_marks(text):
    """拿掉連結、圖片記號，只留純文字（找截止日時用）。"""
    return LINK_MARK.sub(r"\2", IMG_MARK.sub(" ", text))


END_MARKS = re.compile(r"下一則|上一則|回列表|上一篇|下一篇|回上一頁|返回列表|相關文章|Copyright|版權所有")
META_LINE = re.compile(r"^[\s\-*・]*(日期|單位|分類|點閱)\s*[:：]")


def cut_tail(text, start=0):
    """把「上一則／下一則／回列表／頁尾」開始的部分切掉（從那一行的開頭切，才不會留下半個連結）。"""
    m = END_MARKS.search(text, start)
    if not m:
        return text
    cut = text.rfind("\n", 0, m.start()) + 1
    return text[:cut if cut >= start else m.start()]


def detail_text(soup, title, base_url):
    """取出公告內頁的內文：從標題開始，到「上一則／下一則／頁尾」之前。"""
    text = page_text(soup, base_url)
    start = text.find(title)
    if start == -1:
        start = text.find(title[:12])
    if start > 0:
        text = text[start:]
    return cut_tail(text, min(len(title), len(text)))


def clean_content(text, title=""):
    """整理要顯示在頁面上的內文：去掉重複的標題、「日期／單位／點閱」那幾行，太長就截斷。"""
    raw = [line.strip(" ｜") for line in text.split("\n")]
    raw = [line for line in raw if line]
    for n, line in enumerate(raw[:8]):               # 開頭幾行若有「點閱 : 170」，連同前面的資訊行一起去掉
        if re.match(r"點閱\s*[:：]", strip_marks(line)):
            raw = raw[n + 1:]
            break
    lines, size = [], 0
    for line in raw:
        plain = strip_marks(line).strip()
        if plain == title or META_LINE.match(plain):
            continue
        if not lines and plain and title.startswith(plain.rstrip(". …")):
            continue                                 # 麵包屑裡被截短的標題
        lines.append(line)
        size += len(plain)
        if size > MAX_CONTENT_CHARS:
            lines.append("…（內容較長，其餘請點標題看原公告）")
            break
    return "\n".join(lines)


# ═════════════ 3. 從文字裡找出報名截止日 ═════════════
# 公告的截止日都寫在內文裡，寫法很多種，例如：
#   報名期間：即日起至 2026/11/08 止      即日起至115年10月31日截止
#   歡迎投遞履歷(10/31收件截止)           請於115/11/15前寄送履歷
#   【校內截止收件時間】115年10月22日(四)17時
# 做法：把文字切成一小段一小段，只看「跟報名有關」的段落，
#       再挑出後面接著「止／截止／前」或前面是「至／截止／期限」的日期。

DATE_RE = re.compile(r"""
      (?<![\d/.]) (?P<y>20\d{2}|1\d{2}) \s*[./\-年]\s* (?P<m>\d{1,2}) \s*[./\-月]\s* (?P<d>\d{1,2}) \s*日? (?!\d)
    | (?<![\d/.年]) (?P<m2>\d{1,2}) \s*月\s* (?P<d2>\d{1,2}) \s*日?
    | (?<![\d/.])  (?P<m3>\d{1,2}) / (?P<d3>\d{1,2}) (?![\d/])
""", re.X)

REG_WORDS = re.compile(r"報名|申請|收件|投遞|履歷|截止|期限|即日起|應徵|受理|徵件")      # 跟報名有關
REG_LABEL = re.compile(r"(?:報名|申請|收件|受理|投遞|徵件)\s*(?:期間|時間|日期|時程)|截止|期限")   # 可當作小標題
SKIP_WORDS = re.compile(r"實習期間|實習時間|實習日期|工作期間|面試|筆試|測驗|說明會|報到|錄取|放榜|活動日期|活動時間")

# 日期「後面」緊接著：(星期) 時間 止／截止／前
AFTER_CUE = re.compile(r"""^\s*
    (?:[（(][^（()）]{1,8}[)）])? \s*
    (?:(?:上午|中午|下午|晚上|午夜)\s*)?
    (?:\d{1,2}\s*(?:[:：]\s*\d{2}|時|點)(?:\s*\d{1,2}\s*分)?)? \s*
    (?:[（(][^（()）]{1,8}[)）])? \s*
    (?:收件|報名|申請|投遞)? \s*
    (?:截止|為止|止|以前|之前|前(?![往後述面者]))
""", re.X)
# 日期「前面」緊接著：至／～
BEFORE_CUE = re.compile(r"(?:至|迄|[~～])\s*(?:同年)?\s*$")


def roc_to_ad(year):
    """民國年轉西元年（115 → 2026）。"""
    return year + 1911 if year < 1000 else year


def dates_in(seg, pub=None, ref=None):
    """找出一段文字裡的所有日期，回傳 [(開始位置, 結束位置, date), ...]。
    沒寫年份的日期會自動推算：同一段前面有日期就沿用它的年份；否則依公告發布日（pub）；
    連發布日都沒有時，依內文裡其他有寫年份的日期（ref）或最接近今天的那一年。"""
    found, prev = [], None
    for m in DATE_RE.finditer(seg):
        try:
            if m.group("y"):
                d = date(roc_to_ad(int(m.group("y"))), int(m.group("m")), int(m.group("d")))
            else:
                mon = int(m.group("m2") or m.group("m3"))
                day = int(m.group("d2") or m.group("d3"))
                if prev:
                    d = date(prev.year, mon, day)
                    if d < prev:
                        d = date(prev.year + 1, mon, day)
                elif pub:
                    d = date(pub.year, mon, day)
                    if d < pub - timedelta(days=90):
                        d = date(pub.year + 1, mon, day)
                elif ref:                                  # 截止日應該在實習開始之前
                    d = date(ref.year, mon, day)
                    if d > ref:
                        d = date(ref.year - 1, mon, day)
                else:
                    options = [date(TODAY.year + k, mon, day) for k in (-1, 0, 1)]
                    d = min(options, key=lambda x: abs((x - TODAY).days))
        except ValueError:                                 # 不是合理的日期（例如 24/7）
            continue
        found.append((m.start(), m.end(), d))
        prev = d
    return found


def find_deadline(text, pub=None):
    """回傳文字中的報名截止日（date），找不到就回傳 None。有好幾個時取最晚的。"""
    text = re.sub(r"https?://\S+", " ", strip_marks(text))     # 網址裡的數字不要當成日期
    ref = None
    if pub is None:                                        # 沒有發布日：找內文裡最早一個有寫年份的日期
        explicit = [date(roc_to_ad(int(m.group("y"))), int(m.group("m")), int(m.group("d")))
                    for m in DATE_RE.finditer(text) if m.group("y") and _valid(m)]
        ref = min(explicit) if explicit else None

    candidates, label, prev = [], "", ""
    for seg in re.split(r"[\n。；;，,！!｜|]", text):
        seg = seg.strip()
        if not seg:
            continue
        dates = dates_in(seg, pub, ref)
        self_reg, self_skip = bool(REG_WORDS.search(seg)), bool(SKIP_WORDS.search(seg))
        lab_reg, lab_skip = bool(REG_LABEL.search(label)), bool(SKIP_WORDS.search(label))

        # 「請使用線上投遞履歷，請於2026/11/20(四)前上線填寫」：報名的字眼在前一句，日期在這一句
        if (dates and not self_reg and not self_skip
                and REG_WORDS.search(prev) and not SKIP_WORDS.search(prev)):
            candidates += [d for start, end, d in dates if AFTER_CUE.match(seg[end:])]
        prev = seg

        if dates and (self_reg or (lab_reg and not lab_skip and not self_skip)):
            hits = []
            for start, end, d in dates:
                near = seg[max(0, start - 12):start]
                if (AFTER_CUE.match(seg[end:]) or BEFORE_CUE.search(seg[:start])
                        or re.search(r"截止|期限", near)
                        or (start < 4 and re.search(r"截止|期限", label))):
                    hits.append(d)
            # 「報名期間：9/24-10/5」這種沒有「止」字的範圍：取最後一個日期
            if not hits and len(dates) >= 2 and (REG_LABEL.search(seg[:dates[0][0]]) or lab_reg):
                hits.append(dates[-1][2])
            candidates += hits

        # 沒有日期的短句（例如「【申請期間】」）當成下一段的小標題
        label = seg if (not dates and len(seg) <= 20) else ""

    if pub:                                                # 截止日不可能早於公告發布日
        candidates = [d for d in candidates if d >= pub]
    return max(candidates) if candidates else None


def _valid(m):
    try:
        date(roc_to_ad(int(m.group("y"))), int(m.group("m")), int(m.group("d")))
        return True
    except ValueError:
        return False


def new_item(source, title, url, pub=None, summary=""):
    """每一則實習都整理成同樣的格式。"""
    return {"source": source, "title": title, "url": url, "pub": pub, "summary": summary,
            "text": None,            # 內頁全文（找截止日用）；None 表示還沒抓
            "content": "",           # 要顯示在頁面上的詳細內容
            "deadline_hint": None,   # 來源網站直接給的截止日（財金系總覽表）
            "current": False}        # True：來源已確認是本學期的職缺，不套用「太舊就不列」


# ═════════════ 4. 來源 1：企管系「實習與徵才」 ═════════════
MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}


def ba_is_article(url):
    """公告內頁的網址長得像 /Education/公告標題/，分類頁則是 /Education/cate/分類/。"""
    p = path_of(url)
    return (same_site(url, BA_BASE) and p.startswith("/Education/")
            and not p.startswith("/Education/cate/") and p != "/Education/")


def ba_parse_date(text):
    """企管系網站的日期寫成「05 2026 Oct」。"""
    m = re.search(r"(\d{1,2})\s*(\d{4})\s*([A-Za-z]{3})", text)
    if not m or m.group(3).title() not in MONTHS:
        return None
    try:
        return date(int(m.group(2)), MONTHS[m.group(3).title()], int(m.group(1)))
    except ValueError:
        return None


def ba_parse_items(soup, page_url):
    items = []
    for a in soup.find_all("a", href=True):
        url = urljoin(page_url, a["href"])
        if not ba_is_article(url):
            continue
        box = a.find_parent("li") or a.parent              # 每則公告是一個 <li>

        # 首頁會把各種分類混在一起，每則公告下方有分類連結；有的話就只留我們要的分類
        cates = [path_of(urljoin(page_url, c["href"])) for c in box.find_all("a", href=True)]
        cates = [c for c in cates if c.startswith("/Education/cate/")]
        if cates and not any(BA_CATEGORY in c for c in cates):
            continue

        heading = box.find(["h2", "h3", "h4", "h5"])
        title = (a.get("title") or (heading.get_text(strip=True) if heading else "")
                 or a.get_text(strip=True)).strip()
        if not title:
            continue
        p = box.find("p")
        summary = p.get_text(" ", strip=True) if p else ""
        items.append(new_item("企管系", title, url, ba_parse_date(box.get_text(" ", strip=True)),
                              "" if summary == title else summary))
    return items


def ba_next_pages(soup, page_url):
    """找分頁連結：網址以分類頁開頭、但不是分類頁本身的連結。"""
    base, pages = path_of(BA_CATEGORY_URL), []
    for a in soup.find_all("a", href=True):
        url = urljoin(page_url, a["href"]).split("#")[0]
        if not same_site(url, BA_BASE):
            continue
        p, q = path_of(url), urlsplit(url).query
        if p.startswith(base) and (p != base or q) and url not in pages:
            pages.append(url)
    return pages


def crawl_ba():
    results, queue, visited = {}, [BA_CATEGORY_URL], set()
    while queue and len(visited) < BA_MAX_PAGES:
        url = queue.pop(0)
        if url in visited:
            continue
        visited.add(url)
        try:
            soup = fetch(url)
        except requests.RequestException as e:
            print(f"[企管系] 略過 {unquote(url)}：{e}")
            continue
        for item in ba_parse_items(soup, url):
            results.setdefault(item["url"], item)
        queue += [u for u in ba_next_pages(soup, url) if u not in visited]
        print(f"[企管系] {unquote(url)}（累計 {len(results)} 則）")

    if not results:      # 備援：分類頁抓不到時，改從首頁「最新消息」抓（固定列出最新 10 則）
        print("[企管系] 分類頁沒有抓到資料，改從首頁擷取")
        for item in ba_parse_items(fetch(BA_BASE), BA_BASE):
            results.setdefault(item["url"], item)
    return list(results.values())


# ═════════════ 5. 來源 2：國貿系「實習與徵才」 ═════════════
def crawl_inttrade():
    """國貿系網站是 WordPress，可以直接向它要 JSON 格式的文章資料（含全文），不用解析網頁。"""
    api = INTTRADE_BASE + "wp-json/wp/v2/"
    cats = fetch_json(f"{api}categories?slug={INTTRADE_CATEGORY_SLUG}")
    if not cats:
        print("[國貿系] 找不到分類")
        return []
    items = []
    for page in range(1, INTTRADE_MAX_PAGES + 1):
        url = (f"{api}posts?categories={cats[0]['id']}&per_page=20&page={page}"
               "&_fields=id,date,link,title,content")
        try:
            posts = fetch_json(url)
        except requests.RequestException:                  # 超過最後一頁時 WordPress 會回錯誤
            break
        for p in posts:
            title = BeautifulSoup(p["title"]["rendered"], "html.parser").get_text(" ", strip=True)
            item = new_item("國貿系", title, p["link"], date.fromisoformat(p["date"][:10]))
            item["text"] = page_text(BeautifulSoup(p["content"]["rendered"], "html.parser"), p["link"])
            item["content"] = clean_content(item["text"], title)
            items.append(item)
        print(f"[國貿系] 第 {page} 頁（累計 {len(items)} 則）")
        if len(posts) < 20:
            break
    return items


# ═════════════ 6. 來源 3：財金系「實習職缺總覽」 ═════════════
def fin_parse_overview(soup, page_url):
    """總覽公告裡是一張表格：單位｜職缺｜實習資格｜報名方式｜實習資訊(連結)｜報名截止日期。
    每一列就是一個實習職缺。"""
    default = ["單位", "職缺", "實習資格", "報名方式", "實習資訊", "報名截止日期"]
    rows = {}
    for table in soup.find_all("table"):
        header = None
        for tr in table.find_all("tr"):
            cells = tr.find_all(["td", "th"])
            texts = [c.get_text(" ", strip=True) for c in cells]
            link_at = next((j for j, c in enumerate(cells)
                            if c.find("a", href=re.compile(r"detail\.php"))), None)
            if link_at is None:
                if any("截止" in t for t in texts):
                    header = texts                         # 表頭列
                continue
            names = header if header and len(header) >= len(cells) else default
            url = urljoin(page_url, cells[link_at].find("a", href=re.compile(r"detail\.php"))["href"])
            find_col = lambda word: next((j for j, n in enumerate(names) if word in n), None)
            dl_at = find_col("截止") if find_col("截止") is not None else len(cells) - 1
            job_at = find_col("職缺") or 1

            block = []                                     # 這一列的資訊，會放在詳細內容最上面
            for j, cell in enumerate(cells):
                if j == link_at or j >= len(names) or not texts[j]:
                    continue
                rich = page_text(copy.copy(cell), page_url).replace("\n", " ").strip(" ｜")
                block.append(f"{names[j]}：{rich}")
            found = dates_in(texts[dl_at]) if dl_at < len(texts) else []
            hint = max(d for _, _, d in found) if found else None

            if url in rows:                                # 同一則公告在表格裡出現兩列：合併
                old = rows[url]
                old["table"] += "\n" + "\n".join(block)
                old["deadline_hint"] = max(filter(None, [old["deadline_hint"], hint]), default=None)
                continue
            unit, job = texts[0], (texts[job_at] if job_at < len(texts) else "")
            item = new_item("財金系", f"{unit}｜{job}", url, summary=f"職缺：{job}")
            item.update(deadline_hint=hint, current=True, unit=unit, job=job, table="\n".join(block))
            rows[url] = item
    return list(rows.values())


def fin_parse_detail(soup, url):
    """財金系內頁的開頭是：標題 → 日期(2026-09-30) → 單位 → 分類 → 點閱。
    回傳 (標題, 發布日, 內文)；格式不符時回傳 None。"""
    lines = page_text(soup, url).split("\n")
    for n, line in enumerate(lines[1:], start=1):
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})(?!\d)", strip_marks(line).strip())
        if m:
            title = strip_marks(lines[n - 1]).strip()
            text = cut_tail("\n".join(lines[n - 1:]), len(title))
            return title, date(*map(int, m.groups())), text
    return None


def fin_add_detail(item):
    """抓財金系公告內頁。這個網站的網址有時會顯示成另一則公告，所以要核對標題裡有沒有單位名稱。"""
    url = item["url"]
    other = url.replace("&cid=15", "") if "cid=15" in url else url + "&cid=15"
    key = re.sub(r"\s", "", item.get("unit", ""))[:2]      # 例如「玉山金控」→「玉山」
    for try_url in (url, other):
        try:
            parsed = fin_parse_detail(fetch(try_url), try_url)
        except requests.RequestException as e:
            print(f"  [財金系] 內頁抓取失敗：{item['title'][:20]}…（{e}）")
            continue
        if parsed and key in parsed[0]:
            item["title"], item["pub"], item["text"] = parsed
            item["url"] = try_url
            return
    item["text"] = ""                                      # 內頁對不上：只用總覽表的資訊


def crawl_fin():
    home = fetch(FIN_BASE)
    overview, seeds = None, {}
    for a in home.find_all("a", href=re.compile(r"news/detail\.php")):
        text, url = a.get_text(" ", strip=True), urljoin(FIN_BASE, a["href"])
        if "實習" in text and any(w in text for w in OVERVIEW_WORDS):
            overview = overview or url
        elif "cid=15" in url or "實習" in text:
            seeds.setdefault(url, text)

    items = []
    if overview:
        items = fin_parse_overview(fetch(overview), overview)
        print(f"[財金系] 實習職缺總覽表：{len(items)} 個職缺")
    else:
        print("[財金系] 首頁找不到「實習職缺總覽」，只抓首頁列出的實習公告")
    fin_id = lambda u: (re.search(r"id=(\d+)", u) or [u, u])[1]      # 網址裡的公告編號
    known = {fin_id(i["url"]) for i in items}
    for url, text in seeds.items():                        # 首頁上有、但總覽表還沒列的新公告
        if fin_id(url) not in known:
            items.append(new_item("財金系", re.sub(r"^【置頂】", "", text), url))

    items = items[:FIN_MAX_ITEMS]
    for n, item in enumerate(items, start=1):
        fin_add_detail(item)
        body = clean_content(item["text"], item["title"])
        item["content"] = "\n".join(filter(None, [item.get("table", ""), body]))
        item["text"] = item["text"] or item.get("table", "")
        if n % 10 == 0:
            print(f"[財金系] 內頁 {n}/{len(items)}")
    return items


# ═════════════ 7. 來源 4：經濟系「實習資訊」 ═════════════
def crawl_econ():
    """經濟系把所有公告寫在同一頁，每則以一個標題（h2）開頭，沒有各自的內頁與發布日。"""
    soup = fetch(ECON_URL)
    # 依序找內容區：一般文章區 → article → main → 整頁；挑第一個裡面有 2 個以上標題的
    candidates = [soup.select_one(".entry-content"), soup.find("article"), soup.find("main"), soup]
    area, heads = soup, []
    for cand in candidates:
        if cand is None:
            continue
        for level in ("h2", "h3"):
            found = cand.find_all(level)
            if len(found) >= 2:
                area, heads = cand, found
                break
        if heads:
            break
    if not heads:                                          # 找不到標題時印出線索，方便判斷網站是不是改版或擋連線
        title = soup.title.get_text(strip=True) if soup.title else "（無標題）"
        print(f"[經濟系] 頁面上找不到標題：網頁標題「{title}」，長度 {len(str(soup))} 字元，"
              f"開頭：{soup.get_text(' ', strip=True)[:150]}")
    for h in heads:
        h.insert_before("\n" + SPLIT_MARK)                 # 在每個標題前做記號，等一下用它切開
    items = []
    for chunk in page_text(area, ECON_URL).split(SPLIT_MARK)[1:]:
        lines = [l for l in cut_tail(chunk).split("\n") if l.strip()]
        if len(lines) < 2:
            continue
        title, body = strip_marks(lines[0]).strip(), "\n".join(lines[1:])
        if title.startswith("【活動】"):
            continue                                       # 活動不是職缺
        if "實習" not in title + strip_marks(body) and not title.startswith("【徵才】"):
            continue
        item = new_item("經濟系", title, ECON_URL)
        item["text"], item["content"] = body, clean_content(body, title)
        items.append(item)
    if heads and not items:                                # 有標題卻一則都沒抓到：印出前幾個標題，方便除錯
        print(f"[經濟系] 找到 {len(heads)} 個標題但沒有職缺，前幾個：",
              [h.get_text(" ", strip=True)[:20] for h in heads[:5]])
    print(f"[經濟系] {len(items)} 則實習／徵才")
    return items


# ═════════════ 8. 來源 5：實習與學生成就中心「實習最速報」 ═════════════
ISAC_DATE_RE = re.compile(r"日期\s*[:：]\s*(\d{4})-(\d{1,2})-(\d{1,2})")


def isac_parse_items(soup, page_url):
    """清單上每則公告都有一行「日期 : 2026-10-05 …」。
    從每個連結往外層找，找到「剛好只含一個日期」的區塊，那就是一則公告。"""
    items, seen_boxes = [], set()
    for a in soup.find_all("a", href=True):
        title = a.get_text(" ", strip=True)
        url = urljoin(page_url, a["href"])
        is_detail = "detail.php" in url
        if not title or not (is_detail or not same_site(url, ISAC_BASE)):
            continue                                       # 只要公告內頁，或直接連到校外的公告
        box = a.parent
        while box is not None and not ISAC_DATE_RE.search(box.get_text(" ", strip=True)):
            box = box.parent
        if box is None or id(box) in seen_boxes:
            continue
        text = box.get_text(" ", strip=True)
        found = ISAC_DATE_RE.findall(text)
        if len(found) != 1:                                # 含好幾個日期 → 是整個清單，不是單一則
            continue
        seen_boxes.add(id(box))
        y, m, d = map(int, found[0])
        summary = ISAC_DATE_RE.split(text.split(title, 1)[-1])[0].strip(" *")   # 標題之後、日期之前的文字
        items.append(new_item("實習中心", title, url, date(y, m, d), summary))
    return items


def crawl_isac():
    results = {}
    for page in range(1, ISAC_MAX_PAGES + 1):
        url = ISAC_LIST_URL.format(page=page)
        try:
            items = isac_parse_items(fetch(url), url)
        except requests.RequestException as e:
            print(f"[實習中心] 略過第 {page} 頁：{e}")
            break
        new = [i for i in items if i["url"] not in results]
        for item in new:
            results[item["url"]] = item
        print(f"[實習中心] 第 {page} 頁（累計 {len(results)} 則）")
        # 清單由新到舊排列：沒有新資料，或整頁都是很舊的公告，就不用再往後翻
        if not new or all((TODAY - i["pub"]).days > MAX_AGE_DAYS for i in new):
            break
    return list(results.values())


# ═════════════ 8-2. 來源 6：GIFT 全球專業實習聯盟平臺「學期間實習」 ═════════════
GIFT_LABELS = ["職缺分類", "職務類型", "特別身份", "膳宿", "獎學金"]


def gift_field(text, label):
    """取出「職缺分類 Vacancies Type 學期間實習」這種欄位的值（到下一個欄位名稱為止）。"""
    others = "|".join(l for l in GIFT_LABELS if l != label) + "|詳細|儲存"
    m = re.search(rf"{label}\s*(?:[A-Za-z /]+)?[:：]?\s*(.+?)\s*(?={others}|$)", text)
    return m.group(1).strip(" ：:") if m else ""


def gift_parse(soup, page_url):
    """每則職缺的開頭是一個標題（h3），底下依序是公司、地點、職缺分類、職務類型、膳宿、獎學金。
    「詳細」按鈕是 JavaScript，沒有單獨的網址，所以消息來源連到列表頁。"""
    items = []
    for h in soup.find_all(["h3", "h4"]):
        title = h.get_text(" ", strip=True)
        box = h.parent
        while box is not None and "職缺分類" not in box.get_text(" ", strip=True):
            box = box.parent
        if not title or box is None or len(box.find_all(h.name)) != 1:
            continue                                   # 找不到欄位，或包到好幾則（是整個清單）
        text = box.get_text(" ", strip=True)
        lines = [l.strip() for l in box.get_text("\n", strip=True).split("\n") if l.strip()]
        after = lines[lines.index(title) + 1:] if title in lines else lines
        company = after[0] if after else ""
        place = after[1] if len(after) > 1 and not after[1].startswith("職缺分類") else ""
        kind, job, pay = gift_field(text, "職缺分類"), gift_field(text, "職務類型"), gift_field(text, "獎學金")
        a = h.find("a", onclick=re.compile(r"loadvacdatamore\(\d+\)"))
        job_id = re.search(r"\d+", a["onclick"]).group() if a else ""
        item = new_item("GIFT實習平台", title, GIFT_JOB_URL.format(id=job_id) if job_id else page_url, None,
                        "｜".join(filter(None, [company, place, job, pay])))
        item["unit"], item["job"] = company, job       # 職缺欄顯示「職務類型」（財務、資訊…）
        item["content"] = "\n".join(f"{k}：{v}" for k, v in
                                    [("公司", company), ("地點", place), ("職缺分類", kind),
                                     ("職務類型", job), ("膳宿", gift_field(text, "膳宿")), ("薪資／獎學金", pay)] if v)
        item["text"] = item["content"]
        item["gift_id"] = job_id
        items.append(item)
    return items


def gift_detail(item):
    """讀「詳細」視窗，取出【實習時間】，例如「2027/01 ~ 2027/06」。
    實習已經開始（開始月份早於本月）的職缺，學期中已經來不及報名，就不列出來。"""
    if not item.get("gift_id"):
        return True
    try:
        resp = SESSION.post(GIFT_DETAIL_URL, data={"job_id": item["gift_id"]}, timeout=20)
        resp.raise_for_status()
    except requests.RequestException as e:
        print(f"[GIFT] 讀不到 {item['title']} 的詳細內容：{e}")
        return True
    resp.encoding = resp.apparent_encoding or "utf-8"
    text = re.sub(r"\s+", " ", BeautifulSoup(resp.text, "html.parser").get_text(" "))
    m = re.search(r"【實習時間】\s*(\d{4})\s*/\s*(\d{1,2})\s*~\s*(\d{4})\s*/\s*(\d{1,2})", text)
    time.sleep(DELAY)
    if not m:
        return True
    y1, m1, y2, m2 = map(int, m.groups())
    item["period"] = f"{y1}/{m1:02d}~{y2}/{m2:02d}"
    item["content"] += f"\n實習時間：{item['period']}"
    if (y1, m1) < (TODAY.year, TODAY.month):
        return False                                   # 已經開始實習了
    item["current"] = True                             # 還沒開始的梯次：沒有截止日也照樣列出
    return True


def crawl_gift():
    items, seen = [], set()
    for page in range(1, GIFT_MAX_PAGES + 1):
        url = GIFT_LIST_URL.format(page=page)
        try:
            found = gift_parse(fetch(url), url)
        except requests.RequestException as e:
            print(f"[GIFT] 略過第 {page} 頁：{e}")
            break
        key = lambda i: (i["title"], i["unit"], i["content"])     # 同名職缺可能只差在職務類型
        new = [i for i in found if key(i) not in seen]
        for i in new:
            seen.add(key(i))
        items += new
        print(f"[GIFT] 第 {page} 頁（累計 {len(items)} 則）")
        if not new:
            break
    kept = [i for i in items if gift_detail(i)]
    print(f"[GIFT] 讀完詳細內容：{len(kept)} 則還沒開始實習，{len(items) - len(kept)} 則已開始、不列出")
    return kept


# ═════════════ 9. 整理每一則的欄位：實習名稱、職缺、實習期間、產業類別 ═════════════
def clean_title(title):
    """標題去掉學期標記與截止日，留下實習名稱。
    例：【115-2 實習資訊】上銀科技 2027 學期型實習計畫 – 11/01（日） 申請截止 → 上銀科技 2027 學期型實習計畫"""
    s = re.sub(r"^\s*(?:【[^】]*(?:\d{3}-\d|置頂)[^】]*】\s*[-－]?\s*)+", "", title)       # 【115-2實習】【置頂】
    s = re.sub(r"\s*[-–—－]\s*\d{1,2}/\d{1,2}\s*[（(][^）)]*[）)]\s*申請截止\s*$", "", s)    # – 11/01（日） 申請截止
    s = re.sub(r"\s*[（(][^()（）]*(?:截止|止)\s*[)）]\s*$", "", s)                         # (10/18申請截止)
    return s.strip(" -–") or title


PERIOD_LABEL = re.compile(r"實習(期間|日期|起迄|時間|時段)\s*(?:[（(][^）)]{0,12}[）)])?\s*[:：】｜]\s*(.+)")
PERIOD_OK = re.compile(r"[\d一二三四五六七八九十]|學期|暑|寒")


def find_period(text, title):
    """找實習期間：先看內文的「實習期間：…」，沒有的話再看標題裡的「(116/2-116/6)」。"""
    lines = strip_marks(text or "").split("\n")
    for want in (("期間", "日期", "起迄"), ("時間", "時段")):      # 「實習時間」常是每天的上班時段，放在後面才找
        for line in lines:
            m = PERIOD_LABEL.search(line)
            if not m or m.group(1) not in want:
                continue
            full = m.group(2).strip(" ｜")
            first = re.split(r"[，,；;。｜]", full)[0].strip()
            if "時間" in want and not re.search(r"[/年月]", first):
                continue                                   # 「實習時間：8:30~16:30」不是實習期間
            for value in (first, full):                    # 優先選有數字的寫法，例如「三個月，擬安排於2027/2/22~5/22」
                if re.search(r"\d", value):
                    return value[:40]
            if PERIOD_OK.search(first):
                return first[:40]
    m = re.search(r"(\d{3,4}/\d{1,2}(?:/\d{1,2})?)/?\s*[-~～–]\s*(\d{1,4}/\d{1,2}(?:/\d{1,2})?)", title)
    return f"{m.group(1)} ~ {m.group(2)}" if m else ""


POS_LABEL = re.compile(r"(?:實習職缺|招募職缺|職缺名稱|職缺內容|實習職務|招募職務|職位名稱|開放職缺|徵才職缺)\s*[:：】]\s*(.*)")


def find_positions(text):
    """找職缺名稱：「招募職缺：設備工程師、製造工程師」，或「實習職缺：」底下的 1. 2. 清單。"""
    lines = [l.strip(" ｜*") for l in strip_marks(text or "").split("\n")]
    for n, line in enumerate(lines):
        m = POS_LABEL.search(line)
        if not m:
            continue
        value = m.group(1).strip(" ｜*")
        if value and not value.startswith(("如附", "詳")):
            return value[:60]
        picks = []
        for nxt in lines[n + 1:n + 12]:
            mm = re.match(r"^[(（]?\d+[.、)）]\s*(.{2,25})$", nxt)
            if mm:
                picks.append(mm.group(1).strip())
            if len(picks) >= 4:
                break
        if picks:
            return "、".join(picks)
    return ""


# 產業類別：依序比對，先符合的為準（例如「國泰世華銀行 資訊實習」算金融業）
INDUSTRIES = [
    ("文化藝術", ["美術館", "博物館", "藝術", "文化", "生態館", "科學館", "展覽"]),
    ("公部門", ["政府", "國稅局", "財政部", "關務署", "公所", "國營", "農業部", "林業", "水利局",
               "衛生福利部", "管制署", "檔案管理局", "國家發展委員會", "經濟部", "機場"]),
    ("金融業", ["銀行", "商銀", "金控", "證券", "期貨", "保險", "人壽", "壽險", "產險", "產物", "信託",
               "金融", "二信", "合作金庫", "中租", "迪和", "投信", "理財", "富邦", "國泰", "玉山", "元大",
               "兆豐", "永豐銀", "永豐金", "永豐商銀", "永豐證券", "永豐期貨", "中國信託", "中信", "第一金", "華南", "聯邦", "安泰", "板信", "三信",
               "群益", "南山", "新光"]),
    ("科技業", ["科技", "電子", "半導體", "資訊", "軟體", "數智", "鼎新", "Google", "美光", "聯華電子",
               "聯電", "京元", "上銀", "凌網", "晶圓", "資安", "AI", "台積電", "台積", "緯創", "創資通"]),
    ("製造業", ["製造", "工業", "機械", "電機", "化工", "食品", "紡織", "鞋", "興農", "新武", "賜昌", "設備", "成衣", "飛雁", "永豐餘", "禾寶", "手套"]),
    ("商業服務", ["貿易", "國貿", "會計師", "事務所", "顧問", "行銷", "零售", "物流", "餐飲", "飯店", "觀光", "人力", "航運", "海空運", "UNIQLO", "服飾", "百貨"]),
]
INDUSTRY_NAMES = [name for name, _ in INDUSTRIES] + ["其他"]

# 不列出的產業類別：本站以商管為主，實習中心、GIFT 的美術館、博物館、表演藝術等文化藝術類實習，
# 企管系學生較難認列學分，所以不放。企管系自己公告的不受影響；
# 若某則被誤判成文化藝術，可在 manual_overrides.csv 的「類別」欄改成其他類別。
EXCLUDE_INDUSTRIES = {"文化藝術"}


def classify_industry(item):
    """先看名稱、單位、職缺、摘要；都沒有線索再看內文開頭。"""
    if any(w in item["title"] for w in OVERVIEW_WORDS):
        return "其他"
    head = " ".join([item["name"], item.get("unit", ""), item["positions"], item["summary"]])
    noise = re.compile(r"(招募|徵才|實習|校園|宣傳|活動)資訊")             # 「招募資訊」的「資訊」不代表資訊業
    for text in (noise.sub("", head), noise.sub("", strip_marks(item["content"])[:400])):
        for name, words in INDUSTRIES:
            if any(w in text for w in words):
                return name
    return "其他"


FULLTIME_WORDS = ["徵才", "儲備", "新鮮人", "正職", "校園招募", "公開招募", "招考"]   # 標題有這些字、但沒有「實習」→ 正職分頁


def job_kind(item):
    title = clean_title(item["title"])                     # 去掉「【115-2 實習資訊】」這類前綴再判斷
    if "實習" not in title and any(w in title for w in FULLTIME_WORDS):
        return "fulltime"
    return "intern"


def add_fields(item):
    item["kind"] = job_kind(item)
    item["name"] = clean_title(item["title"])
    item["positions"] = item.get("job") or find_positions(item["text"] or item["content"])
    item["period"] = find_period(item["text"] or item["content"], item["title"])
    item["industry"] = classify_industry(item)
    if not item["content"] and item["summary"]:
        item["content"] = item["summary"]                  # 抓不到內文時，至少放清單上的摘要
    return item



# ═════════════ 10. 抓詳細內容、判斷截止日、決定要不要列出 ═════════════
SOURCES = [("企管系", crawl_ba), ("國貿系", crawl_inttrade), ("財金系", crawl_fin),
           ("經濟系", crawl_econ), ("實習中心", crawl_isac),
           ("GIFT實習平台", crawl_gift)]
CAMPUS_SITES = [BA_BASE, ISAC_BASE]


def add_details(item):
    """補上詳細內容（還沒有的話就點進內頁抓），並找出報名截止日。
    截止日的優先順序：來源網站直接給的 → 標題與摘要 → 內文。"""
    age = (TODAY - item["pub"]).days if item["pub"] else 0
    if (item["text"] is None and age <= MAX_AGE_DAYS
            and any(same_site(item["url"], s) for s in CAMPUS_SITES)):
        try:
            item["text"] = detail_text(fetch(item["url"]), item["title"], item["url"])
            item["content"] = clean_content(item["text"], item["title"])
        except requests.RequestException as e:
            print(f"  內頁抓取失敗：{item['title'][:20]}…（{e}）")

    if any(w in item["title"] for w in OVERVIEW_WORDS):
        item["deadline"] = None                            # 彙整表沒有單一的截止日
    else:
        item["deadline"] = (item["deadline_hint"]
                            or find_deadline(item["title"] + "\n" + item["summary"], item["pub"])
                            or find_deadline(item["text"] or "", item["pub"]))
    return add_fields(item)


def status_of(item):
    """open＝可報名、closed＝已截止、unknown＝未標示截止日。"""
    if item["deadline"] is None:
        return "unknown"
    return "open" if item["deadline"] >= TODAY else "closed"


def should_list(item):
    """決定這一則要不要出現在頁面上。"""
    if any(w in item["title"] for w in OVERVIEW_WORDS):
        return True
    state = status_of(item)
    age = (TODAY - item["pub"]).days if item["pub"] else 0
    if state == "open" or item["current"] and state == "unknown":
        return True
    if state == "closed":
        if EXPIRED_KEEP_DAYS is not None and (TODAY - item["deadline"]).days > EXPIRED_KEEP_DAYS:
            return False
        return item["current"] or age <= MAX_AGE_DAYS
    return age <= NO_DEADLINE_KEEP_DAYS                    # 未標示截止日：太舊就不列


# ═════════════ 10-2. 手動補充表 ═════════════
# manual_overrides.csv 的欄位（第一列是標題，用 Excel 編輯後另存成「CSV UTF-8」）：
#   來源          要補哪個來源的資料，例如「企管系」
#   標題關鍵字    公告標題裡有這幾個字的那一則，例如「鼎新」
#   截止報名日期  例如 2026/11/08；沒有就留空
#   截止日顯示文字 選填：想顯示成別的寫法時填，例如「2026/11 月中前」（狀態仍依上一欄的日期判斷）
#   實習期間      例如 2027/02/22~2027/06/25；沒有就留空（目前頁面不顯示這欄，先保留備用）
#   實習職缺      選填：想改掉頁面上顯示的名稱時填
#   類別          選填：想改掉自動判斷的類別時填，例如「金融業」
#   來源網址      選填：表裡這一則在網站上找不到時，會新增成一則，連結到這個網址
def load_manual():
    for folder in (os.getcwd(), os.path.dirname(os.path.abspath(__file__))):
        path = os.path.join(folder, MANUAL_CSV)
        if os.path.exists(path):
            with open(path, encoding="utf-8-sig", newline="") as f:
                rows = [{k.strip(): (v or "").strip() for k, v in r.items() if k} for r in csv.DictReader(f)]
            print(f"讀取手動補充表 {path}：{len(rows)} 筆")
            return rows
    return []


def parse_manual_date(text):
    found = dates_in(text or "")
    return found[0][2] if found else None


def apply_manual(items):
    """用手動補充表覆蓋程式抓到的資料；表裡有、但網站上沒抓到的，新增成一則。"""
    for row in load_manual():
        src, key = row.get("來源", ""), row.get("標題關鍵字", "")
        if not key:
            continue
        hits = [i for i in items if (not src or i["source"] == src) and key in i["title"]]
        if not hits and row.get("來源網址"):
            new = new_item(src or "手動新增", row.get("實習職缺") or key, row["來源網址"])
            new.update(deadline=None, name=new["title"], positions="", period="", industry="其他", current=True)
            items.append(new)
            hits = [new]
        if not hits:
            print(f"  手動補充表：找不到「{src}／{key}」這一則")
        for i in hits:
            d = parse_manual_date(row.get("截止報名日期"))
            if d or row.get("截止報名日期") == "":
                i["deadline"] = d if d else i["deadline"]
            if row.get("截止日顯示文字"):
                i["deadline_text"] = row["截止日顯示文字"]
            if row.get("實習期間"):
                i["period"] = row["實習期間"]
            if row.get("實習職缺"):
                i["name"] = row["實習職缺"]
            if row.get("類別"):
                i["industry"] = row["類別"]
            i["current"] = True                            # 手動確認過的，不套用「太舊就不列」
    return items


# ═════════════ 10-3. 報名限制（其他系的職缺可能只開放該系學生） ═════════════
# eligibility_notes.csv 的欄位：
#   來源        例如「財金系」
#   標題關鍵字  公告標題裡有這幾個字的那一則，例如「華南期貨」
#   報名限制    顯示在職缺名稱下面的文字。以 ❌ 開頭＝已確認企管系不能報（紅色），其他＝要再確認（橘色）
#               留空＝已確認不限科系，不顯示任何標示（也不會被自動判斷標上）
# 表裡沒有的國貿、財金、經濟系職缺，程式會讀公告內文，看到下面這些字樣就自動標「⚠️ 公告可能限定科系」。
OTHER_DEPTS = ("國貿系", "財金系", "經濟系")
LIMIT_RE = re.compile(r"僅限.{0,10}系|限本系|本系學生|(?:財金|國貿|經濟)系大[三四]|東海(?:財金|國貿|經濟)系"
                      r"|需求科系|限.{0,8}相關科系|需配合系所實習課程|必須選修.{0,8}實習學分")
AUTO_LIMIT = "⚠️ 公告可能限定科系或要配合該系實習課，請先洽詢公司是否開放企管系"


def load_eligibility():
    for folder in (os.getcwd(), os.path.dirname(os.path.abspath(__file__))):
        path = os.path.join(folder, ELIGIBILITY_CSV)
        if os.path.exists(path):
            with open(path, encoding="utf-8-sig", newline="") as f:
                return [{k.strip(): (v or "").strip() for k, v in r.items() if k} for r in csv.DictReader(f)]
    return []


def apply_eligibility(items):
    """先套用人工確認的報名限制，其他系的職缺再用內文關鍵字自動判斷。"""
    rows = load_eligibility()
    for i in items:
        i["limit"] = ""
        hit = next((r for r in rows if r.get("標題關鍵字") and (not r.get("來源") or r["來源"] == i["source"])
                    and r["標題關鍵字"] in i["title"] + i.get("name", "")), None)
        if hit:
            i["limit"] = hit.get("報名限制", "")
        elif i["source"] in OTHER_DEPTS:
            text = i.get("text") or ""
            if any(text[max(m.start() - 1, 0)] != "不" for m in LIMIT_RE.finditer(text)):
                i["limit"] = AUTO_LIMIT
    return items


def limit_html(item):
    text = item.get("limit") or ""
    if not text:
        return ""
    return f'<span class="limit {"no" if text.startswith("❌") else "warn"}">{html.escape(text)}</span>'


def collect():
    items = []
    for name, crawl in SOURCES:
        try:
            items += crawl()
        except Exception as e:                             # 一個網站失敗不影響其他網站
            print(f"[{name}] 擷取失敗，先略過：{e}")
    nospace = lambda t: re.sub(r"\s+", "", t)              # 比對時忽略空白（有些標題用全形或不換行空白）
    items = [i for i in items if not any(nospace(w) in nospace(i["title"]) for w in EXCLUDE_TITLE_WORDS)]
    print(f"\n共 {len(items)} 則公告，開始補詳細內容與報名截止日…")
    items = [add_details(i) for i in items]
    items = apply_eligibility(apply_manual(items))
    items = [i for i in items
             if not (i["industry"] in EXCLUDE_INDUSTRIES and i["source"] != "企管系")]
    kept = [i for i in items if should_list(i)]
    # 排序：可報名（快截止的在最上面）→ 未標示截止日（新的在前）→ 已截止（剛截止的在前）
    order = {"open": 0, "unknown": 1, "closed": 2}

    def sort_key(i):
        state = status_of(i)
        if state == "open":
            return (0, i["deadline"].toordinal())
        if state == "unknown":
            return (1, -(i["pub"] or date.min).toordinal())
        return (2, -i["deadline"].toordinal())
    kept.sort(key=sort_key)
    counts = {s: sum(status_of(i) == s for i in kept) for s in order}
    print(f"列出 {len(kept)} 則：可報名 {counts['open']}、未標示截止日 {counts['unknown']}、已截止 {counts['closed']}")
    return kept


# ═════════════ 11. 產生頁面：只列「實習職缺、職缺來源（附網址）、截止報名日期」 ═════════════
SOURCE_LINKS = [("企管系", BA_CATEGORY_URL), ("國貿系", INTTRADE_PAGE_URL), ("財金系", FIN_LIST_URL),
                ("經濟系", ECON_URL), ("實習中心", ISAC_LIST_URL.format(page=1)),
                ("GIFT實習平台", GIFT_LIST_URL.format(page=1))]


# 外部實習平台：這些網站需要登入、禁止程式抓取，或內容由網頁程式產生，所以只放網址與說明。
# 以後要新增網站，照格式加一行：(名稱, 網址, 這個網站在做什麼, 適合找什麼, [(按鈕文字, 網址), ...])
INTERN_SITES = [
    ("LinkedIn（領英）", "https://www.linkedin.com/jobs/",
     "國際性的求職與職涯社群平台，可以建立英文個人檔案、追蹤公司，外商與跨國企業常在這裡刊登實習。",
     "外商行銷、業務開發、人資、財務實習；科技業的專案管理、供應鏈、數據分析；四大會計師事務所。多數要英文履歷。",
     [("台灣・實習職缺", "https://www.linkedin.com/jobs/search/?keywords=%E5%AF%A6%E7%BF%92&location=Taiwan&f_E=1"),
      ("台中・Intern", "https://www.linkedin.com/jobs/search/?keywords=intern&location=Taichung%2C%20Taiwan&f_E=1"),
      ("商管類・Business Intern", "https://www.linkedin.com/jobs/search/?keywords=business%20intern&location=Taiwan&f_E=1")]),
    ("104 實習專區", "https://www.104.com.tw/jobs/main/students?tab=intern",
     "台灣最大的人力銀行 104 為學生設的實習專區，可依地區、產業、寒暑假或學期實習篩選。",
     "本土企業各類實習，職缺數量最多，台中的職缺也不少。", []),
    ("1111 人力銀行", "https://www.1111.com.tw/search/job?ks=%E5%AF%A6%E7%BF%92",
     "台灣主要人力銀行之一，連結是以「實習」為關鍵字的搜尋結果。",
     "本土企業實習、工讀與兼職，服務業與中南部職缺較多。", []),
    ("Cake", "https://www.cake.me/jobs",
     "求職與作品集平台，可以做線上履歷，職缺以新創與科技公司為主。",
     "新創公司的行銷、產品、營運、設計實習，可搜尋「intern」或「實習」。", []),
    ("Yourator", "https://www.yourator.co/jobs",
     "專注新創與科技產業的徵才平台，每則職缺都有詳細的公司介紹。",
     "新創公司的業務開發、行銷企劃、數據分析實習，可搜尋「實習」。", []),
    ("RICH 職場體驗網", "https://rich.yda.gov.tw/",
     "教育部青年發展署的職場體驗平台，辦理大專生公部門見習、暑期社區職場體驗等計畫。",
     "公部門與非營利組織的見習、暑期職場體驗，多有計畫補助。", []),
    ("GIFT 全球專業實習聯盟平臺", "https://www.gift.org.tw/index.php/welcome/jobs?s_1=2",
     "教育部補助的實習媒合平台，分寒暑假、學期間、海外實習等類別。",
     "本頁「校內實習職缺」已自動整理還沒開始的學期間實習；寒暑假與海外實習請到網站查看。", []),
]

# 正職網站：只放網址與說明
JOB_SITES = [
    ("東海大學人才網", "https://jobs.thu.edu.tw/",
     "東海大學的求才求職平台，企業直接對東海學生與校友刊登正職、兼職與實習職缺。",
     "想找重視東海學生的企業，畢業前後都適用。網站禁止程式抓取，請直接到網站查看。", []),
    ("台灣就業通", "https://job.taiwanjobs.gov.tw/",
     "勞動部的官方求職網站，職缺來自各地就業服務站與企業，也會公告各地徵才活動。",
     "各地區正職、徵才博覽會資訊，還有政府的就業補助與職訓課程。", []),
    ("事求人（公部門徵才）", "https://www.dgpa.gov.tw/submenu?uid=33",
     "行政院人事行政總處的機關徵才系統，各政府機關公開徵求約聘、約僱人員都刊登在這裡。",
     "公部門約聘僱工作，不一定需要考取國考。", []),
    ("104 人力銀行", "https://www.104.com.tw/",
     "台灣最大的人力銀行，正職職缺數量最多，可設定條件訂閱新職缺通知。",
     "各產業正職、新鮮人職缺；也可以看公司評價與薪資資訊。", []),
    ("1111 人力銀行", "https://www.1111.com.tw/",
     "台灣主要人力銀行之一，有正職、兼職與新鮮人專區。",
     "各產業正職與兼職，中南部職缺較多。", []),
]


SITE_COLORS = ["#1e6f9e"]   # 卡片主色：企管系藍（「適合找」區塊用企管系橘）


def badge_text(name):
    """卡片左上角圓章的字：LinkedIn→in、104…→104、其他取第一個字。"""
    if name.startswith("LinkedIn"):
        return "in"
    m = re.match(r"\d+", name)
    return m.group() if m else name[0]


def site_cards(sites):
    """網站卡片：標題就是連到來源網址的超連結，下面是網站說明、適合找什麼、快速搜尋按鈕。"""
    out = []
    for n, (name, url, what, good_for, buttons) in enumerate(sites):
        color = SITE_COLORS[n % len(SITE_COLORS)]
        host = urlsplit(url).netloc.replace("www.", "")
        btns = "".join(f'<a class="btn" href="{html.escape(u, quote=True)}" target="_blank" rel="noopener">{html.escape(t)}</a>'
                       for t, u in buttons)
        out.append(f"""      <div class="site" style="--c: {color}">
        <div class="site-head"><span class="badge" aria-hidden="true">{html.escape(badge_text(name))}</span>
        <h3><a href="{html.escape(url, quote=True)}" target="_blank" rel="noopener">{html.escape(name)}</a> <span class="host">{html.escape(host)}</span></h3></div>
        <p>{html.escape(what)}</p>
        <p class="good"><b>適合找</b>{html.escape(good_for)}</p>{f'<p>{btns}</p>' if btns else ''}
      </div>""")
    return "\n".join(out)


SIMPLE_PAGE = """<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>東海實習職缺一覽</title>
<style>
  /* 所有樣式都限定在 #thu-intern 裡面：整段貼進 WordPress、放進 iframe 或單獨開啟，卡片樣式都一樣，也不會影響網站其他地方 */
  /* 配色取自東海大學企管系系徽：主色藍 #1E6F9E、輔助色橘 #F19F33 */
  body.thu-page {{ margin: 0; background: #1e6f9e; }}
  #thu-intern {{ --ba-blue: #1e6f9e; --ba-blue-dark: #155279; --ba-orange: #f19f33; --head: #134a6c; --line: #dbe6ee; --text: #1f2a33; --sub: #5f6b76;
                display: block; box-sizing: border-box; max-width: 1100px; margin: 0 auto; text-align: left;
                font-family: "Microsoft JhengHei", "PingFang TC", sans-serif; font-size: clamp(15px, 0.55vw + 10px, 20px);   /* 螢幕越寬字越大 */
                line-height: 1.55; color: var(--text); background: var(--ba-blue); border-radius: 10px;
                padding: 1.5em max(16px, env(safe-area-inset-left, 0px)); }}
  #thu-intern *, #thu-intern *::before, #thu-intern *::after {{ box-sizing: border-box; }}
  #thu-intern p {{ margin: 0; padding: 0; }}
  #thu-intern a {{ text-decoration: underline; box-shadow: none; }}
  #thu-intern h1.ti-title {{ font-family: inherit !important; font-size: 1.5em !important; line-height: 1.3 !important; font-weight: 700; letter-spacing: normal !important; text-transform: none !important;
                            margin: 0 0 6px !important; padding: 0 0 6px; color: #fff !important; background: none; border: none; border-bottom: 4px solid var(--ba-orange); display: inline-block; }}
  #thu-intern .meta {{ color: #dcebf5; font-size: .88em; margin: 0 0 14px; }}
  #thu-intern .lead {{ color: #eaf3f9; font-size: .92em; margin: 0 0 10px; }}
  /* 頁首說明 */
  #thu-intern .notice {{ background: #fff; color: var(--text); border-radius: 8px; border-left: 5px solid var(--ba-orange); padding: 10px 16px; margin: 0 0 14px; font-size: .9em; }}
  #thu-intern .notice p {{ margin: 4px 0; }}
  #thu-intern .notice b {{ color: var(--head); }}
  #thu-intern .notice details {{ margin: 6px 0 0; padding: 6px 0 0; border: none; border-top: 1px dashed #d6e2ea; background: none; }}
  #thu-intern .notice summary {{ display: list-item; cursor: pointer; font-weight: 700; color: var(--ba-blue); padding: 4px 0; margin: 0; background: none; border: none; font-size: 1em; }}
  #thu-intern .notice summary::after {{ content: "（點一下展開）"; font-weight: 400; color: var(--sub); font-size: .85em; }}
  #thu-intern .notice details[open] summary::after {{ content: "（收合）"; }}
  /* 篩選列 */
  #thu-intern .bar {{ color: #fff; display: flex; flex-wrap: wrap; gap: 10px 16px; align-items: center; margin: 0 0 12px; font-size: .92em; }}
  #thu-intern .bar label {{ display: inline-flex; align-items: center; gap: 6px; margin: 0; color: #fff; font-size: 1em; font-weight: 400; }}
  #thu-intern .bar select {{ font: inherit; font-size: .95em; width: auto; height: auto; min-height: 0; margin: 0; padding: 4px 8px; line-height: 1.4;
                            border: 1px solid #b9cfdf; border-radius: 6px; background: #fff; color: var(--text); box-shadow: none; }}
  #thu-intern .bar .f-q {{ flex: 1 1 100%; width: 100%; min-width: 0; height: auto; margin: 0; font: inherit; font-size: 1em; line-height: 1.4; padding: .45em .7em;
                          border: 2px solid #b9cfdf; border-radius: 8px; background: #fff; color: var(--text); box-shadow: none; }}
  #thu-intern .bar .f-q:focus {{ outline: 3px solid var(--ba-orange); outline-offset: 0; border-color: var(--ba-orange); }}
  #thu-intern .bar input[type=checkbox] {{ accent-color: var(--ba-orange); width: 1.15em; height: 1.15em; margin: 0; padding: 0; vertical-align: -.2em; appearance: auto; position: static; opacity: 1; }}
  #thu-intern .bar .f-shown {{ color: #dcebf5; }}
  /* 卡片：不論放在哪裡、視窗寬窄、放大縮小，都維持一則一張卡片（加 !important 是為了不被網站主題的表格樣式蓋掉） */
  #thu-intern table, #thu-intern tbody, #thu-intern tr, #thu-intern td {{ display: block !important; width: 100% !important; max-width: 100% !important;
                                                                        float: none !important; position: static !important; height: auto !important; }}
  #thu-intern thead {{ display: none !important; }}
  #thu-intern table {{ border: none !important; border-collapse: separate !important; border-spacing: 0 !important; margin: 0 !important; padding: 0 !important;
                      background: transparent !important; box-shadow: none !important; table-layout: auto !important; font-size: .93em; }}
  #thu-intern tbody {{ border: none !important; background: transparent !important; margin: 0 !important; padding: 0 !important; }}
  #thu-intern tr {{ background: #fff !important; border: none !important; border-left: 4px solid var(--ba-orange) !important; border-radius: 8px !important;
                   padding: 6px 0 !important; margin: 0 0 8px !important; box-shadow: none !important; }}
  #thu-intern tr.closed {{ border-left-color: #b0b6bd !important; }}
  #thu-intern tr:hover {{ background: #f4f8fb !important; }}
  #thu-intern tr[hidden] {{ display: none !important; }}
  #thu-intern td {{ border: none !important; background: transparent !important; padding: 3px 12px !important; margin: 0 !important;
                   text-align: left !important; vertical-align: top !important; white-space: normal !important; color: var(--text); font-size: 1em; line-height: 1.55; }}
  #thu-intern td::before {{ content: attr(data-label) "："; color: var(--sub); font-size: .88em; font-weight: 400; }}
  #thu-intern tr.closed td, #thu-intern tr.closed td a {{ color: #5f6670 !important; }}
  #thu-intern td a {{ color: var(--ba-blue); font-weight: 600; display: inline-block; padding: .45em .25em; margin: -.45em -.25em; background: none; border: none; }}   /* 手機上比較好點 */
  #thu-intern td a:hover {{ color: var(--ba-blue-dark); }}
  #thu-intern .tag {{ display: inline-block; padding: 0 8px; border-radius: 999px; font-size: .86em; background: #eaf2f8; color: var(--ba-blue); font-weight: 600; }}
  #thu-intern tr.closed .tag {{ background: #e4e7ea; color: #5f6670; }}
  #thu-intern .soon {{ color: #c62828; font-weight: 600; }}
  #thu-intern .open {{ color: #9a5300; font-weight: 600; }}
  #thu-intern .muted {{ color: var(--sub); }}
  #thu-intern .off {{ display: inline-block; padding: 0 6px; border-radius: 4px; background: #6b7280; color: #fff; font-size: .84em; }}
  /* 報名限制：顯示在職缺名稱下面 */
  #thu-intern .limit {{ display: block; width: fit-content; max-width: 100%; margin: 4px 0 0; padding: 2px 8px; border-radius: 6px; font-size: .88em; font-weight: 600; line-height: 1.5; }}
  #thu-intern .notice .limit {{ display: inline; margin: 0; }}
  #thu-intern .limit.no {{ background: #fdecea; color: #b71c1c; }}
  #thu-intern .limit.warn {{ background: #fff4e5; color: #8a4f00; }}
</style>
</head>
<body class="thu-page">
<main id="thu-intern" class="thu-intern">
  <h1 class="ti-title">東海實習職缺一覽</h1>
  <div class="meta">更新時間 {updated}・每則資訊都附上來源網址，報名前請以原始公告為準</div>
  <section class="panel" id="intern" role="tabpanel">
    <p class="lead">整理企管、國貿、財金、經濟系、實習與學生成就中心及 GIFT 的實習公告，每天自動更新。共 {count} 則（可報名 {n_open}、未標示截止日 {n_unknown}、已截止 {n_closed}）。</p>
    <div class="notice">
      <p><b>⚠️ 報名前請注意：</b>國貿、財金、經濟系公告的職缺，有些只開放該系學生。請先看原公告的資格，並<b>洽詢該公司或職缺窗口是否開放企管系學生報名</b>。標 <span class="limit no" >❌</span> 的已確認限定科系，標 <span class="limit warn" >⚠️</span> 的資格需要再確認。國貿、財金、經濟系的職缺如果<b>沒有標示</b>，代表公司在該系公告上沒有寫出限制條件，但不代表一定開放，仍請先洽詢公司，錄取後再向企管系辦確認學分認列。</p>
      <details>
        <summary>📌 為什麼也放其他系的職缺？為什麼同一個職缺會出現好幾次？</summary>
        <p><b>為什麼放其他系的職缺：</b>大部分商管類實習不限科系，只看企管系的公告會錯過很多機會，所以這裡一併整理管理學院各系、實習與學生成就中心和 GIFT 平台的實習。</p>
        <p><b>為什麼不合併重複的職缺：</b>同一家公司常在好幾個系同時公告，但各系公告的報名資格、名額、截止日和學分認列可能不同，例如有的只限該系學生。為了避免合併後漏掉這些差異，每個系的公告都分開列出，方便對照自己能報哪一則。</p>
        <p><b>實習學分：</b>公告寫的「可認列系上實習學分」是指該系的學分，想抵企管系的實習學分，請先問企管系辦。</p>
      </details>
    </div>
    <div class="bar" data-list="intern-list">
      <input type="search" class="f-q" id="intern-q" placeholder="搜尋職缺名稱、公司或關鍵字，例如：銀行、行銷、台中">
      <label>來源 <select class="f-src" id="intern-src"><option value="">全部（{count}）</option>
{options}
      </select></label>
      <label>類別 <select class="f-ind" id="intern-ind"><option value="">全部（{count}）</option>
{ind_options}
      </select></label>
      <label><input type="checkbox" class="f-hide" id="intern-hide"> 隱藏已截止</label>
      <span class="f-shown muted"></span>
    </div>
    <table>
      <thead><tr><th>實習職缺</th><th>職缺來源</th><th>截止報名日期</th><th>類別</th></tr></thead>
      <tbody id="intern-list">
{rows}
      </tbody>
    </table>
  </section>

</main>
<script>
  // 篩選：每個分頁的篩選列只控制自己的清單
  document.querySelectorAll("#thu-intern .bar[data-list]").forEach(function (bar) {{
    var list = document.getElementById(bar.dataset.list),
        src = bar.querySelector(".f-src"), ind = bar.querySelector(".f-ind"),
        hide = bar.querySelector(".f-hide"), shown = bar.querySelector(".f-shown"), q = bar.querySelector(".f-q"),
        rows = Array.prototype.slice.call(list.querySelectorAll(":scope > tr[data-source]"));
    rows.forEach(function (r) {{ r._text = r.textContent.replace(/\\s+/g, " ").toLowerCase(); }});
    function apply() {{
      var n = 0, keep = [], drop = [], words = q.value.trim().toLowerCase().split(/\\s+/).filter(Boolean);
      rows.forEach(function (r) {{
        var ok = words.every(function (w) {{ return r._text.indexOf(w) >= 0; }}) &&
                 (!src.value || r.dataset.source === src.value) &&
                 (!ind.value || r.dataset.industry === ind.value) &&
                 (!hide.checked || r.dataset.status !== "closed");
        r.hidden = !ok; if (ok) {{ n++; keep.push(r); }} else drop.push(r);
      }});
      keep.concat(drop).forEach(function (r) {{ list.appendChild(r); }});   // 符合條件的排在最上面
      shown.textContent = "顯示 " + n + " 則";
    }}
    src.onchange = ind.onchange = hide.onchange = apply; q.oninput = apply; apply();
  }});
</script>
</body>
</html>
"""

SIMPLE_ROW = """      <tr data-source="{source}" data-industry="{industry}" data-status="{status}"{cls}>
        <td data-label="實習職缺">{name}{limit}</td>
        <td data-label="職缺來源"><a href="{url}" target="_blank" rel="noopener">{source}</a></td>
        <td data-label="截止報名日期">{deadline}</td>
        <td data-label="類別"><span class="tag">#{industry}</span></td>
      </tr>"""


def simple_deadline(item):
    d, state = item["deadline"], status_of(item)
    if state == "unknown":
        return '<span class="muted">未標示</span>'
    if state == "closed":
        return f'{html.escape(item.get("deadline_text") or d.isoformat())} <span class="off">已截止</span>'
    left = (d - TODAY).days
    shown = html.escape(item.get("deadline_text") or d.isoformat())
    about = "約" if item.get("deadline_text") else ""        # 手動填的說明文字（例如「11 月中前」）只能估算天數
    return f'<span class="{"soon" if left <= 7 else "open"}">{shown}（{about}剩 {left} 天）</span>'


def build_simple_html(all_items):
    items = [i for i in all_items if i.get("kind", "intern") == "intern"]
    jobs = [i for i in all_items if i.get("kind") == "fulltime"]
    job_rows = [SIMPLE_ROW.format(
        source=html.escape(i["source"]), status=status_of(i),
        cls=' class="closed"' if status_of(i) == "closed" else "",
        name=html.escape(i["name"]), url=html.escape(i["url"], quote=True),
        industry=html.escape(i["industry"]), limit="",
        deadline=simple_deadline(i)).replace('data-label="實習職缺"', 'data-label="職缺"').replace(
            f'<td data-label="職缺">{html.escape(i["name"])}</td>',
            f'<td data-label="職缺"><a class="title" href="{html.escape(i["url"], quote=True)}" target="_blank" rel="noopener">{html.escape(i["name"])}</a></td>')
        for i in jobs]
    opt = lambda pool, key, names, pre="": "\n".join(
        f'      <option value="{html.escape(n)}">{pre}{html.escape(n)}（{sum(i[key] == n for i in pool)}）</option>'
        for n in names if any(i[key] == n for i in pool))
    job_options = opt(jobs, "source", [n for n, _ in SOURCE_LINKS])
    job_ind_options = opt(jobs, "industry", INDUSTRY_NAMES, "#")
    rows = [SIMPLE_ROW.format(
        source=html.escape(i["source"]), status=status_of(i),
        cls=' class="closed"' if status_of(i) == "closed" else "",
        name=html.escape(i["name"]), url=html.escape(i["url"], quote=True),
        industry=html.escape(i["industry"]), limit=limit_html(i),
        deadline=simple_deadline(i)) for i in items]
    options = "\n".join(
        f'      <option value="{html.escape(n)}">{html.escape(n)}（{sum(i["source"] == n for i in items)}）</option>'
        for n, _ in SOURCE_LINKS if any(i["source"] == n for i in items))
    ind_options = "\n".join(
        f'      <option value="{html.escape(n)}">#{html.escape(n)}（{sum(i["industry"] == n for i in items)}）</option>'
        for n in INDUSTRY_NAMES if any(i["industry"] == n for i in items))
    return SIMPLE_PAGE.format(
        count=len(items), options=options, ind_options=ind_options, rows="\n".join(rows),
        job_count=len(jobs), job_options=job_options, job_ind_options=job_ind_options, job_rows="\n".join(job_rows) or '      <tr><td>目前沒有正職公告</td></tr>',
        intern_sites=site_cards(INTERN_SITES), job_sites=site_cards(JOB_SITES),
        n_open=sum(status_of(i) == "open" for i in items),
        n_unknown=sum(status_of(i) == "unknown" for i in items),
        n_closed=sum(status_of(i) == "closed" for i in items),
        updated=datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M"))


def write_csv(items, path):
    """輸出 CSV（Excel 可直接開啟，中文不會變亂碼）。"""
    import csv
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["類型", "職缺", "產業類別", "職缺來源", "來源網址", "截止報名日期", "狀態", "報名限制"])
        label = {"open": "可報名", "closed": "已截止", "unknown": "未標示截止日"}
        for i in items:
            w.writerow(["正職" if i.get("kind") == "fulltime" else "實習", i["name"], i["industry"], i["source"], i["url"],
                        i.get("deadline_text") or (i["deadline"].isoformat() if i["deadline"] else ""),
                        label[status_of(i)], i.get("limit", "")])


if __name__ == "__main__":
    data = collect()
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(build_simple_html(data))
    write_csv(data, CSV_OUTPUT)
    print(f"\n已輸出到 {OUTPUT} 與 {CSV_OUTPUT}")
