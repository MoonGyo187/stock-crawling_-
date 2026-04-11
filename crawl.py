import requests
import pandas as pd
import time
import re
import json
import os
import sys
from bs4 import BeautifulSoup

# ============================
# ⚙️ 설정
# ============================
STOCK_CODE = sys.argv[1]  # 실행 시 종목코드 받기
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)
EXISTING_FILE = f"{DATA_DIR}/naver_board_{STOCK_CODE}.xlsx"
DELAY = 0.5

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://finance.naver.com/",
}
API_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://m.stock.naver.com/",
    "Accept": "application/json",
}

session = requests.Session()
session.get("https://finance.naver.com", headers=HEADERS)
session.get(f"https://finance.naver.com/item/board.naver?code={STOCK_CODE}", headers=HEADERS)

# ============================
# 📂 기존 파일 로드
# ============================
existing_nids = set()
df_old = pd.DataFrame()

if os.path.exists(EXISTING_FILE):
    df_old = pd.read_excel(EXISTING_FILE)
    if "nid" in df_old.columns:
        existing_nids = set(df_old["nid"].astype(str))
    print(f"✅ 기존 데이터 {len(df_old)}개 로드")
else:
    print("✅ 새로 시작")

# ============================
# 🔍 목록 파싱
# ============================
def parse_posts(soup):
    posts = []
    for row in soup.select("table.type2 tr"):
        cols = row.select("td")
        if len(cols) < 5:
            continue
        date_text = cols[0].get_text(strip=True)
        if not date_text or not re.match(r'\d{4}', date_text):
            continue
        a_tag = cols[1].select_one("a")
        if not a_tag:
            continue
        nid_match = re.search(r'nid=(\d+)', a_tag["href"])
        if not nid_match:
            continue
        posts.append({
            "날짜시간": date_text,
            "제목": a_tag.get_text(strip=True),
            "nid": int(nid_match.group(1)),
        })
    return posts

# ============================
# 📄 본문 API
# ============================
def get_post_detail(nid):
    try:
        url = f"https://m.stock.naver.com/front-api/discussion/detail?id={nid}"
        res = session.get(url, headers=API_HEADERS)
        data = res.json()
        if not data.get("isSuccess"):
            return "", 0, 0, 0
        result = data["result"]
        content = ""
        try:
            content_json = json.loads(result.get("contentJsonSwReplaced", "{}"))
            components = content_json.get("document", {}).get("components", [])
            texts = []
            for component in components:
                for paragraph in component.get("value", []):
                    for node in paragraph.get("nodes", []):
                        val = node.get("value", "")
                        if val:
                            texts.append(val)
            content = "\n".join(texts)
        except:
            content_html = result.get("contentHtml", "")
            soup = BeautifulSoup(content_html, "html.parser")
            content = soup.get_text(separator="\n", strip=True)
        return content, result.get("viewCount", 0), result.get("recommendCount", 0), result.get("notRecommendCount", 0)
    except:
        return "", 0, 0, 0

# ============================
# 🚀 실행
# ============================
new_posts = []
stop_flag = False

for page in range(1, 101):
    print(f"📄 {page}페이지 수집 중...")
    url = f"https://finance.naver.com/item/board.naver?code={STOCK_CODE}&page={page}"
    res = session.get(url, headers=HEADERS)
    soup = BeautifulSoup(res.content, "html.parser", from_encoding="cp949")
    posts = parse_posts(soup)

    if not posts:
        print("  → 게시글 없음. 종료.")
        break

    for post in posts:
        if str(post["nid"]) in existing_nids:
            print(f"  🛑 기존 데이터 발견! 중단 ({post['날짜시간']} | {post['제목'][:20]})")
            stop_flag = True
            break

        content, views, recommend, not_recommend = get_post_detail(post["nid"])
        post["내용"] = content
        post["추천수"] = recommend
        post["비추천수"] = not_recommend
        post["조회수"] = views
        post["제목+내용"] = post["제목"] + " " + content
        new_posts.append(post)
        print(f"  ✅ {post['날짜시간']} | {post['제목'][:20]}")
        time.sleep(DELAY)

    if stop_flag:
        break

    time.sleep(DELAY)

# ============================
# 📊 저장
# ============================
if new_posts:
    df_new = pd.DataFrame(new_posts, columns=[
        "날짜시간", "제목", "내용", "추천수", "비추천수", "조회수", "제목+내용", "nid"
    ])

    if not df_old.empty:
        df_final = pd.concat([df_new, df_old]).drop_duplicates("nid").reset_index(drop=True)
    else:
        df_final = df_new

    df_final = df_final.sort_values("날짜시간", ascending=False).reset_index(drop=True)
    df_final.to_excel(EXISTING_FILE, index=False, engine="openpyxl")
    print(f"\n✅ 새로 추가: {len(new_posts)}개")
    print(f"✅ 전체 누적: {len(df_final)}개")
    print(f"💾 저장: {EXISTING_FILE}")
else:
    print("\n✅ 새로운 게시글 없음!")
