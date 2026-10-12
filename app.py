import streamlit as st
import requests
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import xml.etree.ElementTree as ET

st.set_page_config(page_title="表演藝術演出節目統計系統", page_icon="🎭", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 26px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .sub-title { font-size: 14px; color: #4B5563; margin-bottom: 20px; }
    .stTabs [data-baseweb="tab-list"] { gap: 15px; }
    .stTabs [data-baseweb="tab"] { font-size: 16px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🎭 表演藝術與展演售票節目統計系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">支援 OPENTIX、UDN、寬宏、年代、KKTIX、ACCUPASS 官網直連統計</div>', unsafe_allow_html=True)

PLATFORM_OPTIONS = {
    "🎟️ OPENTIX (官網直連)": "OPENTIX",
    "🎪 udn 售票網 (官網直連)": "UDN",
    "⚡ KKTIX (活動通道)": "KKTIX",
    "🎸 寬宏售票 KHAM (官網直連)": "KHAM",
    "🎟️ 年代售票 ERA (官網直連)": "ERA",
    "🎫 ACCUPASS (活動通專屬)": "ACCUPASS"
}

CATEGORY_MAP_UDN = {
    "戲劇": "116",
    "音樂": "77",
    "舞蹈": "100",
    "親子": "129"
}

BUILTIN_FESTIVALS = [
    "TIFA", "NTT-TIFA", "KSAF", "TTTF", "Fringe", "藝穗節", "台北藝術節", "臺北藝術節",
    "新點子實驗場", "衛武營小時光", "夏日童樂會", "國際劇場藝術節", "臺灣戲曲藝術節",
    "大稻埕戲苑傳統戲曲季", "新北兒童藝術節", "夏至藝術節", "秋天藝術節", "高雄春天藝術節",
    "台南藝術節", "臺南藝術節", "花蓮太平洋左岸藝術節", "嘉義市瘋科技藝術節"
]

WELL_KNOWN_GROUPS = [
    "果陀劇場", "綠光劇團", "明華園", "唐美雲歌仔戲團", "國光劇團", "當代傳奇劇場",
    "故事工廠", "春河劇團", "台南人劇團", "全民大劇團", "相聲瓦舍", "無垢舞蹈劇場",
    "雲門舞集", "布拉瑞揚舞團", "國家交響樂團", "台北市立交響樂團", "高雄市交響樂團",
    "台北民族舞團", "薪傳歌仔戲劇團", "一心戲劇團", "秀琴歌劇團", "紙風車劇團",
    "如果兒童劇團", "偶偶偶劇團", "鞋子兒童實驗劇團", "卡米地", "薩泰爾", "漫才少爺",
    "FOCA福爾摩沙馬戲團", "驚喜製造", "衛武營", "國家兩廳院", "臺中國家歌劇院"
]

def extract_festival(title, desc="", user_keywords=[]):
    combined = f"{title} {desc[:120]}"
    for kw in user_keywords:
        kw = kw.strip()
        if kw and (kw.lower() in combined.lower()):
            return kw
    for fest in BUILTIN_FESTIVALS:
        pattern = re.compile(rf"(?:20\d\d\s*)?{re.escape(fest)}", re.IGNORECASE)
        m = pattern.search(title)
        if m:
            return m.group(0).strip()
        if fest.lower() in combined.lower():
            return fest
    symbol_match = re.search(r"^[【\[［]([^】\]］]+)[】\]］]", title)
    if symbol_match:
        content = symbol_match.group(1).strip()
        if any(term in content for term in ["節", "季", "祭", "週", "Festival", "系列", "展演"]):
            return content
    dash_match = re.search(r"^([^－—\-]+)[－—\-]", title)
    if dash_match:
        content = dash_match.group(1).strip()
        if any(term in content for term in ["節", "季", "祭", "週", "Festival", "小時光", "童樂會"]):
            if len(content) <= 25:
                return content
    suffix_pattern = r"(?:20\d\d\s*)?[\u4e00-\u9fa5A-Za-z0-9]{2,12}(?:藝術節|藝術季|戲劇節|舞蹈節|音樂節|音樂祭|文化季|藝穗節|生活節|劇節|舞季|戲曲季|戲曲節|兒童節|Festival)"
    m_suffix = re.search(suffix_pattern, title, re.IGNORECASE)
    if m_suffix:
        found = m_suffix.group(0).strip()
        if len(found) >= 4:
            return found
    return "單檔演出 (一般檔期)"

def extract_performer(show_title):
    match = re.search(r"[【\[［]([^】\]］]+)[】\]］]", show_title)
    if match:
        candidate = match.group(1).strip()
        if any(term in candidate for term in ["劇團", "舞團", "樂團", "合唱團", "工作室", "劇會", "劇坊", "劇社", "製作", "馬戲"]):
            return candidate

    match_book = re.search(r"^([^\d《]+)《", show_title)
    if match_book:
        candidate = match_book.group(1).strip()
        if len(candidate) <= 15 and any(term in candidate for term in ["團", "社", "坊", "會", "卡米地"]):
            return candidate

    for group in WELL_KNOWN_GROUPS:
        if group in show_title:
            return group

    return "詳見售票頁"

with st.sidebar:
    st.header("⚙️ 統計篩選條件")
    selected_platform_label = st.selectbox("1. 選擇售票平台：", options=list(PLATFORM_OPTIONS.keys()), index=0)
    target_platform = PLATFORM_OPTIONS[selected_platform_label]

    selected_cats = st.multiselect(
        "2. 選擇節目種類：",
        options=["戲劇", "音樂", "舞蹈", "親子", "戲曲", "其他"],
        default=["戲劇", "音樂", "舞蹈", "親子"]
    )

    today = datetime.date.today()
    default_end = today + datetime.timedelta(days=180)

    st.write("3. 演出日期區間：")
    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("開始日期", today)
    with c2:
        end_date = st.date_input("結束日期", default_end)

    st.write("4. 補充藝術節/季關鍵字 (選填)：")
    custom_festivals_input = st.text_input("以逗號分隔，如：Sound Wave, 某某祭", "")
    custom_fest_list = [k.strip() for k in custom_festivals_input.split(",") if k.strip()]

    start_btn = st.button("🚀 開始擷取並產生統計", type="primary", use_container_width=True)

# ----------------- 各平台直連模組 -----------------

@st.cache_data(ttl=1800, show_spinner=False)
def fetch_opentix():
    api_url = "https://csm.api.opentix.life/programs"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.opentix.life/"}
    all_programs = []
    for page in range(1, 11):
        try:
            r = requests.get(api_url, headers=headers, params={"page": page, "limit": 25}, timeout=12)
            if r.status_code == 200:
                items = r.json().get("result", {}).get("data", [])
                if not items: break
                all_programs.extend(items)
            else: break
        except Exception: break
    return all_programs

def fetch_udn(selected_categories):
    udn_records = []
    base_url = "https://tickets.udnfunlife.com"
    headers = {"User-Agent": "Mozilla/5.0"}
    for cat_name in ["戲劇", "音樂", "舞蹈", "親子"]:
        if cat_name in selected_categories:
            cat_id = CATEGORY_MAP_UDN[cat_name]
            url = f"https://tickets.udnfunlife.com/application/UTK01/UTK0101_03.aspx?Category={cat_id}&kdid=cateList"
            try:
                r = requests.get(url, headers=headers, timeout=12)
                if r.status_code == 200:
                    soup = BeautifulSoup(r.text, "lxml")
                    for a in soup.find_all("a", href=re.compile(r"PRODUCT_ID=")):
                        full_link = urljoin(base_url, a.get("href", ""))
                        raw_text = a.get_text(" ", strip=True)
                        if len(raw_text) < 4: continue
                        date_match = re.search(r"(\d{4}[./-]\d{1,2}[./-]\d{1,2}(?:\s*~\s*\d{4}[./-]\d{1,2}[./-]\d{1,2})?)", raw_text)
                        date_str = date_match.group(1).strip() if date_match else "即期 (詳見售票頁)"
                        title = raw_text[:date_match.span()[0]].strip() if date_match else raw_text
                        udn_records.append({
                            "售票通路": "UDN 售票網",
                            "節目分類": cat_name,
                            "節目名稱": title,
                            "演出場次時間": date_str,
                            "演出地點": "詳見售票頁",
                            "售票狀況/連結": full_link
                        })
            except Exception: pass
    return udn_records

def fetch_kham():
    records = []
    headers = {"User-Agent": "Mozilla/5.0"}
    urls = [
        ("https://kham.com.tw/application/UTK01/UTK0101_06.aspx?TYPE=1&CATEGORY=205", "音樂/戲劇"),
        ("https://kham.com.tw/application/UTK01/UTK0101_06.aspx?TYPE=1&CATEGORY=80", "藝文展演")
    ]
    for url, cat in urls:
        try:
            r = requests.get(url, headers=headers, timeout=12)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, "html.parser")
                for a in soup.find_all("a", href=re.compile(r"PRODUCT_ID=")):
                    title = a.get_text(" ", strip=True)
                    if len(title) < 2: continue
                    link = urljoin("https://kham.com.tw/application/UTK01/", a.get("href", ""))
                    records.append({
                        "售票通路": "寬宏售票",
                        "節目分類": "音樂" if "音樂" in cat else "戲劇",
                        "節目名稱": title,
                        "演出場次時間": "即期 (詳見售票頁)",
                        "演出地點": "全台巡演",
                        "售票狀況/連結": link
                    })
        except Exception: pass
    return records

def fetch_era():
    records = []
    headers = {"User-Agent": "Mozilla/5.0"}
    urls = [
        "https://ticket.com.tw/application/UTK01/UTK0101_.aspx",
        "https://ticket.com.tw/application/UTK02/UTK0201_01.aspx?CATEGORY=1"
    ]
    for url in urls:
        try:
            r = requests.get(url, headers=headers, timeout=12)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, "html.parser")
                for a in soup.find_all("a", href=re.compile(r"PRODUCT_ID=")):
                    title = a.get("title", "") or a.get_text(" ", strip=True)
                    if not title or len(title) < 2:
                        img = a.find("img")
                        if img and img.get("alt"): title = img.get("alt").strip()
                    if len(title) < 2: continue
                    link = urljoin("https://ticket.com.tw/application/UTK01/", a.get("href", ""))
                    records.append({
                        "售票通路": "年代售票",
                        "節目分類": "戲劇",
                        "節目名稱": title,
                        "演出場次時間": "即期 (詳見售票頁)",
                        "演出地點": "全台巡演",
                        "售票狀況/連結": link
                    })
        except Exception: pass
    return records

@st.cache_data(ttl=1800, show_spinner=False)
def fetch_kktix():
    records = []
    headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}
    try:
        r = requests.get("https://kktix.com/events.atom", headers=headers, timeout=12)
        if r.status_code == 200:
            root = ET.fromstring(r.content)
            ns = {'atom': 'http://www.w3.org/2005/Atom'}
            for entry in root.findall('atom:entry', ns):
                t_elem = entry.find('atom:title', ns)
                l_elem = entry.find('atom:link', ns)
                p_elem = entry.find('atom:published', ns)
                summary_elem = entry.find('atom:summary', ns)

                title = t_elem.text.strip() if t_elem is not None and t_elem.text else ""
                link = l_elem.attrib.get('href', '') if l_elem is not None else ""
                
                # 時間解析
                date_str = "即期 (詳見售票頁)"
                if p_elem is not None and p_elem.text:
                    try:
                        raw_dt = p_elem.text
                        date_str = raw_dt[:10] + " " + raw_dt[11:16]
                    except Exception:
                        pass

                # 自動依標題判斷分類
                cat = "其他"
                if any(w in title for w in ["劇", "戲曲", "讀劇", "喜劇", "脫口秀", "漫才"]):
                    cat = "戲劇"
                elif any(w in title for w in ["音樂", "演唱會", "樂團", "演奏", "聲樂", "合唱", "LIVE"]):
                    cat = "音樂"
                elif any(w in title for w in ["舞", "芭蕾", "街舞"]):
                    cat = "舞蹈"
                elif any(w in title for w in ["親子", "兒童"]):
                    cat = "親子"

                if title:
                    records.append({
                        "售票通路": "KKTIX",
                        "節目分類": cat,
                        "節目名稱": title,
                        "演出場次時間": date_str,
                        "演出地點": "詳見售票頁",
                        "售票狀況/連結": link
                    })
    except Exception:
        pass
    return records

# ----------------- 執行流程 -----------------

if start_btn:
    if not selected_cats:
        st.warning("⚠️ 請至少勾選一種節目種類！")
    elif start_date > end_date:
        st.error("❌ 開始日期不能晚於結束日期！")
    else:
        status_box = st.empty()
        pbar = st.progress(0)
        detail_records = []

        status_box.info(f"正在連線【{selected_platform_label}】官網即時擷取資料...")

        if target_platform == "OPENTIX":
            programs = fetch_opentix()
            pbar.progress(50)
            for prog in programs:
                title = prog.get("name", "").strip()
                cat = prog.get("displayCategory", "其他")
                matched_cat = "其他"
                for c in ["戲劇", "音樂", "舞蹈", "親子", "戲曲"]:
                    if c in cat:
                        matched_cat = c
                        break
                if matched_cat not in selected_cats and "其他" not in selected_cats:
                    continue

                prog_id = prog.get("id", "")
                link = f"https://www.opentix.life/event/{prog_id}" if prog_id else "https://www.opentix.life"
                cities = "、".join(prog.get("cities", [])) if prog.get("cities") else "全台巡演"
                fest = extract_festival(title, "", custom_fest_list)
                performer = extract_performer(title)
                events = prog.get("events", [])
                if events:
                    for ev in events:
                        ev_start = ev.get("startDateTime")
                        ev_str = "即期 (詳見售票頁)"
                        if ev_start:
                            try:
                                ts_v = ev_start / 1000 if ev_start > 10000000000 else ev_start
                                ev_dt = datetime.datetime.fromtimestamp(ts_v)
                                if not (start_date <= ev_dt.date() <= end_date): continue
                                ev_str = ev_dt.strftime("%Y-%m-%d %H:%M")
                            except Exception: pass
                        detail_records.append({
                            "售票通路": "OPENTIX",
                            "所屬藝術節/系列": fest,
                            "節目分類": matched_cat,
                            "演出單位": performer,
                            "節目名稱": title,
                            "演出場次時間": ev_str,
                            "演出地點": cities,
                            "售票狀況/連結": link
                        })
                else:
                    detail_records.append({
                        "售票通路": "OPENTIX",
                        "所屬藝術節/系列": fest,
                        "節目分類": matched_cat,
                        "演出單位": performer,
                        "節目名稱": title,
                        "演出場次時間": "即期 (詳見售票頁)",
                        "演出地點": cities,
                        "售票狀況/連結": link
                    })

        elif target_platform == "UDN":
            shows = fetch_udn(selected_cats)
            pbar.progress(60)
            for s in shows:
                s["所屬藝術節/系列"] = extract_festival(s["節目名稱"], "", custom_fest_list)
                s["演出單位"] = extract_performer(s["節目名稱"])
                detail_records.append(s)

        elif target_platform == "KHAM":
            shows = fetch_kham()
            pbar.progress(60)
            for s in shows:
                s["所屬藝術節/系列"] = extract_festival(s["節目名稱"], "", custom_fest_list)
                s["演出單位"] = extract_performer(s["節目名稱"])
                detail_records.append(s)

        elif target_platform == "ERA":
            shows = fetch_era()
            pbar.progress(60)
            for s in shows:
                s["所屬藝術節/系列"] = extract_festival(s["節目名稱"], "", custom_fest_list)
                s["演出單位"] = extract_performer(s["節目名稱"])
                detail_records.append(s)

        elif target_platform == "KKTIX":
            shows = fetch_kktix()
            pbar.progress(60)
            for s in shows:
                # 類別篩選
                if s["節目分類"] in selected_cats or "其他" in selected_cats:
                    s["所屬藝術節/系列"] = extract_festival(s["節目名稱"], "", custom_fest_list)
                    s["演出單位"] = extract_performer(s["節目名稱"])
                    detail_records.append(s)

        elif target_platform == "ACCUPASS":
            headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.accupass.com/"}
            for kw in ["戲劇", "音樂", "演出", "脫口秀", "舞蹈"]:
                try:
                    url = f"https://api.accupass.com/v3/search/events?keyword={kw}&page=0&size=20"
                    r = requests.get(url, headers=headers, timeout=8)
                    if r.status_code == 200:
                        for ev in r.json().get("events", []):
                            t = ev.get("title", "").strip()
                            detail_records.append({
                                "售票通路": "ACCUPASS",
                                "所屬藝術節/系列": extract_festival(t, "", custom_fest_list),
                                "節目分類": "其他",
                                "演出單位": ev.get("orgName", "主辦單位"),
                                "節目名稱": t,
                                "演出場次時間": ev.get("eventStartDateTime", "詳見售票頁"),
                                "演出地點": ev.get("address", "") or "詳見活動頁",
                                "售票狀況/連結": f"https://www.accupass.com/event/{ev.get('eventId', '')}"
                            })
                except Exception: pass
            pbar.progress(60)

        pbar.progress(85)

        if not detail_records:
            status_box.empty()
            pbar.empty()
            st.warning(f"💡 在所選的【{selected_platform_label}】平台與條件下未找到節目。建議放寬節目種類（勾選「其他」）或調整日期。")
        else:
            df_details = pd.DataFrame(detail_records)
            summary_list = []
            grouped = df_details.groupby("節目名稱")

            for prog_name, group in grouped:
                first_row = group.iloc[0]
                times = sorted(list(set(group["演出場次時間"].tolist())))
                venues = list(dict.fromkeys([v for v in group["演出地點"].tolist() if v and v != "詳見售票頁"]))
                venues_str = "、".join(venues) if venues else first_row["演出地點"]

                m_d = re.search(r"(\d{4}[./-]\d{1,2}[./-]\d{1,2})", times[0]) if times else None
                first_play_date = m_d.group(1).replace(".", "-").replace("/", "-") if m_d else "即期"
                date_range_str = f"{times[0][:10]} ~ {times[-1][:10]}" if len(times) > 1 else (times[0] if times else "詳見售票頁")
                total_shows = len(group)

                summary_list.append({
                    "售票通路": first_row["售票通路"],
                    "所屬藝術節/系列": first_row["所屬藝術節/系列"],
                    "節目分類": first_row["節目分類"],
                    "演出單位": first_row["演出單位"],
                    "節目名稱": prog_name,
                    "首演日期": first_play_date,
                    "演出日期區間": f"{date_range_str} (共 {total_shows} 場)" if total_shows > 1 else date_range_str,
                    "巡迴/演出場館": venues_str,
                    "售票狀況/主要連結": first_row["售票狀況/連結"]
                })

            df_summary = pd.DataFrame(summary_list)

            fest_first_dates = {}
            for fest, group in df_summary.groupby("所屬藝術節/系列"):
                fest_first_dates[fest] = group["首演日期"].min()

            df_summary["_fest_sort_date"] = df_summary.apply(
                lambda r: r["首演日期"] if r["所屬藝術節/系列"] == "單檔演出 (一般檔期)" else fest_first_dates.get(r["所屬藝術節/系列"], r["首演日期"]),
                axis=1
            )
            df_summary.sort_values(
                by=["_fest_sort_date", "所屬藝術節/系列", "首演日期", "節目名稱"],
                ascending=[True, True, True, True],
                inplace=True
            )
            df_summary.drop(columns=["_fest_sort_date"], inplace=True)
            df_details.sort_values(by=["演出場次時間", "所屬藝術節/系列", "節目名稱"], inplace=True)

            pbar.progress(95)
            status_box.info("正在產出 Excel 報表...")

            output = io.BytesIO()
            wb = openpyxl.Workbook()

            ws1 = wb.active
            ws1.title = "演出節目總表(合併)"
            ws1.append(list(df_summary.columns))
            for r in df_summary.to_dict(orient="records"):
                ws1.append(list(r.values()))

            ws2 = wb.create_sheet(title="場次明細表(展開)")
            ws2.append(list(df_details.columns))
            for r in df_details.to_dict(orient="records"):
                ws2.append(list(r.values()))

            def format_sheet(ws):
                header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
                center_align = Alignment(horizontal="center", vertical="center")
                left_align = Alignment(horizontal="left", vertical="center")
                border = Border(
                    left=Side(style='thin', color='E5E7EB'),
                    right=Side(style='thin', color='E5E7EB'),
                    top=Side(style='thin', color='E5E7EB'),
                    bottom=Side(style='thin', color='E5E7EB')
                )
                for col_idx in range(1, ws.max_column + 1):
                    c = ws.cell(row=1, column=col_idx)
                    c.fill = header_fill
                    c.font = header_font
                    c.alignment = center_align

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.border = border
                        cell.font = Font(name="微軟正黑體", size=10)
                        col_name = ws.cell(row=1, column=cell.column).value
                        if col_name in ["售票通路", "所屬藝術節/系列", "節目分類", "首演日期", "演出場次時間", "演出日期區間"]:
                            cell.alignment = center_align
                        else:
                            cell.alignment = left_align

                for col in ws.columns:
                    col_letter = get_column_letter(col[0].column)
                    max_len = max(len(str(c.value or '')) for c in col)
                    ws.column_dimensions[col_letter].width = max(min(max_len * 1.6, 50), 14)

            format_sheet(ws1)
            format_sheet(ws2)

            wb.save(output)
            excel_bytes = output.getvalue()

            pbar.progress(100)
            pbar.empty()
            status_box.empty()

            st.success(f"🎉 統計完成！共自【{selected_platform_label}】擷取 **{len(df_summary)}** 檔演出節目、**{len(df_details)}** 個場次紀錄。")

            st.download_button(
                label=f"📥 下載【{selected_platform_label}】統計 Excel 報表",
                data=excel_bytes,
                file_name=f"表演藝術演出統計_{target_platform}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            tab1, tab2 = st.tabs(["📊 演出節目總表 (去重合併)", "📝 場次明細表 (展開)"])
            with tab1:
                st.dataframe(df_summary, use_container_width=True, hide_index=True)
            with tab2:
                st.dataframe(df_details, use_container_width=True, hide_index=True)
