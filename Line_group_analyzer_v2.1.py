import streamlit as st
import re
from collections import Counter, defaultdict
from datetime import datetime
import pandas as pd
import gspread
import traceback

# 設定網頁標題
st.set_page_config(page_title="LINE 社群進階分析工具 (Final)", layout="wide")

# ==========================================
# 基礎防呆與格式驗證
# ==========================================
def validate_desktop_format(content):
    lines = content.split('\n')
    has_desktop_date = False
    for i in range(min(50, len(lines))):
        line = lines[i].strip()
        if not line: continue
        if "Saved on:" in line or "[LINE] Chat" in line or "\t" in line:
            return False, "❌ **格式錯誤！** 偵測到手機版格式，請上傳「電腦版 LINE」匯出的 txt！"
        if re.match(r'^\d{4}\.\d{2}\.\d{2}\s+星期[一二三四五六日]$', line):
            has_desktop_date = True
    return has_desktop_date, "success"

# ==========================================
# Google Sheets 核心同步邏輯 (支援多日依序更新)
# ==========================================
def sync_multiple_days_to_sheet(daily_summary):
    try:
        gc = gspread.service_account(filename='creds.json')
        # 你的試算表 ID
        spreadsheet_id = "1lQJ8bLUjcVBBSYL7t7C5WRuoN0JEP_ErMEmy4uaEq5c"
        sh = gc.open_by_key(spreadsheet_id)
        wks = sh.worksheet("03_訊息數量")
        
        results = []
        for date_str, stats in sorted(daily_summary.items()):
            # 每次循環重新取得 A 欄清單，確保列數正確
            dates_in_col_a = wks.col_values(1)
            row_index = -1
            for i, cell_value in enumerate(dates_in_col_a, 1):
                if str(cell_value).strip() == date_str:
                    row_index = i
                    break
            
            total_msgs = int(stats['msgs'])
            unique_users = int(stats['users'])
            
            if row_index != -1:
                # 更新現有列
                cell_range = f"B{row_index}:C{row_index}"
                values = [[total_msgs, unique_users]]
                try:
                    wks.update(range_name=cell_range, values=values)
                except:
                    wks.update(cell_range, values)
                results.append(f"✅ 更新：{date_str} (第 {row_index} 列)")
            else:
                # 新增列
                new_row = [date_str, total_msgs, unique_users]
                wks.append_row(new_row)
                results.append(f"✨ 新增：{date_str} (最後一列)")
        
        return True, results
    except Exception as e:
        if "200" in str(e): return True, ["✅ 同步成功！"]
        return False, [f"❌ 寫入失敗：{str(e)}"]
# ==========================================
# 核心解析邏輯 (共用)
# ==========================================
def analyze_data(content, start_date, end_date, mode="activity", target_keyword=None, match_type="contains"):
    lines = content.split('\n')
    date_pattern = re.compile(r'^(\d{4}\.\d{2}\.\d{2})\s+星期')
    # 智慧姓名辨識：處理包含一個空格的名字
    smart_name_regex = r'([A-Za-z0-9._-]+(?:\s[A-Za-z0-9._-]+)?|[^\s]+)'
    msg_pattern = re.compile(rf'^(\d{{2}}:\d{{2}})\s+{smart_name_regex}\s+(.*)$')
    
    daily_results = defaultdict(lambda: {'msgs': 0, 'users': set()})
    user_counts = Counter()
    current_date, current_user, current_buffer = None, None, []

    def process_buffer():
        nonlocal current_user, current_date, current_buffer
        if current_user and current_date and start_date <= current_date <= end_date:
            full_msg = "\n".join(current_buffer).strip().strip('"')
            date_key = current_date.strftime("%Y-%m-%d")
            
            if mode == "activity":
                daily_results[date_key]['msgs'] += 1
                daily_results[date_key]['users'].add(current_user)
                user_counts[current_user] += 1
            elif mode == "keyword" and target_keyword:
                is_match = (match_type == "exact" and full_msg == target_keyword) or \
                           (match_type == "contains" and target_keyword in full_msg)
                if is_match:
                    user_counts[current_user] += 1

    for line in lines:
        line = line.strip()
        if not line: continue
        
        d_match = date_pattern.match(line)
        if d_match:
            process_buffer()
            current_date = datetime.strptime(d_match.group(1), "%Y.%m.%d")
            current_user, current_buffer = None, []
            continue
            
        m_match = msg_pattern.match(line)
        if m_match:
            process_buffer()
            _, user, msg = m_match.groups()
            current_user, current_buffer = user, [msg]
        elif current_user:
            current_buffer.append(line)
            
    process_buffer()
    
    # 💡 修正重點：將 users 從「姓名清單(set)」轉換為「數量(int)」
    final_daily_results = {}
    for d, data in daily_results.items():
        final_daily_results[d] = {
            'msgs': data['msgs'],
            'users': len(data['users'])
        }
        
    return final_daily_results, user_counts

# ==========================================
# 彈出確認視窗 (st.dialog)
# ==========================================
@st.dialog("確認寫入雲端試算表")
def confirm_sync_dialog(daily_summary):
    st.warning("程式將按照分析時段，逐一更新或新增數據：")
    preview = []
    for d, s in sorted(daily_summary.items()):
        preview.append({"日期": d, "訊息量": s['msgs'], "參與人數": s['users']})
    st.table(pd.DataFrame(preview))
    
    if st.button("確認執行多日同步", use_container_width=True, type="primary"):
        with st.spinner("同步中..."):
            success, results = sync_multiple_days_to_sheet(daily_summary)
            if success:
                st.success("同步成功！")
                for r in results: st.write(r)
            else:
                st.error(results[0])
        if st.button("完成並關閉"): st.rerun()

# ==========================================
# 主介面佈局
# ==========================================
with st.sidebar:
    st.header("⚙️ 基礎設定")
    uploaded_file = st.file_uploader("1. 上傳 LINE 匯出對話 (txt)", type="txt")
    today = datetime.now().date()
    date_range = st.date_input("2. 選擇分析時間範圍", value=(today, today))
    st.markdown("---")
    app_mode = st.radio("🔍 選擇功能：", ["📈 整體活躍度分析", "🎯 關鍵字搜尋"])

# 處理日期範圍轉 datetime
if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    start_d = datetime.combine(date_range[0], datetime.min.time())
    end_d = datetime.combine(date_range[1], datetime.max.time())
    range_str = f"{date_range[0]} ~ {date_range[1]}"
else:
    start_d = datetime.combine(date_range, datetime.min.time())
    end_d = datetime.combine(date_range, datetime.max.time())
    range_str = f"{date_range}"

# --- 1. 活躍度分析功能 ---
if app_mode == "📈 整體活躍度分析":
    st.subheader(f"📈 活躍度趨勢分析 ({range_str})")
    if st.button("🚀 開始深度分析", use_container_width=True):
        if not uploaded_file:
            st.warning("⚠️ 請先上傳檔案！")
        else:
            content = uploaded_file.read().decode("utf-8-sig")
            is_valid, msg = validate_desktop_format(content)
            if not is_valid: st.error(msg)
            else:
                daily_res, user_res = analyze_data(content, start_d, end_d, mode="activity")
                if not daily_res: st.info("該時段內沒有對話紀錄。")
                else:
                    st.session_state['daily_summary'] = daily_res
                    t_msgs = sum(d['msgs'] for d in daily_res.values())
                    t_users = len(user_res)
                    
                    c1, c2 = st.columns(2)
                    c1.metric("時段內總訊息量", f"{t_msgs} 則")
                    c2.metric("參與人數 (不重複)", f"{t_users} 人")
                    
                    st.subheader("🏆 時段內總結發言排行榜")
                    df_rank = pd.DataFrame([{"排名": i, "使用者": u, "總發言數": c} for i, (u, c) in enumerate(user_res.most_common(), 1)])
                    st.dataframe(df_rank, use_container_width=True, hide_index=True)
                    
                    st.subheader("🗓️ 每日統計明細")
                    df_daily = pd.DataFrame.from_dict(daily_res, orient='index').sort_index()
                    df_daily.columns = ['訊息量', '參與人數']
                    st.dataframe(df_daily, use_container_width=True)

    if 'daily_summary' in st.session_state:
        st.markdown("---")
        st.subheader("📤 雲端資料庫同步")
        st.write(f"目前分析範圍：`{range_str}`")
        if st.button("🆙 準備同步至 Excel (依日期逐列更新)", use_container_width=True):
            confirm_sync_dialog(st.session_state['daily_summary'])

# --- 2. 關鍵字搜尋功能 ---
elif app_mode == "🎯 關鍵字搜尋":
    st.subheader(f"🎯 關鍵字搜尋 ({range_str})")
    kw_input = st.text_input("請輸入關鍵字：", value="我奶粉我驕傲")
    match_mode = st.radio("比對規則：", ["包含比對", "完全比對"])
    
    if st.button("🚀 執行關鍵字分析", use_container_width=True):
        if not uploaded_file: st.warning("⚠️ 請先上傳檔案！")
        else:
            content = uploaded_file.read().decode("utf-8-sig")
            is_valid, msg = validate_desktop_format(content)
            if not is_valid: st.error(msg)
            else:
                m_type = "exact" if "完全" in match_mode else "contains"
                _, user_res = analyze_data(content, start_d, end_d, mode="keyword", target_keyword=kw_input.strip(), match_type=m_type)
                
                c3, c4 = st.columns(2)
                c3.metric("符合關鍵字的人數", f"{len(user_res)} 人")
                c4.metric("符合關鍵字的訊息量", f"{sum(user_res.values())} 則")
                
                df_kw = pd.DataFrame([{"排名": i, "使用者": u, "符合次數": c} for i, (u, c) in enumerate(user_res.most_common(), 1)])
                st.dataframe(df_kw, use_container_width=True, hide_index=True)
