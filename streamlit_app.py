import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

# 網頁基本設定 (支援手機 RWD 響應式佈局)
st.set_page_config(page_title="建程包租代管簡易系統", layout="wide")
st.title("🏠 建程包租代管簡易系統")

if "rental_db" not in st.session_state:
    st.session_state.rental_db = None

# ==================== 第一步：匯入 Excel ====================
st.subheader("📥 第一步：匯入與關聯你的 Excel 檔案")
uploaded_base_file = st.file_uploader("請上傳你的 Excel 檔案 (.xlsx 或 .xlsm)", type=["xlsx", "xlsm"])

if uploaded_base_file:
    try:
        uploaded_bytes = uploaded_base_file.getvalue()
        excel_file = pd.ExcelFile(io.BytesIO(uploaded_bytes), engine='openpyxl')
        sheet_names = excel_file.sheet_names
        
        # 讓使用者手動挑選主帳務表 (租金表)
        finance_sheet = st.selectbox("請選擇【當月費用帳務】工作表", sheet_names, index=0)
        
        # 在背景全自動搜尋包含「虛擬帳號」的工作表
        auto_info_sheet = None
        for name in sheet_names:
            if "虛擬帳號" in name:
                auto_info_sheet = name
                break
        
        if not auto_info_sheet:
            auto_info_sheet = sheet_names[min(1, len(sheet_names)-1)]
            
        if st.button("🔄 載入工作表並執行智慧關聯"):
            with st.spinner("正在智慧搜尋標題列並進行資料解析..."):
                
                # 地毯式搜尋「當月費用帳務」的標題列到底在第幾列
                found_finance_header = 0
                for skip_rows in range(10):  # 掃描前 10 列
                    test_df = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=finance_sheet, header=skip_rows, nrows=2, engine='openpyxl')
                    test_cols = [str(c).strip() for c in test_df.columns]
                    if "房客姓名" in test_cols:
                        found_finance_header = skip_rows
                        break
                
                # 地毯式搜尋「虛擬帳號底冊」的標題列到底在第幾列
                found_info_header = 0
                for skip_rows in range(10):
                    test_df = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=auto_info_sheet, header=skip_rows, nrows=2, engine='openpyxl')
                    test_cols = [str(c).strip() for c in test_df.columns]
                    if "房客" in test_cols or "房客姓名" in test_cols:
                        found_info_header = skip_rows
                        break
                
                # 用自動找到的精準列數來加載資料
                df_finance = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=finance_sheet, header=found_finance_header, engine='openpyxl')
                df_info = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=auto_info_sheet, header=found_info_header, engine='openpyxl')
                
                # 清理欄位空白
                df_finance.columns = df_finance.columns.str.strip().astype(str)
                df_info.columns = df_info.columns.str.strip().astype(str)
                
                # 房客姓名欄位智慧相容更名
                if "房客" in df_info.columns and "房客姓名" not in df_info.columns:
                    df_info = df_info.rename(columns={"房客": "房客姓名"})
                
                if "房客姓名" not in df_finance.columns:
                    st.error(f"❌ 錯誤：自動掃描前 10 列後，在【{finance_sheet}】內依然找不到【房客姓名】欄位標題！")
                else:
                    df_finance["房客姓名"] = df_finance["房客姓名"].fillna("").astype(str).str.strip()
                    df_info["房客姓名"] = df_info["房客姓名"].fillna("").astype(str).str.strip()
                    
                    # 智慧業務欄位配對
                    agent_col = next((c for c in df_finance.columns if c in ["房客所屬人", "負責業務", "業務姓名", "業務", "專員"]), None)
                    if agent_col:
                        df_finance = df_finance.rename(columns={agent_col: "房客所屬人"})
                    else:
                        df_finance["房客所屬人"] = "未分配業務"
                        
                    # 執行跨表關聯
                    info_cols = [col for col in ["房客姓名", "房客虛擬帳號", "房東姓名", "應付房東金額"] if col in df_info.columns]
                    df_info_clean = df_info[info_cols].drop_duplicates(subset=["房客姓名"])
                    df_merged = pd.merge(df_finance, df_info_clean, on="房客姓名", how="left")
                    
                    # 設定今天日期 (2026-07-15)
                    today_now = pd.Timestamp(datetime.today().date())
                    
                    # === 智慧收租天數動態計算 ===
                    if "房客租金支付日" in df_merged.columns:
                        df_merged["房客租金支付日_dt"] = pd.to_datetime(df_merged["房客租金支付日"], errors='coerce')
                        df_merged["房客租金支付日_顯示"] = df_merged["房客租金支付日_dt"].dt.strftime("%Y-%m-%d").fillna(df_merged["房客租金支付日"].astype(str))
                        df_merged["距離支付日天數"] = (df_merged["房客租金支付日_dt"].dt.tz_localize(None) - today_now).dt.days
                    else:
                        df_merged["房客租金支付日_顯示"] = "未設定"
                        df_merged["距離支付日天數"] = np.nan
                    
                    # 智慧費用型態轉換與自動加總
                    fee_cols = ["租金收入", "管理費", "水費", "清潔費", "電費", "其他費用", "存摺房客匯入", "差異"]
                    for col in fee_cols:
                        if col in df_merged.columns:
                            df_merged[col] = pd.to_numeric(df_merged[col], errors='coerce').fillna(0)
                        else:
                            df_merged[col] = 0
                            
                    df_merged["房客總應付"] = df_merged["租金收入"] + df_merged["管理費"] + df_merged["水費"] + df_merged["清潔費"] + df_merged["電費"] + df_merged["其他費用"]
                    
                    # 逾期天數與原因欄位強轉型
                    if "逾期天數" in df_merged.columns:
                        df_merged["逾期天數_數字"] = pd.to_numeric(df_merged["逾期天數"], errors='coerce').fillna(1)
                    else:
                        df_merged["逾期天數_數字"] = 1
                        
                    df_merged["原因_文字"] = df_merged["原因"].fillna("").astype(str).str.strip() if "原因" in df_merged.columns else ""
                    df_merged["差異說明_文字"] = df_merged["差異說明"].fillna("").astype(str).str.strip() if "差異說明" in df_merged.columns else ""
                    
                    st.session_state.rental_db = df_merged
                    st.success(f"🎉 智慧自動精準定位！主表成功對齊第 {found_finance_header + 1} 列、底冊對齊第 {found_info_header + 1} 列，串接完美成功！")
                    
    except Exception as e:
        st.error(f"檔案讀取失敗，錯誤訊息: {e}")

# ==================== 第二步：業務專屬智慧預警看板 ====================
if st.session_state.rental_db is not None:
    st.markdown("---")
    st.subheader("👥 第二步：業務專屬催繳與收租預警看板")
    
    st.session_state.rental_db["房客所屬人"] = st.session_state.rental_db["房客所屬人"].fillna("未分配業務").astype(str).str.strip()
    all_agents = sorted([a for a in st.session_state.rental_db["房客所屬人"].unique() if a != ""])
    selected_agent = st.selectbox("請選擇【房客所屬人】切換專屬業務介面", all_agents)
    
    agent_df = st.session_state.rental_db[st.session_state.rental_db["房客所屬人"] == selected_agent].copy()
    
    tab1, tab2 = st.tabs([
        f"🚨 {selected_agent} 的【已逾期未繳費用明細】", 
        f"📅 {selected_agent} 的【下期 7 天內即將收租預警】"
    ])
    
    # --- 🚨 分頁一：已逾期未繳費用明細 ---
    with tab1:
        st.markdown("##### 🎯 篩選條件：【逾期天數 ≦ 0】且原因為【未繳】或【部分繳清】之房客：")
        cond1 = agent_df["原因_文字"].str.contains("未繳|部分繳清")
        cond2 = agent_df["逾期天數_數字"] <= 0
        unpaid_agent = agent_df[cond1 & cond2].copy()
        
        if not unpaid_agent.empty:
            def format_overdue(days):
                if pd.isna(days) or days > 0: return "未到期"
                elif days == 0: return "⏳ 今天到期"
                else: return f"🚨 已逾期 {abs(int(days))} 天"
            unpaid_agent["逾期狀態"] = unpaid_agent["逾期天數_數字"].apply(format_overdue)
            
            clean_show_cols = ["房客姓名", "月份", "出租承租標的地址", "房客租金支付日_顯示", "逾期狀態", "差異", "差異說明_文字", "原因_文字"]
            rename_dict = {
                "房客租金支付日_顯示": "📅 房客租金支付日",
                "差異說明_文字": "📢 Excel差異說明 (會計備註)", 
                "原因_文字": "📝 欠費狀態 (原因)"
            }
            exist_cols = [c for c in clean_show_cols if c in unpaid_agent.columns]
            st.dataframe(unpaid_agent[exist_cols].rename(columns=rename_dict), use_container_width=True)
        else:
            st.success(f"✨ 太棒了！**{selected_agent}** 負責的房客中，目前沒有任何符合條件的逾期欠費案件！")
            
    # --- 📅 分頁二：下期即將收租預警 ---
    with tab2:
        st.markdown("##### 🔔 智慧提醒：依據【房客租金支付日】前 7 天(含)內，且【原因】不為已繳清者自動篩選：")
        
        if "距離支付日天數" in agent_df.columns:
            rent_reminder_df = agent_df[
                (agent_df["距離支付日天數"] >= 0) & 
                (agent_df["距離支付日天數"] <= 7) &
                (agent_df["原因_文字"] != "已繳清")
            ].copy()
            
            if not rent_reminder_df.empty:
                def get_pay_alert_memo(days):
                    if days == 0: return "⏳ 今天就是繳費日！請注意確認入帳"
                    elif days == 1: return "⏳ 明天即將繳費！可傳訊息提醒"
                    else: return f"🔔 倒數 {int(days)} 天繳費 (請於前7天注意提醒)"
                        
                rent_reminder_df["收租行動指引"] = rent_reminder_df["距離支付日天數"].apply(get_pay_alert_memo)
                
                pay_cols = ["房客姓名", "月份", "出租承租標的地址", "房客租金支付日_顯示", "原因_文字", "收租行動指引"]
                exist_pay_cols = [c for c in pay_cols if c in rent_reminder_df.columns]
                st.dataframe(rent_reminder_df[exist_pay_cols].rename(columns={"房客租金支付日_顯示": "📅 房客租金支付日", "原因_文字": "📝 欠費狀態 (原因)"}), use_container_width=True)
            else:
                st.success(f"✅ **{selected_agent}** 旗下目前沒有 7 天內「即將到期且尚未繳清」的收租案件！")
        else:
            st.warning("Excel 中找不到【房客租金支付日】欄位。")

    # ==================== 功能三：原始綜合帳務明細看板 ====================
    st.markdown("---")
    st.subheader("📊 第三步：會計匯入總檔資料顯示結果 (完整備份與下載)")
    
    drop_cols = ["逾期天數_數字", "原因_文字", "差異說明_文字", "房客租金支付日_dt"]
    clean_db = st.session_state.rental_db.drop(columns=drop_cols, errors='ignore')
    
    st.dataframe(clean_db, use_container_width=True)
    
    # 建立下載記憶體二進位流
    output = io.BytesIO()
    clean_db.to_excel(output, index=False, sheet_name='對帳更新結果', engine='openpyxl')
    processed_data = output.getvalue()

    # 下載按鈕安全傳遞
    st.download_button(
        label="📥 下載處理後的 Excel 檔案",
        data=processed_data,
        file_name="對帳更新結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
