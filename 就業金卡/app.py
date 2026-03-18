import streamlit as st
from pathlib import Path
import importlib

# =========================================================
# 1) 儀表板頁面基本設定
# =========================================================
st.set_page_config(
    page_title="就業金卡數據儀表板",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義 CSS 讓介面更專業
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stTable { background-color: white; }
    </style>
    """, unsafe_allow_html=True)

# 定位資料夾路徑
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR # 因為來源 Excel 就在同個資料夾

# =========================================================
# 2) 定義報表選單
# =========================================================
# 這裡定義顯示名稱與對應的檔案名稱 (不含 .py)
REPORTS = {
    "--- 📊 趨勢與分佈圖表 ---": None,
    "圖1: 累計核發人次趨勢": "figure_01",
    "圖2: 累計持卡人次-按領域分": "figure_02",
    "圖3: 累計持卡人次-十大國別": "figure_03",
    "圖4: 累計核發人次-年齡及性別": "figure_04",
    "圖5: 有效持卡人次-按領域分": "figure_05",
    "圖6: 有效持卡人次-十大國別": "figure_06",
    "圖7: 有效持卡人次-年齡及性別": "figure_07",
    
    "--- 📄 數據統計明細表 ---": None,
    "表1: 歷年累計許可人次 (總覽)": "table_01",
    "表2: 歷年累計許可人次 (初次申請)": "table_02",
    "表3: 歷年累計許可人次 (領域分)": "table_03",
    "表4: 累計許可人次 (領域x性別)": "table_04",
    "表5: 累計許可人次 (國別x性別)": "table_05",
    "表6: 有效持卡十大國別 (國別x領域)": "table_06",
    "表7: 累計許可人次 (年齡x性別)": "table_07",
    "表8: 歷年有效許可人次 (外籍x港澳)": "table_08",
    "表9: 有效許可人次 (領域x境內外)": "table_09",
    "表10: 有效許可人次 (有效領域x性別)": "table_10",
    "表11: 有效許可人次 (有效國別x境內外)": "table_11",
    "表12: 有效許可人次 (有效國別x領域)": "table_12",
    "表13: 有效許可人次 (有效國別x性別)": "table_13",
    "表14: 有效許可人次 (有效年齡x性別)": "table_14",
}

# =========================================================
# 3) 側邊欄導覽設計
# =========================================================
st.sidebar.title("🏆 就業金卡報表系統")
st.sidebar.markdown("請選擇下方報表進行檢視")

selected_name = st.sidebar.selectbox(
    "報表清單",
    options=list(REPORTS.keys()),
    index=1 # 預設選中第一個報表 (圖1)
)

st.sidebar.divider()
st.sidebar.info(f"💡 目前偵測資料夾：\n`{BASE_DIR.name}`")

# =========================================================
# 4) 報表動態載入與渲染
# =========================================================
module_name = REPORTS[selected_name]

if module_name:
    try:
        # 動態匯入選中的模組 (例如：import figure_01)
        # 這裡假設 app.py 跟所有的 figure_xx.py 在同一個資料夾
        report_module = importlib.import_module(module_name)
        
        # 呼叫我們在每一支程式最後寫的渲染入口
        report_module.render_streamlit(DATA_DIR)
        
    except ModuleNotFoundError:
        st.error(f"找不到模組檔案: `{module_name}.py`，請確認檔案是否存在。")
    except Exception as e:
        st.error(f"啟動報表時發生錯誤：{e}")
        st.exception(e) # 顯示更詳細的錯誤資訊供工程師調校
else:
    # 如果選到的是分隔線文字
    st.title("歡迎使用就業金卡數據儀表板")
    st.markdown("""
    ### 👈 請從左側選單選擇要查看的數據報表
    本系統會自動讀取資料夾中最新的 Excel 檔案，並即時產出數據分析結果。
    
    - **圖表區**：包含累計與有效人次的趨勢、國別與領域分佈。
    - **表格區**：提供詳細的歷年統計、交叉分析及佔比數據。
    """)