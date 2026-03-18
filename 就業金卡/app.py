import streamlit as st
from pathlib import Path
import importlib
import sys # ✨ 新增：用來處理搜尋路徑

# =========================================================
# 1) 自動修正路徑 (最關鍵的一步)
# =========================================================
# 獲取目前 app.py 所在的資料夾路徑
current_dir = Path(__file__).parent.absolute()

# 如果目前的資料夾路徑不在 Python 的搜尋清單中，就把它加進去
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

# 設定資料夾路徑
DATA_DIR = current_dir 

st.set_page_config(
    page_title="就業金卡數據儀表板",
    page_icon="🏆",
    layout="wide"
)

# =========================================================
# 2) 定義報表選單 (維持原樣)
# =========================================================
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
# 3) 側邊欄與載入邏輯
# =========================================================
st.sidebar.title("🏆 就業金卡報表系統")
selected_name = st.sidebar.selectbox("請選擇報表", options=list(REPORTS.keys()), index=1)

module_name = REPORTS[selected_name]

if module_name:
    try:
        # ✨ 現在有了 sys.path.insert，這裡一定找得到檔案
        report_module = importlib.import_module(module_name)
        report_module.render_streamlit(DATA_DIR)
        
    except ModuleNotFoundError as e:
        st.error(f"❌ 找不到模組檔案: `{module_name}.py`")
        st.info(f"偵測路徑: `{current_dir}`")
        st.debug(f"詳細錯誤: {e}")
    except Exception as e:
        st.error(f"啟動報表時發生錯誤：{e}")
else:
    st.title("歡迎使用就業金卡數據儀表板")
    st.markdown("### 👈 請從左側選單選擇要查看的數據報表")