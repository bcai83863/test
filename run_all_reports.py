import subprocess
import time
import sys
from pathlib import Path

# 檢查是否已安裝合併套件
try:
    from docx import Document
    from docxcompose.composer import Composer
except ImportError:
    print("❌ 缺少合併套件！請在終端機執行：pip install docxcompose")
    sys.exit(1)

# =========================================================
# 0) 設定路徑與任務清單
# =========================================================
# 自動抓取這支程式所在位置，並尋找同層級的「就業金卡」資料夾
BASE = Path(__file__).parent / "就業金卡"
OUT = BASE / "_output"
OUT.mkdir(parents=True, exist_ok=True)

# 依序執行的任務清單
TASKS = [
    "table_01.py", "table_02.py", "figure_01.py",
    "table_03.py", "figure_02.py", "table_04.py",
    "table_05.py", "figure_03.py", "table_06.py",
    "table_07.py", "figure_04.py", "table_08.py",
    "table_09.py", "figure_05.py", "table_10.py",
    "table_11.py", "figure_06.py", "table_12.py",
    "table_13.py", "table_14.py", "figure_07.py"
]

def run_task_and_catch_docx(script_name):
    """執行單一腳本，並找出它剛生成的 Word 檔案路徑"""
    script_path = BASE / script_name
    if not script_path.exists():
        print(f"❌ 找不到檔案：{script_path.name}，跳過。")
        return False, None

    print(f"▶️ 正在執行：{script_name} ...")
    
    # 記錄執行前的時間 (減去1秒緩衝，避免時間差)
    start_time = time.time() - 1 
    
    # 執行子程式
    result = subprocess.run(["python", str(script_path)], cwd=BASE)

    if result.returncode == 0:
        # 尋找 _output 資料夾內，剛剛(執行後)才生成的 docx 檔案
        newest_docx = None
        max_mtime = start_time
        
        for f in OUT.glob("*.docx"):
            if "完整合併報告" in f.name: 
                continue # 避開我們自己生成的總報告
            
            mtime = f.stat().st_mtime
            if mtime > max_mtime:
                max_mtime = mtime
                newest_docx = f
                
        return True, newest_docx
    else:
        print(f" ❌ 執行出錯 (錯誤碼: {result.returncode})")
        return False, None

def main():
    print("======================================================")
    print("      就業金卡自動化報表系統 - 一鍵生成與合併工具")
    print("======================================================")
    print(f"預計執行任務總數：{len(TASKS)}")
    
    success_count = 0
    generated_docs = [] # 用來收集所有產出的 Word 路徑

    # 1. 逐一跑腳本
    for i, task in enumerate(TASKS, 1):
        print(f"\n[{i}/{len(TASKS)}] ", end="")
        success, docx_path = run_task_and_catch_docx(task)
        
        if success:
            success_count += 1
            if docx_path:
                generated_docs.append(docx_path)
                print(f"   └─ 捕捉到檔案: {docx_path.name}")
        else:
            cont = input(f"\n⚠️ {task} 執行失敗，是否繼續執行下一項？(y/n): ").lower()
            if cont != 'y':
                print("停止執行。")
                break

    print("\n======================================================")
    print(f"腳本執行結束！成功產出：{len(generated_docs)} 份 Word 檔案")
    
    # 2. 開始合併所有收集到的 Word 檔案
    if generated_docs:
        print("\n🔄 正在將所有報表整併成單一 Word 檔案...")
        try:
            # 將第一份文件當作「基底」
            master_doc = Document(generated_docs[0])
            composer = Composer(master_doc)
            
            # 將後續的文件依序貼在後面
            for doc_path in generated_docs[1:]:
                doc = Document(doc_path)
                # 插入分頁符號 (每一張表/圖換新的一頁)
                master_doc.add_page_break() 
                composer.append(doc)
            
            # 儲存最終大檔案
            timestamp = time.strftime("%Y%m%d_%H%M")
            merged_output_path = OUT / f"就業金卡_完整合併報告_{timestamp}.docx"
            composer.save(merged_output_path)
            
            print(f"🎉 大功告成！合併報告已儲存至：\n📂 {merged_output_path}")
            
        except Exception as e:
            print(f"❌ 合併過程發生錯誤：{e}")
    
    print("======================================================")

if __name__ == "__main__":
    main()