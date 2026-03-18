import subprocess
import time
import sys
from pathlib import Path

# =========================================================
# 0) 設定路徑與任務清單
# =========================================================
# 自動抓取這支程式所在位置，並尋找同層級的「就業金卡」資料夾
BASE = Path(__file__).parent / "就業金卡"
TASKS = [
    "table_01.py", "table_02.py", "figure_01.py",
    "table_03.py", "figure_02.py", "table_04.py",
    "table_05.py", "figure_03.py", "table_06.py",
    "table_07.py", "figure_04.py", "table_08.py",
    "table_09.py", "figure_05.py", "table_10.py",
    "table_11.py", "figure_06.py", "table_12.py",
    "table_13.py", "table_14.py", "figure_07.py"
]

def run_task(script_name):
    """執行單一 Python 腳本"""
    # 組合完整路徑：C:\Users\mrtbh\實驗\table_01.py
    script_path = BASE / script_name
    
    if not script_path.exists():
        print(f"❌ 找不到檔案：{script_path.name}，跳過。")
        return False

    print(f"▶️ 正在執行：{script_name} ...", end="", flush=True)
    start_time = time.time()

    # 關鍵修正：加入 cwd=BASE，這會讓子腳本在執行時把「實驗」資料夾當作當前目錄
    # 這樣它們才能順利找到資料夾裡的 Excel 檔案
    result = subprocess.run(["python", str(script_path)], cwd=BASE)

    end_time = time.time()
    if result.returncode == 0:
        print(f" ✅ 完成! (耗時 {end_time - start_time:.2f} 秒)")
        return True
    else:
        print(f" ❌ 執行出錯 (錯誤碼: {result.returncode})")
        return False

def main():
    print("======================================================")
    print("      就業金卡自動化報表系統 - 全自動執行工具")
    print("======================================================")
    print(f"預計執行任務總數：{len(TASKS)}\n")

    success_count = 0
    for i, task in enumerate(TASKS, 1):
        print(f"[{i}/{len(TASKS)}] ", end="")
        if run_task(task):
            success_count += 1
        else:
            # 如果失敗，詢問是否繼續
            cont = input(f"\n⚠️ {task} 執行失敗，是否繼續執行下一項？(y/n): ").lower()
            if cont != 'y':
                print("停止執行。")
                break
        print("-" * 30)

    print("\n======================================================")
    print(f"工作結束！成功執行：{success_count}/{len(TASKS)}")
    print(f"所有檔案已輸出至: {BASE / '_output'}")
    print("======================================================")

if __name__ == "__main__":
    main()