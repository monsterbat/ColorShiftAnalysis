import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import numpy as np

def select_file(file_type):
    """彈出選擇檔案的視窗，並回傳檔案路徑與檔名"""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        return file_path, os.path.basename(file_path)  # 回傳完整路徑和檔名
    return None, None

def read_excel_data(file_path):
    """讀取 Excel 檔案中的 `Result` 分頁，取出 B3 到 E143 的數據"""
    if not file_path:
        return None
    
    # 檢查 Excel 檔案內的所有工作表名稱
    xls = pd.ExcelFile(file_path)
    print(f"可用的工作表: {xls.sheet_names}")  # Debug 用
    
    # 檢查是否有 "Result" 分頁
    if 'Result' not in xls.sheet_names:
        print(f"錯誤: '{file_path}' 不包含 'Result' 工作表")
        return None

    # 讀取 `Result` 分頁
    df = pd.read_excel(file_path, sheet_name='Result', header=None)

    # **檢查 DataFrame 結構**
    print("📊 DataFrame 預覽（前10行）：")
    print(df.head(10))  # Debug：顯示前幾行數據
    print(f"DataFrame 大小: {df.shape}")  # Debug：顯示表格大小

    # **確認實際索引範圍**
    if df.shape[1] < 5:
        print("⚠️ 警告: Excel 檔案的列數少於 5，可能數據沒讀完整！")
    
    # 取出 B3 到 E143 的數據 (0-based index，所以 B3 → row=2, col=1，E → col=4)
    data = df.iloc[2:143, 1:5].values  
    return data


def calculate_ijkl_values(u_data, v_data):
    """計算 I、J、K、L 列的數值，並回傳結果"""
    if u_data is None or v_data is None:
        print("❌ 錯誤: 沒有有效的 U 或 V 數據")
        return None

    # 轉換為 NumPy 陣列
    u_array = np.array(u_data, dtype=float)
    v_array = np.array(v_data, dtype=float)

    # 檢查尺寸是否正確
    if u_array.shape[1] < 4 or v_array.shape[1] < 4:
        print("❌ 錯誤: U 或 V 數據的列數不足")
        return None

    # 計算 I、J、K、L 列
    I_values = np.sqrt(u_array[:, 0] ** 2 + v_array[:, 0] ** 2)
    J_values = np.sqrt(u_array[:, 1] ** 2 + v_array[:, 1] ** 2)
    K_values = np.sqrt(u_array[:, 2] ** 2 + v_array[:, 2] ** 2)
    L_values = np.sqrt(u_array[:, 3] ** 2 + v_array[:, 3] ** 2)

    # 組成 DataFrame
    ijkl_df = pd.DataFrame({
        'I': I_values,
        'J': J_values,
        'K': K_values,
        'L': L_values
    })

    return ijkl_df

def create_new_excel(u_data, v_data):
    """創建新 Excel 檔案，將 U 和 V 的數據貼上，並計算 I、J、K、L 列的數值"""
    if u_data is None or v_data is None:
        print("❌ 錯誤: 沒有讀取到有效數據")
        return
    
    # 檢查數據
    print("🔍 U 資料框尺寸:", u_data.shape)
    print("🔍 V 資料框尺寸:", v_data.shape)

    # 建立 DataFrame
    u_df = pd.DataFrame(u_data, columns=['A', 'B', 'C', 'D'])
    v_df = pd.DataFrame(v_data, columns=['E', 'F', 'G', 'H'])

    # **計算 I、J、K、L 數據**
    ijkl_df = calculate_ijkl_values(u_data, v_data)
    if ijkl_df is None:
        print("❌ 錯誤: 無法計算 I、J、K、L 數據")
        return

    print("📝 修正後 U 檔案的前 5 行:\n", u_df.head())
    print("📝 修正後 V 檔案的前 5 行:\n", v_df.head())
    print("📝 計算出的 IJKL 數據:\n", ijkl_df.head())

    # 創建新 Excel 並寫入數據
    with pd.ExcelWriter("output.xlsx", engine='openpyxl') as writer:
        # U 資料貼在 A4 ~ D144
        u_df.to_excel(writer, index=False, header=False, startrow=3, startcol=0)

        # V 資料貼在 E4 ~ H144
        v_df.to_excel(writer, index=False, header=False, startrow=3, startcol=4)

        # IJKL 資料貼在 I4 ~ L144
        ijkl_df.to_excel(writer, index=False, header=False, startrow=3, startcol=8)

    print("✅ 資料處理完成，已儲存至 output.xlsx")

    
def main():
    """主程式：選擇 U 和 V 檔案，讀取並寫入 Excel"""
    root = tk.Tk()
    root.withdraw()  # 隱藏 Tkinter 主視窗

    print("🔹 選擇 U 檔案")
    u_file, u_filename = select_file("U")
    if not u_file:
        print("❌ 未選擇 U 檔案，程式結束")
        return
    
    print(f"✅ U 檔案: {u_filename}")

    print("🔹 選擇 V 檔案")
    v_file, v_filename = select_file("V")
    if not v_file:
        print("❌ 未選擇 V 檔案，程式結束")
        return
    
    print(f"✅ V 檔案: {v_filename}")

    print("📥 讀取 Excel 數據...")
    u_data = read_excel_data(u_file)
    v_data = read_excel_data(v_file)

    print("📤 產生新 Excel 檔案...")
    create_new_excel(u_data, v_data)
    print("📤LLLLL",create_new_excel)
    print("🎉 完成！")

if __name__ == "__main__":
    main()