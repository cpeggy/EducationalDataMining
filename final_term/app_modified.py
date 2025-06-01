import pandas as pd
import numpy as np
import os

# === 讀取 & 處理 ===
def read_data(file_path):
    """嘗試多種編碼讀取CSV檔案"""
    encodings = ['cp950', 'big5', 'utf-8', 'gbk', 'latin1']
    for enc in encodings:
        try:
            df = pd.read_csv(file_path, encoding=enc, low_memory=False)
            if df.shape[1] > 0:
                print(f"檔案 {file_path} 使用編碼 {enc} 成功讀取。")
                return df
        except:
            continue
    print(f"❌ 讀取失敗：{file_path}")
    return None

def extract_grade(filename):
    """從檔名中提取年級"""
    return int(filename.split('數學')[1][0])

def load_all_data(file_paths):
    """載入所有年級的數學成績資料"""
    dfs = []
    for f in file_paths:
        print(f"正在處理檔案: {f}")
        g = extract_grade(f)
        df = read_data(f)
        if df is not None:
            # 重新命名欄位
            df = df.rename(columns={
                df.columns[8]: '姓名', 
                df.columns[9]: '性別', 
                df.columns[2]: '學校代碼', 
                df.columns[18]: '總得分率', 
                df.columns[3]: '學校名稱'
            })
            df['年級'] = g
            # 將總得分率轉換為數值
            df['總得分率'] = pd.to_numeric(df['總得分率'], errors='coerce')
            dfs.append(df)
            print(f"✅ {g}年級資料載入成功，共 {len(df)} 筆")
        else:
            print(f"❌ {g}年級資料載入失敗")
    
    if dfs:
        result = pd.concat(dfs, ignore_index=True)
        print(f"✅ 所有資料合併完成，總共 {len(result)} 筆")
        return result
    else:
        print("❌ 沒有成功載入任何資料")
        return None

def main():
    print("開始處理學力測驗成績資料...")
    
    # 設定檔案路徑
    files = [f'113年度_學力測驗_金門縣_數學{i}年級成績_202406.csv' for i in range(3,9)]
    
    # 檢查檔案是否存在
    existing_files = []
    for f in files:
        if os.path.exists(f):
            existing_files.append(f)
            print(f"✓ 找到檔案: {f}")
        else:
            print(f"⚠️ 找不到檔案: {f}")
    
    if not existing_files:
        print("❌ 沒有找到任何學力測驗成績檔案")
        print("請確認檔案路徑是否正確")
        return
    
    # 載入所有資料
    df_scores = load_all_data(existing_files)
    
    if df_scores is not None:
        print(f"✅ 成績資料載入成功，共 {len(df_scores)} 筆")
        
        # 檢查資料結構
        print(f"資料欄位: {list(df_scores.columns)}")
        print(f"學校數量: {df_scores['學校名稱'].nunique()}")
        print(f"年級分布:\n{df_scores['年級'].value_counts().sort_index()}")
        
        # 計算每所學校的平均得分率
        school_avg_scores = df_scores.groupby('學校名稱')['總得分率'].mean().reset_index()
        
        # 顯示結果
        print(f"\n各學校平均得分率:")
        print("-" * 50)
        for _, row in school_avg_scores.iterrows():
            print(f"{row['學校名稱']:<20}: {row['總得分率']:.3f}")
        
        # 輸出到CSV
        output_path = "test_scores.csv"
        school_avg_scores.to_csv(output_path, index=False, encoding='utf-8')
        print(f"\n✅ 學力測驗成績已匯出至 {output_path}")
        print(f"   共 {len(school_avg_scores)} 所學校的資料")
        
        # 顯示CSV檔案內容確認
        print(f"\n{output_path} 檔案內容:")
        print(school_avg_scores.to_string(index=False))
        
    else:
        print("❌ 未能載入任何成績資料")

if __name__ == '__main__':
    main()