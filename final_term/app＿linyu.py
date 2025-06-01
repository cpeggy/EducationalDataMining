import matplotlib
matplotlib.use('Agg')
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import warnings
import sys
import re
warnings.filterwarnings('ignore')

from matplotlib.font_manager import FontProperties

# 指定中文字型路徑
font_path = "/Users/Peggy/Library/Fonts/NotoSansTC-Regular.ttf"
my_font = FontProperties(fname=font_path)

# === 讀取學力測驗成績 ===
def read_data(file_path):
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
    return int(filename.split('數學')[1][0])

def load_all_data(file_paths):
    dfs = []
    for f in file_paths:
        g = extract_grade(f)
        df = read_data(f)
        if df is not None:
            df = df.rename(columns={
                df.columns[8]: '姓名', 
                df.columns[9]: '性別', 
                df.columns[2]: '學校代碼', 
                df.columns[18]: '總得分率', 
                df.columns[3]: '學校名稱'
            })
            df['年級'] = g
            df['總得分率'] = pd.to_numeric(df['總得分率'], errors='coerce')
            dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else None

# === 力宇平台數據處理 ===
def ultra_robust_convert_to_numeric(df, cols):
    for col in cols:
        def clean_number(x):
            x = str(x).replace('，', '').replace(',', '').replace('．', '.').strip()
            x = re.sub(r'[^\d\.]+', '', x)
            try:
                return float(x) if x else 0.0
            except:
                return 0.0
        df[col] = df[col].apply(clean_number)
    return df

def filter_non_zero(df, cols):
    condition = (df[cols] != 0).any(axis=1)
    return df[condition]

def get_liyou_usage_data_from_excel(excel_path):
    """從力宇教育Excel檔案中提取平台使用數據"""
    try:
        target_cols = [
            '自我練習平均成績', '自我測驗卷數', '自我練習題目數',
            '完成老師指派影片數', '自我點播影片數', '累積影片總時數'
        ]
        
        print(f"正在讀取力宇教育數據: {excel_path}")
        xls = pd.ExcelFile(excel_path)
        combined_df = []

        print(f"Excel檔案包含 {len(xls.sheet_names)} 個分頁:")
        for sheet in xls.sheet_names:
            print(f"  - {sheet}")

        for sheet in xls.sheet_names:
            try:
                df = xls.parse(sheet)
                print(f"\n處理分頁: {sheet}")
                print(f"  原始資料: {len(df)} 筆")
                
                if '科目' in df.columns:
                    # 篩選數學科目的資料
                    df_math = df[df['科目'] == '數學'].copy()
                    print(f"  數學科目資料: {len(df_math)} 筆")
                    
                    if len(df_math) > 0:
                        # 清理數據
                        df_math = ultra_robust_convert_to_numeric(df_math, target_cols)
                        
                        # 過濾掉全為0的資料
                        df_math = filter_non_zero(df_math, target_cols)
                        print(f"  有效資料: {len(df_math)} 筆")
                        
                        if len(df_math) > 0:
                            df_math['學校名稱'] = sheet
                            combined_df.append(df_math)
                        
            except Exception as e:
                print(f"⚠️ 無法處理分頁 {sheet}：{e}")

        if not combined_df:
            print("❌ 沒有找到有效的力宇平台數據")
            return pd.DataFrame()

        df_all = pd.concat(combined_df, ignore_index=True)
        print(f"\n✅ 合併後的力宇數據: {len(df_all)} 筆")

        # 計算平台使用量和相關指標
        df_all['平台使用量'] = df_all[['完成老師指派影片數', '自我練習題目數', '自我測驗卷數']].sum(axis=1)
        df_all['自我練習平均成績'] = pd.to_numeric(df_all['自我練習平均成績'], errors='coerce').fillna(0)
        
        # 計算影片使用指標
        df_all['影片使用量'] = df_all[['完成老師指派影片數', '自我點播影片數']].sum(axis=1)
        df_all['累積影片總時數'] = pd.to_numeric(df_all['累積影片總時數'], errors='coerce').fillna(0)

        return df_all[['學校名稱', '年級', '平台使用量', '自我練習平均成績', '影片使用量', '累積影片總時數', 
                      '完成老師指派影片數', '自我點播影片數', '自我練習題目數', '自我測驗卷數']]
    except Exception as e:
        print(f"處理Excel檔案時發生錯誤: {e}")
        return pd.DataFrame()

def main():
    print("=== 力宇教育平台數據分析程式 ===\n")
    
    # 1. 載入學力測驗成績
    print("步驟1: 載入學力測驗成績...")
    files = [f'113年度_學力測驗_金門縣_數學{i}年級成績_202406.csv' for i in range(3,9)]
    df_scores = load_all_data(files)

    if df_scores is None:
        print("❌ 未能載入學力測驗成績資料")
        return

    print(f"✅ 學力測驗成績載入成功，共 {len(df_scores)} 筆")
    
    # 計算每所學校的平均得分率
    school_avg_scores = df_scores.groupby('學校名稱')['總得分率'].mean().reset_index()
    
    # 輸出學力測驗成績CSV
    school_avg_scores.to_csv("test_scores.csv", index=False, encoding='utf-8')
    print(f"✅ 學力測驗成績已匯出至 test_scores.csv，共 {len(school_avg_scores)} 所學校")

    # 2. 載入力宇教育平台數據
    print("\n步驟2: 載入力宇教育平台數據...")
    excel_file = "力宇教育-時數報表113.08.01-113.12.31 (2).xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ 找不到力宇教育Excel檔案: {excel_file}")
        print("請確認檔案路徑是否正確")
        return
    
    df_liyou = get_liyou_usage_data_from_excel(excel_file)
    
    if df_liyou.empty:
        print("❌ 無法載入力宇平台數據")
        return

    print(f"✅ 力宇平台數據載入成功，共 {len(df_liyou)} 筆")
    
    # 顯示力宇數據統計
    print(f"\n=== 力宇平台數據統計 ===")
    print(f"涵蓋學校數: {df_liyou['學校名稱'].nunique()}")
    print(f"涵蓋年級: {sorted(df_liyou['年級'].unique())}")
    print(f"平台使用量範圍: {df_liyou['平台使用量'].min():.1f} - {df_liyou['平台使用量'].max():.1f}")
    print(f"自我練習平均成績範圍: {df_liyou['自我練習平均成績'].min():.1f} - {df_liyou['自我練習平均成績'].max():.1f}")
    
    # 3. 計算每個學校的平台使用平均值
    print("\n步驟3: 計算各學校平台使用統計...")
    liyou_school_stats = df_liyou.groupby('學校名稱').agg({
        '平台使用量': 'mean',
        '自我練習平均成績': 'mean',
        '影片使用量': 'mean',
        '累積影片總時數': 'mean',
        '完成老師指派影片數': 'mean',
        '自我點播影片數': 'mean',
        '自我練習題目數': 'mean',
        '自我測驗卷數': 'mean'
    }).round(2).reset_index()
    
    # 輸出力宇平台數據CSV
    liyou_school_stats.to_csv("liyou_platform_data.csv", index=False, encoding='utf-8')
    print(f"✅ 力宇平台數據已匯出至 liyou_platform_data.csv，共 {len(liyou_school_stats)} 所學校")
    
    # 顯示力宇平台數據
    print(f"\n=== 各學校力宇平台使用統計 ===")
    print(liyou_school_stats.to_string(index=False))
    
    # 4. 嘗試合併數據進行初步分析
    print(f"\n步驟4: 合併數據進行初步分析...")
    
    # 標準化學校名稱 (處理「縣立」前綴問題)
    liyou_school_stats['學校名稱_標準'] = liyou_school_stats['學校名稱'].str.replace('縣立', '', regex=False)
    school_avg_scores['學校名稱_標準'] = school_avg_scores['學校名稱'].str.replace('含垵湖分校', '', regex=False)
    
    # 合併數據
    merged_preview = pd.merge(
        liyou_school_stats, 
        school_avg_scores, 
        on='學校名稱_標準', 
        how='inner',
        suffixes=('_力宇', '_測驗')
    )
    
    if len(merged_preview) > 0:
        print(f"✅ 成功匹配 {len(merged_preview)} 所學校的數據:")
        for _, row in merged_preview.iterrows():
            print(f"  - {row['學校名稱_力宇']} <-> {row['學校名稱_測驗']}")
        
        # 計算初步相關性
        if len(merged_preview) >= 3:
            usage_corr = merged_preview['平台使用量'].corr(merged_preview['總得分率'])
            practice_corr = merged_preview['自我練習平均成績'].corr(merged_preview['總得分率'])
            
            print(f"\n=== 初步相關性分析 ===")
            print(f"平台使用量 vs 測驗得分率相關係數: {usage_corr:.3f}")
            print(f"自我練習成績 vs 測驗得分率相關係數: {practice_corr:.3f}")
    else:
        print("⚠️ 沒有成功匹配的學校數據")
        print("請檢查學校名稱是否一致")
    
    print(f"\n=== 程式執行完成 ===")
    print("產生的檔案:")
    print("  - test_scores.csv (學力測驗成績)")
    print("  - liyou_platform_data.csv (力宇平台數據)")
    print("接下來可以執行視覺化程式進行詳細分析")

if __name__ == '__main__':
    main()
