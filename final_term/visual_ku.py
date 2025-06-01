import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os
from matplotlib.font_manager import FontProperties

# 設定中文字型
font_path = "/Users/Peggy/Library/Fonts/NotoSansTC-Regular.ttf"
my_font = FontProperties(fname=font_path)

def normalize_school_names(platform_df, test_df):
    """標準化學校名稱以便合併"""
    
    # 創建平台數據的副本，移除「縣立」前綴
    platform_normalized = platform_df.copy()
    platform_normalized['學校名稱_標準'] = platform_normalized['學校名稱'].str.replace('縣立', '', regex=False)
    
    # 創建測驗成績的副本
    test_normalized = test_df.copy()
    test_normalized['學校名稱_標準'] = test_normalized['學校名稱'].copy()
    
    # 處理特殊案例：賢庵國小含垵湖分校 -> 賢庵國小
    test_normalized['學校名稱_標準'] = test_normalized['學校名稱_標準'].str.replace('含垵湖分校', '', regex=False)
    
    # 顯示標準化後的名稱
    print("=== 標準化後的學校名稱 ===")
    print("平台數據標準化名稱:")
    for orig, norm in zip(platform_normalized['學校名稱'], platform_normalized['學校名稱_標準']):
        if orig != norm:
            print(f"  {orig} -> {norm}")
    
    print("\n測驗成績標準化名稱:")
    for orig, norm in zip(test_normalized['學校名稱'], test_normalized['學校名稱_標準']):
        if orig != norm:
            print(f"  {orig} -> {norm}")
    
    return platform_normalized, test_normalized

def main():
    print("開始分析平台使用與學力測驗成績的關係...")
    
    # 1. 讀取a.py產生的平台數據CSV
    platform_csv_path = "platform_data.csv"
    if not os.path.exists(platform_csv_path):
        print(f"❌ 找不到平台數據檔案 {platform_csv_path}")
        print("請先執行 a.py 並確保它輸出 CSV 檔案")
        return
    
    # 讀取平台數據
    platform_df = pd.read_csv(platform_csv_path)
    
    # 將百分比字串轉換為數值
    def safe_convert_percentage(x):
        if pd.isna(x) or x == 'N/A':
            return np.nan
        if isinstance(x, str):
            return float(x.rstrip('%'))
        return float(x)
    
    platform_df['整體數學完成率'] = platform_df['整體數學完成率'].apply(safe_convert_percentage)
    platform_df['平均數學正答率'] = platform_df['平均數學正答率'].apply(safe_convert_percentage)
    
    print(f"✅ 平台數據載入成功，共 {len(platform_df)} 筆")
    
    # 2. 讀取app.py產生的學力測驗成績CSV
    test_scores_path = "test_scores.csv"
    if not os.path.exists(test_scores_path):
        print(f"❌ 找不到學力測驗成績檔案 {test_scores_path}")
        print("請先執行 app.py 並確保它輸出 CSV 檔案")
        return
    
    # 讀取學力測驗成績
    test_df = pd.read_csv(test_scores_path)
    print(f"✅ 學力測驗成績載入成功，共 {len(test_df)} 筆")
    
    # 3. 標準化學校名稱
    platform_normalized, test_normalized = normalize_school_names(platform_df, test_df)
    
    # 4. 使用標準化名稱合併數據
    merged_df = pd.merge(
        platform_normalized, 
        test_normalized, 
        on='學校名稱_標準', 
        how='inner',
        suffixes=('_平台', '_測驗')
    )
    
    print(f"\n✅ 成功合併的學校 ({len(merged_df)} 所):")
    for _, row in merged_df.iterrows():
        print(f"  - {row['學校名稱_平台']} <-> {row['學校名稱_測驗']}")
    
    # 移除缺失值
    merged_df = merged_df.dropna(subset=['整體數學完成率', '平均數學正答率', '總得分率'])
    
    if len(merged_df) < 3:
        print("⚠️ 合併後的數據太少（少於3個學校），可能影響分析可靠性")
        if len(merged_df) == 0:
            print("❌ 沒有成功合併的數據")
            return
    
    print(f"✅ 最終用於分析的數據共有 {len(merged_df)} 筆")
    
    # 5. 計算相關係數
    completion_corr = np.corrcoef(merged_df['整體數學完成率'], merged_df['總得分率'])[0, 1]
    accuracy_corr = np.corrcoef(
        merged_df['平均數學正答率'].dropna(), 
        merged_df[merged_df['平均數學正答率'].notna()]['總得分率']
    )[0, 1]
    
    print(f"✓ 平台完成率與測驗得分率的相關係數: {completion_corr:.3f}")
    print(f"✓ 平台正答率與測驗得分率的相關係數: {accuracy_corr:.3f}")
    
    # 6. 創建視覺化（散點圖）
    plt.figure(figsize=(16, 7))
    
    # 完成率與得分率關係
    plt.subplot(1, 2, 1)
    sns.regplot(x='整體數學完成率', y='總得分率', data=merged_df, scatter_kws={'alpha':0.7})
    
    # 添加學校標籤
    for _, row in merged_df.iterrows():
        plt.annotate(row['學校名稱_標準'], 
                  (row['整體數學完成率'], row['總得分率']),
                  xytext=(5, 5), textcoords='offset points',
                  fontproperties=my_font, fontsize=9, alpha=0.8)
    
    plt.title('康軒平台任務完成率與學力測驗得分率關係', fontproperties=my_font, fontsize=14)
    plt.xlabel('平台整體數學完成率 (%)', fontproperties=my_font, fontsize=12)
    plt.ylabel('學力測驗平均得分率', fontproperties=my_font, fontsize=12)
    plt.text(0.05, 0.95, f'相關係數: {completion_corr:.3f}', 
            transform=plt.gca().transAxes, fontproperties=my_font, fontsize=12,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8))
    
    # 正答率與得分率關係（過濾掉NaN值）
    valid_data = merged_df.dropna(subset=['平均數學正答率'])
    
    plt.subplot(1, 2, 2)
    if len(valid_data) > 0:
        sns.regplot(x='平均數學正答率', y='總得分率', data=valid_data, scatter_kws={'alpha':0.7})
        
        # 添加學校標籤
        for _, row in valid_data.iterrows():
            plt.annotate(row['學校名稱_標準'], 
                      (row['平均數學正答率'], row['總得分率']),
                      xytext=(5, 5), textcoords='offset points',
                      fontproperties=my_font, fontsize=9, alpha=0.8)
        
        plt.text(0.05, 0.95, f'相關係數: {accuracy_corr:.3f}', 
                transform=plt.gca().transAxes, fontproperties=my_font, fontsize=12,
                bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8))
    else:
        plt.text(0.5, 0.5, '無有效的正答率數據', ha='center', va='center',
                transform=plt.gca().transAxes, fontproperties=my_font, fontsize=14)
    
    plt.title('康軒平台任務正答率與學力測驗得分率關係', fontproperties=my_font, fontsize=14)
    plt.xlabel('平台平均數學正答率 (%)', fontproperties=my_font, fontsize=12)
    plt.ylabel('學力測驗平均得分率', fontproperties=my_font, fontsize=12)
    
    plt.tight_layout()
    plt.savefig('康軒平台表現與學力測驗關係.png', dpi=300, bbox_inches='tight')
    print("✅ 已生成平台表現與學力測驗關係圖")
    
    # 7. 創建多維氣泡圖（如果有足夠的有效數據）
    valid_for_bubble = merged_df.dropna(subset=['整體數學完成率', '平均數學正答率', '總得分率'])
    
    if len(valid_for_bubble) >= 3:
        plt.figure(figsize=(10, 8))
        
        # 氣泡圖：X軸=完成率，Y軸=正答率，氣泡大小=測驗得分率
        sizes = valid_for_bubble['總得分率'] * 500  # 調整氣泡大小
        
        scatter = plt.scatter(valid_for_bubble['整體數學完成率'], 
                            valid_for_bubble['平均數學正答率'], 
                            s=sizes, 
                            alpha=0.7,
                            c=valid_for_bubble['總得分率'], 
                            cmap='viridis',
                            edgecolors='black',
                            linewidth=0.5)
        
        # 添加學校標籤
        for _, row in valid_for_bubble.iterrows():
            plt.annotate(row['學校名稱_標準'], 
                      (row['整體數學完成率'], row['平均數學正答率']),
                      xytext=(5, 5), textcoords='offset points',
                      fontproperties=my_font, fontsize=10, alpha=0.9)
        
        plt.title('康軒平台完成率、正答率與學力測驗成績的多維關係', fontproperties=my_font, fontsize=16)
        plt.xlabel('平台整體數學完成率 (%)', fontproperties=my_font, fontsize=14)
        plt.ylabel('平台平均數學正答率 (%)', fontproperties=my_font, fontsize=14)
        
        # 添加顏色條
        cbar = plt.colorbar(scatter)
        cbar.set_label('學力測驗平均得分率', fontproperties=my_font, fontsize=12)
        
        # 添加氣泡大小說明
        plt.text(0.02, 0.98, '氣泡大小代表測驗得分率', 
                transform=plt.gca().transAxes, fontproperties=my_font, fontsize=10,
                bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8),
                verticalalignment='top')
        
        plt.tight_layout()
        plt.savefig('康軒平台表現與學力測驗多維關係.png', dpi=300, bbox_inches='tight')
        print("✅ 已生成平台表現與學力測驗多維關係圖")
    else:
        print("⚠️ 有效數據不足，跳過多維氣泡圖")
    
    # 8. 顯示數據摘要
    print(f"\n=== 分析結果摘要 ===")
    print(f"分析學校數量: {len(merged_df)}")
    print(f"平台完成率範圍: {merged_df['整體數學完成率'].min():.1f}% - {merged_df['整體數學完成率'].max():.1f}%")
    print(f"平台正答率範圍: {valid_data['平均數學正答率'].min():.1f}% - {valid_data['平均數學正答率'].max():.1f}%")
    print(f"測驗得分率範圍: {merged_df['總得分率'].min():.3f} - {merged_df['總得分率'].max():.3f}")
    print("分析完成!")

if __name__ == "__main__":
    main()