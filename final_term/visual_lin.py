import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os
from matplotlib.font_manager import FontProperties

# 設定中文字型
font_path = "/Users/Peggy/Library/Fonts/NotoSansTC-Regular.ttf"
my_font = FontProperties(fname=font_path)

def normalize_school_names(liyou_df, test_df):
    """標準化學校名稱以便合併"""
    
    # 創建力宇數據的副本，移除「縣立」前綴
    liyou_normalized = liyou_df.copy()
    liyou_normalized['學校名稱_標準'] = liyou_normalized['學校名稱'].str.replace('縣立', '', regex=False)
    
    # 創建測驗成績的副本
    test_normalized = test_df.copy()
    test_normalized['學校名稱_標準'] = test_normalized['學校名稱'].copy()
    
    # 處理特殊案例：賢庵國小含垵湖分校 -> 賢庵國小
    test_normalized['學校名稱_標準'] = test_normalized['學校名稱_標準'].str.replace('含垵湖分校', '', regex=False)
    
    # 顯示標準化後的名稱
    print("=== 學校名稱標準化 ===")
    print("力宇平台數據標準化名稱:")
    for orig, norm in zip(liyou_normalized['學校名稱'], liyou_normalized['學校名稱_標準']):
        if orig != norm:
            print(f"  {orig} -> {norm}")
    
    print("\n測驗成績標準化名稱:")
    for orig, norm in zip(test_normalized['學校名稱'], test_normalized['學校名稱_標準']):
        if orig != norm:
            print(f"  {orig} -> {norm}")
    
    return liyou_normalized, test_normalized

def main():
    print("開始分析力宇教育平台使用與學力測驗成績的關係...")
    
    # 1. 讀取力宇平台數據CSV
    liyou_csv_path = "liyou_platform_data.csv"
    if not os.path.exists(liyou_csv_path):
        print(f"❌ 找不到力宇平台數據檔案 {liyou_csv_path}")
        print("請先執行 liyou_app_modified.py 並確保它輸出 CSV 檔案")
        return
    
    # 讀取力宇平台數據
    liyou_df = pd.read_csv(liyou_csv_path)
    print(f"✅ 力宇平台數據載入成功，共 {len(liyou_df)} 筆")
    print(f"力宇數據欄位: {list(liyou_df.columns)}")
    
    # 2. 讀取學力測驗成績CSV
    test_scores_path = "test_scores.csv"
    if not os.path.exists(test_scores_path):
        print(f"❌ 找不到學力測驗成績檔案 {test_scores_path}")
        print("請先執行 liyou_app_modified.py 確保它輸出 CSV 檔案")
        return
    
    # 讀取學力測驗成績
    test_df = pd.read_csv(test_scores_path)
    print(f"✅ 學力測驗成績載入成功，共 {len(test_df)} 筆")
    
    # 3. 標準化學校名稱
    liyou_normalized, test_normalized = normalize_school_names(liyou_df, test_df)
    
    # 4. 使用標準化名稱合併數據
    merged_df = pd.merge(
        liyou_normalized, 
        test_normalized, 
        on='學校名稱_標準', 
        how='inner',
        suffixes=('_力宇', '_測驗')
    )
    
    print(f"\n✅ 成功合併的學校 ({len(merged_df)} 所):")
    for _, row in merged_df.iterrows():
        print(f"  - {row['學校名稱_力宇']} <-> {row['學校名稱_測驗']}")
    
    # 移除缺失值
    merged_df = merged_df.dropna(subset=['平台使用量', '總得分率'])
    
    if len(merged_df) < 3:
        print("⚠️ 合併後的數據太少（少於3個學校），可能影響分析可靠性")
        if len(merged_df) == 0:
            print("❌ 沒有成功合併的數據")
            return
    
    print(f"✅ 最終用於分析的數據共有 {len(merged_df)} 筆")
    
    # 5. 計算相關係數
    usage_corr = merged_df['平台使用量'].corr(merged_df['總得分率'])
    practice_corr = merged_df['自我練習平均成績'].corr(merged_df['總得分率'])
    video_corr = merged_df['影片使用量'].corr(merged_df['總得分率'])
    time_corr = merged_df['累積影片總時數'].corr(merged_df['總得分率'])
    
    print(f"\n=== 相關性分析結果 ===")
    print(f"✓ 平台使用量 vs 測驗得分率相關係數: {usage_corr:.3f}")
    print(f"✓ 自我練習成績 vs 測驗得分率相關係數: {practice_corr:.3f}")
    print(f"✓ 影片使用量 vs 測驗得分率相關係數: {video_corr:.3f}")
    print(f"✓ 累積影片時數 vs 測驗得分率相關係數: {time_corr:.3f}")
    
    # 6. 創建多維度分析視覺化
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    
    # 6.1 平台使用量 vs 測驗得分率
    ax1 = axes[0, 0]
    sns.regplot(x='平台使用量', y='總得分率', data=merged_df, ax=ax1, scatter_kws={'alpha':0.7})
    for _, row in merged_df.iterrows():
        ax1.annotate(row['學校名稱_標準'], 
                  (row['平台使用量'], row['總得分率']),
                  xytext=(5, 5), textcoords='offset points',
                  fontproperties=my_font, fontsize=8, alpha=0.8)
    
    ax1.set_title('平台使用量與學力測驗得分率關係', fontproperties=my_font, fontsize=14)
    ax1.set_xlabel('平台使用量', fontproperties=my_font, fontsize=12)
    ax1.set_ylabel('學力測驗得分率', fontproperties=my_font, fontsize=12)
    ax1.text(0.05, 0.95, f'相關係數: {usage_corr:.3f}', 
            transform=ax1.transAxes, fontproperties=my_font, fontsize=11,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8))
    
    # 6.2 自我練習成績 vs 測驗得分率
    ax2 = axes[0, 1]
    sns.regplot(x='自我練習平均成績', y='總得分率', data=merged_df, ax=ax2, scatter_kws={'alpha':0.7})
    for _, row in merged_df.iterrows():
        ax2.annotate(row['學校名稱_標準'], 
                  (row['自我練習平均成績'], row['總得分率']),
                  xytext=(5, 5), textcoords='offset points',
                  fontproperties=my_font, fontsize=8, alpha=0.8)
    
    ax2.set_title('自我練習成績與學力測驗得分率關係', fontproperties=my_font, fontsize=14)
    ax2.set_xlabel('自我練習平均成績', fontproperties=my_font, fontsize=12)
    ax2.set_ylabel('學力測驗得分率', fontproperties=my_font, fontsize=12)
    ax2.text(0.05, 0.95, f'相關係數: {practice_corr:.3f}', 
            transform=ax2.transAxes, fontproperties=my_font, fontsize=11,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8))
    
    # 6.3 影片使用量 vs 測驗得分率
    ax3 = axes[1, 0]
    sns.regplot(x='影片使用量', y='總得分率', data=merged_df, ax=ax3, scatter_kws={'alpha':0.7})
    for _, row in merged_df.iterrows():
        ax3.annotate(row['學校名稱_標準'], 
                  (row['影片使用量'], row['總得分率']),
                  xytext=(5, 5), textcoords='offset points',
                  fontproperties=my_font, fontsize=8, alpha=0.8)
    
    ax3.set_title('影片使用量與學力測驗得分率關係', fontproperties=my_font, fontsize=14)
    ax3.set_xlabel('影片使用量', fontproperties=my_font, fontsize=12)
    ax3.set_ylabel('學力測驗得分率', fontproperties=my_font, fontsize=12)
    ax3.text(0.05, 0.95, f'相關係數: {video_corr:.3f}', 
            transform=ax3.transAxes, fontproperties=my_font, fontsize=11,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8))
    
    # 6.4 累積影片時數 vs 測驗得分率
    ax4 = axes[1, 1]
    sns.regplot(x='累積影片總時數', y='總得分率', data=merged_df, ax=ax4, scatter_kws={'alpha':0.7})
    for _, row in merged_df.iterrows():
        ax4.annotate(row['學校名稱_標準'], 
                  (row['累積影片總時數'], row['總得分率']),
                  xytext=(5, 5), textcoords='offset points',
                  fontproperties=my_font, fontsize=8, alpha=0.8)
    
    ax4.set_title('累積影片時數與學力測驗得分率關係', fontproperties=my_font, fontsize=14)
    ax4.set_xlabel('累積影片總時數', fontproperties=my_font, fontsize=12)
    ax4.set_ylabel('學力測驗得分率', fontproperties=my_font, fontsize=12)
    ax4.text(0.05, 0.95, f'相關係數: {time_corr:.3f}', 
            transform=ax4.transAxes, fontproperties=my_font, fontsize=11,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8))
    
    plt.tight_layout()
    plt.savefig('力宇平台表現與學力測驗關係.png', dpi=300, bbox_inches='tight')
    print("✅ 已生成力宇平台表現與學力測驗關係圖")
    
    # 7. 創建多維氣泡圖
    plt.figure(figsize=(12, 8))
    
    # 氣泡圖：X軸=平台使用量，Y軸=自我練習成績，氣泡大小=測驗得分率，顏色=影片使用量
    sizes = merged_df['總得分率'] * 1000  # 調整氣泡大小
    
    scatter = plt.scatter(merged_df['平台使用量'], 
                        merged_df['自我練習平均成績'], 
                        s=sizes, 
                        alpha=0.7,
                        c=merged_df['影片使用量'], 
                        cmap='viridis',
                        edgecolors='black',
                        linewidth=0.5)
    
    # 添加學校標籤
    for _, row in merged_df.iterrows():
        plt.annotate(row['學校名稱_標準'], 
                  (row['平台使用量'], row['自我練習平均成績']),
                  xytext=(5, 5), textcoords='offset points',
                  fontproperties=my_font, fontsize=10, alpha=0.9)
    
    plt.title('力宇平台多維使用效果與學力測驗成績關係', fontproperties=my_font, fontsize=16)
    plt.xlabel('平台使用量', fontproperties=my_font, fontsize=14)
    plt.ylabel('自我練習平均成績', fontproperties=my_font, fontsize=14)
    
    # 添加顏色條
    cbar = plt.colorbar(scatter)
    cbar.set_label('影片使用量', fontproperties=my_font, fontsize=12)
    
    # 添加氣泡大小說明
    plt.text(0.02, 0.98, '氣泡大小代表測驗得分率', 
            transform=plt.gca().transAxes, fontproperties=my_font, fontsize=12,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8),
            verticalalignment='top')
    
    plt.tight_layout()
    plt.savefig('力宇平台多維關係分析.png', dpi=300, bbox_inches='tight')
    print("✅ 已生成力宇平台多維關係分析圖")
    
    # 8. 生成詳細的數據統計表
    print(f"\n=== 力宇平台使用與學力測驗分析摘要 ===")
    print(f"分析學校數量: {len(merged_df)}")
    print(f"平台使用量範圍: {merged_df['平台使用量'].min():.1f} - {merged_df['平台使用量'].max():.1f}")
    print(f"自我練習成績範圍: {merged_df['自我練習平均成績'].min():.1f} - {merged_df['自我練習平均成績'].max():.1f}")
    print(f"影片使用量範圍: {merged_df['影片使用量'].min():.1f} - {merged_df['影片使用量'].max():.1f}")
    print(f"累積影片時數範圍: {merged_df['累積影片總時數'].min():.1f} - {merged_df['累積影片總時數'].max():.1f}")
    print(f"測驗得分率範圍: {merged_df['總得分率'].min():.3f} - {merged_df['總得分率'].max():.3f}")
    
    # 9. 創建詳細統計表並輸出
    summary_stats = merged_df.groupby('學校名稱_標準').agg({
        '平台使用量': 'first',
        '自我練習平均成績': 'first', 
        '影片使用量': 'first',
        '累積影片總時數': 'first',
        '總得分率': 'first'
    }).round(3)
    
    summary_stats['平台使用量排名'] = summary_stats['平台使用量'].rank(ascending=False)
    summary_stats['測驗得分率排名'] = summary_stats['總得分率'].rank(ascending=False)
    
    print(f"\n=== 各學校詳細統計 ===")
    print(summary_stats.to_string())
    
    # 輸出詳細統計到CSV
    summary_stats.to_csv('力宇平台與測驗成績統計.csv', encoding='utf-8')
    print(f"\n✅ 詳細統計已輸出至 '力宇平台與測驗成績統計.csv'")
    
    # 10. 找出表現突出的學校
    print(f"\n=== 表現分析 ===")
    
    # 高使用量+高成績的學校
    high_usage = merged_df['平台使用量'] > merged_df['平台使用量'].median()
    high_score = merged_df['總得分率'] > merged_df['總得分率'].median()
    excellent_schools = merged_df[high_usage & high_score]
    
    if len(excellent_schools) > 0:
        print(f"高使用量+高成績學校 ({len(excellent_schools)} 所):")
        for _, school in excellent_schools.iterrows():
            print(f"  - {school['學校名稱_標準']}: 使用量 {school['平台使用量']:.1f}, 得分率 {school['總得分率']:.3f}")
    
    # 低使用量但高成績的學校
    low_usage = merged_df['平台使用量'] <= merged_df['平台使用量'].median()
    efficient_schools = merged_df[low_usage & high_score]
    
    if len(efficient_schools) > 0:
        print(f"\n低使用量但高成績學校 ({len(efficient_schools)} 所):")
        for _, school in efficient_schools.iterrows():
            print(f"  - {school['學校名稱_標準']}: 使用量 {school['平台使用量']:.1f}, 得分率 {school['總得分率']:.3f}")
    
    print("\n分析完成!")
    print("已生成的檔案:")
    print("  - 力宇平台表現與學力測驗關係.png")
    print("  - 力宇平台多維關係分析.png") 
    print("  - 力宇平台與測驗成績統計.csv")

if __name__ == "__main__":
    main()
    