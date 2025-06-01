import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import re
from matplotlib.font_manager import FontProperties

# 設定中文字型
font_path = "/Users/Peggy/Library/Fonts/NotoSansTC-Regular.ttf"
my_font = FontProperties(fname=font_path)

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

def generate_liyou_summary_table():
    """生成力宇平台詳細統計表，類似康軒平台的格式"""
    
    print("=== 力宇教育平台學校使用統計摘要 ===\n")
    
    # 讀取力宇教育Excel檔案
    excel_file = "力宇教育-時數報表113.08.01-113.12.31 (2).xlsx"
    
    try:
        target_cols = [
            '自我練習平均成績', '自我測驗卷數', '自我練習題目數',
            '完成老師指派影片數', '自我點播影片數', '累積影片總時數'
        ]
        
        xls = pd.ExcelFile(excel_file)
        all_schools_summary = []

        print(f"正在處理 {len(xls.sheet_names)} 所學校的力宇平台數據...\n")

        for sheet in xls.sheet_names:
            try:
                df = xls.parse(sheet)
                
                if '科目' in df.columns:
                    # 篩選數學科目
                    df_math = df[df['科目'] == '數學'].copy()
                    
                    if len(df_math) > 0:
                        # 清理數據
                        df_math = ultra_robust_convert_to_numeric(df_math, target_cols)
                        
                        # 計算學校統計
                        school_stats = {
                            '學校名稱': sheet,
                            '總練習題目數': df_math['自我練習題目數'].sum(),
                            '總測驗卷數': df_math['自我測驗卷數'].sum(),
                            '平均練習成績': df_math['自我練習平均成績'].mean() if df_math['自我練習平均成績'].sum() > 0 else 0,
                            '總指派影片數': df_math['完成老師指派影片數'].sum(),
                            '總自選影片數': df_math['自我點播影片數'].sum(),
                            '累積總時數': df_math['累積影片總時數'].sum(),
                            '學生人數': len(df_math)
                        }
                        
                        # 計算衍生指標
                        total_activities = school_stats['總練習題目數'] + school_stats['總測驗卷數']
                        total_videos = school_stats['總指派影片數'] + school_stats['總自選影片數']
                        
                        school_stats.update({
                            '總平台使用量': total_activities + total_videos,
                            '平均每生練習題數': school_stats['總練習題目數'] / school_stats['學生人數'] if school_stats['學生人數'] > 0 else 0,
                            '平均每生影片數': total_videos / school_stats['學生人數'] if school_stats['學生人數'] > 0 else 0,
                            '平均每生時數': school_stats['累積總時數'] / school_stats['學生人數'] if school_stats['學生人數'] > 0 else 0
                        })
                        
                        all_schools_summary.append(school_stats)
                    else:
                        # 沒有數學科目數據的學校
                        all_schools_summary.append({
                            '學校名稱': sheet,
                            '總練習題目數': 0,
                            '總測驗卷數': 0,
                            '平均練習成績': 0,
                            '總指派影片數': 0,
                            '總自選影片數': 0,
                            '累積總時數': 0,
                            '學生人數': 0,
                            '總平台使用量': 0,
                            '平均每生練習題數': 0,
                            '平均每生影片數': 0,
                            '平均每生時數': 0
                        })
                else:
                    # 沒有科目欄位的學校
                    all_schools_summary.append({
                        '學校名稱': sheet,
                        '總練習題目數': 0,
                        '總測驗卷數': 0,
                        '平均練習成績': 0,
                        '總指派影片數': 0,
                        '總自選影片數': 0,
                        '累積總時數': 0,
                        '學生人數': 0,
                        '總平台使用量': 0,
                        '平均每生練習題數': 0,
                        '平均每生影片數': 0,
                        '平均每生時數': 0
                    })
                    
            except Exception as e:
                print(f"處理學校 {sheet} 時發生錯誤: {e}")
                continue

        # 轉換為DataFrame
        summary_df = pd.DataFrame(all_schools_summary)
        
        # 排序（按總平台使用量降序）
        summary_df = summary_df.sort_values('總平台使用量', ascending=False)
        
        # 顯示詳細統計表
        print("=" * 120)
        print("力宇教育平台學校使用摘要：")
        print("=" * 120)
        print(f"{'學校名稱':<15} | {'學生數':<6} | {'練習題數':<8} | {'測驗卷數':<8} | {'平均成績':<8} | {'影片總數':<8} | {'累積時數':<8} | {'平台使用量':<10}")
        print("=" * 120)
        
        for _, row in summary_df.iterrows():
            total_videos = row['總指派影片數'] + row['總自選影片數']
            avg_score = f"{row['平均練習成績']:.1f}" if row['平均練習成績'] > 0 else "N/A"
            
            print(f"{row['學校名稱']:<15} | {row['學生人數']:<6.0f} | {row['總練習題目數']:<8.0f} | {row['總測驗卷數']:<8.0f} | {avg_score:<8} | {total_videos:<8.0f} | {row['累積總時數']:<8.0f} | {row['總平台使用量']:<10.0f}")
        
        print("=" * 120)
        
        # 統計摘要
        active_schools = summary_df[summary_df['總平台使用量'] > 0]
        print(f"\n統計摘要：")
        print(f"總學校數: {len(summary_df)}")
        print(f"有使用力宇平台的學校數: {len(active_schools)}")
        print(f"平台使用率: {len(active_schools)/len(summary_df)*100:.1f}%")
        
        if len(active_schools) > 0:
            print(f"平均每校使用量: {active_schools['總平台使用量'].mean():.1f}")
            print(f"最高使用量學校: {active_schools.iloc[0]['學校名稱']} ({active_schools.iloc[0]['總平台使用量']:.0f})")
            print(f"平均練習成績範圍: {active_schools['平均練習成績'].min():.1f} - {active_schools['平均練習成績'].max():.1f}")
        
        # 輸出CSV檔案
        summary_df.to_csv('力宇平台學校使用統計.csv', index=False, encoding='utf-8')
        print(f"\n✅ 詳細統計已輸出至 '力宇平台學校使用統計.csv'")
        
        # 與康軒平台對比
        print(f"\n=== 力宇 vs 康軒平台使用對比 ===")
        print(f"力宇平台: {len(active_schools)}/{len(summary_df)} 所學校有使用")
        print(f"康軒平台: 4/26 所學校有使用 (根據您提供的數據)")
        print(f"力宇平台使用普及率更高: {len(active_schools)/len(summary_df)*100:.1f}% vs 15.4%")
        
        return summary_df
        
    except FileNotFoundError:
        print(f"❌ 找不到檔案: {excel_file}")
        return None
    except Exception as e:
        print(f"❌ 處理過程發生錯誤: {e}")
        return None

def create_liyou_visualization_like_kangxuan(summary_df):
    """創建類似康軒平台的力宇平台視覺化"""
    
    if summary_df is None or len(summary_df) == 0:
        print("❌ 沒有數據可供視覺化")
        return
    
    # 篩選有使用數據的學校
    active_df = summary_df[summary_df['總平台使用量'] > 0].copy()
    
    if len(active_df) == 0:
        print("❌ 沒有學校使用力宇平台")
        return
    
    # 創建類似康軒的多維分析圖
    plt.figure(figsize=(12, 8))
    
    # 氣泡圖：X軸=平均每生練習題數，Y軸=平均練習成績，氣泡大小=總平台使用量，顏色=平均每生影片數
    scatter = plt.scatter(active_df['平均每生練習題數'], 
                        active_df['平均練習成績'], 
                        s=active_df['總平台使用量'] * 10,  # 調整氣泡大小
                        alpha=0.7,
                        c=active_df['平均每生影片數'], 
                        cmap='viridis',
                        edgecolors='black',
                        linewidth=0.5)
    
    # 添加學校標籤
    for _, row in active_df.iterrows():
        plt.annotate(row['學校名稱'], 
                  (row['平均每生練習題數'], row['平均練習成績']),
                  xytext=(5, 5), textcoords='offset points',
                  fontproperties=my_font, fontsize=10, alpha=0.9)
    
    plt.title('力宇平台使用量、練習成績與影片學習的多維關係', fontproperties=my_font, fontsize=16)
    plt.xlabel('平均每生練習題數', fontproperties=my_font, fontsize=14)
    plt.ylabel('平均練習成績', fontproperties=my_font, fontsize=14)
    
    # 添加顏色條
    cbar = plt.colorbar(scatter)
    cbar.set_label('平均每生影片數', fontproperties=my_font, fontsize=12)
    
    # 添加氣泡大小說明
    plt.text(0.02, 0.98, '氣泡大小代表總平台使用量', 
            transform=plt.gca().transAxes, fontproperties=my_font, fontsize=10,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8),
            verticalalignment='top')
    
    plt.tight_layout()
    plt.savefig('力宇平台使用統計視覺化.png', dpi=300, bbox_inches='tight')
    print("✅ 已生成力宇平台使用統計視覺化圖")

if __name__ == "__main__":
    print("開始生成力宇平台詳細統計...")
    summary_df = generate_liyou_summary_table()
    
    if summary_df is not None:
        print("\n開始生成視覺化圖表...")
        create_liyou_visualization_like_kangxuan(summary_df)
        print("\n✅ 力宇平台統計分析完成！")
    else:
        print("❌ 統計分析失敗")
