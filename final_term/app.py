import matplotlib
matplotlib.use('Agg')
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import warnings
import sys
import re  # 導入re模塊用於正則表達式
warnings.filterwarnings('ignore')
import finalkm


from matplotlib.font_manager import FontProperties

# 指定你的 Noto 字型或 PingFang.ttc 的完整路徑
font_path = "/Users/Peggy/Library/Fonts/NotoSansTC-Regular.ttf"
my_font = FontProperties(fname=font_path)

# === 讀取 & 處理 ===
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
            df = df.rename(columns={df.columns[8]: '姓名', df.columns[9]: '性別', df.columns[2]: '學校代碼', df.columns[18]: '總得分率', df.columns[3]: '學校名稱'})
            df['年級'] = g
            df['總得分率'] = pd.to_numeric(df['總得分率'], errors='coerce')
            dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else None

# === 視覺化1: 年級比較 ===
def visualize_grade_differences(df):
    grade_scores = df.groupby('年級')['總得分率'].mean().reset_index()
    plt.figure(figsize=(10, 6))
    sns.barplot(x='年級', y='總得分率', data=grade_scores, palette='viridis')
    for i, row in grade_scores.iterrows():
        plt.text(i, row['總得分率']+0.02, f'{row["總得分率"]:.2f}', ha='center', fontproperties=my_font)
    plt.title('不同年級的平均數學得分率比較', fontproperties=my_font, fontsize=16)
    plt.xlabel('年級', fontproperties=my_font, fontsize=14)
    plt.ylabel('總得分率', fontproperties=my_font, fontsize=14)
    plt.ylim(0, 1)
    plt.savefig('不同年級數學能力比較.png', dpi=300)

    plt.figure(figsize=(12, 6))
    sns.boxplot(x='年級', y='總得分率', data=df, palette='viridis')
    plt.title('不同年級數學得分率分布', fontproperties=my_font, fontsize=16)
    plt.xlabel('年級', fontproperties=my_font, fontsize=14)
    plt.ylabel('總得分率', fontproperties=my_font, fontsize=14)
    plt.ylim(0, 1)
    plt.savefig('不同年級數學能力分布.png', dpi=300)

# === 視覺化2: 性別差異 ===
def visualize_gender_differences(df):
    df['性別標籤'] = df['性別'].map({1: '男', 2: '女'})
    avg_scores = df.groupby(['年級', '性別標籤'])['總得分率'].mean().reset_index()
    plt.figure(figsize=(10, 6))
    sns.barplot(x='年級', y='總得分率', hue='性別標籤', data=avg_scores, palette='Set1')
    plt.title('不同性別在各年級的平均數學得分率比較', fontproperties=my_font, fontsize=16)
    plt.xlabel('年級', fontproperties=my_font, fontsize=14)
    plt.ylabel('總得分率', fontproperties=my_font, fontsize=14)
    
    # 設定圖例字體
    legend = plt.legend()
    for text in legend.get_texts():
        text.set_fontproperties(my_font)
    
    plt.ylim(0, 1)
    plt.savefig('性別對數學成績的影響.png', dpi=300)

    plt.figure(figsize=(12, 8))
    for g in sorted(df['年級'].unique()):
        plt.subplot(2, 3, g-2)
        sns.violinplot(x='性別標籤', y='總得分率', data=df[df['年級']==g], palette='Set2')
        plt.title(f'{g}年級', fontproperties=my_font, fontsize=14)
        plt.xlabel('性別標籤', fontproperties=my_font, fontsize=12)
        plt.ylabel('總得分率', fontproperties=my_font, fontsize=12)
    plt.tight_layout()
    plt.savefig('各年級性別成績分布.png', dpi=300)

# === 視覺化3: 學校比較 ===
def visualize_school_differences(df):
    top_schools = df['學校名稱'].value_counts().head(10).index
    filtered = df[df['學校名稱'].isin(top_schools)]
    pivot = filtered.groupby(['學校名稱', '年級'])['總得分率'].mean().reset_index().pivot(index='學校名稱', columns='年級', values='總得分率')
    plt.figure(figsize=(14, 10))
    sns.heatmap(pivot, annot=True, cmap='YlGnBu', fmt='.2f', linewidths=.5)
    plt.title('不同學校在各年級的平均數學得分率熱力圖 (前10名學生數學校)', fontproperties=my_font, fontsize=16)
    
    # 設定刻度標籤字體
    for label in plt.gca().get_xticklabels():
        label.set_fontproperties(my_font)
    for label in plt.gca().get_yticklabels():
        label.set_fontproperties(my_font)
    
    # 設定色條標籤字體
    cbar = plt.gcf().axes[-1]  # 取得最後一個axes，即colorbar
    cbar.set_ylabel('平均得分率', fontproperties=my_font, fontsize=12)
    
    plt.savefig('學校間數學表現差異熱力圖.png', dpi=300)

    for g in sorted(df['年級'].unique()):
        school_scores = filtered[filtered['年級']==g].groupby('學校名稱')['總得分率'].mean().sort_values().reset_index()
        plt.figure(figsize=(12, 8))
        bar = sns.barplot(x='總得分率', y='學校名稱', data=school_scores, orient='h', palette='viridis')
        for i, row in school_scores.iterrows():
            bar.text(row['總得分率'] + 0.01, i, f'{row["總得分率"]:.2f}', va='center', fontproperties=my_font)

        plt.title(f'{g}年級各學校的平均數學得分率', fontproperties=my_font, fontsize=16)
        plt.xlabel('總得分率', fontproperties=my_font, fontsize=14)
        plt.ylabel('學校名稱', fontproperties=my_font, fontsize=14)
        
        # 設定刻度標籤字體
        for label in plt.gca().get_xticklabels():
            label.set_fontproperties(my_font)
        for label in plt.gca().get_yticklabels():
            label.set_fontproperties(my_font)
            
        plt.savefig(f'{g}年級學校間數學表現差異.png', dpi=300)

# === finalkm模組功能的內部實現 ===
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

def get_km_usage_data_from_excel(excel_path):
    try:
        # 目標欄位
        target_cols = [
            '自我練習平均成績', '自我測驗卷數', '自我練習題目數',
            '完成老師指派影片數', '自我點播影片數', '累積影片總時數'
        ]
        
        xls = pd.ExcelFile(excel_path)
        combined_df = []

        for sheet in xls.sheet_names:
            try:
                df = xls.parse(sheet)
                if '科目' in df.columns:
                    df = df[df['科目'] == '數學'].copy()
                    df = ultra_robust_convert_to_numeric(df, target_cols)
                    df = filter_non_zero(df, target_cols)
                    df['學校名稱'] = sheet
                    combined_df.append(df)
            except Exception as e:
                print(f"⚠️ 無法處理分頁 {sheet}：{e}")

        if not combined_df:
            return pd.DataFrame()

        df_all = pd.concat(combined_df, ignore_index=True)

        df_all['平台使用量'] = df_all[['完成老師指派影片數', '自我練習題目數', '自我測驗卷數']].sum(axis=1)
        df_all['自我練習平均成績'] = pd.to_numeric(df_all['自我練習平均成績'], errors='coerce').fillna(0)

        return df_all[['學校名稱', '年級', '平台使用量', '自我練習平均成績']]
    except Exception as e:
        print(f"處理Excel檔案時發生錯誤: {e}")
        return pd.DataFrame()

# === 視覺化4: 學生平台使用量（需 df_km 合併） ===
def visualize_student_usage_scatter(df):
    if '平台使用量' in df.columns:
        # 創建一個圖表
        plt.figure(figsize=(14, 10))
        
        # 根據年級添加顏色編碼
        grades = sorted(df['年級'].unique())
        colors = plt.cm.viridis(np.linspace(0, 1, len(grades)))
        
        # 計算總體相關係數
        corr = df['平台使用量'].corr(df['總得分率'])
        
        # 為每個年級繪製散點圖和回歸線
        for i, grade in enumerate(grades):
            grade_data = df[df['年級'] == grade]
            
            # 散點圖
            plt.scatter(grade_data['平台使用量'], grade_data['總得分率'], 
                      color=colors[i], alpha=0.7, s=50, 
                      label=f'{grade}年級 (n={len(grade_data)})')
            
            # 年級回歸線
            if len(grade_data) > 2:  # 至少需要3個點來繪製有意義的回歸線
                sns.regplot(x='平台使用量', y='總得分率', data=grade_data,
                           scatter=False, color=colors[i], line_kws={'linestyle':'--'})
        
        # 添加總體回歸線
        sns.regplot(x='平台使用量', y='總得分率', data=df, 
                  scatter=False, color='black', line_kws={'linewidth':2})
        
        # 計算使用量分組的平均得分率
        bins = [0, 10, 30, 60, 120, df['平台使用量'].max()]
        labels = ['極低', '低', '中', '高', '極高']
        df['使用量分組'] = pd.cut(df['平台使用量'], bins=bins, labels=labels)
        group_means = df.groupby('使用量分組')['總得分率'].mean()
        group_counts = df.groupby('使用量分組').size()
        
        # 在圖表上標出分組平均值
        for i, (group, mean_score) in enumerate(group_means.items()):
            if pd.notna(mean_score):  # 確保平均值不是 NaN
                # 計算每個組的中心點
                center = (bins[i] + bins[i+1]) / 2
                plt.plot(center, mean_score, 'ro', markersize=12, alpha=0.9)
                plt.annotate(f'平均: {mean_score:.2f}\nn={group_counts[group]}', 
                           (center, mean_score), 
                           textcoords="offset points",
                           xytext=(0,10), 
                           ha='center',
                           fontsize=12,
                           fontproperties=my_font,
                           bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="gray", alpha=0.8))
        
        # 設置圖表標題和標籤
        plt.title('平台使用量與平均得分率關係分析', fontsize=16, fontproperties=my_font)
        plt.xlabel('平均平台使用量', fontsize=14, fontproperties=my_font)
        plt.ylabel('平均數學得分率', fontsize=14, fontproperties=my_font)
        plt.grid(True, linestyle='--', alpha=0.7)
        
        # 添加相關係數文字
        plt.annotate(f'整體相關係數 (r): {corr:.3f}', 
                   xy=(0.02, 0.02), 
                   xycoords='axes fraction',
                   fontsize=14,
                   fontproperties=my_font,
                   bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="gray", alpha=0.8))
        
        # 圖例設定
        legend = plt.legend(title='年級分組', fontsize=12)
        plt.setp(legend.get_title(), fontproperties=my_font)
        for text in legend.get_texts():
            text.set_fontproperties(my_font)
        
        # 保存圖表
        plt.tight_layout()
        plt.savefig('平台使用量_vs_成績.png', dpi=300)
        print("✅ 已生成平台使用量與成績關係圖")


# === 主函數 ===
def main():
    files = [f'113年度_學力測驗_金門縣_數學{i}年級成績_202406.csv' for i in range(3,9)]
    df_scores = load_all_data(files)

    if df_scores is not None:
        print(f"✅ 成績資料載入成功，共 {len(df_scores)} 筆")
        visualize_grade_differences(df_scores)
        visualize_gender_differences(df_scores)
        visualize_school_differences(df_scores)
        
        try:
            # 嘗試先導入finalkm模組
            excel_file = "力宇教育-時數報表113.08.01-113.12.31 (2).xlsx"
            try:
                df_km = finalkm.get_km_usage_data_from_excel(excel_file)
            except ImportError:
                print("⚠️ 找不到finalkm模組，使用內部實現...")
                df_km = get_km_usage_data_from_excel(excel_file)
                
            if not df_km.empty:
                print(f"✅ 平台資料載入成功，共 {len(df_km)} 筆")
                df_merged = pd.merge(df_scores, df_km, on=["學校名稱", "年級"], how="inner")
                print(f"✅ 成績與平台資料合併成功，共 {len(df_merged)} 筆")
                visualize_student_usage_scatter(df_merged)
            else:
                print("⚠️ 無法載入平台使用數據或數據為空。")
        except Exception as e:
            print(f"⚠️ 合併平台資料時出錯：{e}")
    else:
        print("❌ 未能載入任何成績資料")

if __name__ == '__main__':
    main()