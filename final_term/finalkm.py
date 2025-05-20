import pandas as pd
import re

# 目標欄位
target_cols = [
    '自我練習平均成績', '自我測驗卷數', '自我練習題目數',
    '完成老師指派影片數', '自我點播影片數', '累積影片總時數'
]

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