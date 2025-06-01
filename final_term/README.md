## 程式說明：

### 1. 數據處理類 (Data Processing)
📄 [a.py - 康軒平台數據處理（原始版）](https://github.com/cpeggy/EducationalDataMining/blob/main/final_term/a.py)

- 功能: 解析康軒平台班級CSV檔案，提取數學任務統計
- 輸入: 康軒平台資料夾中的班級CSV檔案
- 輸出: 控制台顯示統計表格
- 關鍵指標: 總派出任務數、完成任務數、完成率、正答率

📄 [a_copy.py - 康軒平台數據處理（修改版）](https://github.com/cpeggy/EducationalDataMining/blob/main/final_term/a_copy.py)

- 功能: a.py的增強版，增加CSV輸出功能
- 新增: 自動輸出platform_data.csv檔案
- 用途: 為後續視覺化分析提供標準化數據

📄 [app_modified.py - 學力測驗成績處理（簡化版）]()

- 功能: 載入並處理學力測驗成績資料
- 輸入: 113年度學力測驗各年級CSV檔案
- 輸出: test_scores.csv - 各學校平均得分率
- 特色: 多編碼支援、資料驗證

📄 app＿linyu.py - 力宇平台數據處理

- 功能: 從力宇教育Excel報表中提取平台使用數據
- 輸入: 力宇教育時數報表Excel檔案
- 輸出: liyou_platform_data.csv
- 分析指標: 練習題數、測驗卷數、影片使用量、累積時數

📄 a_liyou.py - 力宇平台統計表生成

- 功能: 生成力宇平台詳細使用統計表
- 特色: 類似康軒平台的表格格式
- 額外功能: 包含視覺化氣泡圖生成


### 2. 視覺化分析類 (Visualization & Analysis)
📊 visual_ku.py - 康軒平台視覺化分析

- 功能: 分析康軒平台使用與學力測驗成績關係
- 生成圖表:
  - 任務完成率 vs 測驗得分率散點圖
  - 任務正答率 vs 測驗得分率散點圖
  - 多維氣泡圖

📊 visual_lin.py - 力宇平台視覺化分析

- 功能: 分析力宇平台使用與學力測驗成績關係
- 生成圖表:
  - 平台使用量 vs 測驗得分率
  - 自我練習成績 vs 測驗得分率
  - 影片使用量 vs 測驗得分率
  - 累積影片時數 vs 測驗得分率
  - 多維氣泡圖

### 3. 綜合分析類 (Comprehensive Analysis)
📈 app.py - 原始完整分析程式

- 功能: 學力測驗多維度分析
- 分析維度:
  - 年級差異分析
  - 性別差異分析
  - 學校間比較分析
  - 輸出: 多張分析圖表

📈 c.py - 進階統計分析

- 功能: 深度統計分析與熱力圖生成
- 特色:
  - KDE密度分布圖
  - 班級成績熱力圖
  - 多編碼檔案讀取
  - 中文字型自動設定

### 4. 實用工具類 (Utility Tools)
🔧 b.py - 多檔案批次處理工具
- 功能: 批次處理多個CSV檔案
- 特色:
  - 自動編碼偵測
  - 錯誤處理機制
  - 學校統計彙整
- 用途: 大量數據檔案的快速統計
