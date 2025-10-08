# Label Number Generator

## Label Number Generator_Py v1.0.9

### 執行方式
1. 執行 **Label_number_generator_Py.exe**（應用程式）

### 功能
* 自定義 起始號 & 結束號 (SN)
* 自定義 每箱序號數量 (批次 QR Code)
* 匯出 `.xlsx` / `.csv` / `.xls`
* GUI 介面化操作
* 雙模式選擇：[外箱]、[棧板]

### 適用環境
* Windows 10 22H2 以上版本

---

### 更新紀錄
- v1.0.9 更新輸出 **Office 2003 用格式 `.xls`**
- v1.0.8 更新外箱格式彈性
- v1.0.7 新增自訂棧板號格式 `C/NO.TC4006F5J0251`
- v1.0.6 新增雙模式選擇 [外箱、棧板]
- v1.0.5 修正溢位錯誤
- v1.0.4 移除輸出檔案指定路徑功能
- v1.0.3 調整 UI 排版
- v1.0.2 修正需求排版與 QRCODE 自動換行
- v1.0.1 修正開頭錯誤格式

---

## Label Number Generator_VBA v1.0.4

### 執行方式
1. 開啟 **Label number generator_VBA.xlsm**（巨集 Excel）

### 功能
* 輸入起始序號與結束序號  
* 自訂每箱序號數量  
* 箱號內建 C/NO. 格式，可輸入 `1` 或其它自定義格式  
* 自動產生標籤編號於第一個分頁  
* 標籤機軟體可直接匯入第一個分頁內容  

---

### 更新紀錄
- v1.0.4 修正溢位錯誤  
- v1.0.3 修正函數類性  
- v1.0.2 修正標籤機軟體讀取錯誤  
- v1.0.1 修正長度單格上限計算錯誤  

---

### 環境需求
* 僅支援 **Microsoft Excel 2019 以上版本**

---

### 開啟巨集方法
1. EXCEL → 檔案 → 選項 → 信任中心 → **巨集設定 → 啟用 VBA 巨集**  
2. EXCEL → 檔案 → 選項 → 信任中心 → **信任位置 → 新增位置**  
   - 指定信任的資料夾（包含子層）  
3. 若從 **網路磁碟（SMB）** 開啟：
   - EXCEL → 檔案 → 選項 → 信任中心 → **信任的文件 → 允許信任特定網路文件**  
4. 若仍無法開啟：
   - 下載 [allow_macro_from_internet.reg](https://github.com/bfc8g4v63/Tool/releases/download/v1.0.5/allow_macro_from_internet.reg)  
   - 以系統管理員身分執行並啟用  

---

## 備註
* Python 版與 VBA 版皆可輸出相容格式供標籤機讀取。  
* 推薦以 Py 版進行批次運算、VBA 版用於單機快速輸出。  