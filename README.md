# PyreEQ
![ICON](https://github.com/Chih0321/PyreEQ/blob/main/media/icon.png#pic_left = 100x100)  
PyreEQ提供計算SAP2000(v14)所選群組週期計算，並依`鐵路橋梁耐震設計規範[110年]`進行地震力分配，並assign進SAP2000模型

## 安裝
1. 由`Release`頁面下載[PyreEQ](https://github.com/Chih0321/PyreEQ/releases/tag/v1.0.0)
2. 執行資料夾內`PyreEQ.exe`

## 快速開始
1. ### 模型檔案選擇
    ![STEP1](https://github.com/Chih0321/PyreEQ/blob/main/media/s1.png)  
    - 由選擇按鈕選擇要執行的SAP2000模型(.sdb)  
    - 會第一次執行模型抓取`GROUP`資訊  
2. ### 週期計算  
    ![STEP2](https://github.com/Chih0321/PyreEQ/blob/main/media/s2.png)
    - 選擇三方向欲計算週期群組，Unit-X, Unit-Y可以複選，Unit-z請依標籤選擇上下構群組  
    - 若無選擇或選擇到僅含剛棒(不含質量)群組，程式會因無法計算直接卡死，請重新執行程式  
    ![STEP21](https://github.com/Chih0321/PyreEQ/blob/main/media/s21.png)  
    - 執行`計算週期`按鈕  
    ![STEP22](https://github.com/Chih0321/PyreEQ/blob/main/media/s22.png)  
    - 程式顯示計算所得週期，亦同步輸出計算結果於模型同路徑之`01_period_results.xlsx`  
3. ### 使用者自行計算地震力加速度係數
   - 由對應規範excel計算
4. ### 計算及施加分配地震力
   ![STEP3](https://github.com/Chih0321/PyreEQ/blob/main/media/s3.png)
    - 頁面切換至施加地震力
    - 表格中Groups由週期計算時自動帶入，使用者可更改群組
    - Factors填入計算所得地震力加速度係數   
![STEP31](https://github.com/Chih0321/PyreEQ/blob/main/media/s31.png)
    - 因鐵路110年規範地震力分配為適用規則橋梁以第一振態為主，因此分配力總和會低於總水平力
    - 使用者可調整計算總分配力須達多少總水平力以上，scale方法為直接將結果線性調整至需求
    - Checkbox不勾選時，下方調整百分比不作用
    - 比例預設為須至少為90%總水平力  
![STEP32](https://github.com/Chih0321/PyreEQ/blob/main/media/s32.png)
    - 執行`施加地震力`按鈕
    - 執行結束後SAP模型力量加載完成，計算過程會輸出至`02_eqforce_results.xlsx`   

## 限制條件
1. 地震力計算都以模型`Global`座標系XYZ為計算方向
2. 地震力皆施加在NODE上，若在意桿件TOP CENTER補扭矩請自行施加(preEQ也無)
3. 質量點與力量施加因SAP2000 API開放語法為認joint local axis，因此如果有質量的node建議不轉座標，否則其被當作在`Global`座標XYZ上運動
4. `GROUP`群交界桿件建議不框(SAP是認桿件去選NODE)，只要確定要計算的NODE都有考慮到即可，若都框到，程式以第1筆數值進行計算

## 使用手冊
**施工中**

## 參考資料
1. 鐵路橋梁耐震設計規範
2. [SAP2000 Python API語法](https://github.com/Junjun1guo/pythonInteractSAP2000/blob/main/pythonInterSAP2000.py#L8648)

## ChangeLog
[v1.0.0](https://github.com/Chih0321/PyreEQ/releases/tag/v1.0.0)
- Initial Release