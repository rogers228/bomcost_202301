# 2023/1月 重新建構bom製程 成本表 程式

## 系統架構
1. 發布時同步至nas
2. 採分散式架構，由使用者工作站執行
3. 採用Gui介面，Excel輸出

## 功能選項
1. 輸入品號後，產出BOM 製程成本表
2. 可選項，泵浦不展階
3. 自動製表並檢查、提示、試算

## 資料抓取特殊條件
|  欄位   | 條件  | 抓取 |
|  ----  | ----  | ---- |
|單位    | 採購件P     | 第一順位該產品單位換算首筆,第二順位庫存單位MB004
|        | 加工件M,S,Y | 庫存單位MB004
|        | 銷售件Q | 固定為PCS
|製程單價 | 泵浦(品號6AA開頭 & 不展階) | 售價MB053
|         | 採購件P | 最新進價(本國幣別NTD) MB050
|         | 加工件M,S,Y & 1自製 | 備註 MF023
|         | 加工件M,S,Y & 2托外 | 加工單價 MF018
|單價     |                    | 統一換算為PCS單價
|固定人時 |                     | 僅自製加工件時顯示
|固定機時 |                     | 僅自製加工件時顯示
|工時批量 |                     | 僅自製加工件時顯示

## 資料檢查
系統將自動檢查，符合以下者代表錯誤，將高亮度顯示

|  欄位   | 檢查條件  | 結果 | 備註 |
|  ----  | ----  | ---- |---- |
|品號屬性 | P件不應該有BOM架構，或有BOM應為S件or M件 | 黃 |
|BOM階層 | 最下階應為P, Q件，或應再建立P件為子件 | 黃 |
|換算單位| 加工件M,S,Y途程加工單位是PCS以外的單位卻無換算單位 | 黃 |
|        | 採購件P採購計價單位是PCS以外的單位卻無換算單位 | 黃 |
|        | 採購件P單位換算不符合最新進貨 | 黃 |
|廠商簡稱| 供應商已禁止交易| 黃
|        | 連續最後三次採購進貨者，提示更新為主供應商| 橘
|單位    | 加工件M,S,Y單位&不符合標準途程加工單位| 黃 | 產品途程作業92N001-02
|        | 採購件P單位&不符合標準廠商採購計價單位| 黃 | (第一順位)產品途程作業92N001-03
|        | 採購件P單位&採購進貨最後一筆的計價單位| 黃 | (第二順位)
|製程單價 | 未維護為0者 | 黃 |
|        | 加工件M,S,Y單價不符合標準廠商途程加工單價| 黃 | (第一順位)產品途程作業92N001-01
|        | 加工件M,S,Y單價未更新托外進貨單價| 橘 | (第二順位)
|        | 採購件P單價未更新採購進價| 橘 |

## 試算
1. 試算零件成本逐個加總金額
2. 試算材料KG單價

## 未完成
換算單位分子檢查
採購件 換算單位檢查