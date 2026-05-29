# Excel-AST: 4、拉迪亚码头-二航院_AST

> Semantic representation of cleaned workbook structure

## Workbook Overview

| Metric | Value |
|--------|-------|
| Sheets | 2 |
| Total rows (all sheets) | 2164 |
| Total columns (all sheets) | 12 |

## Sheet: `分项汇总表`

- **Rows:** 16
- **Columns:** 4
- **Range:** 2,4,1,16

### Column Type Distribution

| Type | Count |
|------|-------|
| text | 2 |
| empty | 1 |
| number | 1 |

### Header Preview (first 3 rows)

**Row 1:**  | 拉迪亚码头投标报价表-二航院 |  | 

**Row 2:**  | 序号 | 名称 | 二航院方案（PHC桩）

**Row 3:**  |  |  | 金额（美元）


### Column Analysis

| Col | Letter | Type | Non-empty | Samples |
|-----|--------|------|-----------|---------|
| 1 | A | empty | 0 |  |
| 2 | B | text | 10 | 拉迪亚码头投标报价表-二航院, 序号, B |
| 3 | C | text | 13 | 名称, 准备工作, 疏浚工程 |
| 4 | D | number | 13 | 二航院方案（PHC桩）, 金额（美元）, 88722 |

### Data Preview (first 30 rows)

|  | 拉迪亚码头投标报价表-二航院 |  |  |
| --- | --- | --- | --- |
|  | 序号 | 名称 | 二航院方案（PHC桩） |
|  |  |  | 金额（美元） |
|  | B | 准备工作 | 88722 |
|  | C | 疏浚工程 |  |
|  | D | 地基处理 | 16476810 |
|  | E | 码头 | 50033437 |
|  |  | 护岸 | 23514886 |
|  | F | 路面工程 | 29471443 |
|  | G | 电气 | 14606322 |
|  | H | 房建 | 1119479 |
|  | I | Gate Complex | 1977668 |
|  |  | 直接费小计 | 137288767 |
|  |  | 关税 | 11685729.95 |
|  |  | 项目总计 | 148974496.95 |
|  |  |  |  |

## Sheet: `清单 `

- **Rows:** 2148
- **Columns:** 8
- **Range:** 1,16349,1,2148

### Column Type Distribution

| Type | Count |
|------|-------|
| number | 4 |
| text | 4 |

### Header Preview (first 3 rows)

**Row 1:**  | THE BILL OF QUANTITIES（拉迪亚码头-港湾-二航院） |  |  |  |  |  | 

**Row 2:**  | Item No. | Description2 | 清单说明 | Unit2 | Quantity |  | 

**Row 3:**  |  |  |  |  |  | 综合单价（美元） | 合计（美元）


### Column Analysis

| Col | Letter | Type | Non-empty | Samples |
|-----|--------|------|-----------|---------|
| 1 | A | number | 1008 |  |
| 2 | B | text | 642 | THE BILL OF QUANTITIES（拉迪亚码头-港湾-二航院）, Item No., CLASS A |
| 3 | C | text | 2059 | Description2, PRELIMANARIES |
| 4 | D | text | 2059 | 清单说明, PRELIMANARIES |
| 5 | E | text | 1702 | Unit2 |
| 6 | F | number | 1088 | Quantity |
| 7 | G | number | 2073 | 综合单价（美元） |
| 8 | H | number | 2075 | 合计（美元） |
