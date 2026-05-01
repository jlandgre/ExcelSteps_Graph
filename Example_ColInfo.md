* On sheet ExcelSteps.shtColInfo = "colinfo_" in ColInfo.xlsx workbook
* Example shows mockup metadata for two tables: Columns G and H are table names (e.g. tblRowsCols.TblName attribute)
* Integers in Columns G and H denote variable's inclusion and column order in normalized tbl
* `True` in Column C/IsIndex denotes column is an index column
* Column B to A mapping defines renaming during normalization
* Column I, `rand` is included for use during validation to scramble row order and demonstrate ability to sort by `colInfo.CurTbl` columns G or H
* `Not_in_normalized_tbl` is example of raw data column show for information only or for use in future, calculation capability (e.g will be dropped during normalization)

| A              | B                     | C       | D                        | E     | F             | G          | H          | I        |
| -------------- | --------------------- | ------- | ------------------------ | ----- | ------------- | ---------- | ---------- | -------- |
| VarNameNorm    | VarNameRaw            | IsIndex | Description              | units | data_type_VBA | BR_Example | Second_Tbl | rand     |
| Column_second1 | second_raw1           | True    | Dummy date               |       | Date          |            | 1          | **0.22** |
| Column_second2 | second_raw2           |         | xyz                      |       | Double        |            | 2          | **0.75** |
| Column_second3 | second_raw3           |         | abc                      |       | String        |            | 3          | **0.81** |
| Column_second4 | second_raw4           |         | def                      |       | String        |            | 4          | **0.35** |
| Location       | Locn_Raw              | TRUE    | Store location           |       | String        | 1          |            | **0.53** |
| ProdType       | ProdType_Raw          | TRUE    | Product Type             |       | String        | 2          |            | **0.41** |
| Year           | Year_Raw              | TRUE    | Fiscal year              |       | Long          | 3          |            | **0.06** |
| SerialWeek     | Week_Raw              | TRUE    | Fiscal week              |       | Long          | 4          |            | **0.20** |
| Net_Sales      | Sales_Raw             |         | Net Sales for prodtype   | USD   | Double        | 5          |            | **0.13** |
| Discounts      | Discounts_Raw         |         | Discounts for prodtype   | USD   | Double        | 6          |            | **0.21** |
| Markdowns      | Markdowns_Raw         |         | Markdowns for prodtype   | USD   | Double        | 7          |            | **0.47** |
| COGS           | COGS_Raw              |         | COGS for prodtype        | USD   | Double        | 8          |            | **0.84** |
|                | Not_in_normalized_tbl |         | deleted during normalize |       | Long          | x          |            | **0.76** |

JDL 5/1/26