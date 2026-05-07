* On sheet ExcelSteps.shtColInfo = "colinfo_" in ColInfo.xlsx workbook
* Example shows mockup metadata for two tables: Columns G and H are table names (e.g. tblRowsCols.TblName attribute)
* Integers in Columns G and H denote variable's inclusion and column order in normalized tbl
* `True` in Column C/IsIndex denotes column is an index column
* Column B to A mapping defines renaming during normalization
* `rand` is included for use during validation to scramble row order and demonstrate ability to sort by `colInfo.CurTbl` columns I and J
* `Not_in_normalized_tbl` is example of raw data column shown for information only or for use in future, calculation capability (e.g will be dropped during normalization)
* FillVals column defines mapping for individual values ("BLANK" is special reference to how to fill blank cells. "Locn10":"Locn1" defines replacing "Locn10" string (whole cell value) with "Locn1" )
* FilterVals column defines table filtering by column values --KeepOnly is one option to start but could also have KeepList, KeepExcept, KeepExceptList to filter in different ways

| A              | B                     | C       | D                        | E     | F             | G                            | H                 | I          | J          | K        |
| -------------- | --------------------- | ------- | ------------------------ | ----- | ------------- | ---------------------------- | ----------------- | ---------- | ---------- | -------- |
| VarNameNorm    | VarNameRaw            | IsIndex | Description              | units | data_type_VBA | FillVals                     | FilterVals        | BR_Example | Second_Tbl | rand     |
| Location       | Locn_Raw              | TRUE    | Store location           |       | String        |                              | {KeepOnly:Online} | 1          |            | **0.83** |
| ProdType       | ProdType_Raw          | TRUE    | Product Type             |       | String        | {BLANK:Unknown,Locn10:Locn1} |                   | 2          |            | **0.65** |
| Year           | Year_Raw              | TRUE    | Fiscal year              |       | Long          |                              |                   | 3          |            | **0.21** |
| SerialWeek     | Week_Raw              | TRUE    | Fiscal week              |       | Long          |                              |                   | 4          |            | **0.81** |
| Net_Sales      | Sales_Raw             |         | Net Sales for prodtype   | USD   | Double        |                              |                   | 5          |            | **0.53** |
| Discounts      | Discounts_Raw         |         | Discounts for prodtype   | USD   | Double        |                              |                   | 6          |            | **0.67** |
| Markdowns      | Markdowns_Raw         |         | Markdowns for prodtype   | USD   | Double        |                              |                   | 7          |            | **0.49** |
| COGS           | COGS_Raw              |         | COGS for prodtype        | USD   | Double        |                              |                   | 8          |            | **0.47** |
|                | Not_in_normalized_tbl |         | deleted during normalize |       | Long          |                              |                   | x          |            | **0.45** |
| Column_second1 | second_raw1           |         | Dummy date               |       | Date          |                              |                   |            | 1          | **0.24** |
| Column_second4 | second_raw4           |         | def                      |       | String        |                              |                   |            | 4          | **0.66** |
| Column_second2 | second_raw2           |         | xyz                      |       | Double        |                              |                   |            | 2          | **0.94** |
| Column_second3 | second_raw3           |         | abc                      |       | String        |                              |                   |            | 3          | **0.48** |


JDL 5/1/26; updated 5/7/26