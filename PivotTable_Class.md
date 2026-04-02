## PivotTable Class (ExcelSteps) Reference

This documents the ExcelSteps PivotTable class gives ability to make many types of pivot tables through its `.MakePivotTableProcedure`. The class uses a provisioned ExcelSteps `tblRowsCols` instance (`PivotTable.tblSrc`) for raw data wayfinding.

J.D. Landgrebe/Data Delve LLC, March 31, 2026

## Examples Created by `MakePivotTableProcedure`
(See `tests_XLSteps.xlsm` `Interface` module demos)
Source Data
![[Demo1 and 2_src.png|350]]

`Interface.PivotTableDemo1 sub`
Sum by Category Row Field/SubCategory Column Field (PivotDemo1 sub)
![[Demo1_pvt.png|400]]

`Interface.PivotTableDemo2 sub`
Sum by Category and SubCategory Row Fields/no Column Fields/with Column Totals (PivotDemo2 sub)
![[Demo2_pvt.png|400]]

`Interface.PivotTableDemo3 sub`
Source week+year Data Summary of multiple metrics by Store and Prodtype
Discounts, Markdowns, COGS metrics in separate raw data columns
![[Demo3_src.png|500]]

Row Fields: Store, Prodtype; params: {"DataPivotFieldOrientation":xlRowField,"DataPivotFieldPosition", 2}
Analytes: sum Discounts, sum Markdowns, sum COGS
![[Demo3_pvt.png|550]]

---
## Common Call Patterns

### Simple row/column single-analyte
```vb
rowFields = Array("Category")
colFields = Array("SubCategory")
analytes = Array(Array("Amount", xlSum))

pvt.MakePivotTableProcedure pvt, tblSrc, rowFields, colFields, "PivotOut", analytes
```

### multi-analyte with two column variables
This is the call pattern for the Demo3 example above where there are three metric columns to summarize in row blocks.  The `DataPivotFieldOrientation` param specifies adding the S-Values (Sigma-Values) multi-analyte field as either a row or column field. `DataPivotFieldPosition` specifies its position relative to other fields --interposed between `Store` and `ProdType` in this case.
```vb
rowFields = Array("Store", "Prodtype")
colFields = Array("Week", "Year")
analytes = Array( _
    Array("Discounts", xlSum, "Discounts"), _
    Array("Markdowns", xlSum, "Markdowns"), _
    Array("COGS", xlSum, "COGS") _
)

Set dictParams = ExcelSteps.New_Dictionary
dictParams.Add "DataPivotFieldOrientation", xlRowField
dictParams.Add "DataPivotFieldPosition", 2

pvt.MakePivotTableProcedure pvt, tblSrc, rowFields, colFields, "PivotOut", analytes, _
    dictParams:=dictParams
```

---
## Primary Procedure

```vb
MakePivotTableProcedure(pvt, tblSrc, rowFields, colFields, shtDest, _
    analytes, Optional filters = vbNullString, _
    Optional sortOrder = vbNullString, _
    Optional dictSubT As Object = Nothing, _
    Optional dictParams As Object = Nothing) As Boolean
```

Execution pipeline:
1. InitPivotTable
2. CreatePivotCacheAndTable
3. ValidateFieldSpecs
4. ValidateAnalytes
5. ConfigurePivotFields
6. ApplySortOrder
7. FormatPivotTable
8. ConvertPivotToValues (default True; conditional on dictParams\["isReplaceWithVals"]) )
9. SetOutputRanges

---
## Argument Usage

### Required
- pvt: PivotTable class instance
- tblSrc: source tblRowsCols object; requires tblSrc.rngTable
- rowFields: one field name string or array of row field names
- colFields: one field name string, array of column field names, or vbNullString for row-only layout
- shtDest: destination sheet name for pivot output
- analytes: array of analyte definitions

### Optional
- filters: vbNullString or array of filter specs
- sortOrder: asc|ascending|desc|descending (applies to first column field)
- dictSubT: dictionary of per-field subtotal toggles
- dictParams: dictionary of layout/format overrides

---
## Analytes Array Format

Analytes are two-element arrays: Array(fieldName, aggregation)

* First array element is a field name in raw data
- Second is an XlConsolidationFunction (xlSum, xlCount etc.)

Example:
```vb
analytes = Array( _
    Array("Discounts", xlSum), _
    Array("Markdowns", xlSum), _
    Array("COGS", xlSum))
```

---
## Filters Format

filters supports either:
- "FieldName" (adds a page filter field)
- Array("FieldName", "CurrentPageValue")

Example:
```vb
filters = Array(Array("Store", "Store1"), "Year")
```

---
## dictSubT Options

Purpose: subtotal toggle by field name.

Dictionary keys:
- field name string (for row/column fields)

Dictionary values: True/False: enable/disable field subtotal

Default behavior:
- all row/column field subtotals are off unless overridden by dictSubT

Example:
```vb
Set dictSubT = ExcelSteps.New_Dictionary
dictSubT.Add "Store", True
dictSubT.Add "Prodtype", False
```

---
## dictParams Options

Supported keys:
- RowAxisLayout: xlTabularRow, xlCompactRow, xlOutlineRow (Standard Excel Pivot table layout options
- RepeatAllLabels: xlRepeatLabels, xlDoNotRepeatLabels (Standard Excel Pivot table layout option)
- isRowGrand: True/False
- isColGrand: True/False
- DataPivotFieldOrientation: xlRowField or xlColumnField (when multiple analytes exist)
- DataPivotFieldPosition: numeric position for DataPivotField within that axis
- IsDestNewWorkbook: Toggle create pivot table in new workbook versus as sheet in source data workbook
- isReplaceWithVals: Toggle replacing final pivot table with its values

Default behavior:
- RowAxisLayout = xlTabularRow
- RepeatAllLabels = xlRepeatLabels
- isRowGrand = False
- isColGrand = False
- DataPivotFieldOrientation/Position are not applied unless explicitly provided

Example (See Example 3 above):
```vb
Set dictParams = ExcelSteps.New_Dictionary
dictParams.Add "DataPivotFieldOrientation", xlRowField
dictParams.Add "DataPivotFieldPosition", 2
dictParams.Add "isRowGrand", False
dictParams.Add "isColGrand", False
```

---
## Output Attributes Set by Class

After MakePivotTableProcedure succeeds:
- pvt.wkbkDest: destination workbook
- pvt.wkshtDest: destination worksheet
- pvt.pvtCache: pivot cache
- pvt.pvtTable: pivot table object
- pvt.rngDataOut: data body range pointer (before/after value conversion remap)
- pvt.rngTableOut: used range of output sheet
- pvt.tblOut: tblRowsCols initialized to destination sheet

---
## Practical Notes

- Defaults format the table as normalized rows/columns as much as possible: `RowAxisLayout=xlTabularRow` and `RepeatAllLabels=xlRepeatLabels` results in table that approaches 
- Use `DataPivotFieldOrientation` and `DataPivotFieldPosition` with raw data having multiple analyte/metric columns to place Excel's "Sigma values" field interposed within row or column fields
- If colFields is vbNullString, class builds a row-only pivot.
- sortOrder only applies when at least one column field and one data field exist.
- Destination sheet is recreated each run when name already exists.
- For test stability, validate output values and headers from wkshtDest after ConvertPivotToValues.