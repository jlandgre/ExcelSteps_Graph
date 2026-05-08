## Purpose
Create and validate `ImportParseNorm.cls` for importing, parsing, and normalizing raw data. Instanced as `imptbl` in project code.

## Background

This change follows [[VBA Project Changes]] sequencing except that we develop only the `Change_ChangeName.md` planning note before coding.

`ImportParseNorm` uses a `ColInfo` instance (`colinfo`) and two parameter dictionaries (`dParamsImport`, `dParamsParse`) to drive importing, parsing, and normalization of raw data in a configuration-driven way — replacing the hard-coded approach in `case_studies/2026_0507_ImportCode/BRImport_Example.cls`.

Key design decisions:
- All needed inputs are passed as args to `Init` and set as public attributes (no pre-set attributes before `Init`)
- `files.pfImportFile` (new `ProjFiles` attribute) holds the raw data file path; no separate path arg needed
- `colinfo` may be pre-instanced by the caller; `Init` checks `colinfo Is Nothing` and `colinfo.tbl Is Nothing` — initializes fresh if needed, then always calls `colinfo.SetCurTbl(colinfo, curTbl)`
- `curTbl` is required (not Optional) — always needed to configure `colinfo` for a specific table
- Raw workbook is never mutated in place: `OpenRawData` copies the raw sheet to a new temp workbook, closes the original, then provisions `tblRaw` against the temp workbook
- `tblNorm` lives as a new sheet in the temp workbook
- `WriteNormalized` uses a column-by-column copy approach driven by the integer order values in `colinfo.tbl CurTbl` column — no row-by-row iteration
- `FilterRows` (`KeepOnly`) sorts `tblNorm` by the filter column, uses `FindInRange` to locate first/last matching row, then bulk-deletes non-matching rows above and below in two `.Delete` operations
- `dParamsImport`: `{FileType:xlsx}`; `dParamsParse`: `{RawShape:rowscols}` — minimal keys for initial implementation
- `tblUnstructured` class (future change) needed for non-rowscols raw shapes; `ParseRaw` is a dummy pass-through for now
- Custom project validations (like `ValidateRawDataProcedure` in example) are not part of this class — caller can add those in the driver

## Data I/O Descriptions

**Inputs to `Init`:**

| Arg | Type | Description |
|---|---|---|
| `imptbl` | `ImportParseNorm` | self-reference |
| `colinfo` | `Object` | `ColInfo` instance; may be pre-initialized or `Nothing` |
| `files` | `Object` | `ProjFiles` instance; `files.pfImportFile` required |
| `curTbl` | `String` | Required; passed to `colinfo.SetCurTbl` |
| `dParamsImport` | `Object` | Dictionary; e.g. `{FileType:xlsx}` |
| `dParamsParse` | `Object` | Dictionary; e.g. `{RawShape:rowscols}` |

**Outputs — attributes set after `Init`:**

| Attribute | Value |
|---|---|
| `colinfo` | Initialized `ColInfo` with `CurTbl` and `rngRowsCurTbl` set |
| `files` | `ProjFiles` instance |
| `curTbl` | Table name string |
| `dParamsImport` | Import params dictionary |
| `dParamsParse` | Parse params dictionary |
| `tblRaw`, `tblParsed`, `tblNorm` | `Nothing` until sub-procedures run |

**Outputs — attributes set after sub-procedures:**

| Attribute | Set by | Value |
|---|---|---|
| `tblRaw` | `OpenRawData` | `tblRowsCols` provisioned on copied raw sheet |
| `tblParsed` | `ParseRaw` | Points to `tblRaw` when `RawShape=rowscols` |
| `tblNorm` | `WriteNormalized` | `tblRowsCols` provisioned on new norm sheet in temp workbook |

## Project Architecture

**New class: `ImportParseNorm.cls`**
- Instanced as `imptbl`
- All attributes public; no `Property Get/Let`

**New `ProjFiles` attribute:**
- `pfImportFile As String` — full path to raw data import file; set by caller or `files.Init` when `IsTest = True`

**Public attributes:**
```vb
Public colinfo As Object          ' ColInfo instance
Public files As Object            ' ProjFiles instance
Public curTbl As String           ' Current table name
Public dParamsImport As Object    ' Import params dictionary
Public dParamsParse As Object     ' Parse params dictionary
Public tblRaw As Object           ' tblRowsCols for copied raw sheet
Public tblParsed As Object        ' Points to tblRaw (rowscols) or parsed result
Public tblNorm As Object          ' tblRowsCols for normalized output sheet
```

**Methods:**
- `Public Function Init(imptbl, colinfo, files As Object, curTbl As String, dParamsImport As Object, dParamsParse As Object) As Boolean`
  - Sets all args as public attributes
  - Checks `colinfo Is Nothing` or `colinfo.tbl Is Nothing`; if so calls `colinfo.Init(colinfo, files)`
  - Always calls `colinfo.SetCurTbl(colinfo, curTbl)`
- `Public Function ImportParseNormProcedure(imptbl, colinfo, files As Object, curTbl As String, dParamsImport As Object, dParamsParse As Object) As Boolean`
  - Calls `Init`, then `OpenAndValidateRawProcedure`, `ParseRawProcedure`, `NormalizeParsedProcedure` in sequence
- `Public Function OpenAndValidateRawProcedure(imptbl) As Boolean`
  - Calls `OpenRawData`, `ValidateRawStructure`, `ReplaceVals`
- `Public Function OpenRawData(imptbl) As Boolean`
  - Opens `files.pfImportFile` via `OpenFile` (within-project call, not `ExcelSteps.OpenFile`)
  - Copies sheet 1 to a new workbook; closes original
  - Provisions `tblRaw` against new workbook's sheet; `tblAlias`/`tblUAlias` needed for ByRef `Provision`/`Init` calls
  - `RawShape` read inline: `LCase$(CStr(imptbl.dParamsParse.Item("RawShape", "rowscols")))` — no local variable
- `Public Function ValidateRawStructure(imptbl) As Boolean`
  - Iterates `colinfo.rngRowsCurTbl`; for each row where `VarNameNorm` non-blank, confirms `VarNameRaw` exists in `tblRaw.rngHeader` via `rngTblHeaderVal`; fails fast on first missing column
  - `rngVarNormCol`/`rngVarRawCol` kept as loop-cached ranges
- `Public Function ReplaceVals(imptbl) As Boolean` *(was `FillMissingVals` in plan)*
  - For each row in `colinfo.rngRowsCurTbl` with a non-blank `FillVals` entry, parses JSON-like string into a `New Dictionary`; sorts `tblRaw` by that column; applies value replacements and blank-fills via `ApplyFillMapToSortedColumn`
- `Public Function ApplyFillMapToSortedColumn(imptbl, rngDataCol As Range, dict As Object) As Boolean`
  - Iterates `dict.GetKeys`; for each key calls `ApplyFillKeyToSortedColumn`; early-exits if `dict.Count = 0` or keys array empty
- `Public Function ApplyFillKeyToSortedColumn(imptbl, rngDataCol As Range, ByVal sKey As String, ByVal valFill As Variant) As Boolean`
  - `BLANK` key: uses `SpecialCells(xlCellTypeBlanks)` to fill blanks
  - Other keys: `FindInRange` for first match; native `.Find` with `xlPrevious` for last match; bulk-sets range between them
- `Public Function ParseRawProcedure(imptbl) As Boolean`
  - Checks `dParamsParse` `RawShape` key inline; if `"rowscols"` sets `tblParsed = tblRaw` and exits
  - Non-rowscols: `errs.IsFail(True, 1, ...)` — explicit unimplemented error
- `Public Function NormalizeParsedProcedure(imptbl) As Boolean`
  - Calls `WriteNormalized`, then `FilterRows`
- `Public Function WriteNormalized(imptbl) As Boolean`
  - Validates `tblParsed` is a `tblRowsCols`; delegates to `BuildNormMappings` + `WriteNormFromMappings`
- `Public Function BuildNormMappings(imptbl, ByRef aryNorm() As String, ByRef aryRaw() As String, ByRef maxOrder As Long) As Boolean`
  - Two-pass loop over `colinfo.rngRowsCurTbl`: first pass finds `maxOrder`; second pass fills `aryNorm`/`aryRaw` by order index
  - Fails if `maxOrder = 0` or any ordered `VarNameNorm` row has blank `VarNameRaw`
- `Public Function WriteNormFromMappings(imptbl, aryNorm() As String, aryRaw() As String, ByVal maxOrder As Long) As Boolean`
  - Deletes and recreates `"norm_"` sheet; writes header; copies data columns from `tblParsed` by `VarNameRaw`
  - Provisions `tblNorm` against new sheet; `tblAlias` needed for ByRef `Provision` call
- `Public Function FilterRows(imptbl) As Boolean`
  - Iterates `colinfo.rngRowsCurTbl`; for each row with non-blank `FilterVals`, parses JSON-like string via `New Dictionary`
  - If `dict.Exists("KeepOnly")` calls `ApplyKeepOnlyFilter`; other filter options deferred to future
- `Private Function ApplyKeepOnlyFilter(imptbl, ByVal sVarNorm As String, ByVal keepVal As Variant) As Boolean`
  - Sorts `tblNorm` by `sVarNorm`; locates first/last matching row; bulk-deletes rows above and below in two `.Delete` calls
  - Re-provisions `tblNorm` after deletion; `tblAlias` needed for ByRef `Provision` call

**Factory functions added to `Validation.bas`:**
- `New_ImportParseNorm` — enables cross-workbook instantiation from test suite
- `New_tblUnstructured` — enables cross-workbook instantiation of `tblUnstructured`

**`tblUnstructured.cls` implemented** (was noted as "future change" in plan):
- Attributes: `wkbk`, `sht`, `wksht`, `rngTable`
- `Init(tblU, wkbk, sht)` — sets all attributes; defaults `rngTable` to `UsedRange`

## Test Architecture

Tests housed in `tests_ToolboxClasses.bas`. New `procs.ImportParseNorm` procedure group added to `Procedures.cls`. Follow [[vba-testing-create-new-test-procedure]] skill for all wiring steps.

- **Test section location**: immediately below `TestDriver_ToolBox` for ease of navigation (not at bottom of file)
- **Test data**: two static raw data files in `tests/test_data/`:
  - `BR_Raw_Mockup.xlsx` — `BR_Example` raw columns; includes mixed Location values (Online + others), blank `ProdType_Raw` values, and `Locn10` values to exercise `ReplaceVals` and `FilterRows`
  - `Second_Raw_Mockup.xlsx` — `Second_Tbl` raw columns; no `FillVals`/`FilterVals` (simpler validation case)
- **`files.pfImportFile`**: set by test helper using `tst.wkbkTest.Path` + separator + test data subfolder path
- **Test helper**: `InitImportParseNormTest(tst, imptbl, colinfo, files, dParamsImport, dParamsParse, curTbl)` — shared setup sub; instances all objects and calls `imptbl.Init`; companion `CloseImportParseNormWkbk(imptbl)` closes temp workbook

**Test list (as built):**
- `test_ImportParseNorm_Init` — attributes set; colinfo initialized; SetCurTbl called
- `test_ImportParseNorm_OpenRawData` — `tblRaw` provisioned; original workbook closed
- `test_ImportParseNorm_ValidateRawStructure` — passes for valid raw structure
- `test_ImportParseNorm_BuildNormMappings` — `aryNorm`/`aryRaw` arrays and `maxOrder` correct
- `test_ImportParseNorm_ApplyFillMapToSortedColumn` — dict parsed with unquoted values; blanks and key replacements applied
- `test_ImportParseNorm_ReplaceVals` — value replacements applied in-place on `tblRaw`
- `test_ImportParseNorm_WriteNormalized` — `tblNorm` header matches ordered VarNameNorm; data present
- `test_ImportParseNorm_FilterRows` — `KeepOnly:Online` reduces rows to only Online location rows

*Deferred (not built): `test_ImportParseNorm_ParseRaw`, `test_ImportParseNorm_OpenAndValidateRawProcedure`, `test_ImportParseNorm_NormalizeParsedProcedure`, `test_ImportParseNorm_Procedure`*

## Procedure Outline

**ImportParseNorm Class**
- **`Init`** — set all args as attributes; check/init colinfo; call SetCurTbl
- **`ImportParseNormProcedure`** — call Init, OpenAndValidateRawProcedure, ParseRawProcedure, NormalizeParsedProcedure
- **`OpenAndValidateRawProcedure`** — call OpenRawData, ValidateRawStructure, ReplaceVals
- **`OpenRawData`** — open pfImportFile; copy sheet to new workbook; close original; provision tblRaw (rowscols or tblUnstructured per RawShape)
- **`ValidateRawStructure`** — iterate rngRowsCurTbl; for non-blank VarNameNorm rows confirm VarNameRaw in tblRaw header; fail fast on missing
- **`ReplaceVals`** *(was FillMissingVals)* — iterate rngRowsCurTbl FillVals column; parse JSON-like string into New Dictionary; sort tblRaw; call ApplyFillMapToSortedColumn
- **`ApplyFillMapToSortedColumn`** — iterate dict keys; call ApplyFillKeyToSortedColumn for each; early-exit if dict empty
- **`ApplyFillKeyToSortedColumn`** — BLANK key: SpecialCells blanks fill; other keys: FindInRange first + native Find last; bulk range assignment
- **`ParseRawProcedure`** — if RawShape=rowscols set tblParsed=tblRaw; else explicit IsFail error (unimplemented)
- **`NormalizeParsedProcedure`** — call WriteNormalized, FilterRows
- **`WriteNormalized`** — validate tblParsed; call BuildNormMappings then WriteNormFromMappings
- **`BuildNormMappings`** — two-pass loop over rngRowsCurTbl: find maxOrder, then fill aryNorm/aryRaw by order index
- **`WriteNormFromMappings`** — delete/recreate norm_ sheet; write header; copy data columns from tblParsed; provision tblNorm
- **`FilterRows`** — for each FilterVals row parse dict via New Dictionary; KeepOnly: call ApplyKeepOnlyFilter
- **`ApplyKeepOnlyFilter`** (Private) — sort tblNorm, FindInRange first/last match, bulk-delete above and below; re-provision tblNorm

JDL 5/7/26