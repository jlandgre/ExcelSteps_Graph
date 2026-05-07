## Purpose
Create and validate `ImportParseNorm.cls` for importing, parsing, and normalizing raw data. Instanced as `importtbl` in project code.

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
| `importtbl` | `ImportParseNorm` | self-reference |
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
- Instanced as `importtbl`
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
- `Public Function Init(importtbl, colinfo, files As Object, curTbl As String, dParamsImport As Object, dParamsParse As Object) As Boolean`
  - Sets all args as public attributes
  - Checks `colinfo Is Nothing` or `colinfo.tbl Is Nothing`; if so calls `colinfo.Init(colinfo, files)`
  - Always calls `colinfo.SetCurTbl(colinfo, curTbl)`
- `Public Function ImportParseNormProcedure(importtbl, colinfo, files As Object, curTbl As String, dParamsImport As Object, dParamsParse As Object) As Boolean`
  - Calls `Init`, then `OpenAndValidateRawProcedure`, `ParseRawProcedure`, `NormalizeParsedProcedure` in sequence
- `Public Function OpenAndValidateRawProcedure(importtbl) As Boolean`
  - Calls `OpenRawData`, `ValidateRawStructure`, `FillMissingVals`
- `Public Function OpenRawData(importtbl) As Boolean`
  - Opens `files.pfImportFile` via `ExcelSteps.OpenFile`
  - Copies sheet 1 to a new workbook; closes original
  - Provisions `tblRaw` against new workbook's sheet
- `Public Function ValidateRawStructure(importtbl) As Boolean`
  - Iterates `colinfo.rngRowsCurTbl`; for each row where `VarNameNorm` non-blank, confirms `VarNameRaw` exists in `tblRaw.rngHeader` via `rngTblHeaderVal`; fails fast on first missing column
  - Future: `dParamsParse` can relax this requirement
- `Public Function FillMissingVals(importtbl) As Boolean`
  - For each row in `colinfo.rngRowsCurTbl` with a non-blank `FillVals` entry, parses JSON-like string into a dictionary; applies value replacements and blank-fills in-place on `tblRaw`
- `Public Function ParseRawProcedure(importtbl) As Boolean`
  - Checks `dParamsParse` `RawShape` key; if `"rowscols"` sets `tblParsed = tblRaw` and exits
  - Non-rowscols shapes deferred to future change
- `Public Function NormalizeParsedProcedure(importtbl) As Boolean`
  - Calls `WriteNormalized`, then `FilterRows`
- `Public Function WriteNormalized(importtbl) As Boolean`
  - Builds sorted list of `(orderInt, VarNameNorm, VarNameRaw)` from `colinfo.rngRowsCurTbl` ordered by `CurTbl` integer; skips rows where integer is blank or non-numeric
  - Creates new `"norm_"` sheet in `tblRaw.wkbk`; writes `VarNameNorm` values as header in order
  - For each entry, copies entire data column from `tblRaw` (located by `VarNameRaw`) to corresponding `tblNorm` column
  - Provisions `tblNorm` against new sheet
- `Public Function FilterRows(importtbl) As Boolean`
  - Iterates `colinfo.rngRowsCurTbl`; for each row with non-blank `FilterVals`, parses JSON-like string
  - Implements `KeepOnly` only: sorts `tblNorm` by filter column, uses `FindInRange` to locate first and last matching row, bulk-deletes rows above and below in two `.Delete` calls
  - Other filter options (`KeepList`, `KeepExcept`, `KeepExceptList`) deferred to future

**Factory function:**
- `New_ImportParseNorm` in `Validation.bas` — enables cross-workbook instantiation from test suite

## Test Architecture

Tests housed in `tests_ToolboxClasses.bas`. New `procs.ImportParseNorm` procedure group added to `Procedures.cls`. Follow [[vba-testing-create-new-test-procedure]] skill for all wiring steps.

- **Test section location**: immediately below `TestDriver_ToolBox` for ease of navigation (not at bottom of file)
- **Test data**: two static raw data files in `tests/test_data/`:
  - `BR_Raw_Mockup.csv` — `BR_Example` raw columns; includes mixed Location values (Online + others), blank `ProdType_Raw` values, and `Locn10` values to exercise `FillMissingVals` and `FilterRows`
  - `Second_Raw_Mockup.csv` — `Second_Tbl` raw columns; no `FillVals`/`FilterVals` (simpler validation case)
- **`files.pfImportFile`**: set by `ProjFiles.Init` when `IsTest = True` pointing to test data CSV/xlsx
- **Test helper**: `InitImportParseNormTest(tst, importtbl, colinfo, files, curTbl, dParamsImport, dParamsParse)` — shared setup sub instances all objects and calls `importtbl.Init`

**Test list:**
- `test_ImportParseNorm_Init` — attributes set; colinfo initialized; SetCurTbl called
- `test_ImportParseNorm_OpenRawData` — `tblRaw` provisioned; original workbook closed
- `test_ImportParseNorm_ValidateRawStructure` — passes for valid raw; fails fast on missing VarNameRaw column
- `test_ImportParseNorm_FillMissingVals` — blanks filled; value replacements applied in-place
- `test_ImportParseNorm_ParseRaw` — `tblParsed Is tblRaw` for `RawShape=rowscols`
- `test_ImportParseNorm_WriteNormalized` — `tblNorm` header matches ordered VarNameNorm; data present
- `test_ImportParseNorm_FilterRows` — `KeepOnly:Online` reduces rows to only Online location rows
- `test_ImportParseNorm_OpenAndValidateRawProcedure` — integration: open + validate + fill
- `test_ImportParseNorm_NormalizeParsedProcedure` — integration: write + filter
- `test_ImportParseNorm_Procedure` — end-to-end: correct tblNorm shape, header order, row count

## Procedure Outline

**ImportParseNorm Class**
- **`Init`** — set all args as attributes; check/init colinfo; call SetCurTbl
- **`ImportParseNormProcedure`** — call Init, OpenAndValidateRawProcedure, ParseRawProcedure, NormalizeParsedProcedure
- **`OpenAndValidateRawProcedure`** — call OpenRawData, ValidateRawStructure, FillMissingVals
- **`OpenRawData`** — open pfImportFile; copy sheet to new workbook; close original; provision tblRaw
- **`ValidateRawStructure`** — iterate rngRowsCurTbl; for non-blank VarNameNorm rows confirm VarNameRaw in tblRaw header; fail fast on missing
- **`FillMissingVals`** — iterate rngRowsCurTbl FillVals column; parse JSON-like string; apply replacements in-place on tblRaw
- **`ParseRawProcedure`** — if RawShape=rowscols set tblParsed=tblRaw; else defer (dummy)
- **`NormalizeParsedProcedure`** — call WriteNormalized, FilterRows
- **`WriteNormalized`** — build sorted (orderInt, VarNameNorm, VarNameRaw) list; create norm_ sheet; write header; copy columns from tblParsed by VarNameRaw; provision tblNorm
- **`FilterRows`** — for each FilterVals row parse dict; KeepOnly: sort tblNorm, FindInRange first/last match, bulk-delete above and below

JDL 5/7/26