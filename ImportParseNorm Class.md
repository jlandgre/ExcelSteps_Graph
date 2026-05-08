Version 5/8/26 J.D. Landgrebe/Data Delve LLC

`ImportParseNorm` drives configuration-based import, parsing, and normalization of raw data files. It is instanced as `imptbl` in project code. All configuration is supplied at call time through `Init` arguments — no attributes need to be set before calling `Init`.

The class reads column metadata from a `ColInfo` instance (`colinfo`) and two parameter dictionaries (`dParamsImport`, `dParamsParse`) to control how the raw file is opened, validated, filled, and normalized. This replaces hard-coded per-table import procedures with a single reusable driver.

## Key Design Decisions

- **All inputs passed to `Init`** — caller supplies `colinfo`, `files`, `curTbl`, `dParamsImport`, `dParamsParse`; no attributes are pre-set
- **`colinfo` is lazy-initialized** — `Init` checks `colinfo Is Nothing` and `colinfo.tbl Is Nothing`; initializes fresh if needed, then always calls `colinfo.SetCurTbl(colinfo, curTbl)` to configure for the current table
- **Raw workbook is never mutated in place** — `OpenRawData` copies sheet 1 to a new temp workbook, closes the original, then provisions `tblRaw` against the copy
- **`tblNorm` lives in the same temp workbook** as `tblRaw`, on a new `"norm_"` sheet created by `WriteNormFromMappings`
- **Column copy, not row iteration** — `WriteNormFromMappings` copies entire data columns from `tblParsed` to `tblNorm` using range value assignment; order is driven by integer order values in the `curTbl` column of `colinfo.tbl`
- **`ReplaceVals` uses sort + block ops** — for each `FillVals` entry, sorts `tblRaw` by the target column so equal values are contiguous, then applies replacements in bulk via `SpecialCells` (blanks) or `Find`/`xlPrevious` range (other keys)
- **`FilterRows` (`KeepOnly`) sorts then bulk-deletes** — sorts `tblNorm` by filter column; locates first/last matching rows; deletes non-matching rows above and below in two `.Delete` calls; re-provisions `tblNorm` after deletion

## Public Attributes

```vb
Public colinfo As Object          ' ColInfo instance
Public files As Object            ' ProjFiles instance
Public curTbl As String           ' Current table name
Public dParamsImport As Object    ' Import params dictionary  e.g. {FileType:xlsx}
Public dParamsParse As Object     ' Parse params dictionary   e.g. {RawShape:rowscols}
Public tblRaw As Object           ' tblRowsCols (or tblUnstructured) for copied raw sheet
Public tblParsed As Object        ' Points to tblRaw when RawShape=rowscols
Public tblNorm As Object          ' tblRowsCols for normalized output sheet
```

## Typical Usage

```vb
Dim imptbl As Object, colinfo As Object, files As Object
Dim dImport As Object, dParse As Object

Set imptbl = New ImportParseNorm        ' or New_ImportParseNorm from test suite
Set files = New ProjFiles
If Not files.Init(files, wkbk) Then GoTo ErrorExit

Set dImport = New Dictionary
If Not dImport.ParseStringToDictProcedure("{FileType:xlsx}") Then GoTo ErrorExit

Set dParse = New Dictionary
If Not dParse.ParseStringToDictProcedure("{RawShape:rowscols}") Then GoTo ErrorExit

If Not imptbl.ImportParseNormProcedure(imptbl, colinfo, files, "BR_Example", _
  dImport, dParse) Then GoTo ErrorExit

' imptbl.tblNorm is now provisioned and ready for further processing
```

## Method Summary

| Method | Visibility | Description |
|---|---|---|
| `ImportParseNormProcedure` | Public | Orchestrator: calls Init, OpenAndValidateRawProcedure, ParseRawProcedure, NormalizeParsedProcedure |
| `Init` | Public | Validates args; lazy-inits colinfo; sets all attributes |
| `OpenAndValidateRawProcedure` | Public | Calls OpenRawData, ValidateRawStructure, ReplaceVals |
| `OpenRawData` | Public | Opens raw file; copies to temp workbook; provisions tblRaw |
| `ValidateRawStructure` | Public | Confirms all VarNameRaw columns exist in tblRaw header |
| `ReplaceVals` | Public | Applies FillVals key/value replacements to tblRaw columns |
| `ApplyFillMapToSortedColumn` | Public | Iterates dict keys; dispatches to ApplyFillKeyToSortedColumn |
| `ApplyFillKeyToSortedColumn` | Public | Fills blanks (SpecialCells) or replaces keyed values (Find range) |
| `ParseRawProcedure` | Public | Sets tblParsed=tblRaw for rowscols; errors for unimplemented shapes |
| `NormalizeParsedProcedure` | Public | Calls WriteNormalized then FilterRows |
| `WriteNormalized` | Public | Validates tblParsed; delegates to BuildNormMappings + WriteNormFromMappings |
| `BuildNormMappings` | Public | Two-pass loop; builds aryNorm/aryRaw/maxOrder from colinfo order column |
| `WriteNormFromMappings` | Public | Creates norm_ sheet; writes header; copies columns; provisions tblNorm |
| `FilterRows` | Public | Parses FilterVals dict; dispatches KeepOnly to ApplyKeepOnlyFilter |
| `ApplyKeepOnlyFilter` | Private | Sort + locate first/last match + bulk-delete above/below + re-provision |

## ColInfo Metadata Columns Used

| Column | Used by | Purpose |
|---|---|---|
| `VarNameNorm` | ValidateRawStructure, BuildNormMappings, FilterRows | Normalized output variable name |
| `VarNameRaw` | ValidateRawStructure, ReplaceVals, BuildNormMappings | Source column name in raw file |
| `FillVals` | ReplaceVals | JSON-like replacement map e.g. `{BLANK:Unknown,Locn10:Locn1}` |
| `FilterVals` | FilterRows | JSON-like filter spec e.g. `{KeepOnly:Online}` |
| `curTbl` (column named by curTbl arg) | BuildNormMappings | Integer output order for each variable |

## Parameter Dictionaries

**`dParamsImport`** — controls file opening behavior:

| Key | Values | Default |
|---|---|---|
| `FileType` | `xlsx` | (unused in current implementation) |

**`dParamsParse`** — controls raw data shape:

| Key | Values | Default |
|---|---|---|
| `RawShape` | `rowscols` | `rowscols` |

`RawShape=rowscols` provisions `tblRaw` as a `tblRowsCols` with header in row 1. Other shapes (e.g. column-oriented) provision `tblRaw` as a `tblUnstructured` but `ParseRawProcedure` currently errors for non-rowscols — deferred to future implementation.

## FillVals Format

`FillVals` entries in `colinfo.tbl` use the same JSON-like dictionary string format as the `Dictionary` class:

```
{BLANK:Unknown,Locn10:Locn1,OldCode:NewCode}
```

- Keys are unquoted or quoted; values are unquoted or quoted
- `BLANK` (case-insensitive) targets blank cells using `SpecialCells`
- All other keys are matched with `xlWhole` exact match; replacements are applied as a bulk range assignment to the contiguous block of matching rows (requires column to be sorted first — `ReplaceVals` sorts before calling `ApplyFillMapToSortedColumn`)

## FilterVals Format

```
{KeepOnly:Online}
```

- `KeepOnly` retains only rows where the named `VarNameNorm` column equals the specified value; all other rows are deleted
- Other filter types (`KeepList`, `KeepExcept`, `KeepExceptList`) are deferred to future implementation

## ByRef Alias Pattern

Several methods require a local alias to work around VBA's restriction on passing class attributes directly as `ByRef` arguments:

```vb
' Provision tblRaw (tblRowsCols case)
Set imptbl.tblRaw = New tblRowsCols
Set tblAlias = imptbl.tblRaw
If Not tblAlias.Provision(tblAlias, wkbkTemp, False, sht:=shtRaw, IsSetColRngs:=False) Then GoTo ErrorExit

' Provision tblNorm after WriteNormFromMappings
Set imptbl.tblNorm = New tblRowsCols
Set tblAlias = imptbl.tblNorm
If Not tblAlias.Provision(tblAlias, imptbl.tblParsed.wkbk, False, sht:=normSht, IsSetColRngs:=False) Then GoTo ErrorExit
```

The alias points to the same object as the attribute; it is created solely to satisfy the ByRef call constraint, not as a proxy.

## Factory Function

`New_ImportParseNorm` in `Validation.bas` enables cross-workbook instantiation from the test suite:

```vb
' In test code (separate workbook)
Dim imptbl As Object
Set imptbl = ExcelSteps.New_ImportParseNorm
```

Within the ExcelSteps project itself, use `New ImportParseNorm` directly.
