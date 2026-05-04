## Purpose
Create and validate ColInfo class for storing and utilizing metadata about project data objects (tblRowsCols and mdlScenario).

## Background

This change follows [[VBA Project Changes]] sequencing and is scoped to planning-level architecture before coding.

Project code instances `ColInfo` as `colinfo` and uses its attributes for actions such as importing and normalizing raw data. `ColInfo` is decoupled from any specific project — it references the project global `IsTest` (always declared `Public Boolean` in project workbook, default `False`) directly rather than receiving it as an argument.

ColInfo is based on a project-specific metadata file `ColInfo.xlsx` (see [[Example_ColInfo]]) whose location is specified by `files.pfColInfo` within a `files` instance of `ProjFiles`. The file is always required as a separate xlsx — no sheet fallback.

Within `colinfo.tbl` (a `tblRowsCols` instance), inclusion of a variable in a project table is denoted by an integer in the column named for that table (e.g. `BR_Example`, `Second_Tbl`). The integer specifies column order in the normalized output. Renaming for normalization is defined by the `VarNameNorm` → `VarNameRaw` mapping. Index/key columns are identified by `True` in the `IsIndex` column.

Because `colinfo.tbl` can contain metadata for multiple project tables, usage for a particular table sets `.CurTbl` and sorts the tbl so that rows included in that table are contiguous at the top. `SetCurTbl` is stateless re-entrant — each call re-sorts and resets `.rngRowsCurTbl`.

A companion `TblImport` class (future change) will handle raw data import, parsing, and normalization using `colinfo.tbl` metadata plus `dImportParams` / `dParseParams` dictionaries. The `data_type_VBA` column in ColInfo.xlsx is present for that future use; no yield method for it in this change.

`FindNextChange` utility (for locating value-change boundaries in sorted columns) is deferred to a future change. `rngToExtent` is sufficient here since post-sort non-blank rows are contiguous at top.

## Data I/O Descriptions

**Inputs to `Init`:**
- `colinfo As ColInfo` — self-reference (standard convention)
- `files As Object` — `ProjFiles` instance; `files.pfColInfo` is required (non-empty)
- `Optional curTbl As String` — if passed, calls `SetCurTbl(colinfo, curTbl)` after provisioning

**Implicit inputs used by `Init`:**
- `IsTest` — project global Boolean; consumed by `files.pfColInfo` resolution (already set on `files`)

**Outputs — attributes set after `Init`:**

| Attribute | Value |
|---|---|
| `tbl` | `tblRowsCols` instance provisioned on `"colinfo_"` sheet of opened `ColInfo.xlsx` |
| `CurTbl` | Empty string until `SetCurTbl` called |
| `rngRowsCurTbl` | `Nothing` until `SetCurTbl` called |

**Outputs — attributes set after `SetCurTbl(colinfo, sTbl)`:**

| Attribute | Value |
|---|---|
| `CurTbl` | `sTbl` |
| `rngRowsCurTbl` | Row range of `colinfo.tbl.rngRows` where `CurTbl` column is non-blank; derived via `rngToExtent` after sort |

## Project Architecture

**New class: `ColInfo.cls`**
- Instanced as `colinfo` in project code (e.g. `Set colinfo = New ColInfo`)
- All attributes public; no `Property Get/Let`

**Public attributes:**
- `tbl As Object` — `tblRowsCols` instance for ColInfo.xlsx
- `CurTbl As String` — name of current project table column
- `rngRowsCurTbl As Range` — row range for current table's included rows

**Methods:**
- `Public Function Init(colinfo, files As Object, Optional curTbl As String) As Boolean`
  - Opens `ColInfo.xlsx` via `ExcelSteps.OpenFile(files.pfColInfo, wkbkColInfo)`; fails fast if `files.pfColInfo` is empty
  - Provisions `colinfo.tbl` via `tbl.Init(tbl, wkbkColInfo, "colinfo_")`; no Refresh needed
  - If `curTbl` passed, calls `SetCurTbl(colinfo, curTbl)`
- `Public Function SetCurTbl(colinfo, sTbl As String) As Boolean`
  - Sets `.CurTbl = sTbl`
  - Sorts `colinfo.tbl` ascending by `sTbl` column via `TblSortBy(colinfo.tbl, sTbl)` (blanks sink to bottom)
  - Sets `.rngRowsCurTbl` using `rngToExtent` from first data cell of `sTbl` column in `colinfo.tbl.rngRows`
- `Public Function YieldAryIndices(colinfo) As Variant`
  - Fails fast if `.CurTbl` is empty
  - Iterates `.rngRowsCurTbl`; collects `VarNameNorm` values where `IsIndex = True`
  - Returns `Variant` array (caller assigns to local `Variant` before iterating)
- `Public Function YieldAryMetrics(colinfo) As Variant`
  - Fails fast if `.CurTbl` is empty
  - Iterates `.rngRowsCurTbl`; collects `VarNameNorm` values where `IsIndex` is not `True`
  - Returns `Variant` array
- `Public Function YieldDNormalize(colinfo) As Object`
  - Fails fast if `.CurTbl` is empty
  - Iterates `.rngRowsCurTbl`; adds `VarNameNorm` → `VarNameRaw` pairs to `Dictionary` instance
  - Fails fast (`errs.IsFail`) if any row has blank `VarNameRaw`
  - Returns `Dictionary` instance

**Factory function:**
- `New_ColInfo` in `Validation.bas` — enables cross-workbook instantiation from test suite

## Test Architecture

Tests housed in existing `tests_ToolboxClasses.bas` (shared with ProjFiles tests). New `procs.ColInfo` procedure group added to `Procedures.cls`. Separate `TestingDriver_ColInfo` driver sub in `tests_ToolboxClasses.bas`.

- **Procedure wiring**: Add `Public ColInfo As Object` to `Procedures.cls` declarations; add `Set .ColInfo = New Procedure` + `.ColInfo.Name = "ColInfo"` in `Procedures.Init` under `' tests_ToolboxClasses` group comment. Follow [[vba-testing-create-new-test-procedure]] skill for all wiring steps.
- **Driver**: `TestingDriver_ColInfo` in `tests_ToolboxClasses.bas`; `procs.Init` args: `ThisWorkbook`, `"ColInfo"`, `"Tests_ToolboxClasses"`; enable `procs.ColInfo.Enabled = True`
- **Cross-workbook instantiation**: Add `New_ColInfo` factory function to `Validation.bas`
- **Test data**: `tests/test_data/ColInfo.xlsx` mockup matching [[Example_ColInfo]]; accessed via `files.pfColInfo` with `IsTest = True` and `subdir_tests = "test_data"`
- **Test setup helper**: shared sub `InitColInfoTest(tst, colinfo, files)` — instances `files` and `colinfo`, sets `IsTest = True`, calls `files.Init` then `colinfo.Init`
- **Coverage**: `Init` (with and without `curTbl`); `SetCurTbl` (sort, rngRowsCurTbl extent); `YieldAryIndices` / `YieldAryMetrics` (correct split, correct order); `YieldDNormalize` (correct mapping, blank `VarNameRaw` error); `SetCurTbl` re-entrant (call twice with different table names)
- **Key edge cases**: `curTbl` column all blank after sort; blank `VarNameRaw` triggers fail-fast; `CurTbl` empty when yield called

## Procedure Outline

**ColInfo Class**
- **`Init`** — fail fast if `files.pfColInfo` empty; open `ColInfo.xlsx`; provision `colinfo.tbl` on `"colinfo_"`; if `curTbl` passed call `SetCurTbl`
- **`SetCurTbl`** — set `.CurTbl`; call `TblSortBy(colinfo.tbl, sTbl)`; set `.rngRowsCurTbl` via `rngToExtent` from first data cell of `sTbl` column
- **`YieldAryIndices`** — fail fast if `.CurTbl` empty; ReDim local Variant; iterate `rngRowsCurTbl`; collect `VarNameNorm` where `IsIndex=True`; return array
- **`YieldAryMetrics`** — same as `YieldAryIndices` for non-index rows
- **`YieldDNormalize`** — fail fast if `.CurTbl` empty; instance Dictionary; iterate `rngRowsCurTbl`; fail fast on blank `VarNameRaw`; add `VarNameNorm`→`VarNameRaw`; return dict