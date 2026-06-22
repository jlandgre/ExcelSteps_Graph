# Change: Add Params Column to ExcelSteps Recipe

Status: Draft for ProjectOwner approval  
ProjectOwner: JDL  
AI: GitHub Copilot  
Date: 6/22/26

## Purpose

Add a new `Params` column to the ExcelSteps recipe sheet so recipe rows can specify optional JSON-like parameters without adding a dedicated worksheet column for every new action. The first implementation is limited to `tblRowsCols` refresh and supports horizontal alignment through `{HorizAlign:xlHAlignCenter}`-style params on `Col_Format` and `Col_Insert` rows.

This change creates a narrow working pattern for future params-driven recipe actions while keeping existing recipe columns and existing recipes compatible.

## Background

The ExcelSteps recipe is stored on the `ExcelSteps` sheet. See [[Example_ExcelSteps]]. Each row describes a recipe instruction for refreshing a `tblRowsCols` table or, in other parts of ExcelSteps, an `mdlScenario` model. The recipe format was originally developed for `tblRowsCols`, so the current headers are table-oriented:

- Column A `Sheet`: table/model name or sheet location for default objects.
- Column B `Column`: table column name or Scenario Model variable name.
- Column C `Step`: recipe action to run.
- Columns D:I: fixed parameter columns for existing action types.

Current fixed parameter columns work for established actions such as number format, width, comment, formula, and insert location, but they do not scale well for additional formatting or behavior. `Params` provides a structured extension point using the existing cross-platform `Dictionary.cls` parser documented in [[procPlan_ParseStringToDictProcedure]].

Relevant implementation files:

- `src/Constants.bas`
- `src/Refresh.cls`
- `src/RecipeStep.cls`
- `src/tblRowsCols.cls`
- `tests/tests_tblRowsCols.bas`

## Data I/O Descriptions

### ExcelSteps Recipe Input

Add `Params` as the last recipe header, after `Width`:

`Sheet,Column,Step,Formula/List Name/Sort-by,After End or Rename Column,Keep Formulas,Comment,Number Format,Width,Params`

`Params` is placed at the end because JSON-like strings may be long and should not create an excessively wide interstitial column between frequently edited recipe fields.

### Params String Format

`Params` values use the existing JSON-like parser in `Dictionary.ParseStringToDictProcedure`.

Example:

```text
{HorizAlign:xlHAlignCenter}
```

The parser supports unquoted keys and unquoted string values, so enum names can be written directly. The first implementation does not introduce new parser behavior.

### Supported Step Types

Parse and apply `Params` only when `RecipeStep.sType` is:

- `Col_Format`
- `Col_Insert`

For all other step types, ignore the `Params` cell completely.

### Supported Params Keys

Initial supported key:

- `HorizAlign`: resolves an Excel horizontal alignment enum name and applies it to table data rows.

Initial supported values:

- `xlHAlignCenter`
- `xlHAlignLeft`
- `xlHAlignRight`
- `xlHAlignGeneral`

Unsupported keys and unsupported enum names should fail through the existing recipe-row error flow that places an error message comment on the offending row.

### Output/Effect

For `HorizAlign`, apply alignment to the data cells for the target table variable:

```vb
Intersect(targetColumn.EntireColumn, tbl.rngRows).HorizontalAlignment = Step.HorizAlign
```

The header cell is not affected. This matches the scoping of number format behavior, which applies to data values rather than table headers.

## Project Architecture

### Constants.bas

Update `sHeaderSteps` to append `Params` at the end of the recipe header string.

Use `sHeaderSteps` as the single source of truth for the ExcelSteps recipe header. `Refresh.SetStepsShtHeader` currently hardcodes the same header string locally; replace that local string with `sHeaderSteps`.

Add a new action constant for the params-derived action:

```vb
Public Const sAHorizAlign As String = "Col_HorizAlign"
```

This constant is an internal action token added to `Step.Actions`; it does not need to appear in the ExcelSteps step dropdown list.

### Refresh.cls

`PrepExcelStepsSht` should support two header paths:

- Full setup/reformat path for missing or invalid ExcelSteps sheets.
- Legacy upgrade path for existing valid 9-column recipe sheets that only lack `Params`.

For existing 9-column sheets, the only upgrade action should be to add the new `Params` header at the end with blank values below it. Do not rebuild, clear, or broadly reformat the existing recipe sheet.

Formatting changes:

- `SetStepsShtHeader` splits `sHeaderSteps` instead of a local hardcoded string.
- `FormatStepsShtWidths` adds a width entry for `Params`, using width `32`.
- `FormatStepsShtColumns` wraps the `Params` column.
- The legacy upgrade path should apply equivalent width/wrap formatting to the appended `Params` column.

### tblRowsCols.cls

Add a public recipe column range attribute:

```vb
Public colrngParams As Range
```

Update `SetColRanges` for `TblName = shtSteps` to set:

```vb
Set .colrngParams = .cellHome.Offset(0, 9).EntireColumn
```

The existing offsets for columns A:I should remain unchanged.

### RecipeStep.cls

Add new attributes:

```vb
Public paramsJSON As String
Public paramsDict As Object
Public HorizAlignVal As Variant
```

`Read` should populate the existing attributes first, then parse params only when `sType` is `Col_Format` or `Col_Insert` and the `Params` cell is nonblank.

Recommended helper structure:

- `Params`: read `tblSteps.colrngParams`, initialize `paramsDict`, parse with `Dictionary.ParseStringToDictProcedure`, validate keys, and resolve supported values into dedicated attributes such as `Step.HorizAlignVal`.
- `ParamKeys`: fail if any parsed key is not supported for the current implementation.
- `ParamActions`: add params-derived action tokens such as `sAHorizAlign` to `Step.Actions`.
- `HorizAlign`: apply `Step.HorizAlignVal` to the target data rows when called by `RunActions`.

`RunActions` should dispatch `sAHorizAlign` the same way it dispatches existing action tokens.

Action helper error handling should match existing recipe action helpers such as `SetNumFmt` and `SetWidth`: use the local `SetErrs` pattern and return `False` from `ErrorExit` without calling `errs.RecordErr`, allowing `RunActions` to place the recipe-row comment and route through the existing comment/error flow.

## Test Architecture

Tests should stay in `tests/tests_tblRowsCols.bas`, under the existing `procs.tblRefresh` group in `TestDriver_TblRowsCols`.

Existing recipe mockup/helper:

- `PopulateStepsTblRefresh` in `tests/Populate_Tbl.bas` should be reused for the success tests. It already creates the two needed recipe rows: row 2 is `Col_Format` for `Data_2`, and row 3 is `Col_Insert` for `Data_4` after `Data_3`.
- `test_HorizAlignFormat` can call `PopulateStepsTblRefresh`, then write `{HorizAlign:xlHAlignCenter}` to the row 2 `Params` cell.
- `test_HorizAlignInsert` can call `PopulateStepsTblRefresh`, then write `{HorizAlign:xlHAlignCenter}` to the row 3 `Params` cell.
- `test_UnsupportedParamKey` and `test_UnsupportedHorizAlign` can use the same helper and change only the row 2 or row 3 `Params` value for the failure case.
- `test_AddParamsColumn` should simulate a legacy 9-column ExcelSteps header directly, because `PopulateStepsTblRefresh` assumes an already prepared recipe sheet and is aimed at refresh-row behavior rather than header-upgrade behavior.

Update existing test:

- `test_PrepExcelStepsSht`: expect `tblSteps.nCols = 10` and verify the `Params` header exists at the end.

Add focused tests:

- `test_AddParamsColumn`: create or simulate a valid 9-column ExcelSteps sheet, call `PrepExcelStepsSht`, and assert that only the blank `Params` column is appended.
- `test_HorizAlignFormat`: add `{HorizAlign:xlHAlignCenter}` to a `Col_Format` row and assert `HorizontalAlignment = xlHAlignCenter` on `Intersect(targetColumn.EntireColumn, tbl.rngRows)` after refresh.
- `test_HorizAlignInsert`: add `{HorizAlign:xlHAlignCenter}` to a `Col_Insert` row and assert the inserted column data rows are centered after refresh.
- `test_UnsupportedParamKey`: use a valid params string with an unsupported key and assert refresh fails through the recipe-row error/comment path.
- `test_UnsupportedHorizAlign`: use `HorizAlign` with an unsupported enum name and assert refresh fails through the recipe-row error/comment path.

Do not add malformed JSON-like syntax tests here. Parser syntax validation belongs to the existing Dictionary tests.

## Discussion: Scope

This change is `tblRowsCols` first. `mdlScenario` remains a future-compatible design target but is not part of this implementation. The current concrete flow is `Refresh.RefreshRC` -> `RefreshTblFromRecipe` -> `RecipeStep.Read` -> `RecipeStep.RunActions`, which is the table refresh path. Extending params to `mdlScenario` should be planned separately because the model recipe paths and action surfaces differ.

## Discussion: Failure Policy

For supported step types, `Params` is treated as executable recipe instruction. Therefore:

- Malformed params strings fail through the parser and existing recipe-row error flow.
- Unsupported params keys fail.
- Unsupported enum names fail.

For unsupported step types, `Params` is ignored completely.

This keeps accidental notes or future-looking params harmless on unsupported rows while ensuring typos are caught when params are active.

## Testing Considerations

- Existing recipes without `Params` should continue to refresh after the compatibility path appends the blank `Params` column.
- Existing tests and helper routines that assume 9 recipe columns must be updated to 10 columns where appropriate.
- Tests should assert data-row alignment, not header alignment.
- Failure tests should confirm the existing recipe-row comment/error path is used, not just that a function returns `False`.
- No external test data files are required.
- Cross-workbook object creation in tests should continue to use existing factory functions such as `ExcelSteps.New_tbl`, `ExcelSteps.New_Refresh`, and `ExcelSteps.New_Dictionary` as needed.

## Procedure Outline

### Refresh.PrepExcelStepsSht

Prepare the ExcelSteps recipe sheet, including full initialization for missing/invalid sheets and a narrow compatibility upgrade for existing sheets missing `Params`.

Flow:

1. If the ExcelSteps sheet does not exist, add it and run the full setup path.
2. If the sheet exists but does not have the expected first header value, run the full setup path.
3. If the sheet exists and has the legacy 9-column header without `Params`, append only the `Params` column and apply its basic formatting.
4. Provision `tblSteps` with `IsSetColRngs:=True` so `colrngParams` is available.

### Refresh.SetStepsShtHeader

Write the ExcelSteps recipe header from `Constants.sHeaderSteps` and apply header formatting.

### Refresh.FormatStepsShtWidths

Apply recipe sheet widths, including width `32` for the final `Params` column.

### Refresh.FormatStepsShtColumns

Apply text/wrap formatting to special columns, including wrapping `Params`.

### Refresh.AddParamsColumn

Add the `Params` header at the end of existing valid legacy recipe sheets, leave existing recipe values intact, and format the appended column.

### RecipeStep.Read

Read standard recipe row attributes, create the actions collection, then read and resolve params for supported step types.

### RecipeStep.Params

Read the row's `Params` string, parse it into `paramsDict`, validate keys, resolve values into dedicated `RecipeStep` attributes, and add params-derived actions.

### RecipeStep.ParamKeys

Validate that every key in `paramsDict` is supported for the current implementation.

### RecipeStep.ParamActions

Add params-derived action constants to `Step.Actions`.

### RecipeStep.HorizAlign

Apply the resolved horizontal alignment to the target table column's data rows.

### RecipeStep.RunActions

Dispatch `sAHorizAlign` through the same action loop used for existing recipe actions.

## Approval Gate

ProjectOwner approval of this Change_ note is required before implementation. After approval, implementation should update the project code and tests together, then export the VBA project files from `XLSteps.xlam` and `tests_XLSteps.xlsm`.