---
status: validated
updated: 06-24-2026
type: code_change
description: Add ExcelSteps recipe support for indexed Scenario Model variable blocks
tags:
  - VBA
  - ExcelSteps
  - Scenario_Model
created: 06-24-2026
author: JDL
---

# Change: Lite Scenario Model Variable Blocks

---

## Purpose
Add ExcelSteps recipe support for indexed Scenario Model variable blocks so a Lite model can define one recipe row for a contiguous block of variables whose names share a stem and differ only by numeric suffix.

New recipe instructions:
- `Input_VarBlock` — refreshes indexed input variable rows and their number formats only.
- `Formula_VarBlock` — refreshes indexed calculated variable rows, formulas, number formats, and calculation-cell formatting.

## Background
**Referenced docs and code:**
- [[Example_Mdl_BlockVars]] — mock Lite model examples with placeholders and pre-existing indexed blocks
- [[Example_Mdl_BlockVars_Recipe]] — ExcelSteps recipe mockup for block variables
- [src/mdlScenario.cls](../src/mdlScenario.cls) — `Provision`, `PrepStepsForMdl`, `SetRngFormulaRows`, and `Refresh`
- [src/mdlRow.cls](../src/mdlRow.cls) — single-row model refresh attributes and formula/format methods
- [src/RecipeStep.cls](../src/RecipeStep.cls) — existing params parsing pattern for table recipe rows
- [src/Dictionary.cls](../src/Dictionary.cls) — `ParseStringToDictProcedure` for JSON-like params strings
- [tests/tests_mdlScenario.bas](../tests/tests_mdlScenario.bas) — current Scenario Model and Lite model tests
- [tests/Populate_Mdl.bas](../tests/Populate_Mdl.bas) — current Lite model and ExcelSteps recipe population helpers

**Change-specific framing:**
Lite Scenario Models currently match ExcelSteps recipe rows to model variables by exact variable name in `mdlScenario.SetRngFormulaRows`. That works for ordinary variables but forces every indexed block variable to be listed separately in the recipe. This change lets the recipe list the base placeholder variable once, for example `inputvar_xx`, and use params `{nBlockVars:3}` to expand or maintain `inputvar_1`, `inputvar_2`, and `inputvar_3` in the model.

The capability applies only to non-default Lite Scenario Models, because default models store formula and format instructions in template columns instead of relying on ExcelSteps recipe rows.

## Locked Decisions

1. Implement block-variable orchestration initially by extending `mdlRow.cls` and `mdlScenario.cls`, not by creating a new `MdlBlockVar` class.
2. `mdlScenario` owns model-level discovery, row shaping, and block-aware refresh branching because it owns the model ranges and refresh lifecycle.
3. `mdlRow` owns block-row application after rows exist, including indexed variable/formula substitution and calls to existing row methods.
4. The existing non-block `mdlRow` refresh workflow remains the `Else` path in the block-aware `mdlScenario.Refresh` loop.
5. Revisit a separate `MdlBlockVar` class only if implementation makes `mdlRow` too broad or hard to reason about.
6. Recipe params key is `nBlockVars` and is case-sensitive, matching the requested JSON-like string `{nBlockVars:n}`.
7. `nBlockVars` is required for both `Input_VarBlock` and `Formula_VarBlock` and must be a positive integer.
8. Block recipe variable names must end in `_xx`; the stem is everything before `_xx`.
9. Indexed model variable names are generated as `stem_1` through `stem_n`.
10. Block variables must be contiguous in the model; gaps or interposed variables hard-fail.
11. If the model contains the base `_xx` placeholder, refresh replaces that placeholder row with `nBlockVars` indexed rows.
12. If the model already contains exactly `nBlockVars` indexed rows, refresh preserves those rows and only refreshes formats/formulas.
13. If the model contains more than `nBlockVars` indexed rows, refresh deletes overage rows from bottom up. This works for both instruction types and preserves existing input data in retained `Input_VarBlock` rows.
14. If the model contains fewer than `nBlockVars` indexed rows, refresh deletes the existing block and recreates it from the recipe definition.
15. `Formula_VarBlock` formulas may refer to other block variables using `_xx`; refresh substitutes the current index into every `_xx` token for each row, including structured references such as `@InputVar_xx`.
16. `Input_VarBlock` does not write formulas or apply calculation style; it applies input-row formatting and number format only. Any text in `Formula/List Name/Sort-by` is ignored, matching analogous `Col_Format` behavior.
17. Add a reset utility that deletes existing indexed block rows for a specified Lite mdl and replaces each block with its base `_xx` placeholder row.
18. Add `ResetAllBlockVarsDriver` plus a Boolean API/function for explicit calls and tests, but do not add a menu item in this change.
19. The Change_ note must be approved by ProjectOwner before code implementation begins.

## Data I/O Descriptions

### Recipe Input
ExcelSteps recipe rows for Lite model blocks use the existing recipe columns:

| Column | Meaning for block rows |
|---|---|
| `Sheet` | Lite model name used by `mdl.tblSteps`; must match `mdl.mdlName` |
| `Column` | Base placeholder variable name, ending in `_xx` |
| `Step` | `Input_VarBlock` or `Formula_VarBlock` |
| `Formula/List Name/Sort-by` | Formula template for `Formula_VarBlock`; ignored for `Input_VarBlock` |
| `Comment` | Optional comment text applied to block variable rows |
| `Number Format` | Optional number format applied to model cells |
| `Width` | Optional column width handling remains unchanged if applicable |
| `Params` | Required JSON-like string such as `{nBlockVars:3}` |

### Model Input
The model may be in one of these states before refresh:

1. **Placeholder state:** one base row exists, for example `inputvar_xx`.
2. **Exact indexed state:** contiguous indexed rows already exist from `stem_1` through `stem_n`.
3. **Oversized indexed state:** contiguous indexed rows exist from `stem_1` through `stem_m`, where `m > n`.
4. **Undersized indexed state:** contiguous indexed rows exist from `stem_1` through `stem_m`, where `m < n`.
5. **Invalid state:** `_xx` row and indexed rows both exist, indexed rows are non-contiguous, suffixes are missing/out of sequence, or another variable is interposed in the block.

### Output
After normal `mdl.Refresh`:
- The model contains indexed rows `stem_1` through `stem_n` for each block recipe row.
- Existing input data is preserved for exact and oversized `Input_VarBlock` cases where rows remain in the model.
- Overage indexed rows are removed bottom-up.
- Undersized indexed blocks are deleted and recreated.
- `Formula_VarBlock` rows receive indexed formulas, for example `=@InputVar_xx * 2` becomes `=@InputVar_1 * 2` for `calculatedvar_1`.
- `mdl.rngPopRows`, `mdl.rngFormulaRows`, and Lite recipe dictionaries are refreshed after any row shape changes.

After `ResetAllBlockVars`:
- For each block recipe row, any indexed rows for the block are deleted.
- One base `_xx` placeholder row is left in the model at the block position.
- The placeholder row is ready for a future normal refresh to expand again.

## Project Architecture

### Modified Constants
**Constants.bas**
- Add constants for the new step names:
  - `sAInputVarBlock = "Input_VarBlock"`
  - `sAFormulaVarBlock = "Formula_VarBlock"`
- Add these step names to `sStepList` if validation or UI listing uses that constant.

### Modified Class: mdlScenario.cls
Add Lite-model block orchestration around the current refresh flow.

**New/modified responsibilities:**
- Detect block recipe rows before exact variable-name dictionary matching.
- Parse and validate block params.
- Resolve block location in the model.
- Shape model rows to the requested `nBlockVars` count.
- Re-provision/rebuild model ranges after row insertion/deletion.
- Populate Lite recipe dictionaries with expanded indexed variable names.
- Use a block-aware refresh branch for block rows and keep the existing normal `mdlRow` refresh workflow as the `Else` path for non-block rows.
- Skip indexed rows already consumed by the block branch so they are not refreshed twice.
- Keep normal non-block Lite behavior unchanged.

**Likely new methods:**
- `RefreshBlockVars(mdl) As Boolean` — scan Lite recipe rows and shape all block rows before ordinary refresh.
- `ReadBlockRecipeParams(mdl, rngStepRow, ByRef nBlockVars As Long) As Boolean` — parse `{nBlockVars:n}`.
- `SetBlockVarProps(mdl, r As mdlRow, rngRecipeRow As Range) As Boolean` — initialize block-specific attributes on `mdlRow`.
- `FindBlockRows(mdl, r As mdlRow) As Boolean` — locate `_xx` placeholder or indexed rows.
- `ValidateBlockContiguity(mdl, r As mdlRow) As Boolean` — enforce contiguous, ordered block rows.
- `ShapeBlockRows(mdl, r As mdlRow) As Boolean` — apply placeholder/exact/over/under row logic.
- `ExpandLiteBlockStepDicts(mdl, r As mdlRow) As Boolean` — add indexed formula/format/comment entries to Lite dictionaries.
- `IsBlockRefreshRow(mdl, rngRow, ByRef r As mdlRow) As Boolean` — detect whether the current row starts a block already shaped from a recipe row.
- `ResetAllBlockVars(mdl) As Boolean` — delete indexed block rows and restore base placeholders.

### Modified Class: mdlRow.cls
Extend `mdlRow` with block-specific attributes and helper methods while keeping existing single-row refresh intact.

**Candidate new public attributes:**
- `IsBlockVar As Boolean`
- `IsInputVarBlock As Boolean`
- `IsFormulaVarBlock As Boolean`
- `nBlockVars As Long`
- `sBaseVar As String` — recipe placeholder such as `inputvar_xx`
- `sVarStem As String` — stem such as `inputvar`
- `sFormulaBase As String` — formula template such as `=inputvar_xx * 2`
- `rngBlockRows As Range`
- `rngPlaceholderRow As Range`
- `nExistingBlockVars As Long`

**Candidate helper methods:**
- `InitBlockVar(r, mdl, rngRecipeRow) As Boolean` — set block recipe attributes from ExcelSteps row.
- `IndexedVarName(r, idx As Long) As String` — return `stem_idx`.
- `IndexedFormula(r, idx As Long) As String` — replace `_xx` tokens with `_<idx>` in formula template.
- `RefreshVarBlock(r, mdl) As Boolean` — loop indexed rows in `rngBlockRows`, set indexed `mdlRow` attributes, and call the existing row methods for name/formula/format behavior.
- `IsBlockStep(r) As Boolean` — true for `Input_VarBlock` or `Formula_VarBlock`.

### Modified Lite Refresh Flow
Current flow:
1. `mdl.Provision` prepares `tblSteps` for Lite models.
2. `SetRngFormulaRows` scans current model rows and exact-matches variable names to recipe rows.
3. `Refresh` loops `mdl.rngPopRows`, instances `mdlRow`, and writes formulas/formats.

Planned flow:
1. `mdl.Provision` prepares `tblSteps` for Lite models as today.
2. If `.IsLiteModel`, call `RefreshBlockVars(mdl)` before `SetRngFormulaRows` or before the row loop in `Refresh`.
3. `RefreshBlockVars` shapes rows for every block recipe row and refreshes model ranges when needed.
4. `SetRngFormulaRows` populates ordinary dictionaries plus expanded dictionary entries for block rows.
5. `Refresh` uses a block-aware branch: block rows call `mdlRow.RefreshVarBlock`; non-block rows continue through the existing `mdlRow.Init` / `NameRow` / `WriteRowFormula` / `FormatRow` path.

### Interface/API Additions
**Interface.bas**
- `Public Sub ResetAllBlockVarsDriver()`
  - Driver-style entry point using `SetErrs "driver"`.
  - Targets the active workbook and active sheet/model by default.
  - Initializes/provisions the specified Lite model.
  - Calls the Boolean API/helper.
- `Public Function ResetAllBlockVarsAPI(wkbkR As Workbook, Optional sht, Optional MdlName, Optional Defn) As Boolean`
  - Explicit test/project API.
  - Provisions the requested Lite mdl using existing `mdl.Init`/`mdl.Provision` patterns.
  - Calls `mdl.ResetAllBlockVars(mdl)`.

No ExcelSteps menu item is added in this change.

### Cross-Workbook Factory Functions
No new factory function is required if implementation stays inside `mdlScenario` and `mdlRow`, because tests already use:
- `ExcelSteps.New_mdl`
- `ExcelSteps.New_mdlRow`
- `ExcelSteps.New_tbl`

If a later implementation phase extracts `MdlBlockVar.cls`, add `New_MdlBlockVar` to `Validation.bas` and update this note before coding that branch.

## Test Architecture

### Test Module Location
**tests_mdlScenario.bas** — add tests near the existing Lite model variation tests.

### Procedure Group
Recommended new group: `procs.mdlBlockVars`

Required wiring:
- Add `Public mdlBlockVars As Object` to `Procedures.cls`.
- In `Procedures.Init`, instance and name it under `TestDriver_mdlScenario` groups.
- In `TestDriver_mdlScenario`, add enable toggle and a new block calling the new tests.

Alternative if ProjectOwner wants less wiring: add tests under existing `procs.mdlVariations`. Recommended is the new group because this is a distinct feature surface with several edge cases.

### Test Data/Helpers
**Populate_Mdl.bas**
- Add helper to populate Lite model block examples based on [[Example_Mdl_BlockVars]].
- Add helper to populate ExcelSteps block recipe rows based on [[Example_Mdl_BlockVars_Recipe]].
- Prefer small, focused model layouts with one ordinary input/calculated variable plus block examples.

**Suggested helpers:**
- `PopulateSMdlBlockVars(tst, wksht)`
- `PopulateStepsSMdlBlockVars(wkbk, MdlName, Optional nBlockVars As Long = 3)`
- `InitBlockVarTest(tst, mdl)`

### Test Coverage
| Scenario | Expected outcome |
|---|---|
| Placeholder `Input_VarBlock` | `_xx` row replaced by `stem_1..stem_n`; number formats applied; input values blank or copied only where intentionally seeded |
| `Input_VarBlock` with text in formula column | Text is ignored; refresh still applies input block formats |
| Placeholder `Formula_VarBlock` | `_xx` row replaced by `stem_1..stem_n`; indexed formulas written and calculation style applied |
| Formula references another block with `_xx` | Each formula substitutes matching suffix index, including references such as `@InputVar_xx` |
| Existing exact input block | Existing input values remain; formats refresh |
| Existing exact formula block | Rows remain; formulas/format refresh |
| Existing oversized input block | Rows greater than `nBlockVars` deleted bottom-up; retained input values preserved |
| Existing oversized formula block | Rows greater than `nBlockVars` deleted bottom-up; retained formulas refreshed |
| Existing undersized block | Existing block deleted and recreated with `nBlockVars` indexed rows |
| Non-contiguous indexed block | Refresh hard-fails and marks/records error |
| Missing or invalid `nBlockVars` | Refresh hard-fails |
| Recipe variable lacks `_xx` suffix | Refresh hard-fails |
| `ResetAllBlockVarsAPI` | Indexed rows are replaced by a single `_xx` placeholder row for every block recipe row |
| Existing non-block Lite model tests | Continue passing unchanged |

### Existing Tests Affected
- `test_RefreshSMdl5`, `test_RefreshSMdl6`, `test_RefreshSMdl7` should continue to pass and should not require changes except for any new shared helper behavior.
- `test_mdlRowLiteModel` may need a small addition if `mdlRow.ReadPropsLite` gains block awareness.
- No `tests_tblRowsCols.bas` changes are expected unless shared recipe params parsing is moved into `RecipeStep.cls` for reuse.

### Unit/Integration Scope
- Unit-style tests for helper methods such as `IndexedVarName`, `IndexedFormula`, and params validation if exposed enough to call through `mdlRow`.
- Integration tests through `mdl.Provision` + `mdl.Refresh` for the real block expansion behavior.
- API test through `ExcelSteps.ResetAllBlockVarsAPI`.

## Discussion: Class Boundary

**Decision:** Start by extending existing `mdlRow.cls` rather than adding a new class.

**Rationale:**
`mdlRow` already holds row-level variable metadata, formula strings, number formats, and write/format behavior. Block variables are still model rows, and the simplest first implementation is to add block metadata and helper methods to `mdlRow` while keeping row-shaping orchestration in `mdlScenario`.

The intended boundary is a split wrapper: `mdlScenario` identifies and shapes blocks, then calls a focused `mdlRow.RefreshVarBlock` method for the block rows. That `mdlRow` method loops the indexed rows, substitutes `_xx` with `_<idx>` in both variable names and formulas, and delegates to the existing single-row methods. Ordinary variables remain the `Else` path in `mdlScenario.Refresh` and use the current `mdlRow` workflow unchanged.

**Risk:**
Block row detection and row insertion/deletion are multi-row responsibilities. If `mdlRow` starts owning too much orchestration, extract a later `MdlBlockVar` class before implementation proceeds beyond the first working spike.

## Discussion: Row Count Changes

**Exact match:** Preserve all rows and data; refresh only formulas/formats.

**Too many rows:** Delete from bottom up until `nBlockVars` rows remain. This preserves existing input data in retained rows and avoids index shifts during deletion.

**Too few rows:** Delete the existing undersized block and recreate from the recipe. This avoids ambiguous insertion behavior and ensures all rows are generated from one clean recipe definition.

**Placeholder row:** Treat `_xx` as a template locator and replace it with indexed rows.

## Discussion: Reset Utility

`ResetAllBlockVars` is intentionally separate from normal refresh. Normal refresh should preserve usable model state when possible; reset is an explicit utility for returning a model to recipe-template form.

The utility cycles through the Lite mdl's block recipe rows, finds the corresponding indexed block in the model, deletes the indexed rows, and inserts/restores one base `_xx` variable row at the block location. This is useful before distributing a template or after changing recipe structure.

## Testing Considerations

### Module Structure and Target Procedures
- **Code modules/classes:** `mdlScenario.cls`, `mdlRow.cls`, `Interface.bas`, `Constants.bas`
- **Test modules:** `tests_mdlScenario.bas`, `Populate_Mdl.bas`, `Procedures.cls`
- **Procedure group:** `procs.mdlBlockVars` recommended

### Cross-Workbook Instantiation
Use existing factories:
```vb
Dim mdl As Object: Set mdl = ExcelSteps.New_mdl
Dim r As Object: Set r = ExcelSteps.New_mdlRow
```

No new class factory required unless implementation extracts `MdlBlockVar`.

### Error Cases to Validate
1. Missing `{nBlockVars:n}` params.
2. `nBlockVars` is zero, negative, non-numeric, or not an integer.
3. Recipe variable is `stem` or `stem_1` instead of `stem_xx`.
4. Indexed block has `stem_1`, `stem_3` with missing `stem_2`.
5. Indexed block has another variable interposed between block rows.
6. `Formula_VarBlock` has blank formula text.

### Test Data Requirements
No external xlsx files are required. Use workbook-local test population helpers in `Populate_Mdl.bas`.

### Existing Behavior Protection
Run the full `tests_mdlScenario` driver after implementation, with special attention to ordinary Lite model tests. If row-shaping requires re-provisioning during refresh, verify normal Lite model dictionary matching remains exact and unchanged for non-block rows.

## Procedure Outline

### `mdlScenario.Refresh` [Modified]
```vb
Function Refresh(mdl) As Boolean
    ' Validate variable/scenario names as today
    ' Delete prior range names if needed as today
    ' For Lite models, prepare/shape block variable rows before row loop
    ' Loop through rngPopRows with block-aware branching
    ' If current row starts a block, call mdlRow.RefreshVarBlock and skip consumed rows
    ' Else instance mdlRow for the current row as today
    ' Else branch names row/cell, writes formulas, applies formatting as today
    ' Format template/header columns
    ' Format model class
End Function
```

### `mdlRow.RefreshVarBlock` [New]
```vb
Function RefreshVarBlock(r, mdl) As Boolean
    ' Loop idx from 1 to nBlockVars across rngBlockRows
    ' Instance an ordinary mdlRow for the indexed row
    ' Set indexed variable name and formula before formula write
    ' Call Init / NameRow / WriteRowFormula / FormatRow as applicable
End Function
```

### `mdlScenario.RefreshBlockVars` [New]
```vb
Function RefreshBlockVars(mdl) As Boolean
    ' Exit if not Lite model or no Steps rows
    ' Find recipe rows whose Step is Input_VarBlock or Formula_VarBlock
    ' For each block recipe row:
    '   Instance mdlRow and InitBlockVar from recipe row
    '   Validate nBlockVars and _xx base variable
    '   Locate placeholder or indexed block in mdl.colrngVarNames
    '   Validate contiguity and suffix order
    '   Shape rows according to placeholder/exact/over/under rules
    ' Re-provision/rebuild mdl ranges if any rows changed
End Function
```

### `mdlScenario.SetRngFormulaRows` [Modified]
```vb
Sub SetRngFormulaRows(ByRef mdl)
    ' Initialize Lite dictionaries as today
    ' Populate ordinary exact-match recipe entries as today
    ' For block recipe rows, add indexed dictionary entries:
    '   dStepsNumFormats(stem_i) = recipe Number Format
    '   dStepsComments(stem_i) = recipe Comment
    '   dStepsFormulas(stem_i) = indexed Formula_VarBlock formula
    ' Add formula rows to rngFormulaRows for Formula_VarBlock rows
End Sub
```

### `mdlRow.InitBlockVar` [New]
```vb
Function InitBlockVar(r, mdl, rngRecipeRow) As Boolean
    ' Read base variable name, Step, formula template, num format, comment, params
    ' Set IsInputVarBlock / IsFormulaVarBlock
    ' Parse nBlockVars from params
    ' Validate base variable ends in _xx
    ' Set sVarStem
End Function
```

### `mdlRow.IndexedFormula` [New]
```vb
Function IndexedFormula(r, idx As Long) As String
    ' Return formula template with every _xx token replaced by _<idx>
End Function
```

### `mdlScenario.ResetAllBlockVars` [New]
```vb
Function ResetAllBlockVars(mdl) As Boolean
    ' Require Lite model with prepared tblSteps
    ' For each block recipe row:
    '   Initialize mdlRow block attributes
    '   Locate existing indexed block or placeholder
    '   Delete indexed rows bottom-up
    '   Insert or restore one placeholder row named base _xx
    ' Re-provision/rebuild model ranges
End Function
```

### `Interface.ResetAllBlockVarsAPI` [New]
```vb
Function ResetAllBlockVarsAPI(wkbkR As Workbook, Optional sht, Optional MdlName, Optional Defn) As Boolean
    ' Instance and provision mdlScenario for requested Lite model
    ' Call mdl.ResetAllBlockVars(mdl)
End Function
```

### `Interface.ResetAllBlockVarsDriver` [New]
```vb
Sub ResetAllBlockVarsDriver()
    ' Driver wrapper for active workbook/model
    ' Set application environment
    ' Call ResetAllBlockVarsAPI
    ' Restore application environment
End Sub
```

___
## AI Summary of As-Built Changes

- [ Constants.bas](vscode-file://vscode-app/c:/Users/j.d.landgrebe/AppData/Local/Programs/Microsoft%20VS%20Code/fcf604774b/resources/app/out/vs/code/electron-browser/workbench/workbench.html): ~5 lines changed/added  
    Updated version constants/comments and added `Input_VarBlock` / `Formula_VarBlock` constants plus step-list entries.
- [Interface.bas](vscode-file://vscode-app/c:/Users/j.d.landgrebe/AppData/Local/Programs/Microsoft%20VS%20Code/fcf604774b/resources/app/out/vs/code/electron-browser/workbench/workbench.html): ~40-50 lines added/modified  
    Added `ResetAllBlockVarsAPI`, `ResetAllBlockVarsDriver`, and later adjusted optional `mdlName` handling.
- [mdlRow.cls](vscode-file://vscode-app/c:/Users/j.d.landgrebe/AppData/Local/Programs/Microsoft%20VS%20Code/fcf604774b/resources/app/out/vs/code/electron-browser/workbench/workbench.html): ~100-120 lines added/modified  
    Added block-variable attributes and helper methods: `InitBlockVar`, `IndexedVarName`, `IndexedFormula`, `RefreshVarBlock`.
- [mdlScenario.cls](vscode-file://vscode-app/c:/Users/j.d.landgrebe/AppData/Local/Programs/Microsoft%20VS%20Code/fcf604774b/resources/app/out/vs/code/electron-browser/workbench/workbench.html): ~250-300 lines added/modified  
    Added block-variable orchestration: `dBlockVars`, `RefreshBlockVars`, row finding/shaping helpers, range rebuild, block map/index helpers, reset method, and `Refresh` / `SetRngFormulaRows` integration.
- [Procedures.cls](vscode-file://vscode-app/c:/Users/j.d.landgrebe/AppData/Local/Programs/Microsoft%20VS%20Code/fcf604774b/resources/app/out/vs/code/electron-browser/workbench/workbench.html): ~5-8 lines changed/added  
    Added the `mdlBlockVars` procedure group and initialized it.
- [Populate_Mdl.bas](vscode-file://vscode-app/c:/Users/j.d.landgrebe/AppData/Local/Programs/Microsoft%20VS%20Code/fcf604774b/resources/app/out/vs/code/electron-browser/workbench/workbench.html): ~45-60 lines added/modified  
    Added block-variable fixture helpers: `PopulateSMdlBlockVars` and `PopulateStepsSMdlBlockVars`.
- [tests_mdlScenario.bas](vscode-file://vscode-app/c:/Users/j.d.landgrebe/AppData/Local/Programs/Microsoft%20VS%20Code/fcf604774b/resources/app/out/vs/code/electron-browser/workbench/workbench.html): ~150-190 lines added/modified  
    Added the `MdlBlockVars` driver block and four tests for expansion, oversized shrink, undersized recreate, and reset API. Also includes your later focused-test toggle edits.
 
---

**Status:** Change Completed/validated