---
status: planning
updated: 5/19/26
---

# Change: ImportParseNorm.FilterVals - Add KeepExcept Option

## Purpose
Add `KeepExcept` filter option to `ImportParseNorm.FilterRows()` to allow elimination of one or more values from normalized table rows, complementing the existing `KeepOnly` filter.

## Background
**Referenced docs:**
- [ImportParseNorm Class Architecture](ImportParseNorm%20Class.md)
- [Current FilterVals Implementation](../src/ImportParseNorm.cls) (lines 400-430)
- [Current ApplyKeepOnlyFilter Implementation](../src/ImportParseNorm.cls) (lines 437-483)
- [Test Reference: test_ImportParseNorm_FilterRows](../tests/tests_ToolboxClasses.bas)

**Change-specific framing:**
Current `FilterRows()` only supports `KeepOnly` filter via dict key lookup. The `KeepExcept` option will allow inverse filtering: eliminating rows where a variable (column) matches excluded values. This is useful for removing outliers, invalid statuses, or specific categories while keeping all others.

## Locked Decisions (Pre-Implementation)

1. KeepExcept and KeepOnly are mutually exclusive and must hard-fail if both are present in the same FilterVals object.
2. KeepExcept value must be a quoted scalar list in FilterVals JSON-like text, wrapped in braces.
3. Matching is case-sensitive.
4. KeepExcept tokens are literal; no trimming.
5. Empty KeepExcept value or empty tokens are misformatted and must hard-fail.
6. BLANK keyword handling is exact all-caps BLANK only.
7. BLANK matches only truly empty cells.
8. Non-existent KeepExcept values are no-op and do not error.
9. Duplicate KeepExcept values are allowed as harmless duplicates.
10. If filtering leaves no rows, hard-fail.
11. Zero-row hard-fail is checked after each filter application by re-provisioning and checking tblNorm.rngRows Is Nothing.
12. Filters across colinfo rows apply sequentially (AND semantics).
13. Comma remains the only delimiter for this version.
14. Value comparison remains CStr-based.
15. KeepExcept implementation uses a bottom-up row scan strategy.
16. Test scope is positive-path only for this change: BLANK case and two-value list case.

## Data I/O Descriptions

### Input
- **dict key:** `"KeepExcept"` (mutually exclusive with `KeepOnly`)
- **dict value:** Comma-separated list of values to eliminate
  - Format inside FilterVals: `{KeepExcept:"Val1,Val2"}`
  - Special keyword: `BLANK` (all caps only) = truly empty cells
  - Spaces are literal and not trimmed
  - Empty list or empty tokens are invalid and hard-fail

### Variable
- **sVarNorm:** Variable name (column) from `.colrngVarNorm` to filter on
- **Target:** `imptbl.tblNorm` — normalized table created by `WriteNormalized()`

### Output
- **Result:** Rows matching any excluded value deleted from `imptbl.tblNorm`
- **Side effect:** `tblNorm.Provision()` called after each filter application to re-establish row ranges
- **Validation:** Hard-fail if post-provision `.tblNorm.rngRows Is Nothing`

## Project Architecture

### Modified/New Classes & Methods
**ImportParseNorm.cls:**
- **FilterRows()** (lines 400-430) — MODIFY: Add `ElseIf dict.Exists("KeepExcept")` branch
- **ApplyKeepExceptFilter()** (NEW) — Private function to handle KeepExcept logic
  - Signature: `Private Function ApplyKeepExceptFilter(imptbl, sVarNorm As String, keepExceptVal As Variant) As Boolean`
  - Parse comma-separated values from `keepExceptVal`
  - Use a single bottom-up row scan against target column to delete matching rows
  - Use CStr comparisons
  - Treat BLANK token as truly empty cells only
  - Re-provision at the end of this filter application

**Pattern consistency:**
- Error handling: `SetErrs` / `errs.RecordErr` per vba-boolean-function-error-handling-structure
- Docstring: 3-line format (hyphens, description, author/date)
- ByRef alias: If tblNorm attribute passed ByRef, create alias first
- Array iteration: Use `For Each` over Split result per vba-preferred-loop-code

## Test Architecture

### Test Module Location
**tests_ToolboxClasses.bas** — `ImportParseNorm` Procedures group (same as existing `test_ImportParseNorm_FilterRows`)

### Test Procedure Group
Procedures group: **ImportParseNorm** 
- Existing: `test_ImportParseNorm_FilterRows` (tests KeepOnly with "Online" value)
- New: `test_ImportParseNorm_KeepExcept_BLANK` — Filter with `KeepExcept:"BLANK"`
- New: `test_ImportParseNorm_KeepExcept_BLANK_Online` — Filter with `KeepExcept:"BLANK,Online"`

### Test Data
**File:** `BR_Raw_Mockup_KeepExcept.xlsx` (user-created; test suite will reference)
- Location: `test_data/` folder (same as existing test data)
- Contains: Raw BR import with `Locn_Raw` values including Online, StoreA, StoreB, StoreC, and one blank location row
- Supports BLANK option testing and two-value list testing

### Test Constants
Add to [tests_ToolboxClasses.bas](../tests/tests_ToolboxClasses.bas) (line 16, after `ImportFile_BR_Example`):
```vb
Const ImportFile_BR_KeepExcept As String = "BR_Raw_Mockup_KeepExcept.xlsx"
```

### Test Coverage
| Scenario | Test Name | Data | Expected Outcome |
|----------|-----------|------|------------------|
| BLANK keyword | `test_ImportParseNorm_KeepExcept_BLANK` | `{KeepExcept:"BLANK"}` | 8 rows remain; no blank Location values |
| BLANK + Online list | `test_ImportParseNorm_KeepExcept_BLANK_Online` | `{KeepExcept:"BLANK,Online"}` | 3 rows remain; Location values exactly StoreA, StoreB, StoreC |

### Cross-Workbook Instantiation
None required — all test code uses existing factory functions (`ExcelSteps.New_ImportParseNorm`, `ExcelSteps.New_Dictionary`, etc.)

## Discussion: Mutually Exclusive Filters

**Decision:** KeepExcept and KeepOnly are **mutually exclusive**; only one dict key can be present per FilterRows call.

**Rationale:**
- Logically opposite operations (keep-only vs keep-except)
- No defined precedence if both present
- Callers can structure logic (if/else) to use one or the other per filtering scenario
- Simplifies implementation and avoids ambiguity

**Implementation:** Detect both keys and hard-fail; otherwise execute a single branch.

## Testing Considerations

### Module Structure & Target Procedures
- **Module:** `tests_ToolboxClasses.bas`
- **Driver sub:** `TestingDriver_ToolboxClasses()` (lines 45–80)
- **Procedures group:** `procs.ImportParseNorm` (enable/disable toggle for test group)
- **Test calls added in driver:** Immediately after existing `test_ImportParseNorm_FilterRows` call

### Unit/Integration Coverage
- **Unit:** Positive-path KeepExcept behavior for BLANK and BLANK+Online
- **Integration:** Full import/parse/normalize/filter flow retained in each new test

### Existing Tests Affected
- `test_ImportParseNorm_FilterRows` — **No changes required** (tests KeepOnly; new code branch not executed)
- All other ImportParseNorm tests — **No changes required** (do not exercise FilterRows)

### Test Data Requirements
- **New file:** `BR_Raw_Mockup_KeepExcept.xlsx` (in `test_data/` folder, user-created)
- **Usage pattern:** Call `InitImportParseNormTest`, then override `files.pfImportFile` to this file in each KeepExcept test

### Edge Cases & Boundary Validations
This change intentionally excludes malformed/negative-path tests. Behavior is still defined for implementation:
1. Empty KeepExcept value hard-fails.
2. Empty tokens hard-fail.
3. Non-existent values are no-op.
4. All rows removed hard-fails.

## Procedure Outline

### Entry Point: FilterRows() [Modified]
```vb
Public Function FilterRows(imptbl) As Boolean
    ' Iterate through colinfo.rngRowsCurTbl rows
    ' For each row, check .tbl.colrngFilterVals for filter string
    ' Parse dict via ParseStringToDictProcedure
  ' Hard-fail if KeepOnly and KeepExcept both exist
  ' EXISTING: KeepOnly → ApplyKeepOnlyFilter()
  ' NEW: KeepExcept → ApplyKeepExceptFilter()
  ' After each applied filter: re-provision and hard-fail if tblNorm.rngRows Is Nothing
End Function
```

### Helper: ApplyKeepExceptFilter() [New Private Function]
```vb
Private Function ApplyKeepExceptFilter(imptbl, sVarNorm As String, keepExceptVal As Variant) As Boolean
  ' 1. Validate keepExceptVal not empty
  ' 2. Split by comma, validate no empty tokens
  ' 3. Resolve target column from sVarNorm
  ' 4. Delete matching rows with bottom-up scan
  '    - CStr comparisons
  '    - BLANK token matches truly empty cell only
  ' 5. Return success; caller handles re-provision and zero-row check
End Function
```

### Sub-Methods (Reused from ApplyKeepOnlyFilter)
- `.rngTblHeaderVal()` — Find column range by header name
- `.wksht.Range().EntireRow.Delete` — Delete matching rows
- `.Provision()` — Re-establish ranges post-deletion

### Error Handling Chain
```
FilterRows() 
  → ApplyKeepExceptFilter() 
    → rngTblHeaderVal(), row iteration/deletes
      (each returns Boolean; chain with "If Not ... GoTo ErrorExit")
```

## Test Setup Notes

1. In each new test, call `InitImportParseNormTest` first.
2. Override `files.pfImportFile = files.pathData & ImportFile_BR_KeepExcept`.
3. Find colinfo row for `VarNameNorm = "Location"` dynamically.
4. Write FilterVals using `Intersect(foundRow, colinfo.tbl.colrngFilterVals).Value2`.
5. FilterVals text must include braces, for example `{KeepExcept:"BLANK,Online"}`.
6. Add both new test calls immediately after `test_ImportParseNorm_FilterRows` in the ImportParseNorm driver block.

---

**Status:** Ready for ProjectOwner approval before implementation proceeds.
