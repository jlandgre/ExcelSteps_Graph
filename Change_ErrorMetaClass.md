## Purpose
Consolidate ErrorHandling root and nested message flow around the 3/12 RecordErr refactor intent: keep `RecordErr` as the stable entry point, centralize root message assembly, and preserve current user-facing/developer-facing behavior.

## Background
Current code now includes:
1. `RecordErr` orchestration with conditional reporting gate.
2. `AppendErrMsg` retained as compatibility wrapper.
3. `ProcessRootErrMsg` private helper for root message build.
4. `ErrorMeta` typed metadata state with direct `Errors_` lookup.

Reference documents:
- [[VBA Project Architecture]]
- [[VBA Project Changes]]
- ../Excel_Steps/src/ErrorHandling.cls
- ../Excel_Steps/src/ErrorMeta.cls
- ../Excel_Steps/tests/tests_ErrorHandling.bas
- ../Excel_Steps/tests/Populate_Errs.bas

## Data I/O Descriptions
### Input Data
1. `errs.Locn`, `errs.iCodeLocal`, `errs.ErrParam`, `errs.IsNewErr`, `errs.IsShowMsgs`, `errs.IsTesting`.
2. `Errors_` lookup table in `errs.wkbkE`:
   - Code (col 1)
   - Routine (col 3)
   - Message (col 4)
   - User/developer flag (col 6)

### Output Data
1. `errs.ErrMsg` appended with root or nested content.
2. `errs.IsUserFacing` set for root path behavior mode.
3. `errs.IsNewErr` toggled `False` after root message materialization.

### Behavioral Output
1. User-facing: concise message plus optional `ErrParam`.
2. Developer-facing: `Error <code>; in sub or function, ...` format with trace lines for nested calls.

## Project Architecture
### Modified Classes
1. `ErrorHandling.cls`
   - `RecordErr` is the top-level message/report orchestration point.
   - `AppendErrMsg` remains a compatibility wrapper.
   - `ProcessRootErrMsg` contains root-path build logic.

2. `ErrorMeta.cls`
   - Typed public fields: `Code`, `Routine`, `Message`, `IsUserFacing`, `IsFound`, `IsBaseMessage`, `IsMalformed`.
   - `LoadFromLookup(meta, errs)` performs direct table lookup.
   - `Validate(meta, Locn)` enforces malformed-row normalization.
   - `ToUserMessage` and `ToDeveloperMessage` format output strings.

## Test Architecture
1. Test module: `tests_ErrorHandling.bas`.
2. Procedure groups in `Procedures.cls`:
   - `procs.ErrorHandling`
   - `procs.RecordErr`
3. Fixture source: `Populate_Errs_Default` in `Populate_Errs.bas` on test workbook `Errors_` sheet.
4. Cross-workbook instantiation via existing Validation factory methods (`New_ErrorHandling`, `New_ErrorMeta`).

## Discussion: RecordErr Consolidation (3/12)
1. Keep routine name/signature compatibility for widespread existing call sites.
2. Route root and nested append operations through one callable append path.
3. Use typed metadata state (`meta.IsFound`, `meta.IsMalformed`, `meta.IsUserFacing`) rather than array-index branching for root formatting decisions.
4. Preserve existing wording contracts and reporting gate semantics.

## Testing Considerations
1. Unit-level coverage for `ErrorMeta` lookup/validation/formatters.
2. Integration-level coverage for `RecordErr` root and nested paths.
3. Mock test data requirements:
   - base rows
   - user-facing rows
   - developer-facing rows
   - malformed row examples
4. Existing tests impacted:
   - `AppendErrMsg` assertions updated to align with wrapper behavior and current mock table semantics.
5. Edge cases:
   - not found codes
   - malformed metadata row
   - user-facing nested call should not append trace
   - non-user-facing nested call should append trace

## Procedure Outline
**RecordErr Refactor Procedure**
* **`RecordErr`**: [[procPlan_ErrorMetaClass]]
* `SetCallContext` - Determine driver/Boolean call handling and set caller return state
* `SetLocnAndSuffix` - Set `errs.Locn` and optional nested suffix (`Called by ...`)
* `AppendErrMsg` - Route to root helper or nested append path
* `ProcessRootErrMsg` - Resolve codes, load/validate metadata, format and append root message
* `ReportIfEnabled` - Apply `IsShowMsgs` and driver/testing gate, then call `ReportMsg`
