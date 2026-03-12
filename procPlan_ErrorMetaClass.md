## Procedure Purpose
Define and integrate a new typed helper class, `ErrorMeta`, that encapsulates Errors_ lookup metadata for `ErrorHandling.AppendErrMsg`.

The goal is to remove magic-index Variant array handling and move lookup/not-found/malformed-row validation into a dedicated class-level flow while preserving current user-facing and developer-facing message outputs.

## Procedure Detailed Requirements
### Overall Action
Replace direct `aryErrLookup()` array-index usage in `ErrorHandling.AppendErrMsg` with a typed object flow:
1. `ErrorMeta.LoadFromLookup` reads metadata from `Errors_` based on current `errs.Locn` and `errs.iCodeLocal` context.
2. `ErrorMeta.Validate` verifies required fields and types.
3. `ErrorHandling.AppendErrMsg` selects user-facing or developer-facing message construction through `ErrorMeta` methods.

### Key Behavioral Requirements
1. Not-found handling occurs immediately in lookup path (not deferred to sentinel checks in `AppendErrMsg`).
2. Malformed row handling uses message text: `Malformed Errors_ Row for <Locn>`.
3. Preserve existing message wording behavior for user-facing and developer-facing outputs.
4. Preserve existing nested trace behavior (`Called by ...`) for non-user-facing nested errors.

### Input/Output Context
- Inputs used by lookup path:
  - `errs.Locn`
  - `errs.iCodeLocal`
  - `errs.wkbkE.Sheets(shtErrors)`
- Outputs:
  - Typed fields on `ErrorMeta`
  - Message text consumed by `ErrorHandling.AppendErrMsg`
  - Existing `errs.ErrMsg`, `errs.IsUserFacing`, `errs.IsNewErr` state updates remain in `ErrorHandling` orchestration

## Procedure Method/Sub-Procedure Descriptions
### ErrorMeta.LoadFromLookup
- **Action**: Locate and load one metadata row from `Errors_` using current error context.
- **Inputs**:
  - `errsObj` (ErrorHandling-like context object or direct arguments for Locn/code lookup)
- **Outputs**:
  - Returns `Boolean`
  - Sets typed properties: `Code`, `Routine`, `Message`, `IsUserFacing`, `IsFound`, `IsBaseMessage`
- **Logic Steps**:
  1. Resolve effective report code from base/local context (same logic used today for report code lookup).
  2. Query `Errors_` table using existing lookup utility flow.
  3. If no row is found:
     - Set `IsFound = False`
     - Set `Message = "Msg Not Found"` (or project-approved not-found text)
     - Return True with explicit not-found state.
  4. If row is found:
     - Map row values into typed properties.
     - Set `IsFound = True`.
- **Validation/Error Conditions**:
  - Errors_ sheet missing/unavailable -> return False and route through existing error handling.
  - Unexpected lookup utility failure -> return False.

### ErrorMeta.Validate
- **Action**: Validate loaded metadata fields for required completeness and type correctness.
- **Inputs**: Current `ErrorMeta` property state.
- **Outputs**:
  - Returns `Boolean`.
  - Optionally sets a validation/fallback message state for malformed rows.
- **Logic Steps**:
  1. If `IsFound = False`, skip malformed checks and return True (not-found already handled in lookup path).
  2. Validate required fields:
     - `Code` present/valid numeric
     - `Routine` not blank when required
     - `Message` not blank where expected
     - `IsUserFacing` parseable as Boolean
  3. On malformed data:
     - Set message to `Malformed Errors_ Row for <Locn>`
     - Return True with malformed state indicated (or return False if final implementation chooses hard-fail).
- **Validation/Error Conditions**:
  - Required field blank or malformed Boolean.

### ErrorMeta.ToUserMessage
- **Action**: Return concise user-facing message text.
- **Inputs**:
  - `errParam As String`
- **Outputs**:
  - Message string for user-facing path.
- **Logic Steps**:
  1. Start from validated `Message`.
  2. Append `errParam` when non-empty.
  3. Return concise text only (no trace prefix).

### ErrorMeta.ToDeveloperMessage
- **Action**: Return developer-facing message text with code/context formatting.
- **Inputs**:
  - `errCodeReport As Long`
  - `errParam As String`
- **Outputs**:
  - Developer-formatted message string.
- **Logic Steps**:
  1. Build existing developer message shape:
     - `Error <code>; in sub or function, ...`
  2. Append routine/message segments with line breaks per current behavior.
  3. Append `errParam` where currently expected.

### ErrorHandling.AppendErrMsg (refactor integration)
- **Action**: Orchestrate root-vs-nested message update using `ErrorMeta`.
- **Inputs**:
  - `sMsgSuffix` from `RecordErr`.
- **Outputs**:
  - Updates `errs.ErrMsg`, `errs.IsUserFacing`, `errs.IsNewErr`.
- **Logic Steps**:
  1. If `.IsNewErr`:
     - Load metadata via `ErrorMeta.LoadFromLookup`.
     - Validate metadata via `ErrorMeta.Validate`.
     - Set `.IsUserFacing` from typed metadata state.
     - Build message through user/developer formatter.
     - Append to `.ErrMsg`; set `.IsNewErr = False`.
  2. Else (nested path):
     - Append `sMsgSuffix` only when non-user-facing.

  ## Additional Procedure: RecordErr Consolidation (3/12/26)
  ### Objective
  Refactor `ErrorHandling.RecordErr` into the primary orchestration procedure for error lifecycle, and fold current `AppendErrMsg` branching into this single flow while preserving external behavior and call signatures.

  ### Scope
  1. Keep routine name `RecordErr` unchanged.
  2. Move root/nested message branching decisions to `RecordErr` (or narrowly-scoped private helpers called by `RecordErr`).
  3. Migrate metadata lookup responsibilities from `aryErrLookup` toward `ErrorMeta`.
  4. De-emphasize `iErrNotFound` sentinel checks in message assembly; rely on explicit `meta.IsFound` state.

  ### Detailed Flow Requirements
  1. **Call-context setup**
    - Compute `IsDriverCall = IsMissing(CallingFunction)`.
    - For non-driver Boolean callers, set `CallingFunction = False` immediately.
    - Set `errs.Locn = Locn`.
  2. **Nested path first (early exit style)**
    - If `Not errs.IsNewErr`:
      - If `Not errs.IsUserFacing`, append `Called by <Locn>` line.
      - Skip root lookup/build logic.
  3. **Root path orchestration**
    - Resolve base/report code.
    - Instance `ErrorMeta` and call `LoadFromLookup`.
    - Call `Validate`.
    - If `errs.ErrMsg` already contains text, add line break separator before appending new root message.
    - If `meta.IsFound = False`, append not-found message per current wording rules.
    - Else append message via `meta.ToUserMessage` or `meta.ToDeveloperMessage` based on `meta.IsUserFacing`.
    - Set `errs.IsNewErr = False`.
  4. **Reporting gate**
    - Preserve current reporting condition using `IsShowMsgs` and driver/testing context.
    - Invoke `ReportMsg` only when gate evaluates True.

  ### Method Ownership Changes
  1. `RecordErr` owns control flow for new vs nested handling and report gating.
  2. `ErrorMeta` owns lookup-state semantics (`found`, malformed metadata, message-shape inputs).
  3. `aryErrLookup` becomes migration candidate:
    - either move logic into `ErrorMeta.LoadFromLookup` directly,
    - or retain temporarily as compatibility helper until cleanup pass.
  4. `SetTblELocations` becomes cleanup candidate if no longer needed by lookup path.

  ### Behavioral Parity Constraints
  1. No breaking changes to existing call sites that use `errs.RecordErr "Locn", FunctionName` pattern.
  2. Preserve existing user-facing and developer-facing message text shapes unless explicitly called out by tests.
  3. Preserve non-user-facing nested trace behavior (`Called by ...`).
  4. Preserve `IsShowMsgs`/`IsTesting` behavior in automated tests and driver runs.

  ### Refactor Completion Criteria
  1. `RecordErr` reads as one end-to-end procedure from context setup to conditional `ReportMsg`.
  2. `AppendErrMsg` is either removed or reduced to a thin internal helper with no duplicated branching logic.
  3. Sentinel-based not-found branching in message assembly is replaced by `meta.IsFound` pathing.
  4. Tests confirm parity for:
    - root developer-facing,
    - root user-facing,
    - not-found,
    - malformed row,
    - nested trace,
    - reporting gate behavior.

## Testing Requirements
### Test Module Location
- `tests_ErrorHandling.bas`
- `Populate_Errs.bas` (new helper module for mock `Errors_` data)
- Procedure grouping in `Procedures.cls`: `procs.ErrorHandling`

### Test Setup Pattern
- Use `create_new_test_procedure.md` guidance for:
  1. Module driver sub in `tests_ErrorHandling.bas`
  2. New Procedure attribute/instancing in `Procedures.cls`
- Do not rely on the add-in `*.xlam` `Errors_` table for these tests.
- Use the newly-added `Errors_` sheet in the tests workbook (`tests_XLSteps.xlsm`) as the lookup source.
- Pre-initialize `errs` in tests with `wkbkE = tst.wkbkTest` by calling `SetErrs` with a dummy non-driver `CallingFunction` and passing `tst.wkbkTest`.
- Add a setup helper in `Populate_Errs.bas` to prepare fixture data:
  1. Clear `errs.wkbkE.Sheets(shtErrors)` before each fixture build.
  2. Populate row 1 headers using new `sErrorsHeadings` constant.
  3. Populate default mock rows that cover found and not-found flows.
  4. Allow tests to modify/add rows as needed for malformed and edge-case scenarios.

### Required Test Coverage
1. **Load success**: `LoadFromLookup` maps typed fields correctly from a valid row.
2. **Not found**: lookup against the mock table returns explicit not-found state and not-found message from lookup path.
3. **Malformed row**: row found but missing/malformed required fields -> `Malformed Errors_ Row for <Locn>` behavior.
4. **User-facing message build**: concise output and `errParam` append behavior.
5. **Developer message build**: preserve existing format and line-break behavior.
6. **AppendErrMsg root path**: integrates load/validate/build and toggles `.IsNewErr`.
7. **AppendErrMsg nested path**: appends `Called by ...` only when non-user-facing.
8. **No behavior drift checks**: representative legacy cases produce expected user/developer outputs.

### Edge Cases
1. Errors_ sheet missing.
2. Duplicate code rows.
3. Boolean metadata values stored as string variants (`"TRUE"`, `"FALSE"`, mixed case).
4. Empty/whitespace message cells.
5. Existing non-empty `ErrMsg` before append (line break behavior).

### Success Criteria
1. AppendErrMsg complexity is reduced by moving lookup metadata typing/validation into `ErrorMeta`.
2. Required not-found and malformed-row behaviors are handled in the lookup/validation path.
3. Existing output wording remains unchanged where designated.
4. New tests pass for all coverage categories above.

## References
- [[Change_ErrorMetaClass]]
- [[VBA Project Changes]]
- [[VBA Project Architecture]]
- ../Excel_Steps/src/ErrorHandling.cls
- ../Excel_Steps/src/ErrorMeta.cls
- ../Excel_Steps/tests/tests_ErrorHandling.bas
