## Procedure Purpose
Define the as-coded implementation details for the 3/12 RecordErr consolidation procedure in `ErrorHandling.cls`, including `ErrorMeta` integration and final testing requirements.

## Procedure Detailed Requirements
1. Preserve `RecordErr` API compatibility:
   - `RecordErr(Locn, Optional ByRef CallingFunction)`
2. Preserve caller semantics:
   - set Boolean caller to `False` on error path.
3. Preserve reporting gate behavior:
   - report only when `.IsShowMsgs` and `(driver call or testing context)`.
4. Root error path must:
   - resolve base/report code,
   - load metadata from `Errors_`,
   - validate metadata,
   - append correct user/developer message,
   - set `.IsNewErr = False`.
5. Nested path must append `Called by ...` only for non-user-facing mode.

## Procedure Method/Sub-Procedure Descriptions
### RecordErr
- **Action**: Primary orchestration entry for recording error state, appending message content, and conditional reporting.
- **Inputs**:
  - `Locn As String`
  - `CallingFunction As Boolean ByRef Optional`
- **Outputs**:
  - updates `errs.ErrMsg`, `errs.Locn`, `errs.IsUserFacing`, `errs.IsNewErr`
  - sets optional Boolean caller argument to `False`
- **Logic Steps**:
  1. Compute `IsDriverCall` from `IsMissing(CallingFunction)`.
  2. If non-driver Boolean caller, set `CallingFunction = False`.
  3. Set `errs.Locn = Locn`.
  4. If nested and non-user-facing, prepare `sMsgSuffix = "Called by " & Locn`.
  5. Call `AppendErrMsg(sMsgSuffix)` and exit on failure.
  6. If gate passes, call `ReportMsg`.
- **Validation/Error Conditions**:
  - if root helper fails, exit without reporting.

### AppendErrMsg
- **Action**: Compatibility wrapper that routes root assembly or nested append.
- **Inputs**:
  - `sMsgSuffix As String`
- **Outputs**:
  - `Boolean` success status
  - appends to `errs.ErrMsg`
- **Logic Steps**:
  1. If `.IsNewErr`, call `ProcessRootErrMsg()` and return result.
  2. ElseIf `Not .IsUserFacing`, append `sMsgSuffix & vbCrLf`.
  3. Return `True`.

### ProcessRootErrMsg (Private)
- **Action**: Build and append root message using `ErrorMeta` state.
- **Inputs**:
  - internal `errs` state (`Locn`, `iCodeLocal`, `ErrParam`, `ErrMsg`)
- **Outputs**:
  - `Boolean` success
  - updated `errs.ErrMsg`, `errs.IsUserFacing`, `errs.IsNewErr`
- **Logic Steps**:
  1. Call `GetBaseErrCode` and compute `iCodeReport`.
  2. Instance `meta As ErrorMeta`.
  3. Call `meta.LoadFromLookup(meta, errs)`.
  4. Call `meta.Validate(meta, errs.Locn)`.
  5. Add line break if `ErrMsg` already contains text.
  6. If `Not meta.IsFound`, append not-found wording.
  7. Else:
     - set `errs.IsUserFacing = meta.IsUserFacing`
     - append `meta.ToUserMessage(...)` or `meta.ToDeveloperMessage(...)`.
  8. Set `errs.IsNewErr = False` and return `True`.

### ErrorMeta.LoadFromLookup
- **Action**: Lookup one row from `Errors_` using `errs.iCodeReport`.
- **Inputs**:
  - `meta`
  - `errs`
- **Outputs**:
  - `Boolean`
  - sets `meta` fields: `Code`, `Routine`, `Message`, `IsUserFacing`, `IsFound`, `IsBaseMessage`, `IsMalformed`
- **Logic Steps**:
  1. Read worksheet/ranges from `errs.wkbkE.Sheets(shtErrors)`.
  2. Compute `errs.iCodeReport = errs.iCodeBase + errs.iCodeLocal`.
  3. Find row by code.
  4. If no row, set not-found defaults.
  5. If row found, map field values and parse Boolean-like user flag.

### ErrorMeta.Validate
- **Action**: Validate row completeness and normalize malformed metadata.
- **Inputs**:
  - `meta`
  - `Locn As String`
- **Outputs**:
  - `Boolean`
  - malformed normalization of `meta.Message` and `meta.IsUserFacing`
- **Logic Steps**:
  1. If not found, return `True`.
  2. Mark malformed when required fields missing/invalid.
  3. If malformed, set `meta.Message = "Malformed Errors_ Row for " & Locn` and force `meta.IsUserFacing = False`.

## Testing Requirements
### Test Module Location
1. `tests_ErrorHandling.bas`
2. Procedure groups:
   - `procs.ErrorHandling`
   - `procs.RecordErr`

### Test Setup Pattern
1. Initialize `errs` with `SetErrs` using `wkbkE = tst.wkbkTest`.
2. Seed mock `Errors_` via `Populate_Errs_Default` in `Populate_Errs.bas`.
3. Keep `errs.IsShowMsgs = False` for automation-safe tests.

### Required Test Coverage
1. ErrorMeta mapping success and not-found cases.
2. Malformed-row normalization behavior.
3. User/developer formatter output shape.
4. Append wrapper behavior for root and nested paths.
5. RecordErr Boolean ByRef reset behavior.
6. RecordErr root developer and user-facing flows.
7. RecordErr nested trace append/suppress behavior.

### Edge Cases
1. missing code row
2. malformed row values
3. base-message mapping
4. non-empty existing `ErrMsg` before root append

### Success Criteria
1. Existing `RecordErr` call sites remain compatible.
2. Root and nested message behavior matches current expected outputs.
3. Reporting gate behavior remains unchanged.
4. `tests_ErrorHandling` passes for both `procs.ErrorHandling` and `procs.RecordErr` groups.
