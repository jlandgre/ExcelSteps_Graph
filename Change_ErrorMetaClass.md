## Purpose
Refactor ErrorHandling.AppendErrMsg to reduce complexity by introducing a typed metadata helper class, then use that class to build user-facing and developer-facing messages with clearer validation and fewer branching paths.

Secondary purpose: establish initial unit-test architecture for ErrorHandling behavior that is currently not directly tested.

## Background
Current ErrorHandling.AppendErrMsg uses a Variant array returned from aryErrLookup and indexes by position. This creates three pain points:
1. Readability risk from magic array indices.
2. Limited type validation of lookup metadata.
3. High branching complexity in a single method.

The change introduces a small helper class (working name: ErrorMeta) that encapsulates lookup row metadata and provides typed access plus validation.

This note remains in Change_ scope only. We are not moving to procPlan_ content yet per stage-gate process in [[VBA Project Changes]].

Reference documents:
- [[VBA Project Architecture]]
- [[VBA Project Changes]]
- ../Excel_Steps/src/ErrorHandling.cls
- ../Excel_Steps/src/ErrorHandleUtil.bas
- ../Excel_Steps/src/ErrorMeta.cls
- ../Excel_Steps/tests/tests_ErrorHandling.bas

## Data I/O Descriptions
### Input Data
- Source worksheet: Errors_ table in errs.wkbkE.
- Current lookup metadata fields (logical):
	- Report code (Long)
	- Routine name (String)
	- Message text (String)
	- IsUserFacing flag (Boolean-like)

### Output Data
- ErrorHandling.ErrMsg (String) assembled from lookup metadata plus optional ErrParam.
- ErrorHandling.IsUserFacing (Boolean) set based on lookup metadata for root error.
- ErrorHandling.IsNewErr toggled False after root message is materialized.

### Behavioral Output
- User-facing mode: concise message, no stack trace suffixes.
- Developer mode: includes error code context and nested "Called by" tracing.

## Project Architecture
### Existing Class to Modify
- ErrorHandling.cls
	- Refactor AppendErrMsg orchestration into smaller single-action methods.
	- Keep RecordErr control flow and reporting gates consistent with current IsShowMsgs/IsTesting model.

### New Class (Proposed)
- ErrorMeta.cls (name candidate; alternatives below)
	- Responsibility: typed container for one lookup result plus validation and formatting helpers.

Proposed ErrorMeta properties:
1. Code As Long
2. Routine As String
3. Message As String
4. IsUserFacing As Boolean
5. IsFound As Boolean
6. IsBaseMessage As Boolean

Proposed ErrorMeta methods:
1. LoadFromLookup(errsObj) As Boolean
2. Validate() As Boolean
3. ToUserMessage(errParam As String) As String
4. ToDeveloperMessage(errCodeReport As Long, errParam As String) As String

## Test Architecture
No direct ErrorHandling unit-test module exists today. This change should add one.

Proposed test module:
- tests_ErrorHandling.bas (new)

Proposed Procedures.cls grouping:
- procs.ErrorHandling

`.github/skills/create_new_test_procedure.md` describes is a guide for how to populate tests_ErrorHandling.bas with the module driver sub and when adding the new Procedure attribute/instancing flow in `Procedures.cls`.

Initial test focus (high level):
1. Root error message generation (user-facing vs developer-facing).
2. Missing-code fallback behavior.
3. Nested trace append behavior for non-user-facing errors.
4. Validation behavior when lookup metadata is incomplete/invalid.
5. Reporting gate interaction with IsShowMsgs and IsTesting.

## Discussion: Scope and Risk
### Refactor Focus
Minimize risk by preserving current modular boundaries:
1. Keep RecordErr responsibilities unchanged (location/state/update and report gating).
2. Keep AppendErrMsg as the orchestration point for message construction.
3. Isolate new typed behavior inside ErrorMeta so lookup typing/validation is encapsulated.

### In Scope
1. Introduce ErrorMeta and replace magic-index metadata reads in AppendErrMsg.
2. Preserve current external behavior of RecordErr and ReportMsg.
3. Add targeted tests for AppendErrMsg root and nested paths.

### Out of Scope
1. Broad redesign of ErrorHandling class lifecycle.
2. Changes to Errors_ table schema or lookup source architecture.
3. Any non-essential behavior changes outside message construction.
4. Changes to current user-facing/developer-facing outputs

## Testing Considerations
High-level testing strategy per [[VBA Project Architecture]]:
1. Unit tests for ErrorMeta load/validate/message methods.
2. Unit tests for refactored ErrorHandling.AppendErrMsg orchestration.
3. Integration-style tests from RecordErr root and nested call paths.
4. Explicit tests for IsShowMsgs/IsTesting reporting gate outcomes.

Cross-workbook setup notes:
1. Ensure ErrorHandling tests initialize errs with valid wkbkE containing Errors_.
2. Use helper setup routine in tests_ErrorHandling for repeated initialization patterns.

## Procedure Outline
### ErrorHandling.AppendErrMsg refactor
1. Look up error metadata using ErrorHandling.Locn and iCodeLocal (root-error path only).
2. Validate lookup metadata.
3. Build message via user-facing or developer-facing formatter.
4. Append to ErrMsg and update IsNewErr.
5. Handle nested-call trace append for non-user-facing path.

### New helper procedure/method candidates
1. ErrorMeta.LoadFromLookup
2. ErrorMeta.Validate
3. ErrorMeta.ToUserMessage
4. ErrorMeta.ToDeveloperMessage
5. ErrorHandling.AppendNestedTraceIfNeeded (or similar)

### Procedure outline (plain language)
1. In RecordErr, set Locn and iCodeReport as currently done.
2. In AppendErrMsg for a new error, call ErrorMeta.LoadFromLookup to read Errors_ table metadata from Locn and iCodeLocal.
	- If lookup does not find a matching row, handle this immediately in lookup with a not-found message (do not defer to later sentinel checks).
3. Validate the metadata and choose message format based on IsUserFacing.
4. Append the resulting message text to ErrMsg and set IsNewErr = False.
5. For nested calls, append "Called by ..." only for non-user-facing errors.

## Design Decisions
1. ErrorMeta will store typed fields only. Do not retain the raw lookup array as an ErrorMeta attribute.
2. Validation behavior for malformed metadata:
	- If a row is found for the code but required fields are blank or malformed (for example invalid Boolean), set error text to "Malformed Errors_ Row for Locn".
	- Use this malformed-row message instead of the current generic code-not-found fallback for this case.
3. Code-not-found handling will occur in the lookup path (ErrorMeta.LoadFromLookup or helper it calls), not in AppendErrMsg.
4. Retire the current iErrNotFound sentinel pattern for this flow; use explicit found/not-found outcomes from lookup instead.
## Design Decisions
1. ErrorMeta will store typed fields only. Do not retain the raw lookup array as an ErrorMeta attribute.
2. Validation behavior for malformed metadata:
	- If a row is found for the code but required fields are blank or malformed (for example invalid Boolean), set error text to "Malformed Errors_ Row for Locn".
	- Use this malformed-row message instead of the current generic code-not-found fallback for this case.
3. Code-not-found handling will occur in the lookup path (ErrorMeta.LoadFromLookup or helper it calls), not in AppendErrMsg.
4. Retire the current `iErrNotFound` sentinel pattern for this flow; use explicit found/not-found outcomes from lookup instead.

## Additional Refactor Opportunities (3/12/26)
### High-level Refactor Direction
Refactor `RecordErr` into the single orchestration procedure for error-message lifecycle and reporting, while keeping the routine name unchanged for backward compatibility.

Under this direction:
1. Treat current `AppendErrMsg` logic as implementation detail of `RecordErr` and fold root/nested message branching into one clear flow.
2. Keep `RecordErr` as a Sub-based API surface used by existing callers; do not rename.
3. Move metadata lookup ownership to `ErrorMeta` (including current `aryErrLookup` responsibilities).
4. Replace sentinel-based not-found handling (`iErrNotFound` checks in message assembly) with explicit `meta.IsFound` outcomes.
5. Reduce helper overhead by removing now-redundant lookup plumbing once `ErrorMeta` owns direct table lookup.
6. Normalize early-exit control flow for nested/non-user-facing append behavior to reduce branching depth.

### Intended End State
1. `RecordErr` computes call context, resolves root-vs-nested path, builds/extends `ErrMsg`, and conditionally reports.
2. `ErrorMeta` handles lookup, validation, and message-shape inputs.
3. `errs.IsUserFacing` is retained only if required for cross-method state; otherwise, use local message-mode state derived from `meta` in root flow.
4. Existing user-visible and developer-visible message wording remains stable.

### Candidate `RecordErr` Flow (clean outline)
1. Determine driver call and set `CallingFunction = False` for Boolean callers.
2. Set `errs.Locn`.
3. If nested error:
	- append `Called by ...` only for non-user-facing mode;
	- skip root lookup/build steps.
4. If new root error:
	- resolve base/report code;
	- instance/load/validate `ErrorMeta`;
	- apply not-found/malformed handling from `meta` state;
	- append formatted user/developer message;
	- set `.IsNewErr = False`.
5. Apply reporting gate (`IsShowMsgs` and driver/testing context) and call `ReportMsg` when enabled.
