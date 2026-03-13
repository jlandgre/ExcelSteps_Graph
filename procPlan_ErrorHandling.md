## Procedure Purpose
Refactor ErrorHandling orchestration so RecordErr remains the stable entry point while ErrorsMeta owns lookup-state handling for Errors_ metadata resolution. This plan also defines coordinated updates for ReportWarningMsg and LookupCommentMsg so all error-related use cases share consistent lookup/validation behavior.

## Procedure Detailed Requirements

### Scope and compatibility
1. Keep public entry points callable from existing code without breaking current call patterns:
   - RecordErr(Locn, Optional ByRef CallingFunction)
   - ReportWarningMsg(iCode, Locn, Optional param)
   - LookupCommentMsg(rng, Locn, Optional IsReinitialize = True)
2. Expose orchestration helpers as Public methods so tests workbook code can call and validate them independently.
3. Keep SetErrs-driven call-context behavior as described in Change_ErrorMetaClass:
   - driver mode
   - non-bool mode
   - Boolean function mode
4. Keep report gating semantics based on IsShowMsgs, IsTesting, and driver context.

### Core design shift
1. Move Errors_ lookup and malformed/not-found state logic from array/sentinel branching to ErrorsMeta state flags.
2. Use ErrorsMeta state booleans as error-condition flags:
   - IsBaseNotFound
   - IsCodeNotFound
   - IsMalformed
3. Keep iCode ownership in ErrorHandling (errs.iCodeBase, errs.iCodeLocal, errs.iCodeReport).
4. Allow ErrorsMeta to read/write needed values via errs input where appropriate; do not duplicate persistent code-state ownership inside ErrorsMeta.

### Data object wayfinding requirement
1. ErrorsMeta should use provisioned tblRowsCols instance tblErrs for Errors_:
   - Provision with IsSetColRngs = True
   - Use tblErrs.rngRows and table column range attributes for lookups
2. Avoid ad hoc local range wiring in ErrorHandling where possible.

## Procedure Method/Sub-Procedure Descriptions

### Required Pre-Steps (Must Complete First)
These are prerequisites and should be completed before the main ErrorHandling flow refactor.

1. SetErrs refactor pre-step
- Align SetErrs defaults and caller-context behavior for:
   - driver mode
   - non-bool mode
   - Boolean function mode
- Confirm downstream RecordErr assumptions on IsTesting, IsShowMsgs, and caller Boolean semantics.

2. Errors_ table format cleanup pre-step
- Remove legacy unused sVal column from planning/data contract.
- Standardize Errors_ headings to:
   - iCodeReport
   - Routine
   - Message
   - IsUserFacing
   - VBAProject
- Remove sentinel constant usage for lookup failures and delete `iErrNotFound` from ErrorHandleUtil constants.
- Confirm lookup dependencies and mock fixture setup are aligned to cleaned headings before implementing ErrorsMeta lookup logic.

### RecordErr Refactor Procedure
This is the primary orchestration flow. Helper steps should be implemented as Public methods for direct tests-workbook invocation while preserving the RecordErr signature.

1. RecordErr
- Action: Top-level fatal-error entry point. Owns current and nested message append behavior directly, then applies reporting gate.
- Inputs: Locn, optional ByRef CallingFunction.
- Outputs: updates errs state and optional calling Boolean return.
- Logic steps:
  1. Call SetCallContext to normalize driver/non-driver behavior and set caller Boolean to False where applicable.
  2. Set errs.Locn = Locn for every invocation.

   - Nested branch (errs.IsNewErr = False)
  3. If errs.IsNewErr = False, evaluate nested branch first (special case handling):
     - If errs.IsUserFacing = False, append Called by Locn trace line directly in RecordErr.
     - If errs.IsUserFacing = True, do not append nested trace text.

   4. If errs.IsNewErr = False after nested handling, skip current lookup path and continue to reporting gate evaluation.

   - Current branch (errs.IsNewErr = True)
   5. If errs.IsNewErr = True, enter current assembly path.
   6. Instance ErrorsMeta, call meta.Init(meta, errs), then pass errs by argument into metadata pipeline methods.
  7. Run metadata pipeline in order:
     - call meta.ResolveCodesFromLocn(meta, errs) to resolve base code by Locn on Errors_ and compute errs.iCodeReport
     - call meta.LoadFromLookup(meta, errs) to load row attributes for errs.iCodeReport
     - call meta.Validate(meta, Locn) to normalize malformed and not-found states

   8. If errs.ErrMsg already contains text, append one line break before current text.
   9. Build current message segment for the current error explicitly:
     - base-not-found path
     - report-code-not-found path
     - malformed metadata path
     - user-facing formatted message path
     - developer-facing formatted message path

   10. Append current message segment to errs.ErrMsg (current path only).
  11. Set errs.IsUserFacing from resolved metadata for found/valid rows.

   12. Set errs.IsNewErr = False after current text is appended.
  13. Call ReportIfEnabled to apply IsShowMsgs/driver/testing gate and trigger ReportMsg when enabled.

  Branch intent note:
   - Current means first error event in the active chain (errs.IsNewErr = True).
   - Nested means downstream caller frames after current is established (errs.IsNewErr = False).
   - Nested trace accumulation (Called by ...) applies only when errs.IsUserFacing = False.
   - For errs.IsUserFacing = True, only the current message segment is used; nested trace text is not appended.

2. SetCallContext (public helper)
- Action: Normalize driver/non-driver interpretation and optional caller Boolean handling.
- Inputs: optional CallingFunction.
- Outputs: errs.IsDriver and optional caller return value.
- Logic steps:
  1. If CallingFunction is missing, treat as driver call path.
  2. If Boolean caller argument is present, set caller to False on error path.
  3. Preserve testing/display defaults previously established by SetErrs.

3. ProcessCurrentErrMsg (public helper)
- Action: Build only the current message segment using computed code and ErrorsMeta state.
- Inputs: errs state (Locn, iCodeLocal, ErrParam, existing ErrMsg).
- Outputs: appended current message, IsUserFacing assignment, IsNewErr transition.
- Logic steps:
   1. Instance/load ErrorsMeta, call Init once, and pass errs to each metadata method call.
   2. Call ResolveCodesFromLocn, then LoadFromLookup, then Validate.
   3. Build message content by branch:
     - base/code not found
     - malformed metadata
     - user-facing message
     - developer-facing message
   4. Append current message segment to errs.ErrMsg with prior-message line-break handling.
   5. Set IsNewErr = False.

4. ReportIfEnabled (public helper)
- Action: Apply reporting gate rules and trigger ReportMsg.
- Inputs: errs flags (IsShowMsgs, IsDriver, IsTesting).
- Outputs: conditional call to ReportMsg.
- Logic steps:
  1. Evaluate gate exactly once after append.
  2. Preserve current behavior for driver and testing contexts.

### ErrorsMeta Integration Methods
These methods define lookup responsibilities shared by RecordErr, ReportWarningMsg, and LookupCommentMsg.

1. Init(meta, errs)
- Action: Initialize ErrorsMeta state and provision meta.tblErrs for Errors_ lookups.
- Inputs: meta instance, errs instance (for workbook/context access).
- Outputs: meta.tblErrs as provisioned Object and reset lookup-state flags.
- Logic steps:
   1. Instance `meta.tblErrs As Object` (tblRowsCols).
   2. Provision with: `Provision meta.tblErrs, errs.wkbkE, False, sht:=shtErrors, IsSetColRngs:=True`.
   3. Initialize/clear per-lookup state fields (Code, Routine, Message, IsUserFacing, IsBaseNotFound, IsCodeNotFound, IsMalformed).
   4. Return success/failure for caller gating.

2. ResolveCodesFromLocn(meta, errs)
- Action: Resolve errs.iCodeBase by Locn base-row lookup and compute errs.iCodeReport.
- Inputs: meta instance, errs instance (Locn and iCodeLocal context).
- Outputs: updated errs.iCodeBase and errs.iCodeReport, plus meta.IsBaseNotFound for base-row resolution.
- Logic steps:
    1. Lookup base row by Locn and Base marker using meta.tblErrs.
   2. Set meta.IsBaseNotFound = True when base row is missing.
    3. If base row is found, compute errs.iCodeReport = errs.iCodeBase + errs.iCodeLocal.

3. LoadFromLookup(meta, errs)
- Action: Resolve Errors_ row state for current lookup context.
- Inputs: meta instance, errs instance (including iCodeReport/Locn context).
- Outputs: meta fields set for code/routine/message/user-facing and meta.IsCodeNotFound status.
- Logic steps:
   1. Find row by target lookup key (RecordErr path: errs.iCodeReport) using meta.tblErrs.
   2. Set meta.IsCodeNotFound = True when no row is found.
   3. Map row values into typed fields when row exists.

4. Validate(meta, Locn)
- Action: Detect malformed metadata and normalize fallback state.
- Inputs: meta, Locn.
- Outputs: meta.IsMalformed and normalized message/visibility behavior.
- Logic steps:
   1. If meta.IsBaseNotFound Or meta.IsCodeNotFound, return without malformed mutation.
  2. Check required fields and expected type syntax.
  3. On malformed row, set normalized fallback message and safe visibility behavior.

5. ToUserMessage / ToDeveloperMessage
- Action: Format message payload for caller append.
- Inputs: meta and context values (code, ErrParam, Locn as needed).
- Outputs: final message string.
- Logic steps:
  1. Keep existing wording contracts unless malformed/not-found fallback applies.
  2. Preserve behavior differences between user-facing and developer-facing outputs.

### Coordinated Non-Fatal Use Cases

1. ReportWarningMsg refactor path
- Action: Keep public API; switch lookup path to ErrorsMeta.
- Inputs: iCode, Locn, optional param.
- Outputs: warning text append/report and reinitialize behavior.
- Logic steps:
  1. Set warning context on errs (Locn, iCodeLocal, optional param).
  2. Instance ErrorsMeta and call meta.Init(meta, errs).
  3. Run metadata pipeline in order:
     - call meta.ResolveCodesFromLocn(meta, errs)
     - call meta.LoadFromLookup(meta, errs)
     - call meta.Validate(meta, Locn)
  4. Build warning message branch explicitly:
     - base-not-found fallback warning text
     - code-not-found fallback warning text
     - malformed metadata fallback warning text
     - found-row warning message text
  5. Apply warning-specific visibility rule:
     - force user-facing output for warning path even if row flag is invalid/mis-entered
  6. Append final warning message to Msgs_accum.
  7. If IsShowMsgs = True, report warning message.
  8. Reinitialize errs per existing warning contract.

2. LookupCommentMsg refactor path
- Action: Keep public API; use ErrorsMeta for message lookup.
- Inputs: rng, Locn, optional IsReinitialize.
- Outputs: comment text write and optional reinitialize.
- Logic steps:
  1. Set comment context on errs (Locn and current iCodeLocal context).
  2. Instance ErrorsMeta and call meta.Init(meta, errs).
  3. Run metadata pipeline in order:
     - call meta.ResolveCodesFromLocn(meta, errs)
     - call meta.LoadFromLookup(meta, errs)
     - call meta.Validate(meta, Locn)
  4. Build comment text branch explicitly:
     - base-not-found fallback comment text
     - code-not-found fallback comment text
     - malformed metadata fallback comment text
     - found-row comment text
  5. Apply comment-specific visibility rule:
     - force user-facing output for comment path even if row flag is invalid/mis-entered
  6. Write final comment text to target range.
  7. Reinitialize errs only when IsReinitialize = True.

### SetErrs alignment dependency
Although SetErrs is outside this procedure file, RecordErr flow depends on its finalized defaults. Implementation/testing should treat this as an upstream prerequisite.

### Method Impact Map (Current ErrorHandling.cls)
These lists define planned ownership/impact for methods currently in ErrorHandling.cls.

1. Refactor in ErrorHandling (remain in class)
   - Init
   - RecordErr
   - ReportWarningMsg
   - LookupCommentMsg

2. Delete from ErrorHandling
   - AppendErrMsg
   - aryErrLookup
   - GetBaseErrCode
   - SetTblELocations

3. Delete from ErrorHandleUtil
   - iErrNotFound constant

4. Move lookup responsibilities to ErrorsMeta
   - GetBaseErrCode responsibilities -> ResolveCodesFromLocn
   - aryErrLookup responsibilities -> LoadFromLookup plus ToUserMessage and ToDeveloperMessage
   - SetTblELocations responsibilities -> Init (provision meta.tblErrs)

5. Out of current refactor scope (no planned changes)
   - ReportMsg
   - UpdateMsgsAccum
   - IsFail
   - ShowMessage
   - setMsgTitleAndText
   - ResetWarningsAndErrors
   - ParseNow
   - WriteErrorsToFile

6. Compile safety check
   - Before deleting ErrorHandling lookup helpers, confirm out-of-scope methods (especially ShowMessage) are either migrated to ErrorsMeta or kept working through temporary compatibility wrappers.
   - Confirm no remaining references to iErrNotFound after sentinel removal.

## Testing Requirements

### Test module and procedure groups
Skill hook:
- Use `.Github/skills/create_new_test_procedure.md` when adding tests to `tests_ErrorHandling.bas` and when adding/updating `Procedures.cls` groups so test architecture and setup conventions are applied consistently.

1. Module: tests_ErrorHandling.bas.
2. Procedure groups in Procedures.cls:
   - procs.ErrorHandling
   - procs.ErrorParams

### Test setup requirements
1. Initialize errs against tests workbook Errors_ table.
2. Use mock Errors_ fixture setup per ErrorHandling note guidance.
3. Keep IsShowMsgs disabled for automated runs unless specifically validating display gate behavior.

### Coverage matrix
1. RecordErr current path:
   - found user-facing row
   - found developer-facing row
   - base not found
   - report code not found
   - malformed metadata
2. RecordErr nested path:
   - user-facing suppresses Called by trace
   - developer-facing appends Called by trace
3. ReportWarningMsg:
   - standard lookup
   - not-found fallback
   - malformed normalization
   - forced user-facing behavior when row flag is invalid
4. LookupCommentMsg:
   - found comment write
   - not-found fallback comment
   - malformed fallback comment
   - IsReinitialize True/False behavior
5. SetErrs integration behavior:
   - driver mode defaults
   - non-bool mode defaults
   - Boolean mode defaults
   - caller Boolean reset semantics via RecordErr

### Success criteria
1. Existing public call sites remain compatible.
2. Message behavior remains stable except where explicitly normalized for malformed/not-found handling.
3. No sentinel-driven branching is required in ErrorHandling current logic.
4. Tests pass for both procedure groups.

### Open design checks for ProjectOwner confirmation
1. Confirm whether non-fatal use cases always force IsUserFacing = True when row data is invalid, or only for specific malformed conditions.
2. Confirm fallback wording contracts for:
   - base-not-found
   - code-not-found
   - malformed metadata
3. Confirm final method names for public helper methods if different from this draft.
4. Confirm strategy for out-of-scope methods that currently depend on deleted lookup helpers: migrate now or retain temporary wrappers.
