## Purpose
* Refactor SetErrs for streamlined code and for logical usage of flags to control message display in production and testing usages
* Consolidate ErrorHandling root and nested message flow and separate Errors_ table lookup activities into new `ErrorsMeta` helper class: keep `RecordErr` as the stable entry point, centralize root message assembly, and preserve current user-facing/developer-facing behavior.

## Background

This change follows [[Projects/VBA_Development/VBA Project Changes]] sequencing and is scoped to planning-level architecture before coding. The refactor starts with call-context behavior, then stabilizes Errors_ metadata inputs, then consolidates ErrorHandling flows around a typed lookup helper.

### High-level Refactor Flow
1. **SetErrs behavior alignment (first):**
   Update `ErrorHandleUtil.SetErrs` to match [[ErrorHandling Class]] expectations for `IsTesting` and `IsShowMsgs` defaults based on `CallingFunction` mode:
   - driver call (`"driver"`)
   - non-Boolean/function-name call
   - Boolean call result path

2. **Errors_ table schema cleanup (second):**
   Simplify and standardize lookup inputs before changing lookup logic:
   - remove unused legacy `sVal` column
   - normalize header list to: `iCodeReport,Module,Routine,Message,IsUserFacing,VBAProject`
   - keep table semantics consistent with existing reporting behavior

3. **ErrorHandling flow consolidation (third):**
   * Refactor `errs.RecordErr`, `.ReportWarningMsg`, and `.LookupCommentMsg` so root and nested paths use a consistent metadata-driven decision flow while preserving current user-facing/developer-facing output contracts.
   * Create ErrorsMeta helper extraction
	   * Move Errors_ lookups and lookup-state flags into a dedicated helper object and remove sentinel-driven branching (`iErrNotFound`) from calling logic. Use object flag attributes instead
	   * Add detection of malformed metadata lookup based on expected metadata types/syntax

### Planned ErrorsMeta State and Responsibilities
Use typed fields/flags so consumers branch on state rather than magic values:
1. `Code As Long`
2. `Routine As String`
3. `Message As String`
4. `IsUserFacing As Boolean`
5. `IsNotFound As Boolean`
6. `IsMalformed As Boolean`

For implementation planning, `ErrorsMeta` should:
- resolve base row by `Locn` and compute report code
- resolve final row by computed `errs.iCodeReport`
- distinguish base-not-found vs code-not-found
- detect malformed row data independently
- support all three use cases: `RecordErr`, `ReportWarningMsg`, `LookupCommentMsg`
- normalize warning/comment rows to user-facing behavior when table data is mis-entered

### Table Access Approach
Within helper scope, use a provisioned `tblRowsCols` instance (`tblE`) for `Errors_` with `IsSetColRngs = True` so lookups use `tblRowsCols` column/range attributes instead of local ad hoc range setup.


### Testing (Planning Scope)
Current state: there is no dedicated test coverage for `SetErrs` or the new `ErrorsMeta`-driven ErrorHandling flow.

High-level test direction:
- Add/refresh `tests_ErrorHandling.bas`.
- Use procedure groups `procs.ErrorHandling` and `procs.ErrorParams`.
- Cover three tiers: `ErrorsMeta` unit tests, `SetErrs` mode tests, and `RecordErr` integration tests.
- Validate key behaviors: found/not-found, malformed metadata, user-facing vs developer-facing branching, and warning/comment normalization.
- Use deterministic mock `Errors_` fixtures; see [[ErrorHandling]] for table setup conventions.

Planning gate:
- Keep this section high level in `Change_`.
- Put detailed test matrix and assertions in `procPlan_`.

Reference documents:
- [[VBA Project Architecture]]
- [[Projects/VBA_Development/VBA Project Changes]]
- ../Excel_Steps/src/ErrorHandling.cls
- ../Excel_Steps/src/ErrorMeta.cls
- ../Excel_Steps/tests/tests_ErrorHandling.bas
- .Github/copilot-instructions.md
- .Github/skills/create_new_test_procedure.md

