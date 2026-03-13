## Purpose
* Refactor SetErrs for streamlined code and for logical usage of flags to control message display in production and testing usages
* Consolidate ErrorHandling root and nested message flow and separate Errors_ table lookup activities into new `ErrorsMeta` helper class: keep `RecordErr` as the stable entry point, centralize root message assembly, and preserve current user-facing/developer-facing behavior.

## Background

Refactor ExcelSteps `ErrorHandleUtil.SetErrs` to match [[ErrorHandling Class]] description of `IsTesting` and `IsShowMsgs` default logic based on 1) "driver", "non-bool" or Boolean `CallingFunction` arg.

Referring to ErrorHandling class by its `errs` instanced name, refactor `errs.RecordErr`, `.ReportWarningMsg` and `.LookupCommentMsg` to simplify flows and make consistent across use cases. 
* Start by simplify Errors_ table columns by eliminating unused sVal column E (headings comma-separate list in `Constants.bas` and example table in [[ErrorHandling Class]]). Rename columns to better match programmatic names: "iCodeReport,Routine,Message,IsUserFacing,VBAProject"
* Create object oriented helper structure for lookups of attributes from Errors_ sheet table and flags for error conditions during lookup. Suggested attributes :
1. `Code As Long`
2. ``Routine As String`
3. `Message As String`
4. `IsUserFacing As Boolean`
5. `IsNotFound As Boolean`
6. `IsMalformed As Boolean`

The three used cases mentioned in the purpose, all involve reading a set of values from a found row in the projects Errors_ sheet table.  The current approach uses a sentinel value (`iErrNotFound`) to flag the case where the row is not found. 
* Current code for fatal errors is essentially a step-by-step procedure, but it follows a meandering path through RecordErr which calls AppendErrMsg which calls other sub functions to look up a Locn base row in Errors_ and compute the ErrorHandling.iCodeReport code used to then look up a specific row’s attributes.
* There are two lookups currently. First is looking up Routine based on Locn arg to get base code and compute `errs.iCodeReport`. The second is lookup of metadata for message to report and user versus developer facing handling flag.
* We should streamline RecordErr into a straightforward procedure that  branches appropriately for errors that are either user facing or developer facing as flagged in the Errors_ table.
* Move Errors_ lookups to a new ErrorsMeta helper class that looks up  ErrorHandling.Locn Base row and computes ErrorHandling.iCodeReport
* Detect not found Base row and malformed row data. Separately detect not found iCodeReport in second lookup
* Use the helper class to manage lookups for fatal error reporting from .RecordErr, warning message reporting by ReportWarningMsg and comment from `LookupCommentMsg`
* Use Boolean flag attributes to seed error message generation for the three use cases (Messages could be consistently generated in the helper class and passed to ErrorHandling attributes for reporting)
* Eliminate use of the not found sentinel value in lieu of ErrorsMeta flag attributes
* Within the helper class convert to instancing a `tblRowsCols`, `tblE`, for the Errors_ table. Instead of setting local rngRows and colrng's for the table as currently, provision, the tblE instance with `IsSetColRngs = True` to take advantage of hard-coded colrng attributes for shtErrors (in `tblRowsCols.SetColRngs`) set colrng’s (see xxx tblRowsCols method that does this for the errors sheet). Provision also automatically sets `tblE.rngRows` attribute needed for lookups. 
* Note that RecordErr for fatal errors needs lookup to determine whether user or developer facing but other use cases are inherently user facing
* Mis-entered Errors_ row data for codes used by `.ReportWarningMsg` and `LookupCommentMsg` should be automatically corrected to UserFacing = True if False in the Errors_ table


Instructions about testing
we currently do not have tests for SetErrs or ErrorHandling class. We should populate a new tests_ErrorHandling module with new procedures procs.ErrorHandling and procs.ErrorParams to cover testing of the new helper class

Reference documents:
- [[VBA Project Architecture]]
- [[VBA Project Changes]]
- ../Excel_Steps/src/ErrorHandling.cls
- ../Excel_Steps/src/ErrorMeta.cls
- ../Excel_Steps/tests/tests_ErrorHandling.bas
- .Github/copilot-instructions.md
- .Github/skills/create_new_test_procedure.md

## Data I/O Descriptions


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
