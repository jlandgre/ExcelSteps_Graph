Version 3/13/26 J.D. Landgrebe/Data Delve LLC

ErrorHandling is instanced as global `errs` (from ExcelSteps ErrorHandleUtil), so one shared error state is available to project and test workbooks that reference ExcelSteps. Message text is not hard-coded in procedures; it is looked up from the hidden `Errors_` table in `errs.wkbkE`. The code dates to 2021 origination and is under an MIT license.

`SetErrs` in `ExcelSteps.ErrorHandleUtil` module is the required entry point before using `errs`. It initializes `errs` when needed, configures default mode flags by call context, and optionally initializes a Boolean function return value to `True`.
* `SetErrs "driver"`: use at the top of user-initiated driver subs. This creates or resets context for a full procedure run and enables normal reporting behavior.
* `SetErrs "non-bool"`: use in non-Boolean procedures that participate in the same flow and need access to `errs` state.
* `SetErrs <BooleanFunctionReturn>`: use in Boolean functions. On error, `errs.RecordErr` flips that return value to `False` signifying unsuccessful executon

There are three primary ErrorHandling use cases:
* `errs.RecordErr`: fatal error path for driver + nested Boolean architecture. It records code location context as the sub or function name as a String. It builds and appends root or nested trace messaging. It provides caller-failure signaling.
* `errs.ReportWarningMsg`: non-fatal warning/info path. It resolves warning text from `Errors_`, formats optional parameters, reports to message accumulation (and optionally MsgBox), then reinitializes for next event.
* `errs.LookupCommentMsg`: comment path for worksheet annotations. It resolves message text from `Errors_` (or a fallback failure message) and writes that text as a cell comment.

In practice, a top-level driver calls `SetErrs "driver"` once, then nested procedures/functions reuse that same `errs` instance. In tests, helper setup typically calls `SetErrs` with test workbook context so lookups target the test `Errors_` sheet.

`SetErrs` sets `errs.IsHandle` when initializing. This is on/off switch for error handling in the project. If `False`, user-facing errors are not reported and VBA errors are handled by VBA not by `errs.RecordErr` developer-facing logic.

#### Use Case Options Describing SetErrs utility actions
`errs.IsTesting` and `errs.ShowMsgs` control display of messages. If `SetErrs` is called directly from a Boolean function, test flow is assumed (`IsTesting=True, ShowMsgs=False)`. If called from a driver, the opposite is set by default. `.ShowMsgs` provides a way to override the behavior after `SetErrs` returns.  The list shows the permutations.

**Driver normal**
- First `SetErrs` call: `"driver"`
- `errs` before call: `Nothing`
- Defaults: `IsTesting=False`, `IsShowMsgs=True`
- Expected behavior: show messages from called Boolean functions

**Driver silent**
- First `SetErrs` call: `"driver"`, then manually set `.IsShowMsgs=False`
- `errs` before call: `Nothing`
- Defaults: `IsTesting=False`, `IsShowMsgs=True` (overridden after return)
- Expected behavior: suppress message display from called Boolean functions

**Nested Boolean function**
- First `SetErrs` call: inherited from driver flow (`"driver"` at top level)
- `errs` before call: existing driver context
- Defaults: inherited from driver flow
- Expected behavior: `.RecordErr` sets Boolean return `False`; driver reports as configured

**Default test flow (Boolean)**
- First `SetErrs` call: Boolean function calls `SetErrs`
- `errs` before call: `Nothing`
- Defaults: `IsTesting=True`, `IsShowMsgs=False`
- Expected behavior: no message display

**Demo test flow (Boolean)**
- First `SetErrs` call: pre-init (`"non-bool"` or helper), then set `.IsShowMsgs=True`, then Boolean call
- `errs` before call: `Nothing` on pre-init
- Defaults: `IsTesting=True`, `IsShowMsgs=False` (overridden to `True`)
- Expected behavior: show messages intentionally

**Default test flow (non-bool)**
- First `SetErrs` call: `"non-bool"`
- `errs` before call: `Nothing`
- Defaults: `IsTesting=True`, `IsShowMsgs=False`
- Expected behavior: no message display

**Demo test flow (non-bool)**
- First `SetErrs` call: `"non-bool"`, then set `.IsShowMsgs=True`
- `errs` before call: `Nothing`
- Defaults: `IsTesting=True`, `IsShowMsgs=False` (overridden to `True`)
- Expected behavior: show messages intentionally

#### Error Code Indices
`errs.iCodeLocal` flags a specific error condition within a routine, typically through `errs.IsFail(..., iCodeLocal, ...)`. In Boolean routines, that error path diverts to `ErrorExit`, where `errs.RecordErr` performs message resolution and reporting.

Each routine has one Base row in `Errors_` (`errs.iCodeBase`). The reported lookup code is:
`iCodeReport = errs.iCodeBase + errs.iCodeLocal`

`errs.iCodeLocal = 0` represents unknown/VBA errors, which resolve to the routine's Base row. Base rows should be developer-facing (`iMsgDevUser = FALSE`).

#### Errors_ Table Format
The table below shows the `Errors_` column layout and array indices.
* Base rows are used to locate routine-level message groups by `Routine` lookup.
* `Routine` values for Base rows should be unique across the project.

Table showing Excel Column letters (array indices)

| A (0) | B (1)        | C (2)           | D (3)                                                         | E (4) | F (5)       | G (6)      |
| ----- | ------------ | --------------- | ------------------------------------------------------------- | ----- | ----------- | ---------- |
| iCode | Module       | Routine         | sMsg_String                                                   | sVal  | iMsgDevUser | VBAProject |
| 100   | modInterface | RefreshRCDriver | Base                                                          |       | FALSE       | ExcelSteps |
| 104   | modInterface | RefreshRCDriver | VBA error handling off. Set errs.IsErrorHandle True to enable |       | TRUE        | ExcelSteps |
| 110   | modInterface | RefreshAPI      | Base                                                          |       | FALSE       | ExcelSteps |
| 120   | modInterface | Auto_Open       | Base                                                          |       | FALSE       | ExcelSteps |
| 130   | modInterface | RefreshDriver   | Base                                                          |       | FALSE       | ExcelSteps |

#### `errs` Setup for Testing
- For testing, to avoid using the project's production `Errors_` table, create a mock `Errors_` table in the tests workbook. Generally, such 
- Per `.Github/copilot-instructions.md` tests, destroy `errs` as an initial step to avoid carryover across tests. nitialize `errs` with `wkbkE = tst.wkbkTest` so lookups resolve against the mockup in the tests workbook.
- To initialize `errs`, call `SetErrs` with a dummy non-driver `False` CallingFunction argument and pass `tst.wkbkTest` to set `errs.wkbkE`
- Populate `Errors_` test data through `Populate_Errors_Default` sub residing in a Populate_ module not in a tests_ module.in `Populate_Errs.bas`. Clear the sheet first and write first row headers from `sErrorsHeadings` constant defined in `ExcelSteps.ErrorHandleUtil` module.

#### Example mock Errors_ table

| iCode | Module | Routine  | sMsg_String       | sVal | iMsgDevUser | VBAProject  |
| ----- | ------ | -------- | ----------------- | ---- | ----------- | ----------- |
| 2000  | Mockup | TestProc | Base              |      | FALSE       | ProjExample |
| 2001  | Mockup | TestProc | User visible:     |      | TRUE        | ProjExample |
| 2002  | Mockup | TestProc | Developer detail: |      | FALSE       | ProjExample |
| 3000  | Mockup | BadProc  | Base              |      | FALSE       | ProjExample |
| 3001  | Mockup |          |                   |      | maybe       | ProjExample |
| 4000  | Mockup | UserProc | Base              |      | FALSE       | ProjExample |
| 4001  | Mockup | UserProc | User visible:     |      | TRUE        | ProjExample |

- Row `3001` is intentionally malformed for validation testing.
- `TestProc` with `iCodeLocal = 1` resolves to `iCodeReport` `2001`; with `iCodeLocal = 2` resolves to `2002`.

**Example tests_ Module Fixture and Populate Module Populate_Errs_Default sub**
```
'--------------------------------------------------------------------------------------
' Initialize errs and populate Errors_ fixture rows used by ErrorHandling tests
' JDL 3/11/26
'

Sub SetupErrorsFixture(tst)
    ExcelSteps.SetErrs False, tst.wkbkTest
    
    'Set as False anyway by SetErrs but provides easy toggle
    ExcelSteps.errs.IsShowMsgs = False

    Populate_Errs_Default
End Sub
```

```
'--------------------------------------------------------------------------------------
' Populate mock Errors_ table for ErrorHandling tests
' JDL 3/11/26
'
Sub Populate_Errs_Default()
	Dim wksht As Worksheet

	Set wksht = ExcelSteps.errs.wkbkE.Sheets(ExcelSteps.shtErrors)

	With wksht
		.Cells.Clear

		'Set canonical mock Errors_ headers from project constant
		.Range("A1:G1").Value = Split(ExcelSteps.sErrorsHeadings, ",")

		'Base rows are always developer-facing (iMsgDevUser=False).
		'Base row is used when code resolves to base (e.g., unknown VBA error fallback).
		'For normal flow, iCodeReport is Base + iCodeLocal where iCodeLocal comes from IsFail.

		'Base code row for TestProc (developer-facing)
		.Cells(2, 1).Value = 2000
		.Cells(2, 3).Value = "TestProc"
		.Cells(2, 4).Value = ExcelSteps.sErrBase
		.Cells(2, 6).Value = False

		'User-facing row for TestProc (Base 2000 + iCodeLocal 1 = 2001)
		.Cells(3, 1).Value = 2001
		.Cells(3, 3).Value = "TestProc"
		.Cells(3, 4).Value = "User visible: "
		.Cells(3, 6).Value = True

		'Developer-facing non-base row for TestProc (Base 2000 + iCodeLocal 2 = 2002)
		.Cells(4, 1).Value = 2002
		.Cells(4, 3).Value = "TestProc"
		.Cells(4, 4).Value = "Developer detail: "
		.Cells(4, 6).Value = False

		'Base code row for malformed case (developer-facing)
		.Cells(5, 1).Value = 3000
		.Cells(5, 3).Value = "BadProc"
		.Cells(5, 4).Value = ExcelSteps.sErrBase
		.Cells(5, 6).Value = False

		'Malformed row: blank routine and message plus non-Boolean user flag
		.Cells(6, 1).Value = 3001
		.Cells(6, 3).Value = ""
		.Cells(6, 4).Value = ""
		.Cells(6, 6).Value = "maybe"

		'Base code row for UserProc (developer-facing)
		.Cells(7, 1).Value = 4000
		.Cells(7, 3).Value = "UserProc"
		.Cells(7, 4).Value = ExcelSteps.sErrBase
		.Cells(7, 6).Value = False

		'User-facing row
		.Cells(8, 1).Value = 4001
		.Cells(8, 3).Value = "UserProc"
		.Cells(8, 4).Value = "User visible: "
		.Cells(8, 6).Value = True
	End With
End Sub

```

