## Purpose
Refactor `ErrorHandling.AppendErrMsg` to simplify it and move to object-oriented structure; Initialize testing for ErrorHandling

## Background
Introduce a typed Error record class as a helper for current `ErrorHandling` class This consists of creating class `ErrorMeta` to be called from a refactored `ErrorHandling.AppendErrMsg` function.

We do not currently have testing for ErrorHandling.cls. With this refactoring, we should add a new 


**Chat suggestions**
(move to appropriate Change_ and procPlan_ sections based on keeping Change_ as higher level description and procPlan_ as details for eventual input into code_plan.csv [[VBA Project Changes]])
Suggested fields:

1. `Code As Long`
2. `Routine As String`
3. `Message As String`
4. `IsUserFacing As Boolean`
5. `IsFound As Boolean`
6. `IsBaseMessage As Boolean`

Suggested methods:

1. `LoadFromLookup(errsObj) As Boolean`
2. `Validate() As Boolean`
3. `ToUserMessage(errParam As String) As String`
4. `ToDeveloperMessage(errCode As Long, errParam As String) As String`

Pros:

1. Removes magic indices (`aryMetaData(2)`, etc.).
2. Centralized validation and defaults.
3. Cleaner `AppendErrMsg` orchestration.

Cons:

1. Moderate code changes across lookup consumers.
2. Needs tests around constructor/load paths.


Reference documents:
- [[VBA Project Architecture]] - Standard VBA project structure
- [[VBA Project Changes]] - Planning mode process
- .github/Skills/create_new_test_procedure.md
