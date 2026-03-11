## Purpose
Define a concise, repeatable process for planning and implementing new code in a VBA project and its test suite
## Overall guidance
*  We refer to human and AI collaborator roles as "ProjectOwner" and "AI"
* We use the data object architecture in `copilot-instructions.md` (`tblRowsCols` + `mdlScenario`).
*  The workspace landscape (`ProjName.code-workspace`) for the project and change descriptions is comprised of a `ProjectName` folder and an Obsidian `Graph_ProjectName` folder containing notes with context. They have distinct GitHub repositories, so both have a `.GitHub` folder in their root; however Skill and and `copilot-instructions.md` are in `ProjectName/.Github` only. If `ExcelSteps` is used in the project, its project folder is also part of the workspace.
* We develop three change-related files to plan what is needed in "Planning Mode" for a code change. They are developed sequentially with hard stage gates, and AI should not proceed into developing subsequent documents until **ProjectOwner** approves.
	* `Change_ChangeName.md`: Initiates code change from **ProjectOwner** description in its "Purpose" section. This document is co-developed in Planning Mode by **ProjectOwner** and **AI**. It is a high-level description.
	* `procPlan_ProcedureMethodName.md` (We develop one such document per procedure. It is linked to the outline in the top level `Change_ChangeName.md`. It describes code details, list of methods and their actions etc. 
	* `Code_plan.csv`: Rows/columns table lists methods (rows), arguments, docstrings, qualitative internal logic and testing considerations (columns). This is developed as the last stage of planning changes but, when updated to "as coded" serves as project documentation. See [[Code_Plan]] for example and background.
	
 \*.md files are in the Obsidian graph folder. The third, code plan CSV is in the project folder's `/docs subfolder`.

### Required Sequence (Do Not Skip)
1. `Change_ChangeName.md` only: develop and iterate this note until **ProjectOwner** explicitly confirms it is finalized.
2. `procPlan_*.md` notes: begin only after step 1 approval; write one `procPlan_` note per procedure from the approved `Change_` outline.
3. `Code_plan.csv`: begin only after the relevant `procPlan_` note(s) are approved.

If `Change_` is still under review or has open questions, AI must remain in `Change_` and must not draft `procPlan_` content yet.

During Planning Mode \*.md development and coding, AI should  Read `Change_ChangeName.md` and follow the key links which include references related to the change and background documentation on `ProjectName` code base.
## VBA Project Architecture
Code changes are constructed per the [[VBA Project Architecture]] for project and test code. Legacy projects may contain non-conforming code structures. In planning mode, AI should alert **ProjectOwner** of these  for a decision on including in scope refactoring code, docstrings and/or architecture. 

## `Change_ChangeName.md` (aka `Change_`) Note Sections
1. **Purpose**: Purpose and high level description of the desired change
2. **Background**: Include links to project architecture and other docs as well as background on specific change to design
3. **Data I/O Descriptions**: Descriptions of data source and output architecture and content (including descriptions of method arguments as appropriate). For input/output data sources, discussion should include descriptions of typical data object key variables and data values and how they get mapped into the project by canonical names or other destination descriptions such as target tblRowsCols or mdlScenario data blocks.
4. **Project Architecture**: Description of proposed project architecture including listing of new or modified classes and their roles
5. **Test Architecture**: Description of where tests will be housed (current or new module) in tests_ProjName.xlsm. Listing of new or current Procedure(s) (Procedures.cls) to be used for tests of new code
6. **Discussion: Topic XYZ**: (optional) additional discussion sections on key topics and design decisions as appropriate based on questions raised requirements input. This should include discussion of options considered for key, architecture and input/output decisions
7. **Testing Considerations**: High-level description of test strategy per [[VBA Project Architecture]]. Include:
   - Test module structure: Which existing test module(s) or new tests_UseCaseX.bas module(s) will house tests
   - `Procedures.cls` attributes: Which `Procedure` attribute(s) will group related tests (e.g., procs.ParseData)
   - Test coverage scope: Default is that all procedures/methods require unit tests of individual methods and integration tests of complete procedures
   - Impact on existing tests: Which current tests may be affected by changes and require updates
   - Test data requirements: Describe test data files needed `test_data_xxx`subfolder
   - Cross-workbook instantiation: List factory functions needed in `ProjectName` `Validation.bas` for new classes
   - Edge cases and validation: High-level description of key boundary conditions to test
8. **Procedure Outline**: Outline of overall by listing of individual methods and sub-procedures. See description of architecture in [[VBA Project Architecture]] and example below.

#### Example Procedure Outline
* Example is for a hypothetical `ParseToNormalizedProcedure`. The Outline section shows a proposed sequential flow and contains links to specific `procPlan_ProcedureMethodName.md` notes
* \[\[xyz\]\] denotes Obsidian link to xyz.md
* Descriptions after "-" are drafts of docstrings for those in method code
* CopyToDestProcedure is hypothetical example of a multistep sub-procedure per [[VBA Project Architecture]]. Hypothetically it is being created with intent to call from/re-use it in multiple top level procedures or because it accomplishes a self-contained action that should be validated as a sub part of developing the overall procedure.

	**ParseToNormalizedProcedure** - parse data xyz to new tblRowsCols instance, tblDest
	* **`ParseToNormalizedProcedure`**:  \[\[procPlan_ParseToNormalizedProcedure\]\]
	* `OpenAndValidateInput` - Open and validate input data
	* `DeleteUnusedRowsAndCols` - Delete initial blank rows and columns surrounding data
	* `CopyToDestProcedure`  \[\[procPlan_CopyToDestProcedure\]\] - Sub-procedure to copy normalized data to dest
	    * `SetSrcDataRange` - Set tblSrc data range for copying
	    * `InstanceTblDest` - Instance the parsed data table in destination location
	    * `CopyParsedToDest` - Copy (array move) src data to tblDest
	* `CleanupParseToNormalized` - Close src data and reprovision tblDest


## `procPlan_ProcedureMethodName.md` (aka `procPlan_`) Note Sections
This Procedure-specific plan repeats and expands on the `Change_ChangeName.md` note's Procedure Outline for an individual procedure.

Prerequisite: Do not create or edit any `procPlan_` note until `Change_ChangeName.md` is finalized and approved by **ProjectOwner**.


1. **Procedure Purpose**: Restate the purpose/action of the procedure (based on the link back to the Change_ note's Procedure outline)
2. **Procedure Detailed Requirements**: Detailed prose description of procedure's actions on its inputs and detailed description of its outputs. Both should be in terms of input arguments including project data objects (either `tbls` and/or `mdls` attribute objects or standalone `tblRowsCols` and/or `mdlScenario` instances created to accomplish the procedure)
3. **Procedure Method/Sub-Procedure Descriptions**: Expansion on listing from the Change_ note's Procedure outline. For each method in the procedure, specify:
   - **Action**: What the method does (prose description matching draft docstring)
   - **Inputs**: Argument names with types and expected states (e.g., "tblSrc As tblRowsCols, must be provisioned with .rngHeader set")
   - **Outputs**: Return type (Boolean for functions) and/or modified objects (e.g., "Sets tblDest.rngRows to normalized data range")
   - **Logic Steps**: Numbered sequence of operations with specific attribute/range references:
     - For data object operations, reference specific attributes (e.g., `.rngHeader`, `.colrngModel`, `.ScenModelLoc()`)
     - For ExcelSteps utilities, specify function names (e.g., `ExcelSteps.rngToExtent`, `ExcelSteps.OpenFile`)
     - For range operations, describe source and destination (e.g., "Set rngSrc to Intersect of .wkshtPivot.Rows(1) and .colRngSrc")
   - **Data Flow**: Which attributes/variables are read from inputs and written to outputs, including intermediate variables
   - **Validation/Error Conditions**: Edge cases requiring validation checks (e.g., "Check tblSrc.ncols >= 3 via errs.IsFail"; "Validate file exists with Len(Dir$(pathFile)) > 0")
   - **ExcelSteps Integration**: Specify which ExcelSteps patterns to use per `copilot-instructions.md` (e.g., use `FindInRange` not native `Find`; use `ScenModelLoc` for model lookups)
   - **Sub-Procedure Links**: For multi-step sub-procedures, include child note links to their `procPlan_` notes (e.g., \[\[procPlan_CopyToDestProcedure\]\])
4. **Testing Requirements**: Specification for testing this procedure per [[VBA Project Architecture]]. Identify which test module (existing tests_UseCaseX.bas or new module) and `Procedure` attribute in `Procedures.cls` will group these tests:
   - **Test Module Location**: Specify module name (e.g., "tests_ParseData.bas") and Procedure group (e.g., "procs.ParseToNormalized")
   - **Test Setup Pattern**: Required data object initialization using factory functions from `Validation.bas`:
     - For project classes: `Set obj = VBAProject_ProjectName.New_ClassName`
     - For ExcelSteps classes: `Set dict = ExcelSteps.New_Dictionary`
     - Specify test data requirements (e.g., "Instance tblSrc with 5 columns named A-E; populate 10 rows with pattern X")
   - **Success Criteria**: Assertions validating correct behavior (e.g., "tblDest.ncols = 5", "tblDest.rngRows.Count = 10", "First row value matches expected")
   - **Edge Cases to Test**: Boundary conditions and error paths (e.g., "Empty tblSrc", "Missing required columns", "File not found", "Invalid data types")
   - **Integration/File Tests**: If procedure interacts with external files, specify test data files in `test_data_import/` subfolder and use `tst.wkbkTest.Path` pattern for paths
   - **Helper Subroutines**: If multiple tests share initialization patterns, specify helper subs with `tst` and data objects as arguments