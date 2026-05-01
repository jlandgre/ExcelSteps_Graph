- [x] Try convert setup vs code to prompts subfolder file
- [x] Try create skill for vba-tests-new-testing-procedure. Try in VBA Project Starter
- [x] Write change note for colinfo class and ProjFiles class
- [ ] Move ProjFiles to ExcelSteps project
- [ ] Write Background for ColInfo class
- [ ] grill-me for ColInfo
- [ ] Create Change_ note for ColInfo


## Import, Parse and Normalize
Init
* Init needed mdls and tbls
* Read required settings from params and Settings mdls
* initialize colinfo
	* provision tblColInfo
* Initialize tblInfo (should tblInfo be metadata in tbl instance?)
* Initialize files instance (sets pathData based on IsTest)
BuildImportInfo


import has
* dictIndices and dictMetrics (built from arys). These map raw to normalized names
* aryIndex and aryMetrics from colinfo flagged as included
* tblRaw has paramsImport JSON
	* fName based on files attributes
	* fPath
	* ShtEnum (0=first, 1=named, 2=match regex)
	* Optional shtName, regex
* Optional tblParsed (tblRaw parsed to rowscols)
* tblNorm

Init
OpenRaw
Optional ParseRaw
CreateDictColIdx
ValidateParsedStructure

ProjectFiles class
Set pathdata, pathroot pathtests
pfColInfo


## Colinfo class

**Attributes**
* tbl Object set as tblRowsCols instance in Init
* CurTblMetadata String name of current project table for metadata 
* rngRowsCurTbl Range
* aryIndices
* aryMetrics
* dictNormalize 

**Methods**
* Init(files, Optional curTbl)
	* open at files.pfColInfo if specified otherwise sheet (ExcelSteps.shtColInfo Const = “colinfo_”) in project workbook
	* Optional call .SetCurTbl if specified
- SetCurTbl. Set .CurTbl based on arg input. Sorts .tbl by that Col and sets .rngRowsTbl to be rows with tbl_Name column value = .CurTbl
	* Sort by .curTbl column, Index_order column, Metric_order column
	* SetRngRowsTbl
* YieldAryIndices. Create ary or dict with indices for .curTbl
	* Could create a dictionary with integer order as keys and column names as values?
	* ary inherently has order and can be iterated 
	* Factory function so ary acts as pseudo attribute 
* YieldAryMetrics
	* Same as YieldAryIndex for VarName_Norm
* YieldDictNormalize
	* Dictionary of all VarName_Norm keys and VarName_Raw values


## ImportAndNorm Class
See Python_Modeling_Toolbox code as reference

**Attributes**
tblRaw tblRowsCols if raw data are rowscols
tblParsed
dImportParams
dParseParams

**ImportNormalize Procedure**
* InitImportAndNorm
	* Initializes colinfo if Nothing 
	* Sets colInfo.CurTbl
* OpenRawData
	* Set up for Excel workbook for now but file format can vary
	* Set .tblRaw
* ValidateRawStructure
* CreateTblNorm
	* Write normalized rows
* NormalizeVals
	* ColInfo can contain fill missing or Val substitution instructions for any normalized col that needs
	* `<blank>` key for fill missing with specified 




___

Claude code skills: https://code.claude.com/docs/en/skills

Copilot-instructions.md
> Good content for this file:

- preferred architecture patterns
- naming/style conventions
- testing standards
- error/logging expectations
- dependency preferences
- “do not do this” rules
- how to present answers and diffs

Chat says good file arrangement:
```
.github/
  copilot-instructions.md
  instructions/
    frontend.instructions.md
    backend.instructions.md
    testing.instructions.md
    migrations.instructions.md
  prompts/
    generate-tests.prompt.md
    review-change.prompt.md
    safe-refactor.prompt.md
```

But chat also says that the skills folder can contain named sub folders with a SKILL.md file plus optional scripts/resources, and Copilot can load the skill when relevant.

```
.github/skills/
  write-tests/
    SKILL.md
    jest-template.ts
    examples.md
  safe-refactor/
    SKILL.md
    checklist.md
```

## SKILLS to Write
Preferred overall architecture 
* classes by major use case
* Classes organized with procedure blocks consisting of a procedure function that calls sub-methods to perform the procedure
* Methods shared across procedures at bottom of class
Preferred driver sub shape
Preferred function shape
Preferred test driver sub shape
Preferred test sub shape
test-class-and-attribute-usage. Use of .Assert and .Update methods, .wkbkTest as convenient way to refer to project workbook being tested