Draft of Change_Create_ImportParseNorm_Class

Prompt
>Change_Create_ImportParseNorm_Class.md note gives basic, raw background on a coding project to write ImportParseNorm.cls. Use grill-me skill to ask and answer questions allowing you to fully populate the Change_ note based on Change_Create_ColInfo_Class format as a guide. It's ok to begin modifying the Change_ note based on our grill-me back and forth if this helps you retain the context versus it being in our chat


Cleanup prompts after initial coding:
>Review code in light of vba-function-architecture. If functions exceed approx. 50 lines (FillMissingVals is example) you should refactor to move part into a helper method. In that example, the code within For Each sKey would be good candidate to move to helper method so loop interior becomes just a call to the helper. As you refactor, add tests of the helper methods. Review all class code for such opportunities using the skill as a guide

>Super! Also review for unecessary use of local variables. They make sense if there is big brevity advantage where class instance + attribute chain is long and With/End With doesn't make sense, but otherwise, use existing class attributes instead of simply equating them to local vars
## Purpose
Create and validate ImportParseNorm class for importing, parsing and normalizing raw data (instance as )

## Background

This change follows [[VBA Project Changes]] sequencing and is scoped to planning-level architecture before coding.

We can use the **ColInfo** class and dicts of import params and parse params to specify how to import, parse and normalize raw data such as a rows/columns raw data source.  Parsing is needed to rearrange raw data that are not rows/cols already (e.g. unstructured but repeatable raw format), but it is not needed if the raw data shape is already structured
* The dict of import params, **dParamsImport**, can contain all needed items for importing --such as a param describing the shape of data (e.g. RawShape = "rowscols" or "unstructured" where the latter triggers requiring parsing to reshape into rows/columns). 
* **dParamsImport** can also contain an item that specifies type of raw data file such as Excel or CSV.
* For unstructured data that require parsing, **dParamsParse** can describe parsing details such as whether the data come in row-oriented blocks etc.
* **ColInfo.tbl** can contain metadata on how to map raw variable names to normalized names. It can also include params such as a JSON dict-like string describing how to fill blank values; any needed replacement value mappings and filtering requirements such as "IncludeAllExcept" or "IncludeOnly" dict keys to describe which values to retain post-normalization for individual variables
* The previous custom project code example (**`case_studies/2026_0507_ImportCode/BRImport_Example.cls`**) shows import and custom normalization for project data. This code does not take advantage of the **ColInfo.cls** code for specifying raw to normalized column name mapping as a configuration. It is based on hard-coded variable names for example, but it is an example of opening a rows/columns file, validating the data and normalizing the data. The example code also hard codes dealing with date-related variables

Based on the previous, custom example, envision the following multistep procedure functions for an importtbl instance of the ImportParseNorm class:
* ImportParseNormProcedure (overall procedure; calls Init and individual sub-procedures in sequence
* Init(importtbl, colinfo, files, curTbl)
	* instance colinfo if it doesn't already exist
	* colinfo.SetCurTbl
* OpenAndValidateRawProcedure
	* OpenRawData (based on preset files.pfFileImport and dImportParams item to set type of file Excel, CSV, other); instance and provision .tblRaw (tblRowsCols Provision if dParseParams RawShape = rowscols; tblUnstructured Provision otherwise to set wayfinding via .wkbk, .wksht, .sht, .rngTable attributes)
	* ValidateRawStructure (for now, raw data variables in colInfo.rngRowsCurTbl VarNameRaw column are required if VarNameNorm column is populated)
	* FillMissingVals (based on new FillVals colinfo.tbl column that gives JSON-like mappings for variables)
* ParseRawProcedure
	* Optional based on dParseParams RawShape item --no parse needed if RawShape = rowscols. In this case set .tblParsed = .tblRaw and move on to normalize
	* For now, just create dummy procedure. We will validate initially for rowscols then add other data shapes later that require parsing
* NormalizeParsedProcedure
	* WriteNormalized
		* same result as WriteNormalizedRows in example custom code
		* use colinfo mapping for normalization
		* code example does not use an efficient approach because it conditionally writes line-by-line to deal with custom filtering by custom attribute .locationBR
		* Should consider alternate logic that copies entire table to .tblNorm and then selectively deletes unneeded columns
		* Code also should be all-in on use of .tblRaw and .tblNorm attributes like .wksht instead of unnecessarily using local variables like wkshtNorm for these items.
	* FilterRows 
		* Use new colinfo.tbl FilterVals dict to filter based on KeepOnly, KeepExcept, KeepList or KeepExceptList options
		* Validate for KeepOnly Location="Online" example that subsets mockup raw data to just rows where Location column = "Online"); leave other options for future development

* `ValidateRawDataProcedure` in example code represents custom validations that we don't include in new class but which could be developed and called from the driver that runs ImportParseNormProcedure the 
* Need to create `tblUnstructured` class for describing wayfinding for an unstructured table that is described by .wkbk, .wksht, .sht and .rngTable attributes only (e.g. can have an Init as only method that sets these from args and from .wksht.UsedRange for .rngTable)
* Have successfully used dictionaries such as **dParamsImport** and **dParamsParse** for concepts that have a large number of options. In future, could create **TblInfo** to store metadata as configuration for import and parsing of multiple tables in a project
* Our custom Dictionary.cls has ability to read in JSON-like strings to populate dictionaries, and this makes it easy to store as metadata in **colinfo.tbl** or future **tblinfo.tbl**

Testing:
* create new ImportParseNorm Procedure and add section to bottom of TestDriver_ToolBox (vba-testing-create-new-test-procedure skill)
* For ease of navigation, add section for new testing immediately below TestDriver_ToolBox instead of at bottom of tests_ToolboxClasses.bas