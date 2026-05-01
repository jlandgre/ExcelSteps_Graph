## Purpose
Create and validate ColInfo class for storing and utilizing metadata about project data objects (tblRowsCols and mdlScenario). 

## Background

This change should follow [[VBA Project Changes]] sequencing and is scoped to planning-level architecture before coding.

Project code will instance `ColInfo` as `colinfo` and uses its attributes for taking actions within the project such as importing and normalizing raw data. `ColInfo` is decoupled from any specific project — It references the project global `IsTest` to toggle between test and production modes (always declared `Public Boolean` in project workbook, default `False`) directly rather than receiving it as an argument.

ColInfo is based on use of a project-specific metadata table (see [[Example_ColInfo]]) with default name ColInfo.xlsx that contains rows for normalized variable names in one or more project tables. Its location is specified by `files.pfColInfo` within `files` instance of `ProjFiles`.

Within `colinfo.tbl` (aka `tblRowsCols` instance for the ColInfo.xlsx metadata table), Inclusion in a project table, `tblExample` is denoted by an integer in the `colinfo.tbl` `tblExample` column. The integers specify column order in the normalized data. The table specifies renaming for normalization by the mapping between `VarNameNorm` and `VarNameRaw` columns.

Project tables contain one or more "index" or key columns that are identified by a True toggle value in the `colinfo.tbl` `IsIndex` column. These are analogous to indices in Pandas DataFrames --useful for subsetting, sorting etc. later.

Because `colinfo.tbl` can contain metadata for multiple project tables or models, usage for a particular table, `colInfo.curTbl` implies sorting the colinfo tbl by the `.curTbl` column and setting a `.rngRowsCurTbl` row range for rows included in the table (e.g. non-blank `VarNameNorm` column value).  

ExcelSteps contains RecipeStep class with SortBy function usable for sorting on one or more keys (has non-standard error handling where errors are passed to calling routine for handling by errs.RecordErr). This is validated and preferred for sorting tblRowsCols instances

For this and other subsetting purposes, would it be helpful to create a "FindNextChange" utility that sets a cell range for the location of the next change in value (e.g. from empty to non-empty or from value_a to something else) for the purpose of efficiently locating blocks of contiguous values within a column of a sorted table? This seems like it will be more efficient than iterating through rows of unsorted table etc.

We will get to this later, but my thought is to subsequently create a companion TblImport class (instanced as import) to handle raw data import, parsing and normalization based on the `colinfo.tbl` metadata and `dImportParams` and `dParseParams` dictionaries that specify how to import and parse raw data prior to normalization.

## Brainstorming on Class design:

**Attributes**
* tbl Object set as tblRowsCols instance in Init
* CurTbl String name of current project table for metadata 
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
	* Indices identified by `True` value in `IsIndex` column
	* ary preferred because inherently has order and can be iterated 
	* Factory function so ary acts as pseudo attribute 
* YieldAryMetrics 
	* Same as YieldAryIndex for VarName_Norm (exclude `IsIndex=True` variables)
* YieldDNormalize
	* Dictionary of all VarName_Norm keys and VarName_Raw values

**Testing**
Create new ColInfo Procedure and add section to tests_Toolbox module with its driver sub