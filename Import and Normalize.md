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
