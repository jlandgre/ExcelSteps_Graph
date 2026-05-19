
ImportParseNorm.FilterVals only current filtering option is "KeepOnly". We need to add option for KeepExcept e.g. parsed sFilter dict.Exists("KeepExcept") to eliminate an individual or comma-separated list of values from the .tbl.colrngVarNorm value for that rngRow.

Examples .colrngVarNorm is "Location" variable
After checking for dict.Exists("KeepExcept")
dict.Item("KeepExcept") = "BLANK": Should eliminate all rows where Location column is blank
dict.Item("KeepExcept") = "BLANK,Locn1": Should eliminate all rows where Location column is blank or where = "Locn1" for rngRow row

This logic can be an elseif with .Exists("KeepOnly") since they are mutually exclusive

Should test with BRImport_raw_KeepExcept that contains values "Online", "Store" and BLANK in its Location_BR column

test_ImportParseNorm_FilterRows is a reference on testing