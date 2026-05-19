I think you've got it but to be clear, we read 

```
sFilter = CStr(Intersect(rngRow, .tbl.colrngFilterVals).Value2). 
```

We should run our test on the .colinfo.tbl.colrngVarNorm row = "Location_BR" in .colinfo.rngRowsCurTbl row range, so that's where you need to write the test strings --to that row's .tbl.colrngFilterVals cell