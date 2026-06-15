---
type: jdl_writeup
description: Potential Performance and Robustness Upgrade to ExcelSteps
tags:
  - VBA
  - ExcelSteps
  - errorhandling
created: 06-10-2026
updated: 06-10-2026
author: JDL
---

Client project had "Stack Overflow" issue due to recursively calling `tblRowsCols.Init` for shtErrors when using errs.ShowMessage (looks up message from shtErrors) when errors had ShtCaseError (Sheet actually named "Errors_" when provisioning with shtErrors="errors_").

ErrorHandling currently provisions `shtErrors` table WITH EVERY Error, warning, showmessage etc., This is  problematic for addressing above (obscure) issue because `.ShowMessage` sets `errs.iCodeLocal=1` which gets in the way of using errs to warn about the sheet name case issue within the `.ShowMessage` flow of calls to initialize the errors_ table for message lookup

For both performance and code robustness, an improvement would be to add `ErrorHandling.tbl` attribute and provision `.tbl` when errs is instanced. This would make the table available whenever needed and eliminate the performance hit of doing it every time errs is used to report something to the user.

* Can provision `.tbl` within `errs.Init`
* `errs` aka `errsObj` argument is already passed as arg to `ErrsMeta.Init`, so provisioned `.tbl` would be passed with that
* Can provision `errs.tbl` once `errs.wkbkE` is set in `ErrorHandling.Init`

Possible that similar logic applies to ExcelSteps table --make optional to pre-initialize/provision it and pass as argument to Refresh so that Refresh doesn't need to do this for each, individual refresh (detect if .tblSteps is Nothing and only initialize if so etc.)

**Current `ErrorsMeta.cls` `.Init` that provisions the errors table to do lookups**
```vb
'--------------------------------------------------------------------------------------
' Initialize metadata state and provision Errors_ table wayfinding
' JDL 3/13/26
'
Public Function Init(meta, errsObj) As Boolean
  On Error GoTo ErrorExit

  Set meta.tblErrs = New tblRowsCols
  If Not meta.tblErrs.Provision(meta.tblErrs, errsObj.wkbkE, False, _
    sht:=shtErrors, IsSetColRngs:=True) Then GoTo ErrorExit

  meta.Code = 0
  meta.Routine = ""
  meta.Message = ""
  meta.IsUserFacing = False
  meta.IsBaseNotFound = False
  meta.IsCodeNotFound = False
  meta.IsMalformed = False
  Init = True
  Exit Function

ErrorExit:
  Init = False
End Function

```

**Current ErrorHandleUtil.bas**
```vb
Attribute VB_Name = "ErrorHandleUtil"
'Version 5/14/26 shtErrors lower case
Option Explicit
Public errs As Object

'Error Handling Constants
Public Const shtErrors As String = "errors_"
Public Const sErrBase As String = "Base"
Public Const sVBAErr As String = "Unknown VBA Error"
Public Const sFileErrs As String = "Warnings_and_Errors.txt"
Public Const sSettingErrs As String = "Warnings_Errors"
'-----------------------------------------------------------------------------------------------------
'Initialize local error handling and errs Class
'
''JDL 3/8/23; Modified 10/16/25
'
Sub SetErrs(CallingFunction, Optional wkbkE As Workbook = Nothing)
    Dim IsDriver As Boolean, IsNonBool As Boolean

    IsDriver = (VarType(CallingFunction) = vbString And LCase(CallingFunction) = "driver")
    IsNonBool = (VarType(CallingFunction) = vbString And LCase(CallingFunction) = "non-bool")

    'Initialize errs (In case CallingFunction called directly instead of by local driver sub)
    If (errs Is Nothing) Or IsDriver Then
        Set errs = New ErrorHandling
        
        'Default Errors_ sheet location
        If wkbkE Is Nothing Then Set wkbkE = ThisWorkbook
        
        'True/False = Master switch for enabling error handling in project
        errs.Init wkbkE, IsHandle:=False

        'Set defaults by calling mode
        If IsDriver Then
            errs.IsTesting = False
            errs.IsShowMsgs = True
        Else
            errs.IsTesting = True
            errs.IsShowMsgs = False
        End If

    ElseIf IsDriver Then
        'Driver call refreshes mode flags even if errs already exists
        errs.IsTesting = False
        errs.IsShowMsgs = True
    End If
    
    'Initialize Boolean calling function
    If Not IsDriver And Not IsNonBool Then CallingFunction = True
End Sub
```

**Current `ErrorHandling.Init`**
```vb
'------------------------------------------------------------------------------------------------------
'Initialize errs Class
' Declare errs as global prior to first, user-initiated procedure sub (Public errs as Object)
' Call errs.Init from user-initiated sub after instancing errs (dim errs as New ErrorHandling)
'
'JDL 12/20/22   Modified 9/19/25 add wkbkE arg and logic for setting .wkbkE
'
Public Sub Init(wkbkE, Optional IsHandle)
    With errs
        If IsMissing(IsHandle) Then
            .IsHandle = True
        Else
            .IsHandle = IsHandle
        End If
        
        'Not used anywhere? JDL 9/19/25 - delete attr if no issues found
        'Set .wkbk = ThisWorkbook
        
        'If shtErrors resides in a different wkbk (e.g. in project wkbk) than ErrorHandling class
        Set .wkbkE = ThisWorkbook
        If Not IsMissing(wkbkE) Then Set .wkbkE = wkbkE
        
        .iCodeLocal = 0
        .iCodeBase = 0
        .iCodeReport = 0
        .ErrParam = ""
        .ErrMsg = ""
        .IsUserFacing = False
        .IsNewErr = True
        .IsDriver = False
        .IsShowMsgs = True
    End With
End Sub

```