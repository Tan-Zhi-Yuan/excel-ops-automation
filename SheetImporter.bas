' ========================================================================================
' Module:       SheetImporter
' Description:  Automates the consolidation of specific worksheets from an external workbook
'               into the current file based on user-defined configuration cells.
' Configuration:
'               1. Named Range "Filepath": Full path of source file.
'               2. Named Range "SheetList": Comma-separated list of sheet names (e.g., "Data,Summary").
' Author:       [Tan ZHi Yuan]
' ========================================================================================
Option Explicit

' Configuration Constants
Private Const NAMED_RNG_PATH As String = "Filepath"
Private Const NAMED_RNG_LIST As String = "SheetList"

Sub ImportSheets()
    
    ' --- 1. DECLARE VARIABLES ---
    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim sFilePath As String, sSheetList As String, sSheetName As String
    Dim rngFilePath As Range, rngSheetList As Range
    Dim arrSheetNames As Variant, vSheetName As Variant
    Dim i As Long
    Dim bSheetFound As Boolean
    
    ' --- 2. SETUP AND VALIDATION ---
    Set wbDest = ThisWorkbook
    
    ' A. Retrieve File Path
    On Error Resume Next
    Set rngFilePath = wbDest.Names(NAMED_RNG_PATH).RefersToRange
    On Error GoTo 0
    
    If rngFilePath Is Nothing Then
        MsgBox "Setup Error: Named range '" & NAMED_RNG_PATH & "' not found.", vbCritical
        Exit Sub
    End If
    sFilePath = rngFilePath.Cells(1, 1).Value
    
    ' B. Retrieve Sheet List
    On Error Resume Next
    Set rngSheetList = wbDest.Names(NAMED_RNG_LIST).RefersToRange
    On Error GoTo 0
    
    If rngSheetList Is Nothing Then
        MsgBox "Setup Error: Named range '" & NAMED_RNG_LIST & "' not found.", vbCritical
        Exit Sub
    End If
    sSheetList = rngSheetList.Cells(1, 1).Value
    
    ' C. Validate Inputs
    If sFilePath = "" Or Dir(sFilePath) = "" Then
        MsgBox "Error: Source file not found at path:" & vbCrLf & sFilePath, vbCritical
        Exit Sub
    End If
    
    If sSheetList = "" Then
        MsgBox "Error: Sheet list is empty.", vbCritical
        Exit Sub
    End If
    
    ' D. Parse Sheet List
    arrSheetNames = Split(sSheetList, ",")
    
    ' --- 3. PREPARE DESTINATION (DELETE OLD SHEETS) ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False ' Suppress "Delete Sheet" warnings
    
    For Each vSheetName In arrSheetNames
        sSheetName = Trim(CStr(vSheetName))
        If sSheetName <> "" Then
            On Error Resume Next
            wbDest.Sheets(sSheetName).Delete
            On Error GoTo 0
        End If
    Next vSheetName
    
    ' --- 4. IMPORT PROCESS ---
    On Error GoTo ErrorHandler
    
    ' Open Source File (Read Only to prevent locking/saving changes)
    Set wbSource = Workbooks.Open(sFilePath, ReadOnly:=True)
    bSheetFound = False
    
    ' Iterate Backwards to preserve sheet order when copying to "Before:=Sheet1"
    For i = UBound(arrSheetNames) To LBound(arrSheetNames) Step -1
        sSheetName = Trim(CStr(arrSheetNames(i)))
        
        If sSheetName <> "" Then
            On Error Resume Next
            wbSource.Sheets(sSheetName).Copy Before:=wbDest.Sheets(1)
            
            If Err.Number = 0 Then bSheetFound = True
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    Next i
    
    ' --- 5. CLEANUP & REPORTING ---
    wbSource.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    If bSheetFound Then
        MsgBox "Import Complete. Sheets updated successfully.", vbInformation
    Else
        MsgBox "Process finished, but no matching sheets were found in the source file.", vbExclamation
    End If
    
    Exit Sub

ErrorHandler:
    ' Safety cleanup in case of crash
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Runtime Error: " & Err.Description, vbCritical
End Sub
