Option Explicit

' ==============================================================================
' Module:       Report_Updater
' Description:  Automates the weekly/monthly reporting cycle.
'               1. READS a configuration table ("Source") to map sheets.
'               2. OPENS a source workbook defined in a named range ("Link1").
'               3. SEARCHES for specific data blocks ("ACCURACY REPORT SUMMARY").
'               4. IMPORTS valid data columns into the destination tables.
'               5. EXPORTS a "Clean" (XLSX) version, removing macros and config sheets.
'
' Author:       [Your Name]
' Date:         2025-12-09
' ==============================================================================

Sub UpdateReport_And_SaveClean()
    
    ' --- CONFIGURATION ---
    Const VISIBLE_DATA_COLS As Integer = 6  ' Number of recent columns to keep visible
    Const FIXED_COLS As Integer = 1         ' Number of static columns on the left
    Const SEARCH_KEYWORD As String = "ACCURACY REPORT SUMMARY"
    Const MAP_TABLE_NAME As String = "Source"
    Const FILE_PATH_RANGE As String = "Link1"
    ' ---------------------

    ' Object Variables
    Dim wbThis As Workbook, wbSource As Workbook, wbNew As Workbook
    Dim wsMain As Worksheet, wsDest As Worksheet, wsSource As Worksheet
    Dim tblMap As ListObject, tblDest As ListObject
    Dim newCol As ListColumn
    Dim rLink1 As Range, rFound As Range, rTestRange As Range, cell As Range
    
    ' String Variables
    Dim sPath As String, sSourceSheet As String, sDestSheet As String, sAddition As String
    Dim sSearchTerm As String, strCleanPath As String, colLetter As String
    
    ' Logic Variables
    Dim x As Long, i As Long, k As Long
    Dim firstAddress As String
    Dim foundRows As Collection
    Dim isValidBlock As Boolean, hasNonZero As Boolean
    Dim totalCols As Long
    
    ' Performance Optimization
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbThis = ThisWorkbook
    Set wsMain = wbThis.ActiveSheet ' Assumes button is on the "Source"/Config sheet
    
    ' =========================================================
    ' PHASE 1: VALIDATION & SETUP
    ' =========================================================
    
    ' 1. Validate File Path (Link1)
    On Error Resume Next
    Set rLink1 = wbThis.Names(FILE_PATH_RANGE).RefersToRange
    If rLink1 Is Nothing Then Set rLink1 = wsMain.Range(FILE_PATH_RANGE)
    If rLink1 Is Nothing Then Set rLink1 = wsMain.Range("B1") ' Fallback
    On Error GoTo 0
    
    If rLink1 Is Nothing Or Dir(rLink1.Value) = "" Then
        MsgBox "Error: Invalid file path defined in '" & FILE_PATH_RANGE & "'.", vbCritical
        GoTo Cleanup
    End If
    
    sPath = rLink1.Value
    
    ' 2. Validate Mapping Table
    On Error Resume Next
    Set tblMap = wsMain.ListObjects(MAP_TABLE_NAME)
    On Error GoTo 0
    If tblMap Is Nothing Then
        MsgBox "Error: Configuration table '" & MAP_TABLE_NAME & "' not found.", vbCritical
        GoTo Cleanup
    End If

    ' 3. Open Source Workbook (Read Only)
    Set wbSource = Workbooks.Open(Filename:=sPath, ReadOnly:=True)
    
    ' =========================================================
    ' PHASE 2: DATA IMPORT LOOP
    ' =========================================================
    
    For x = 1 To tblMap.ListRows.Count
        
        ' Read Config Row
        sDestSheet = Trim(tblMap.DataBodyRange.Cells(x, 1).Value)
        sSourceSheet = Trim(tblMap.DataBodyRange.Cells(x, 2).Value)
        
        ' Optional "Addition" search criteria
        sAddition = ""
        On Error Resume Next
        sAddition = Trim(tblMap.ListColumns("Addition").DataBodyRange.Cells(x).Value)
        On Error GoTo 0
        
        If sDestSheet <> "" And sSourceSheet <> "" Then
            ' Ensure both sheets exist before proceeding
            If SheetExists(wbThis, sDestSheet) And SheetExists(wbSource, sSourceSheet) Then
                
                Set wsDest = wbThis.Sheets(sDestSheet)
                Set wsSource = wbSource.Sheets(sSourceSheet)
                
                If wsDest.ListObjects.Count > 0 Then
                    Set tblDest = wsDest.ListObjects(1)
                    
                    ' Build dynamic search string
                    sSearchTerm = "*" & SEARCH_KEYWORD & "*" & sAddition & "*"
                    
                    ' Find all occurrences of the keyword
                    Set foundRows = New Collection
                    With wsSource.Columns("A")
                        Set rFound = .Find(What:=sSearchTerm, LookIn:=xlValues, LookAt:=xlPart)
                        If Not rFound Is Nothing Then
                            firstAddress = rFound.Address
                            Do
                                foundRows.Add rFound.Row
                                Set rFound = .FindNext(rFound)
                            Loop While rFound.Address <> firstAddress
                        End If
                    End With
                    
                    ' Search Bottom-Up for the most recent Valid Data Block
                    For i = foundRows.Count To 1 Step -1
                        ' Define the data block relative to the found header (Rows +1 to +7)
                        Set rTestRange = wsSource.Range(wsSource.Cells(foundRows(i) + 1, "C"), _
                                                        wsSource.Cells(foundRows(i) + 7, "C"))
                        
                        ' Validation: Check for Errors, "NAN", and ensure at least one non-zero value
                        isValidBlock = True
                        hasNonZero = False
                        
                        For Each cell In rTestRange
                            If IsError(cell.Value) Then isValidBlock = False: Exit For
                            If UCase(Trim(cell.Text)) = "NAN" Then isValidBlock = False: Exit For
                            If IsNumeric(cell.Value) And cell.Value <> 0 Then hasNonZero = True
                        Next cell
                        
                        If isValidBlock And hasNonZero Then
                            ' -- IMPORT DATA --
                            Set newCol = tblDest.ListColumns.Add
                            rTestRange.Copy
                            newCol.DataBodyRange.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
                            
                            ' -- ADD CALCULATED TOTAL (Formula Injection) --
                            ' Adds formula: =X9+X4+X3
                            colLetter = Split(newCol.Range.Cells(1).Address(True, False), "$")(0)
                            newCol.DataBodyRange.Cells(8, 1).Formula = "=" & colLetter & "9+" & colLetter & "4+" & colLetter & "3"
                            
                            newCol.Range.EntireColumn.AutoFit
                            
                            ' -- ROLLING WINDOW VISIBILITY --
                            ' Hides older columns to keep the view clean
                            totalCols = tblDest.ListColumns.Count
                            If totalCols > (FIXED_COLS + VISIBLE_DATA_COLS) Then
                                For k = (FIXED_COLS + 1) To (totalCols - VISIBLE_DATA_COLS)
                                    tblDest.ListColumns(k).Range.EntireColumn.Hidden = True
                                Next k
                                For k = (totalCols - VISIBLE_DATA_COLS + 1) To totalCols
                                    tblDest.ListColumns(k).Range.EntireColumn.Hidden = False
                                Next k
                            End If
                            
                            Exit For ' Stop after finding the first valid block
                        End If
                    Next i
                End If
            End If
        End If
    Next x
    
    wbSource.Close SaveChanges:=False

    ' =========================================================
    ' PHASE 3: EXPORT CLEAN XLSX
    ' =========================================================
    
    ' Generate new filename
    strCleanPath = Left(wbThis.FullName, InStrRev(wbThis.FullName, ".") - 1) & ".xlsx"
    
    ' Copy sheets to new workbook
    wbThis.Sheets.Copy
    Set wbNew = ActiveWorkbook
    
    ' Delete the Config/Macro Sheet
    On Error Resume Next
    wbNew.Sheets(wsMain.Name).Delete
    On Error GoTo 0
    
    ' Break External Links for security
    Dim vLinks As Variant
    vLinks = wbNew.LinkSources(Type:=xlLinkTypeExcelLinks)
    If Not IsEmpty(vLinks) Then
        Dim L As Integer
        For L = 1 To UBound(vLinks)
            wbNew.BreakLink Name:=vLinks(L), Type:=xlLinkTypeExcelLinks
        Next L
    End If
    
    ' Save Clean File
    On Error Resume Next
    wbNew.SaveAs Filename:=strCleanPath, FileFormat:=xlOpenXMLWorkbook
    
    If Err.Number = 0 Then
        MsgBox "Report Update Successful!" & vbNewLine & vbNewLine & _
               "1. Data imported from source." & vbNewLine & _
               "2. Clean file saved: " & strCleanPath, vbInformation, "Success"
    Else
        MsgBox "Data updated, but failed to save clean file." & vbNewLine & _
               "Error: " & Err.Description, vbCritical
    End If
    On Error GoTo 0

Cleanup:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

' Helper Function
Function SheetExists(wb As Workbook, sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
