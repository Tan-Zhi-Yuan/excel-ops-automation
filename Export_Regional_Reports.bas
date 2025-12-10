Option Explicit

' ==============================================================================
' Module:       Report Extraction & Updater
' Description:  Automates the weekly/monthly reporting cycle.
'
' Core Features:
'   1. Multi-Source Handling: Loops through file paths defined in "Link1" & "Link2".
'   2. Data Extraction: Scans for "ACCURACY REPORT SUMMARY" blocks in source files.
'   3. Table Synchronization: Updates main reporting tables with rolling window visibility.
'   4. Graph Data Sync: Transposes extracted data into a separate "Graph" sheet,
'      adds date stamps (YYYYMM), resizes tables, and enforces FIFO row visibility.
'   5. Sanitization: Exports a clean XLSX copy stripped of macros and links.
' ==============================================================================

Sub UpdateReport_And_SaveClean()
    ' --- PART 1: DECLARATIONS ---
    Dim wbThis As Workbook, wbSource As Workbook, wbNew As Workbook
    Dim wsMain As Worksheet, wsDest As Worksheet, wsSource As Worksheet
    Dim tblMap As ListObject, tblDest As ListObject
    Dim newCol As ListColumn
    Dim rLink As Range, rFound As Range
    Dim rTestRange As Range, cell As Range
    Dim sPath As String, sSourceSheet As String, sDestSheet As String, sAddition As String
    Dim sSearchTerm As String, strCleanPath As String
    Dim x As Long, i As Long, k As Long
    Dim colLetter As String
    Dim firstAddress As String
    Dim foundRows As Collection
    Dim isValidBlock As Boolean
    Dim totalCols As Long
    Dim lastSourceCol As Long
    Dim rDataStart As Range
    
    ' Variables for Looping Links
    Dim vLinks As Variant, vItem As Variant
    
    ' Variables for Graph Sheet Logic
    Dim wsGraph As Worksheet
    Dim tblGraph As ListObject
    Dim rForecast As Range
    Dim rInsertion As Range
    Dim sGraphSheet As String
    Dim rIndex As Long
    
    ' --- SETTINGS ---
    Const VISIBLE_DATA_COLS As Integer = 6
    Const FIXED_COLS As Integer = 1
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbThis = ThisWorkbook
    Set wsMain = wbThis.ActiveSheet ' Assumes button is on the "Source" sheet
    
    ' =========================================================
    ' PART 2: THE DATA UPDATE PROCESS
    ' =========================================================
    
    ' 0. Find Mapping Table
    On Error Resume Next
    Set tblMap = wsMain.ListObjects("Source")
    On Error GoTo 0
    If tblMap Is Nothing Then
        MsgBox "STOP: Table named 'Source' not found.", vbCritical
        GoTo Cleanup
    End If
    
    ' Define the links to process
    vLinks = Array("Link1", "Link2")
    
    ' START LOOP THROUGH FILES
    For Each vItem In vLinks
        
        ' 1. Find Link Range
        Set rLink = Nothing
        On Error Resume Next
        Set rLink = wbThis.Names(vItem).RefersToRange
        If rLink Is Nothing Then Set rLink = wsMain.Range(vItem)
        On Error GoTo 0
        
        ' Check if Link exists
        If rLink Is Nothing Then GoTo NextFile
        If Dir(rLink.Value) = "" Then GoTo NextFile
        
        sPath = rLink.Value
        Set wbSource = Workbooks.Open(Filename:=sPath, ReadOnly:=True)
        
        ' 2. Loop Through Sheets
        For x = 1 To tblMap.ListRows.Count
            
            sDestSheet = Trim(tblMap.DataBodyRange.Cells(x, 1).Value)
            sSourceSheet = Trim(tblMap.DataBodyRange.Cells(x, 2).Value)
            
            sAddition = ""
            On Error Resume Next
            sAddition = Trim(tblMap.ListColumns("Addition").DataBodyRange.Cells(x).Value)
            On Error GoTo 0
            
            If sDestSheet <> "" And sSourceSheet <> "" Then
                If SheetExists(wbThis, sDestSheet) And SheetExists(wbSource, sSourceSheet) Then
                    
                    Set wsDest = wbThis.Sheets(sDestSheet)
                    Set wsSource = wbSource.Sheets(sSourceSheet)
                    
                    If wsDest.ListObjects.Count > 0 Then
                        Set tblDest = wsDest.ListObjects(1)
                        
                        sSearchTerm = "*ACCURACY REPORT SUMMARY*" & sAddition & "*"
                        
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
                        
                        For i = foundRows.Count To 1 Step -1
                            
                            ' A. Find Correct Data Column
                            Dim rowToCheck As Long
                            rowToCheck = foundRows(i) + 8
                            lastSourceCol = wsSource.Cells(rowToCheck, wsSource.Columns.Count).End(xlToLeft).Column
                            If lastSourceCol < 3 Then lastSourceCol = 3
                            
                            ' B. Grab 8 Rows of Data
                            Set rDataStart = wsSource.Cells(foundRows(i) + 1, lastSourceCol)
                            Set rTestRange = rDataStart.Resize(8, 1)
                            
                            isValidBlock = True
                            For Each cell In rTestRange
                                If IsError(cell.Value) Then isValidBlock = False: Exit For
                            Next cell
                            
                            If isValidBlock Then
                                
                                ' STEP 1: PREPARE MAIN TABLE (Ensure 9 rows)
                                Do While tblDest.ListRows.Count < 9
                                    tblDest.ListRows.Add
                                Loop
                                
                                ' STEP 2: PASTE 8 ROWS TO MAIN TABLE
                                Set newCol = tblDest.ListColumns.Add
                                rTestRange.Copy
                                newCol.DataBodyRange.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
                                
                                ' STEP 3: ADD FORMULA TO ROW 9
                                colLetter = Split(newCol.Range.Cells(1).Address(True, False), "$")(0)
                                newCol.DataBodyRange.Cells(9, 1).Formula = _
                                    "=" & colLetter & "8+" & colLetter & "3+" & colLetter & "2+" & colLetter & "4"
                                
                                newCol.Range.EntireColumn.AutoFit
                                
                                ' STEP 4: HIDE OLD COLUMNS
                                totalCols = tblDest.ListColumns.Count
                                If totalCols > (FIXED_COLS + VISIBLE_DATA_COLS) Then
                                    For k = (FIXED_COLS + 1) To (totalCols - VISIBLE_DATA_COLS)
                                        tblDest.ListColumns(k).Range.EntireColumn.Hidden = True
                                    Next k
                                    For k = (totalCols - VISIBLE_DATA_COLS + 1) To totalCols
                                        tblDest.ListColumns(k).Range.EntireColumn.Hidden = False
                                    Next k
                                End If
                                
                                ' =========================================================
                                ' PART 2B: UPDATE GRAPH SHEET
                                ' =========================================================
                                sGraphSheet = sDestSheet & " GRAPH"
                                
                                If SheetExists(wbThis, sGraphSheet) Then
                                    Set wsGraph = wbThis.Sheets(sGraphSheet)
                                    
                                    ' 1. Find "forecast" in Column A
                                    Set rForecast = wsGraph.Columns("A").Find(What:="forecast", LookIn:=xlValues, LookAt:=xlPart)
                                    
                                    If Not rForecast Is Nothing Then
                                        
                                        ' 2. Find Insertion Point (Next Empty Cell)
                                        If rForecast.Offset(1, 0).Value = "" Then
                                            Set rInsertion = rForecast.Offset(1, 0)
                                        Else
                                            Set rInsertion = rForecast.End(xlDown).Offset(1, 0)
                                        End If
                                        
                                        ' 3. PASTE DATA (Transposed)
                                        rInsertion.Resize(1, 9).Value = _
                                            Application.Transpose(newCol.DataBodyRange.Resize(9, 1).Value)
                                            
                                        ' --- ADD YYYYMM to Column J ---
                                        wsGraph.Cells(rInsertion.Row, "J").Value = Format(Date, "yyyymm")
                                        ' ------------------------------
                                        
                                        ' 4. FORMAT PERCENTAGES (Columns 4 & 5)
                                        rInsertion.Offset(0, 3).NumberFormat = "0.00%"
                                        rInsertion.Offset(0, 4).NumberFormat = "0.00%"
                                        
                                        ' 5. FORCE TABLE RESIZE
                                        On Error Resume Next
                                        Set tblGraph = wsGraph.ListObjects(1)
                                        On Error GoTo 0
                                        
                                        If Not tblGraph Is Nothing Then
                                            Dim rngNewTable As Range
                                            Set rngNewTable = wsGraph.Range( _
                                                tblGraph.Range.Cells(1, 1), _
                                                wsGraph.Cells(rInsertion.Row, tblGraph.Range.Columns(tblGraph.Range.Columns.Count).Column) _
                                            )
                                            tblGraph.Resize rngNewTable
                                            
                                            ' 6. HIDE FIRST VISIBLE ROW (FIFO)
                                            If tblGraph.ListRows.Count > 1 Then
                                                For rIndex = 1 To tblGraph.ListRows.Count - 1
                                                    If tblGraph.ListRows(rIndex).Range.EntireRow.Hidden = False Then
                                                        tblGraph.ListRows(rIndex).Range.EntireRow.Hidden = True
                                                        Exit For
                                                    End If
                                                Next rIndex
                                            End If
                                        End If
                                    End If
                                End If
                                ' =========================================================
                                
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
        Next x
        
        wbSource.Close SaveChanges:=False
        
NextFile:
    Next vItem

    ' =========================================================
    ' PART 3: SAVE CLEAN VERSION
    ' =========================================================
    
    strCleanPath = Left(wbThis.FullName, InStrRev(wbThis.FullName, ".") - 1) & ".xlsx"
    
    wbThis.Sheets.Copy
    Set wbNew = ActiveWorkbook
    
    On Error Resume Next
    wbNew.Sheets(wsMain.Name).Delete
    On Error GoTo 0
    
    Dim vLinksEx As Variant
    vLinksEx = wbNew.LinkSources(Type:=xlLinkTypeExcelLinks)
    If Not IsEmpty(vLinksEx) Then
        Dim L As Integer
        For L = 1 To UBound(vLinksEx)
            wbNew.BreakLink Name:=vLinksEx(L), Type:=xlLinkTypeExcelLinks
        Next L
    End If
    
    On Error Resume Next
    wbNew.SaveAs Filename:=strCleanPath, FileFormat:=xlOpenXMLWorkbook
    
    If Err.Number = 0 Then
        MsgBox "Process Complete!" & vbNewLine & vbNewLine & _
               "1. Data updated from multiple sources." & vbNewLine & _
               "2. Graph Sheets updated with new forecast data." & vbNewLine & _
               "3. Clean file saved: " & strCleanPath, vbInformation
    Else
        MsgBox "Error saving clean file: " & Err.Description, vbCritical
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
