' ========================================================================================
' Module:       TableAutoAutomation
' Description:  Scans worksheets for unstructured data blocks, converts them to ListObjects
'               (Excel Tables), standardizes column names, and applies dynamic filtering
'               logic based on adjacent header text.
' Author:       [Tan ZHi Yuan]
' ========================================================================================
Option Explicit

' Configuration Constants
Private Const SHEET_PREFIX As String = "T"
Private Const TABLE_START_KEYWORD As String = "Item Code*"
Private Const COL_FILTER_TARGET_INDEX As Long = 11
Private Const COL_TO_RENAME_FROM As String = "%"
Private Const COL_TO_RENAME_TO As String = "percentage"
Private Const TABLE_WIDTH_COLS As Long = 12

Public Sub ProcessAllDataSheets()
    ' Entry point for the automation.
    ' Loops through specific sheets (T1-T9) and triggers processing.
    
    Dim i As Long
    Dim ws As Worksheet
    Dim sheetName As String
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    For i = 1 To 9
        sheetName = SHEET_PREFIX & i
        If SheetExists(sheetName) Then
            Set ws = ThisWorkbook.Sheets(sheetName)
            ProcessSingleSheet ws
        End If
    Next i
    
    MsgBox "Automation Complete: Tables created, sanitized, and filtered.", vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub ProcessSingleSheet(ws As Worksheet)
    ' Scans a single worksheet for data blocks to convert to tables.
    
    Dim lastRow As Long
    Dim i As Long
    Dim rng As Range
    Dim tbl As ListObject
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 1 To lastRow
        ' Check for the start of a data block
        If LCase(ws.Cells(i, 1).Value) Like LCase(TABLE_START_KEYWORD) Then
            
            ' Define the range of the table
            Set rng = FindBlockRange(ws, i, lastRow)
            
            ' Create or retrieve the table
            Set tbl = EnsureTableExists(ws, rng)
            
            If Not tbl Is Nothing Then
                ' Apply Business Logic
                ApplyStandardFilters tbl
                StandardizeHeaders tbl
                ApplySmartForecastFilter ws, tbl
            End If
        End If
    Next i
End Sub

Private Function FindBlockRange(ws As Worksheet, startRow As Long, limitRow As Long) As Range
    ' Helper: Determines the extent of a contiguous data block.
    Dim endRow As Long
    endRow = startRow + 1
    
    Do While endRow <= limitRow And ws.Cells(endRow, 1).Value <> ""
        endRow = endRow + 1
    Loop
    
    Set FindBlockRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow - 1, TABLE_WIDTH_COLS))
End Function

Private Function EnsureTableExists(ws As Worksheet, rng As Range) As ListObject
    ' Helper: safely adds a ListObject or returns the existing one.
    On Error Resume Next
    Set EnsureTableExists = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    If EnsureTableExists Is Nothing Then Set EnsureTableExists = rng.ListObject
    On Error GoTo 0
End Function

Private Sub ApplyStandardFilters(tbl As ListObject)
    ' Step 2: Apply baseline filtering (Column 11 = "0")
    If tbl.ListColumns.Count >= COL_FILTER_TARGET_INDEX Then
        tbl.Range.AutoFilter Field:=COL_FILTER_TARGET_INDEX, Criteria1:="0"
    End If
End Sub

Private Sub StandardizeHeaders(tbl As ListObject)
    ' Step 3: Rename specific columns for consistency (e.g., "%" -> "percentage")
    Dim col As ListColumn
    
    On Error Resume Next
    Set col = tbl.ListColumns(COL_TO_RENAME_FROM)
    On Error GoTo 0
    
    If Not col Is Nothing Then
        On Error Resume Next ' Suppress error if "percentage" already exists (e.g. from previous run)
        col.Name = COL_TO_RENAME_TO
        On Error GoTo 0
    End If
End Sub

Private Sub ApplySmartForecastFilter(ws As Worksheet, tbl As ListObject)
    ' Step 4: Context-aware filtering based on headers located above the table.
    
    Dim headerContext As String
    Dim targetCol As ListColumn
    Dim filterCriteria As String
    
    ' A. Detect Context (Sales < Forecast vs Sales > Forecast)
    headerContext = ScanHeaderArea(ws, tbl)
    
    ' B. Identify which column to filter (handles variations like %, percentage, diff)
    Set targetCol = FindPercentageColumn(tbl)
    
    ' C. Apply Logic
    If Not targetCol Is Nothing Then
        Select Case headerContext
            Case "LESS"
                tbl.Range.AutoFilter Field:=targetCol.Index, Criteria1:=">100%"
            Case "MORE"
                tbl.Range.AutoFilter Field:=targetCol.Index, Criteria1:="<-100%"
        End Select
    End If
End Sub

Private Function ScanHeaderArea(ws As Worksheet, tbl As ListObject) As String
    ' Scans the 3 rows immediately above the table for keywords.
    Dim searchRng As Range
    Dim cell As Range
    
    On Error Resume Next
    Set searchRng = ws.Range(tbl.HeaderRowRange.Offset(-3, 0), tbl.HeaderRowRange.Offset(-1, 0))
    On Error GoTo 0
    
    If searchRng Is Nothing Then Exit Function
    
    For Each cell In searchRng
        If InStr(1, cell.Value, "SALES < FORECAST", vbTextCompare) > 0 Then
            ScanHeaderArea = "LESS"
            Exit Function
        ElseIf InStr(1, cell.Value, "SALES > FORECAST", vbTextCompare) > 0 Then
            ScanHeaderArea = "MORE"
            Exit Function
        End If
    Next cell
End Function

Private Function FindPercentageColumn(tbl As ListObject) As ListColumn
    ' Looks for the target column using various acceptable naming conventions.
    Dim col As ListColumn
    Dim cleanName As String
    
    For Each col In tbl.ListColumns
        cleanName = UCase(Trim(col.Name))
        Select Case cleanName
            Case "PERCENTAGE", "PERCENTAGE2", "%", "%2", "DIFF"
                Set FindPercentageColumn = col
                Exit Function
        End Select
    Next col
End Function

Private Function SheetExists(sheetName As String) As Boolean
    ' Utility to check sheet existence without error throwing.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
