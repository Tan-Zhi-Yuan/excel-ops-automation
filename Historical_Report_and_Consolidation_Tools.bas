Option Explicit

' ==============================================================================
' Module:       Historical_Report_and_Consolidation_Tools
' Description:  A suite of tools for managing monthly accuracy reporting.
'
' Contents:
'   1. Copy_Historical_Accuracy_Reports:
'      - Automates the retrieval of tabs from the last 6 months of files.
'      - Handles directory navigation and year-rollover date logic.
'
'   2. Consolidate_Filtered_Data:
'      - Merges the retrieved tabs into a single master dataset.
'      - Filters rows against a "Master Item List" using Hash Maps (Dictionary).
'      - Remaps columns (ETL) and cleans data types.
'
' Author:       Tan Zhi Yuan
' ==============================================================================

' ==============================================================================
' TOOL 1: HISTORICAL REPORT COMPILER
' ==============================================================================
Sub Copy_Historical_Accuracy_Reports()

    ' --- CONFIGURATION ---
    ' Update this path to your specific network or local directory
    Const ROOT_PATH As String = "C:\Path\To\Procurement\Reports\"
    Const SHEET_TO_COPY As String = "ASN"
    Const FILE_PARTIAL_MATCH As String = "accuracy report" ' Keyword to find the file
    ' ---------------------

    Dim currentDate As Date
    Dim monthFolders(5) As String ' Array for the last 6 month folder names (YYYY-MM)
    Dim i As Long
    Dim folderPath As String, fileName As String
    Dim targetWorkbook As Workbook, sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim foundFile As Boolean

    Set targetWorkbook = ThisWorkbook
    
    ' 1. Calculate the last 6 months (handling year rollovers)
    ' We start from the 1st of the current month to ensure accurate DateAdd subtraction
    currentDate = DateSerial(Year(Date), Month(Date), 1)

    For i = 0 To 5 ' 0 = Last Month, 5 = 6 Months ago
        ' DateAdd("m", -1) moves back one month. We offset by (i + 1).
        monthFolders(i) = Format(DateAdd("m", -(i + 1), currentDate), "YYYY-MM")
    Next i

    ' 2. Loop through the generated folder names
    For i = LBound(monthFolders) To UBound(monthFolders)
        folderPath = ROOT_PATH & monthFolders(i) & "\"
        foundFile = False
        
        ' Check if directory exists
        If Dir(folderPath, vbDirectory) <> "" Then
            ' Find the first Excel file matching the keyword
            fileName = Dir(folderPath & "*" & FILE_PARTIAL_MATCH & "*.xls*")

            ' 3. Process the file if found
            Do While fileName <> ""
                Set sourceWorkbook = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
                
                ' Safely attempt to set the sheet
                On Error Resume Next
                Set sourceSheet = sourceWorkbook.Sheets(SHEET_TO_COPY)
                On Error GoTo 0

                If Not sourceSheet Is Nothing Then
                    ' Copy sheet to the end of the current workbook
                    sourceSheet.Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)
                    
                    ' Rename sheet: "2023-10 ASN" to avoid duplicates
                    On Error Resume Next
                    targetWorkbook.Sheets(targetWorkbook.Sheets.Count).Name = monthFolders(i) & " " & SHEET_TO_COPY
                    On Error GoTo 0
                    
                    foundFile = True
                End If
                
                sourceWorkbook.Close SaveChanges:=False
                
                ' Stop after finding the first valid report for this month
                Exit Do
            Loop
            
            If Not foundFile Then Debug.Print "Report missing for: " & monthFolders(i)
        Else
            Debug.Print "Directory missing: " & folderPath
        End If
    Next i

    MsgBox "Historical report compilation complete.", vbInformation, "Success"

End Sub

' ==============================================================================
' TOOL 2: FILTERED DATA CONSOLIDATOR
' ==============================================================================
Sub Consolidate_Filtered_Data()

    ' --- CONFIGURATION ---
    Const MERGED_SHEET_NAME As String = "Merged Filtered Data"
    Const ITEM_LIST_TABLE_NAME As String = "Itemlist"
    Const ITEM_LIST_COLUMN_HEADER As String = "Item Code"
    Const DATA_ITEM_COLUMN As String = "A"
    Const MONTH_SHEET_KEYWORD As String = "ASN" ' Identifies sheets to process
    ' ---------------------

    ' Define Final Headers (A to K)
    Dim HeaderArray As Variant
    HeaderArray = Array("Item Code", "Sales Month", "Desc", "Warehouse", "Sales", "Forecast", "%", "Difference", "Adjusted Forecast", "%", "Difference")
    
    ' Column Mapping: Which source columns to grab? (1-based index)
    ' We skip Col 3 (Supplier) and map the rest.
    Dim SourceColMap As Variant
    SourceColMap = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 11)

    Dim targetWorkbook As Workbook, mergedSheet As Worksheet, ws As Worksheet
    Dim itemTable As ListObject
    Dim itemCodes As Object ' Scripting.Dictionary
    Dim i As Long, j As Long, k As Long
    Dim lastRowSource As Long, lastRowTarget As Long
    Dim headerCopied As Boolean

    Set targetWorkbook = ThisWorkbook
    Set itemCodes = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler

    ' 1. Load Filter Criteria (Item Codes) into Dictionary
    Set itemTable = FindTable(ITEM_LIST_TABLE_NAME)
    If itemTable Is Nothing Then
        MsgBox "Table '" & ITEM_LIST_TABLE_NAME & "' not found.", vbCritical
        Exit Sub
    End If
    
    ' Efficiently load unique codes
    Dim itemColRange As Range
    Set itemColRange = itemTable.ListColumns(ITEM_LIST_COLUMN_HEADER).DataBodyRange
    For i = 1 To itemColRange.Rows.Count
        Dim code As String
        code = Trim(UCase(itemColRange.Cells(i, 1).Value))
        If code <> "" And Not itemCodes.Exists(code) Then itemCodes.Add code, True
    Next i
    
    ' 2. Prepare Target Sheet
    On Error Resume Next
    Set mergedSheet = targetWorkbook.Sheets(MERGED_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If mergedSheet Is Nothing Then
        Set mergedSheet = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        mergedSheet.Name = MERGED_SHEET_NAME
    Else
        mergedSheet.Cells.Clear
    End If

    ' 3. Iterate Through Sheets
    For Each ws In targetWorkbook.Worksheets
        ' Skip the config sheet and the target sheet itself
        If ws.Name = itemTable.Parent.Name Or ws.Name = MERGED_SHEET_NAME Then GoTo NextSheet
        
        ' Process only "ASN" sheets
        If InStr(1, ws.Name, MONTH_SHEET_KEYWORD, vbTextCompare) > 0 Then
            
            lastRowSource = ws.Range(DATA_ITEM_COLUMN & Rows.Count).End(xlUp).Row
            If lastRowSource <= 1 Then GoTo NextSheet

            ' Extract Sales Month from Tab Name (e.g. "2025-10" -> "10")
            Dim salesMonth As String
            If InStr(ws.Name, "-") > 0 Then
                salesMonth = Mid(ws.Name, InStr(ws.Name, "-") + 1, 2)
            Else
                salesMonth = "00"
            End If
            
            ' Add Headers (Once)
            If Not headerCopied Then
                mergedSheet.Range("A1").Resize(1, UBound(HeaderArray) + 1).Value = HeaderArray
                headerCopied = True
            End If
            
            ' Loop Rows
            For i = 2 To lastRowSource
                Dim currentCode As String
                currentCode = Trim(UCase(ws.Cells(i, Columns(DATA_ITEM_COLUMN).Column).Value))
                
                ' Filter Check
                If itemCodes.Exists(currentCode) Then
                    lastRowTarget = mergedSheet.Cells(Rows.Count, "A").End(xlUp).Row
                    
                    ' Build Row Array (11 Columns)
                    Dim RowData(1 To 11) As Variant
                    
                    k = 1
                    For j = LBound(SourceColMap) To UBound(SourceColMap)
                        Dim sourceColIndex As Long
                        sourceColIndex = SourceColMap(j)
                        
                        Dim cellValue As Variant
                        cellValue = ws.Cells(i, sourceColIndex).Value
                        
                        ' Data Cleaning
                        If IsEmpty(cellValue) Or UCase(cellValue) = "NA" Then cellValue = 0
                        If sourceColIndex = 4 Then cellValue = Val(cellValue) ' Clean Warehouse ID
                        
                        ' Map Data (Handling the insertion of Sales Month at Index 2)
                        If k = 1 Then
                            RowData(1) = cellValue ' Item Code
                        ElseIf k >= 2 Then
                            RowData(k + 1) = cellValue ' Shift remaining data right
                        End If
                        k = k + 1
                    Next j
                    
                    ' Inject derived Sales Month
                    RowData(2) = salesMonth
                    ' Re-assert Item Code
                    RowData(1) = ws.Cells(i, 1).Value
                    
                    ' Bulk Write Row
                    mergedSheet.Cells(lastRowTarget + 1, 1).Resize(1, UBound(RowData)).Value = RowData
                End If
            Next i
        End If
NextSheet:
    Next ws

    mergedSheet.Activate
    MsgBox "Consolidation Complete.", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ==============================================================================
' SHARED HELPER FUNCTIONS
' ==============================================================================
Function FindTable(tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set FindTable = lo
            Exit Function
        End If
    Next ws
End Function
