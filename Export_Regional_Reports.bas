Option Explicit

' ==============================================================================
' Module:       Export_Regional_Reports
' Description:  Generates a clean "Client-Facing" or "External" report workbook.
'               - Copies specific required sheets 
'               - Copies Raw Data needed for comparison.
'               - Saves the new file in the same directory as the source.
'
' Author:       [Tan ZHi Yuan]
' ==============================================================================

Sub Create_Regional_PO_Extract()
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Dim defaultSheetName As String
    Dim targetPath As String
    
    ' --- CONFIGURATION: EDIT SHEET NAMES HERE ---
    Const FILE_NAME_OUTPUT As String = "Regional_PO_Status_Export.xlsx"
    
    ' Source Sheet Names (Must match your workbook exactly)
    Const SRC_SHEET_TRACKING_1 As String = "AU PO Tracking"
    Const SRC_SHEET_TRACKING_2 As String = "NZ PO Tracking"
    Const SRC_SHEET_DATA_1 As String = "AU PO status"
    Const SRC_SHEET_DATA_2 As String = "NZ PO status"
    Const SRC_SHEET_INSTRUCT As String = "INSTRUCTIONS"
    
    ' Target Sheet Names (Clean names for the new file)
    Const TGT_SHEET_DATA_1 As String = "Raw Data AU"
    Const TGT_SHEET_DATA_2 As String = "Raw Data NZ"
    ' --------------------------------------------

    ' 1. Optimization & Safety Checks
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False ' Suppress "Delete Sheet" warnings

    Set wbSource = ThisWorkbook

    ' Ensure source is saved so we can get the path
    If wbSource.Path = "" Then
        MsgBox "Please save this workbook before running the export.", vbExclamation, "Save Required"
        GoTo CleanUp
    End If
    
    targetPath = wbSource.Path

    ' 2. Create New Workbook
    Set wbTarget = Workbooks.Add
    
    ' Track default sheet to delete later
    If wbTarget.Worksheets.Count > 0 Then
        defaultSheetName = wbTarget.Worksheets(1).Name
    End If

    ' 3. Copy "Tracking" Sheets (Preserve Formatting)
    ' We copy these BEFORE the first sheet of the new workbook
    On Error Resume Next ' Handle cases where sheets might be missing
    wbSource.Worksheets(SRC_SHEET_TRACKING_1).Copy Before:=wbTarget.Worksheets(1)
    wbSource.Worksheets(SRC_SHEET_TRACKING_2).Copy Before:=wbTarget.Worksheets(2)
    On Error GoTo ErrorHandler

    ' 4. Copy "Data" Sheets (Values Only - No Formatting/Formulas)
    
    ' -- Process Region 1 Data --
    Set wsTarget = wbTarget.Worksheets.Add(After:=wbTarget.Worksheets(wbTarget.Worksheets.Count))
    wsTarget.Name = TGT_SHEET_DATA_1
    wbSource.Worksheets(SRC_SHEET_DATA_1).UsedRange.Copy
    wsTarget.Range("A1").PasteSpecial xlPasteValues
    wsTarget.Cells.ClearFormats ' Strip formatting for raw data cleanliness
    
    ' -- Process Region 2 Data --
    Set wsTarget = wbTarget.Worksheets.Add(After:=wbTarget.Worksheets(wbTarget.Worksheets.Count))
    wsTarget.Name = TGT_SHEET_DATA_2
    wbSource.Worksheets(SRC_SHEET_DATA_2).UsedRange.Copy
    wsTarget.Range("A1").PasteSpecial xlPasteValues
    wsTarget.Cells.ClearFormats

    ' -- Copy Instructions --
    wbSource.Worksheets(SRC_SHEET_INSTRUCT).Copy After:=wbTarget.Worksheets(wbTarget.Worksheets.Count)

    ' 5. Clean Up New Workbook
    Application.CutCopyMode = False
    
    ' Remove the default blank sheet created by Excel
    If defaultSheetName <> "" Then
        On Error Resume Next
        wbTarget.Worksheets(defaultSheetName).Delete
        On Error GoTo ErrorHandler
    End If
    
    ' Set focus to the first sheet
    wbTarget.Worksheets(1).Activate

    ' 6. Save and Close
    ' Note: DisplayAlerts=False will automatically overwrite existing files
    wbTarget.SaveAs Filename:=targetPath & "\" & FILE_NAME_OUTPUT, FileFormat:=xlOpenXMLWorkbook
    wbTarget.Close SaveChanges:=False

    ' 7. Completion
    Application.DisplayAlerts = True
    MsgBox "Export Complete!" & vbCrLf & vbCrLf & _
           "File created: " & FILE_NAME_OUTPUT & vbCrLf & _
           "Location: " & targetPath, vbInformation, "Success"

CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set wbSource = Nothing
    Set wbTarget = Nothing
    Set wsTarget = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during export:" & vbCrLf & Err.Description, vbCritical, "Macro Error"
    Resume CleanUp
End Sub
