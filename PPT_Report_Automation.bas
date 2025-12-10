Option Explicit

' ==============================================================================
' Module:       PPT_Report_Automation
' Description:  Automates the weekly/monthly reporting update of a PowerPoint deck.
'
' Core Features:
'   1. Cross-App Control: Attaches to or launches PowerPoint via OLE Automation.
'   2. Configuration-Driven: Uses an Excel Table ("Source2") to map Excel Sheets
'      to specific PowerPoint Slides (via Slide Title).
'   3. Resilient Update: Locates the data based on a keyword ("ACCURACY REPORT SUMMARY").
'   4. Intelligent Paste: Captures the position/size of the old data object before
'      deleting it, ensuring the new pasted data lands in the exact same spot.
'
' Author:       [Your Name]
' Date:         2025-12-10
' ==============================================================================

Sub Update_PowerPoint_Reports()
    
    ' --- CONFIGURATION ---
    Const CONFIG_SHEET_NAME As String = "Macro"
    Const CONFIG_TABLE_NAME As String = "Source2"
    Const PPT_PATH_NAME As String = "Power1" ' Named range containing the PowerPoint file path
    Const DATA_ANCHOR_KEYWORD As String = "ACCURACY REPORT SUMMARY"
    ' ---------------------

    ' Object Variables
    Dim pptApp As Object, pptPres As Object, targetSlide As Object, pptSlide As Object
    Dim pptShape As Object, s As Object
    Dim wb As Workbook
    Dim wsConfig As Worksheet, wsData As Worksheet
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim copyRange As Range, foundCell As Range
    
    ' Value Variables
    Dim pptPath As String
    Dim sheetName As String, slideMatchStr As String
    Dim oldTop As Single, oldLeft As Single, oldWidth As Single
    Dim i As Long
    
    Set wb = ThisWorkbook
    
    ' 1. INITIALIZATION & VALIDATION
    On Error Resume Next
    Set wsConfig = wb.Sheets(CONFIG_SHEET_NAME)
    Set tbl = wsConfig.ListObjects(CONFIG_TABLE_NAME)
    pptPath = wb.Names(PPT_PATH_NAME).RefersToRange.Value
    On Error GoTo 0
    
    If tbl Is Nothing Or pptPath = "" Then
        MsgBox "Configuration Error: Check '" & CONFIG_TABLE_NAME & "' on '" & CONFIG_SHEET_NAME & "' sheet or the '" & PPT_PATH_NAME & "' file path.", vbCritical
        Exit Sub
    End If
    
    ' 2. OPEN POWERPOINT APPLICATION & PRESENTATION
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application") ' Attempt to attach to open instance
    If Err.Number <> 0 Then Set pptApp = CreateObject("PowerPoint.Application") ' Create new instance
    On Error GoTo 0
    pptApp.Visible = True
    
    On Error Resume Next
    Set pptPres = pptApp.Presentations.Open(pptPath)
    If Err.Number <> 0 Then
        MsgBox "Could not open PowerPoint file at path: " & pptPath, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 3. LOOP THROUGH MAPPING TABLE
    For Each tblRow In tbl.ListRows
        sheetName = Trim(tblRow.Range(1, 1).Value)
        slideMatchStr = Trim(tblRow.Range(1, 2).Value)
        
        If sheetName <> "" And slideMatchStr <> "" Then
            
            ' --- A. PREPARE EXCEL DATA (Find Range) ---
            Set wsData = Nothing
            On Error Resume Next
            Set wsData = wb.Sheets(sheetName)
            On Error GoTo 0
            
            If Not wsData Is Nothing Then
                Set foundCell = wsData.Cells.Find(What:=DATA_ANCHOR_KEYWORD, LookAt:=xlPart)
                
                If Not foundCell Is Nothing Then
                    ' Define copy area (CurrentRegion is usually the fastest way to get contiguous data)
                    Set copyRange = foundCell.CurrentRegion
                Else
                    Debug.Print "Marker '" & DATA_ANCHOR_KEYWORD & "' not found on sheet: " & sheetName
                    GoTo NextIteration
                End If
            Else
                Debug.Print "Sheet not found: " & sheetName
                GoTo NextIteration
            End If
            
            ' --- B. FIND TARGET SLIDE (Search by Title) ---
            Set targetSlide = Nothing
            For Each pptSlide In pptPres.Slides
                If pptSlide.Shapes.HasTitle Then
                    If InStr(1, pptSlide.Shapes.Title.TextFrame.TextRange.Text, slideMatchStr, vbTextCompare) > 0 Then
                        Set targetSlide = pptSlide
                        Exit For
                    End If
                End If
            Next pptSlide
            
            ' --- C. CLEANUP OLD & PASTE NEW ---
            If Not targetSlide Is Nothing Then
                
                ' Default position captures (in case old shape is not found)
                oldTop = 100: oldLeft = 50: oldWidth = 600
                
                ' Delete Old Loop (Captures position before deleting)
                For i = targetSlide.Shapes.Count To 1 Step -1
                    Set s = targetSlide.Shapes(i)
                    ' Delete Picture (Type 13) or Table (Type 19) or our specifically tagged shape
                    If s.Type = 13 Or s.Type = 19 Or s.Name = "MacroTable" Then
                        oldTop = s.Top
                        oldLeft = s.Left
                        oldWidth = s.Width
                        s.Delete
                    End If
                Next i
                
                ' 1. COPY FRESH DATA
                copyRange.Copy
                DoEvents
                
                ' 2. PASTE AS ENHANCED METAFILE (Non-Select, Stable Paste)
                On Error Resume Next
                ' DataType:=2 is ppPasteEnhancedMetafile, or use 10 for Linked Excel Chart
                Set pptShape = targetSlide.Shapes.PasteSpecial(DataType:=2)(1) 
                On Error GoTo 0
                
                ' 3. POSITION & RENAME
                If Not pptShape Is Nothing Then
                    With pptShape
                        .Name = "MacroTable"
                        .LockAspectRatio = msoTrue ' Excel equivalent of -1
                        .Top = oldTop
                        .Left = oldLeft
                        .Width = oldWidth
                    End With
                End If
                
                Application.CutCopyMode = False
            End If
        End If
NextIteration:
    Next tblRow
    
    ' 4. FINAL CLEANUP
    pptPres.Save
    ' pptPres.Close ' Optionally close the presentation
    
    MsgBox "PowerPoint Update Complete!", vbInformation
    
End Sub
