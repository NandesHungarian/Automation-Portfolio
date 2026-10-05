' ========================================================================================
' SCRIPT: Incoming Coverage Report Aggregator
'
' DESCRIPTION:
' This VBA macro serves as the counterpart to the Report Distributor. It automates the 
' consolidation of returned regional coverage files. It scans a designated "Incoming" folder, 
' iterates through all submitted workbooks, extracts the raw data from the input forms, 
' and aggregates everything into a centralized master sheet.
'
' PROCESS FLOW:
'   1. Clears the existing data in the Master Consolidation sheet.
'   2. Scans the target directory for incoming Excel workbooks (*.xlsx).
'   3. Opens each file and validates the presence of the required data input sheet.
'   4. Unhides all columns to ensure complete data extraction.
'   5. Identifies the data boundaries and appends the values to the master sheet.
'   6. Closes the source files without saving to preserve their original state.
'
' CONFIGURATION: Update "INCOMING_FOLDER" and sheet names before use.
' ========================================================================================

Option Explicit

Sub ConsolidateIncomingReports()
    
    ' Error handling setup
    On Error GoTo ErrorHandler

    Dim wsMaster As Worksheet
    Dim wsSource As Worksheet
    Dim fileName As String
    Dim lastRowMaster As Long
    Dim lastRowSource As Long
    Dim sourceRange As Range
    Dim wbSource As Workbook
    
    ' ---------------------------------------------------------
    ' CONFIGURATION BLOCK
    ' Set folder paths and sheet names here
    Const INCOMING_FOLDER As String = "C:\YourCompany\TradeDept\CoverageReports\Incoming\"
    Const MASTER_SHEET_NAME As String = "Consolidated_Data" ' Replaces "Bejövő táblák"
    Const INPUT_SHEET_NAME As String = "Input_Form"         ' Replaces "Kitöltendő"
    ' ---------------------------------------------------------

    ' Prevent screen flickering and speed up the macro
    Application.ScreenUpdating = False

    ' Set reference to the Master destination sheet
    Set wsMaster = ThisWorkbook.Sheets(MASTER_SHEET_NAME)

    ' Clear previous data from row 2 downwards on the Master sheet
    wsMaster.Rows("2:" & wsMaster.Rows.Count).Clear

    ' Loop through all Excel files in the specified incoming folder
    fileName = Dir(INCOMING_FOLDER & "*.xlsx")
    
    Do While fileName <> ""
        ' Open the incoming file
        Set wbSource = Workbooks.Open(INCOMING_FOLDER & fileName)
        
        ' Check if the expected input sheet exists in the file
        On Error Resume Next
        Set wsSource = wbSource.Sheets(INPUT_SHEET_NAME)
        On Error GoTo ErrorHandler ' Restore standard error handling
        
        If wsSource Is Nothing Then
            MsgBox "The file does not contain the required '" & INPUT_SHEET_NAME & "' sheet: " & fileName, vbExclamation, "Missing Sheet"
            wbSource.Close False
            GoTo SkipFile ' Skip to the next file if the sheet is missing
        End If

        ' Remove filters and unhide rows/columns so no data is missed during copy
        If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
        wsSource.Cells.EntireColumn.Hidden = False
        wsSource.Cells.EntireRow.Hidden = False

        ' Find the last row with data in Column A of the source file
        lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
        
        If lastRowSource >= 2 Then
            ' Define the range to copy (Columns A to Z, down to the last row)
            Set sourceRange = wsSource.Range("A2:Z" & lastRowSource)
        Else
            wbSource.Close False
            GoTo SkipFile ' Skip if there is no data in the file
        End If

        ' Find the next available blank row in the Master sheet
        lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row + 1
        
        ' Transfer values directly (faster than Copy/Paste)
        wsMaster.Cells(lastRowMaster, 1).Resize(sourceRange.Rows.Count, sourceRange.Columns.Count).Value = sourceRange.Value

        ' Close the incoming file without saving any changes
        wbSource.Close False

SkipFile:
        ' Reset the source variables and get the next file
        Set wsSource = Nothing
        Set wbSource = Nothing
        fileName = Dir
    Loop

    ' Remove any active filters on the Master sheet for a clean view
    If wsMaster.AutoFilterMode Then wsMaster.AutoFilterMode = False

    ' Re-enable screen updating
    Application.ScreenUpdating = True

    MsgBox "Data successfully imported and consolidated!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Stop instead of continuing with half-imported data, and close the open source file
    Application.ScreenUpdating = True
    MsgBox "An error occurred while processing '" & fileName & "':" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "The import was stopped. Please check the file and run the macro again.", vbCritical, "Macro Error"
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close False
End Sub
