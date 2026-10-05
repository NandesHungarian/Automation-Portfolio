' ========================================================================================
' SCRIPT: FCA Period Average Calculator
'
' DESCRIPTION:
' This VBA macro automates the extraction and calculation of period-average commodity 
' prices from external daily calculation workbooks. It dynamically injects location-specific 
' parameters into the external file to update its formulas, parses local textual date 
' ranges, and computes precise average prices for predefined periods.
'
' PROCESS FLOW:
'   1. Identifies and opens the target daily calculation file based on today's date.
'   2. Injects specific parameters (e.g., factory location) into the calculation sheet.
'   3. Forces an application recalculation to ensure formulas reflect the new parameters.
'   4. Parses local string-based date ranges (e.g., "2026. január") into valid Date serials.
'   5. Calculates the average price using a high-speed, in-memory variant array.
'   6. Writes the results back to the main indication sheet and resets the external file.
'
' TECHNICAL NOTES:
' - Uses Variant Arrays to load entire ranges into memory, bypassing slow cell-by-cell loops.
' - Contains a custom Date Parser to handle localized Hungarian month names seamlessly.
' ========================================================================================

Option Explicit

' -- MAIN PROCEDURE --
Sub CalculatePeriodAverages()
    ' --- Optimize Performance ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Dim wbIndication As Workbook
    Dim wsIndication As Worksheet
    Dim indicationPeriod As String
    Dim fromDate As Date, toDate As Date
    Dim parts() As String
    Dim FolderPath As String
    Dim fileName As String
    Dim fileDateStr As String
    Dim wbKalk As Workbook
    Dim wsCostDetail As Worksheet
    Dim wsPriceCalc As Worksheet
    Dim avgVal As Double
    Dim cell As Range
    
    ' ---------------------------------------------------------
    ' CONFIGURATION BLOCK
    ' Set paths, file prefixes, and sheet names here
    Const SOURCE_FOLDER As String = "C:\YourCompany\Pricing_Calculations\Product_Folder\"
    Const FILE_PREFIX As String = "DailyCalc_"
    Const INDICATION_SHEET As String = "FCA_Indication"
    Const COST_SHEET As String = "Costs detailed"
    Const CALC_SHEET As String = "Price Calculation"
    Const TARGET_LOCATION As String = "Location_A" ' Replaces specific factory name
    ' ---------------------------------------------------------
    
    ' --- Setup Indication Workbook and Sheet ---
    Set wbIndication = ThisWorkbook
    Set wsIndication = wbIndication.Sheets(INDICATION_SHEET)
    
    ' --- Setup Folder Path ---
    FolderPath = SOURCE_FOLDER
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    
    ' --- Determine Expected Filename (based on today's date) ---
    fileDateStr = Format(Date, "yyyy MM dd")
    fileName = FILE_PREFIX & fileDateStr & ".xlsx"
    
    ' --- Open or Connect to the Calculation Workbook ---
    Set wbKalk = GetWorkbook(fileName, FolderPath)
    If wbKalk Is Nothing Then
        MsgBox "The file '" & fileName & "' could not be found in the specified folder:" & vbCrLf & FolderPath, vbCritical, "File Not Found"
        GoTo Cleanup
    End If
    
    ' --- Set Worksheets ---
    Set wsCostDetail = wbKalk.Sheets(COST_SHEET)
    Set wsPriceCalc = wbKalk.Sheets(CALC_SHEET)

    ' --- Inject required parameters before calculation ---
    wsCostDetail.Range("S12").Value = TARGET_LOCATION
    wsCostDetail.Range("V31").Value = 5
    
    ' Force application calculation to ensure the injected parameters are reflected in formulas
    Application.Calculate
    
    ' --- Process predefined cell ranges (Periods) ---
    For Each cell In wsIndication.Range("D3:D5, D15:D18")
        indicationPeriod = Trim(cell.Value)
        
        If indicationPeriod = "" Then
            cell.Offset(0, 7).Value = "no data"
            GoTo NextCell
        End If
        
        parts = Split(indicationPeriod, "-")
        If UBound(parts) <> 1 Then
             cell.Offset(0, 7).Value = "format error"
             GoTo NextCell
        End If
        
        ' Parse string periods into actual Date serials
        fromDate = ParseYearMonthToDate(Trim(parts(0)))
        toDate = ParseYearMonthToDate(Trim(parts(1)))
        
        If fromDate = 0 Or toDate = 0 Then
            cell.Offset(0, 7).Value = "date error"
            GoTo NextCell
        End If

        ' Include the whole end month (e.g. "január - május" runs until 31 May)
        toDate = DateSerial(Year(toDate), Month(toDate) + 1, 0)
        
        ' --- Calculate Average using the optimized Array function ---
        avgVal = CalculateAverage_Optimized(wsPriceCalc, fromDate, toDate)
        
        If avgVal = -1 Then
            cell.Offset(0, 7).Value = "no data"
        Else
            cell.Offset(0, 7).Value = avgVal
        End If
        
NextCell:
    Next cell
    
    MsgBox "Macro executed successfully!", vbInformation, "Success"

Cleanup:
    ' --- Restore initial values in external workbook ---
    If Not wsCostDetail Is Nothing Then
        wsCostDetail.Range("S12").Value = "START"
        wsCostDetail.Range("W9").Value = 0
        
        ' Calculate one last time to reset external dependencies
        Application.Calculate
    End If
    
    ' Restore Application Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Macro Error"
    Resume Cleanup
End Sub


' --- OPTIMIZED FUNCTION: Uses Memory Arrays for high-speed calculation ---
Function CalculateAverage_Optimized(ws As Worksheet, fromDate As Date, toDate As Date) As Double
    Dim startRow As Long, endRow As Long
    Dim sumVal As Double, countVal As Long
    Dim vData As Variant
    Dim i As Long
    Dim currDate As Variant
    Dim currVal As Variant

    startRow = 14
    endRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If endRow < startRow Then
        CalculateAverage_Optimized = -1
        Exit Function
    End If

    ' Load the entire target range into memory for significantly faster processing
    vData = ws.Range("A" & startRow & ":BA" & endRow).Value
    
    sumVal = 0
    countVal = 0

    ' Loop through the memory array
    For i = 1 To UBound(vData, 1)
        currDate = vData(i, 1) ' Column A
        currVal = vData(i, 53) ' Column BA
        
        If IsDate(currDate) Then
            If currDate >= fromDate And currDate <= toDate Then
                If IsNumeric(currVal) And Not IsEmpty(currVal) Then
                    sumVal = sumVal + currVal
                    countVal = countVal + 1
                End If
            End If
        End If
    Next i

    If countVal > 0 Then
        CalculateAverage_Optimized = sumVal / countVal
    Else
        CalculateAverage_Optimized = -1
    End If

End Function


' --- HELPER FUNCTION: Safely Open or Connect to Workbook ---
Function GetWorkbook(fileName As String, FolderPath As String) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set GetWorkbook = Application.Workbooks(fileName)
    On Error GoTo 0
    
    If GetWorkbook Is Nothing Then
        If Dir(FolderPath & fileName) <> "" Then
            Set GetWorkbook = Workbooks.Open(FolderPath & fileName, UpdateLinks:=False)
        End If
    End If
End Function


' --- HELPER FUNCTION: String to Date Parser ---
Function ParseYearMonthToDate(strYM As String) As Date
    Dim yearPart As Integer
    Dim monthPart As Integer
    Dim monthsArr As Variant
    Dim i As Integer
    Dim monthName As String

    On Error GoTo ErrHandler
    
    ' Note: Using local Hungarian month names here deliberately, as the input source data strings are in Hungarian (e.g., "2026. január")
    monthsArr = Array("január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember", "október", "november", "december")
    
    yearPart = CInt(Trim(Split(strYM, ".")(0)))
    monthName = LCase(Trim(Split(strYM, " ")(1)))
    
    monthPart = 0
    For i = 0 To 11
        If monthName Like monthsArr(i) & "*" Then
            monthPart = i + 1
            Exit For
        End If
    Next i
    
    If monthPart = 0 Then
        MsgBox "Failed to parse month string: " & strYM, vbCritical, "Parsing Error"
        ParseYearMonthToDate = 0
        Exit Function
    End If
    
    ParseYearMonthToDate = DateSerial(yearPart, monthPart, 1)
    Exit Function

ErrHandler:
    MsgBox "Error in ParseYearMonthToDate function: " & Err.Description, vbCritical, "Macro Error"
    ParseYearMonthToDate = 0
End Function


' --- HELPER FUNCTION: Check if Sheet Exists ---
Function SheetExists(wb As Workbook, shName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(shName)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function
