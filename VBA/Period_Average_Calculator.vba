' ========================================================================================
' SCRIPT: Soybean Meal Period Average Calculator (Soy Ratio)
'
' DESCRIPTION:
' Sunflower meal and rapeseed meal prices are compared to the soybean meal price for the
' same periods (e.g. January - May). This macro takes the periods written as text in the
' indication workbook, finds the matching daily soybean meal prices in the daily soy
' calculation file and writes the period averages back next to each period.
'
' PROCESS FLOW:
'   1. Identifies and opens the daily soy calculation file based on today's date.
'   2. Saves the current input values, then sets the location and parameter cells.
'   3. Forces a recalculation so the formulas reflect the new inputs.
'   4. Parses text periods (e.g. "2026. január - 2026. május") into whole-month date ranges.
'   5. Averages the daily prices in each range using an in-memory array.
'   6. Writes the results next to the periods and restores the original input values.
'
' NO HARD-CODED CELLS:
' - Single cells are referenced by workbook-level Named Ranges (Formulas > Name Manager),
'   so they keep working when rows or columns are inserted.
' - The date and price columns are located by their header text, not by column letter.
' ========================================================================================

Option Explicit

' -- MAIN PROCEDURE --
Sub CalculatePeriodAverages()
    ' --- Optimize Performance ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler

    Dim wbIndication As Workbook
    Dim wbKalk As Workbook
    Dim wsPriceCalc As Worksheet
    Dim rngPeriods As Range, rngResults As Range
    Dim rngLocation As Range, rngParameter As Range
    Dim origLocation As Variant, origParameter As Variant
    Dim inputsChanged As Boolean
    Dim periodCells As Collection, resultCells As Collection
    Dim indicationPeriod As String
    Dim fromDate As Date, toDate As Date
    Dim parts() As String
    Dim FolderPath As String
    Dim fileName As String
    Dim avgVal As Double
    Dim cell As Range, outCell As Range
    Dim headerRow As Long, dateCol As Long, priceCol As Long
    Dim i As Long

    ' ---------------------------------------------------------
    ' CONFIGURATION BLOCK
    ' Paths, file prefix, sheet name, Named Ranges and column headers
    Const SOURCE_FOLDER As String = "C:\YourCompany\Pricing_Calculations\Product_Folder\"
    Const FILE_PREFIX As String = "DailyCalc_"
    Const CALC_SHEET As String = "Price Calculation"
    Const TARGET_LOCATION As String = "Location_A" ' Replaces specific factory name
    Const PARAMETER_VALUE As Long = 5

    ' Named Ranges in the daily soy calculation file
    Const NAME_LOCATION As String = "Calc_Location"    ' location selector cell
    Const NAME_PARAMETER As String = "Calc_Parameter"  ' extra calculation input cell

    ' Named Ranges in this (indication) workbook. Both must have the same layout.
    Const NAME_PERIODS As String = "Soy_Periods"       ' text periods, e.g. "2026. január - 2026. május"
    Const NAME_RESULTS As String = "Soy_Averages"      ' cells receiving the averages

    ' Column headers in the price calculation sheet
    Const DATE_HEADER As String = "Date"
    Const PRICE_HEADER As String = "SBM Price"
    ' ---------------------------------------------------------

    Set wbIndication = ThisWorkbook
    Set rngPeriods = GetNamedRange(wbIndication, NAME_PERIODS)
    Set rngResults = GetNamedRange(wbIndication, NAME_RESULTS)
    If rngPeriods Is Nothing Or rngResults Is Nothing Then GoTo Cleanup

    Set periodCells = CellsInOrder(rngPeriods)
    Set resultCells = CellsInOrder(rngResults)
    If periodCells.Count <> resultCells.Count Then
        MsgBox "'" & NAME_PERIODS & "' and '" & NAME_RESULTS & "' must contain the same number of cells.", vbCritical, "Setup Error"
        GoTo Cleanup
    End If

    ' --- Setup Folder Path ---
    FolderPath = SOURCE_FOLDER
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"

    ' --- Determine Expected Filename (based on today's date) ---
    fileName = FILE_PREFIX & Format(Date, "yyyy MM dd") & ".xlsx"

    ' --- Open or Connect to the Calculation Workbook ---
    Set wbKalk = GetWorkbook(fileName, FolderPath)
    If wbKalk Is Nothing Then
        MsgBox "The file '" & fileName & "' could not be found in the specified folder:" & vbCrLf & FolderPath, vbCritical, "File Not Found"
        GoTo Cleanup
    End If

    Set wsPriceCalc = wbKalk.Sheets(CALC_SHEET)
    Set rngLocation = GetNamedRange(wbKalk, NAME_LOCATION)
    Set rngParameter = GetNamedRange(wbKalk, NAME_PARAMETER)
    If rngLocation Is Nothing Or rngParameter Is Nothing Then GoTo Cleanup

    ' --- Locate the date and price columns by header text ---
    If Not FindHeader(wsPriceCalc, DATE_HEADER, headerRow, dateCol) Then GoTo Cleanup
    If Not FindHeader(wsPriceCalc, PRICE_HEADER, 0, priceCol) Then GoTo Cleanup

    ' --- Save original inputs, then inject the required parameters ---
    origLocation = rngLocation.Value
    origParameter = rngParameter.Value
    inputsChanged = True
    rngLocation.Value = TARGET_LOCATION
    rngParameter.Value = PARAMETER_VALUE

    ' Force application calculation to ensure the injected parameters are reflected in formulas
    Application.Calculate

    ' --- Process each period ---
    For i = 1 To periodCells.Count
        Set cell = periodCells(i)
        Set outCell = resultCells(i)
        indicationPeriod = Trim(cell.Value)

        If indicationPeriod = "" Then
            outCell.Value = "no data"
            GoTo NextCell
        End If

        parts = Split(indicationPeriod, "-")
        If UBound(parts) <> 1 Then
             outCell.Value = "format error"
             GoTo NextCell
        End If

        ' Parse string periods into actual Date serials
        fromDate = ParseYearMonthToDate(Trim(parts(0)))
        toDate = ParseYearMonthToDate(Trim(parts(1)))

        If fromDate = 0 Or toDate = 0 Then
            outCell.Value = "date error"
            GoTo NextCell
        End If

        ' Include the whole end month (e.g. "január - május" runs until 31 May)
        toDate = DateSerial(Year(toDate), Month(toDate) + 1, 0)

        ' --- Calculate Average using the optimized Array function ---
        avgVal = CalculateAverage_Optimized(wsPriceCalc, headerRow + 1, dateCol, priceCol, fromDate, toDate)

        If avgVal = -1 Then
            outCell.Value = "no data"
        Else
            outCell.Value = avgVal
        End If

NextCell:
    Next i

    MsgBox "Macro executed successfully!", vbInformation, "Success"

Cleanup:
    ' --- Restore the original input values in the external workbook ---
    If inputsChanged Then
        rngLocation.Value = origLocation
        rngParameter.Value = origParameter
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


' --- Average calculation: reads the columns into memory arrays, much faster than reading cell by cell ---
Function CalculateAverage_Optimized(ws As Worksheet, firstDataRow As Long, dateCol As Long, priceCol As Long, _
                                    fromDate As Date, toDate As Date) As Double
    Dim lastRow As Long
    Dim sumVal As Double, countVal As Long
    Dim vDates As Variant, vPrices As Variant
    Dim i As Long

    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row
    If lastRow <= firstDataRow Then
        CalculateAverage_Optimized = -1
        Exit Function
    End If

    ' Load both columns into memory for significantly faster processing
    vDates = ws.Range(ws.Cells(firstDataRow, dateCol), ws.Cells(lastRow, dateCol)).Value
    vPrices = ws.Range(ws.Cells(firstDataRow, priceCol), ws.Cells(lastRow, priceCol)).Value

    For i = 1 To UBound(vDates, 1)
        If IsDate(vDates(i, 1)) Then
            If vDates(i, 1) >= fromDate And vDates(i, 1) <= toDate Then
                If IsNumeric(vPrices(i, 1)) And Not IsEmpty(vPrices(i, 1)) Then
                    sumVal = sumVal + vPrices(i, 1)
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


' --- HELPER FUNCTION: Get a workbook-level Named Range, with a clear message if missing ---
Function GetNamedRange(wb As Workbook, rangeName As String) As Range
    On Error Resume Next
    Set GetNamedRange = wb.Names(rangeName).RefersToRange
    On Error GoTo 0
    If GetNamedRange Is Nothing Then
        MsgBox "Named Range '" & rangeName & "' was not found in " & wb.Name & "." & vbCrLf & _
               "Create it under Formulas > Name Manager.", vbCritical, "Setup Error"
    End If
End Function


' --- HELPER FUNCTION: Find a column by its header text ---
Function FindHeader(ws As Worksheet, headerText As String, ByRef headerRow As Long, ByRef headerCol As Long) As Boolean
    Dim found As Range
    Set found = ws.Cells.Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If found Is Nothing Then
        MsgBox "Column header '" & headerText & "' was not found on sheet '" & ws.Name & "'.", vbCritical, "Setup Error"
        FindHeader = False
    Else
        headerRow = found.Row
        headerCol = found.Column
        FindHeader = True
    End If
End Function


' --- HELPER FUNCTION: List the cells of a (possibly multi-area) range in reading order ---
Function CellsInOrder(rng As Range) As Collection
    Dim result As New Collection
    Dim area As Range, c As Range
    For Each area In rng.Areas
        For Each c In area.Cells
            result.Add c
        Next c
    Next area
    Set CellsInOrder = result
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
