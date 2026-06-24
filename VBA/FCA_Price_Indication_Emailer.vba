Sub SendCommodityPrices_InsertPicture()
    ' Error handling
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim imagePath As String
    Dim chartObj As ChartObject
    Dim rng As Range
    Dim logSheet As Worksheet
    Dim lastRow As Long
    Dim currentDate As String
    Dim currentTime As String
    Dim inspector As Object
    Dim wordDoc As Object
    Dim selection As Object
    
    ' Variables for smart search
    Dim findRng As Range
    Dim foundCell As Range
    Dim srcRow As Long

    ' ---------------------------------------------------------
    ' CONFIGURATION
    ' Set your specific email, sheet names, and email subject here
    Const SEND_TO_EMAIL As String = "trading.team@example.com"
    Const SOURCE_SHEET_NAME As String = "FCA_Indication"
    Const LOG_SHEET_NAME As String = "Log"
    Const SUBJECT_PREFIX As String = "Factory Commodity Prices - "
    ' ---------------------------------------------------------

    currentDate = Format(Date, "yyyy.mm.dd")
    currentTime = Format(Now, "hh:nn:ss")

    Set ws = ThisWorkbook.Sheets(SOURCE_SHEET_NAME)
    
    ' Column Z is the search range for the logging criteria
    Set findRng = ws.Range("Z:Z")
    
    ' The range to be exported as a picture (modify if table size changes!)
    Set rng = ws.Range("A23:F30")

    ' Set temporary image path and delete if it already exists
    imagePath = Environ("TEMP") & "\PriceTable.png"
    If Dir(imagePath) <> "" Then Kill imagePath

    ' Copy the range as a picture and export it via a temporary ChartObject
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Set chartObj = ws.ChartObjects.Add(Left:=rng.Left, Top:=rng.Top, Width:=rng.Width, Height:=rng.Height)
    chartObj.Activate
    chartObj.Chart.Paste
    chartObj.Chart.Export fileName:=imagePath, FilterName:="PNG"
    chartObj.Delete

    ' Initialize Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    With OutlookMail
        .To = SEND_TO_EMAIL
        .Subject = SUBJECT_PREFIX & currentDate
        .Display

        Set inspector = .GetInspector
        Set wordDoc = inspector.WordEditor
        Set selection = wordDoc.Application.selection

        ' Write Email Body
        selection.TypeText Text:="Dear Team," & vbCrLf & vbCrLf
        selection.TypeText Text:="Please find today's prices below. For larger volume requests, please coordinate with our trading desk." & vbCrLf & vbCrLf

        ' Insert the exported image inline
        selection.InlineShapes.AddPicture fileName:=imagePath, LinkToFile:=False, SaveWithDocument:=True
        selection.TypeParagraph
        selection.TypeParagraph
    End With

    ' --- START LOGGING ---
    Set logSheet = ThisWorkbook.Sheets(LOG_SHEET_NAME)
    lastRow = logSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

    With logSheet
        ' 1. Basic Data
        .Cells(lastRow, "A").Value = currentDate
        .Cells(lastRow, "B").Value = currentTime
        .Cells(lastRow, "C").Value = SEND_TO_EMAIL
        .Cells(lastRow, "D").Value = SUBJECT_PREFIX & currentDate
        .Cells(lastRow, "E").Value = Environ("Username")

        ' --- 2. PRODUCT A (SFM) DATA ---
        
        ' SFM Position 1 -> Columns N, O, P, Q
        Set foundCell = findRng.Find(What:="SFM_1", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "N").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "O").Value = ws.Cells(srcRow, 7).Value
            .Cells(lastRow, "P").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "Q").Value = ws.Cells(srcRow, 9).Value
        End If

        ' SFM Position 2 -> Columns R, S, T, U
        Set foundCell = findRng.Find(What:="SFM_2", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "R").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "S").Value = ws.Cells(srcRow, 7).Value
            .Cells(lastRow, "T").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "U").Value = ws.Cells(srcRow, 9).Value
        End If
        
        ' SFM Position 3 -> Columns V, W, X, Y
        Set foundCell = findRng.Find(What:="SFM_3", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "V").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "W").Value = ws.Cells(srcRow, 7).Value
            .Cells(lastRow, "X").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "Y").Value = ws.Cells(srcRow, 9).Value
        End If

        ' SFM Position 4 -> Columns Z, AA, AB, AC
        Set foundCell = findRng.Find(What:="SFM_4", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "Z").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "AA").Value = ws.Cells(srcRow, 7).Value
            .Cells(lastRow, "AB").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "AC").Value = ws.Cells(srcRow, 9).Value
        End If

        ' --- 3. PRODUCT B (RSM) DATA ---
        
        ' RSM Position 3 -> Columns AT, AU, AV, AW
        Set foundCell = findRng.Find(What:="RSM_3", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "AT").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "AU").Value = ws.Cells(srcRow, 6).Value
            .Cells(lastRow, "AV").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "AW").Value = ws.Cells(srcRow, 9).Value
        End If

        ' RSM Position 4 -> Columns AX, AY, AZ, BA
        Set foundCell = findRng.Find(What:="RSM_4", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "AX").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "AY").Value = ws.Cells(srcRow, 6).Value
            .Cells(lastRow, "AZ").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "BA").Value = ws.Cells(srcRow, 9).Value
        End If

        ' RSM Position 5 -> Columns BB, BC, BD, BE
        Set foundCell = findRng.Find(What:="RSM_5", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "BB").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "BC").Value = ws.Cells(srcRow, 6).Value
            .Cells(lastRow, "BD").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "BE").Value = ws.Cells(srcRow, 9).Value
        End If
        
        ' RSM Position 6 -> Columns BF, BG, BH, BI
        Set foundCell = findRng.Find(What:="RSM_6", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "BF").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "BG").Value = ws.Cells(srcRow, 6).Value
            .Cells(lastRow, "BH").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "BI").Value = ws.Cells(srcRow, 9).Value
        End If
        
        ' RSM Position 7 -> Columns BJ, BK, BL, BM
        Set foundCell = findRng.Find(What:="RSM_7", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "BJ").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "BK").Value = ws.Cells(srcRow, 6).Value
            .Cells(lastRow, "BL").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "BM").Value = ws.Cells(srcRow, 9).Value
        End If
        
        ' RSM Position 8 -> Columns BN, BO, BP, BQ
        Set foundCell = findRng.Find(What:="RSM_8", LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            srcRow = foundCell.Row
            .Cells(lastRow, "BN").Value = ws.Cells(srcRow, 5).Value
            .Cells(lastRow, "BO").Value = ws.Cells(srcRow, 6).Value
            .Cells(lastRow, "BP").Value = ws.Cells(srcRow, 8).Value
            .Cells(lastRow, "BQ").Value = ws.Cells(srcRow, 9).Value
        End If

    End With

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description & vbCrLf & "Error at line: " & Erl, vbCritical, "Macro Error"
    Resume CleanExit

End Sub
