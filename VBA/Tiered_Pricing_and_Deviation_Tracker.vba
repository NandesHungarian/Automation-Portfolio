' ========================================================================================
' SCRIPT: Tiered Pricing Email Dispatcher & Logger
'
' DESCRIPTION:
' This VBA macro automates the daily distribution of localized commodity prices to different 
' client tiers (Standard, Broker, and End-User) using a predefined Outlook template. 
' It dynamically injects distinct pricing tables into the email body and logs the quoted 
' prices for specific delivery periods, along with the price deviations against a baseline 
' reference point, into a tracking sheet for audit and reporting.
'
' PROCESS FLOW:
'   1. Locates the Outlook email template across multiple possible network/local drives.
'   2. Initializes Outlook and sets up the generic parameters (Date, User, Subject).
'   3. Generates 3 separate BCC emails for each client tier by inserting specific data ranges.
'   4. Uses advanced Clipboard Error Handling (Retry Logic) to prevent COM/Clipboard locks 
'      (Error 4605) during the cross-application table pasting process (Excel to WordEditor).
'   5. Calculates the deviation (Spread/Premium) between quoted prices and the reference 
'      point for different delivery periods (e.g., ASO, NDJ, FMA, MJJ) and logs them.
'
' CONFIGURATION: Update Named Ranges and template paths before execution.
' ========================================================================================

Option Explicit

Public Sub SendTieredPricingEmails()

    ' === Variable Declarations ===
    Dim logSheet As Worksheet, ws As Worksheet
    Dim userCode As String, logSubject As String, todayDate As String
    Dim outlookApp As Object
    Dim templatePath As String
    
    ' ---------------------------------------------------------
    ' CONFIGURATION
    Const LOCATION_NAME As String = "Factory" ' Replaces specific location like Bruck
    Const TEMPLATE_NAME As String = "Daily_Prices_Template.oft"
    ' ---------------------------------------------------------
    
    ' === Enable Error Handling ===
    On Error GoTo ErrorHandler

    ' === Basic Setup ===
    Set ws = ThisWorkbook.Sheets("Rates")
    Set logSheet = ThisWorkbook.Sheets("Log")
    userCode = Environ("username")
    todayDate = Format(Date, "dd mm yyyy")
    logSubject = LOCATION_NAME & " prices " & todayDate

    ' === Locate Template File using Helper Function ===
    templatePath = FindTemplatePath(TEMPLATE_NAME)
    If templatePath = "" Then
        MsgBox "Cannot find the Outlook template file in any of the known directories!", vbCritical, "File Not Found"
        Exit Sub
    End If

    ' === Initialize Outlook Application (Singleton approach) ===
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler ' Restore standard error handling
    
    If outlookApp Is Nothing Then
        MsgBox "Outlook application is not available!", vbCritical, "Initialization Error"
        Exit Sub
    End If

    ' =======================================================
    ' ===           EMAIL GENERATION SECTION              ===
    ' =======================================================

    ' 1. Email: Standard Partners (Sent as BCC)
    Call CreateTieredEmail(outlookApp, templatePath, _
                          ws.Range("Standard_ProductA"), _
                          ws.Range("Standard_ProductB"), _
                          ws.Range("Standard_Email").Value, _
                          todayDate, LOCATION_NAME, _
                          True) ' True = Use BCC

    ' 2. Email: Broker Partners (Sent as BCC)
    Call CreateTieredEmail(outlookApp, templatePath, _
                          ws.Range("Broker_ProductA"), _
                          ws.Range("Broker_ProductB"), _
                          ws.Range("Broker_Email").Value, _
                          todayDate, LOCATION_NAME, _
                          True) 
                          
    ' 3. Email: End-User Partners (Sent as BCC)
    Call CreateTieredEmail(outlookApp, templatePath, _
                          ws.Range("EndUser_ProductA"), _
                          ws.Range("EndUser_ProductB"), _
                          ws.Range("EndUser_Email").Value, _
                          todayDate, LOCATION_NAME, _
                          True) 

    ' =======================================================
    ' ===                 LOGGING SECTION                 ===
    ' =======================================================
    Dim lastRow As Long
    lastRow = logSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

    ' --- Read values from Named Ranges for different periods ---
    ' ASO = Aug-Sep-Oct | NDJ = Nov-Dec-Jan | FMA = Feb-Mar-Apr | MJJ = May-Jun-Jul
    Dim valPriceASO As Double, valRef_ASO As Double
    Dim valPriceNDJ As Double, valRef_NDJ As Double
    Dim valPriceFMA As Double, valRef_FMA As Double
    Dim valPriceMJJ As Double, valRef_MJJ As Double

    valPriceASO = ws.Range("Price_ASO").Value
    valRef_ASO = ws.Range("LR_ASO").Value
    
    valPriceNDJ = ws.Range("Price_NDJ").Value
    valRef_NDJ = ws.Range("LR_NDJ").Value
    
    valPriceFMA = ws.Range("Price_FMA").Value
    valRef_FMA = ws.Range("LR_FMA").Value
    
    valPriceMJJ = ws.Range("Price_MJJ").Value
    valRef_MJJ = ws.Range("LR_MJJ").Value

    ' --- Write Data to Log Sheet ---
    With logSheet
        ' Basic Data (Columns A-D)
        .Cells(lastRow, "A").Value = Date
        .Cells(lastRow, "B").Value = Time
        .Cells(lastRow, "C").Value = logSubject
        .Cells(lastRow, "D").Value = userCode

        ' FMA Period Data (K, L, M) -> Period Price, Reference Point, Deviation vs Reference
        .Cells(lastRow, "K").Value = valPriceFMA
        .Cells(lastRow, "L").Value = valRef_FMA
        .Cells(lastRow, "M").Value = valPriceFMA - valRef_FMA

        ' MJJ Period Data (N, O, P) -> Period Price, Reference Point, Deviation vs Reference
        .Cells(lastRow, "N").Value = valPriceMJJ
        .Cells(lastRow, "O").Value = valRef_MJJ
        .Cells(lastRow, "P").Value = valPriceMJJ - valRef_MJJ
        
        ' ASO Period Data (Q, R, S) -> Period Price, Reference Point, Deviation vs Reference
        .Cells(lastRow, "Q").Value = valPriceASO
        .Cells(lastRow, "R").Value = valRef_ASO
        .Cells(lastRow, "S").Value = valPriceASO - valRef_ASO
        
        ' NDJ Period Data (T, U, V) -> Period Price, Reference Point, Deviation vs Reference
        .Cells(lastRow, "T").Value = valPriceNDJ
        .Cells(lastRow, "U").Value = valRef_NDJ
        .Cells(lastRow, "V").Value = valPriceNDJ - valRef_NDJ

        ' Clear any lingering comments in older ranges for safety
        On Error Resume Next
        .Cells(lastRow, "E").Comment.Delete
        .Cells(lastRow, "F").Comment.Delete
        On Error GoTo 0
        
    End With

    MsgBox "Emails successfully drafted with tables injected." & vbNewLine & "Pricing data has been logged.", vbInformation, "Success"

' === Cleanup and Exit ===
ExitProcedure:
    Set outlookApp = Nothing
    Set ws = Nothing
    Set logSheet = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during execution!" & vbNewLine & vbNewLine & _
           "Error Code: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description, vbCritical, "Runtime Error"
    GoTo ExitProcedure
End Sub


' =========================================================================================
' === HELPER SUB: DRAFTS A SINGLE EMAIL WITH DATA TABLES                                ===
' === (Includes advanced Clipboard retry logic to prevent Error 4605 during paste)      ===
' =========================================================================================
Private Sub CreateTieredEmail(ByVal outlookApp As Object, ByVal templatePath As String, _
                             ByVal rngTable1 As Range, ByVal rngTable2 As Range, _
                             ByVal recipientAddress As String, ByVal todayDate As String, _
                             ByVal locationName As String, _
                             Optional ByVal useBCC As Boolean = True)
                             
    Dim mailItem As Object, inspector As Object, wdDoc As Object, findObj As Object
    Dim attempts As Integer

    ' Create new email from standard template
    Set mailItem = outlookApp.CreateItemFromTemplate(templatePath)
    If mailItem Is Nothing Then Exit Sub

    mailItem.Display
    DoEvents

    Set inspector = mailItem.GetInspector
    Set wdDoc = inspector.WordEditor
    
    If Not wdDoc Is Nothing Then
        ' --- Insert First Table with Clipboard Protection ---
        Set findObj = wdDoc.Content.Find
        With findObj
            .ClearFormatting
            .Text = "{table1}"
            If .Execute Then
                attempts = 0
                ' Retry loop to bypass Windows Clipboard lock issues
                Do While attempts < 3
                    On Error Resume Next
                    rngTable1.Copy
                    findObj.Parent.PasteExcelTable False, False, False
                    If Err.Number = 0 Then
                        On Error GoTo 0
                        Exit Do ' Successful paste
                    End If
                    Err.Clear
                    On Error GoTo 0
                    attempts = attempts + 1
                    DoEvents
                    Application.Wait (Now + TimeValue("0:00:01")) ' Wait 1 second and retry
                Loop
                Application.CutCopyMode = False
            End If
        End With

        ' --- Insert Second Table with Clipboard Protection ---
        Set findObj = wdDoc.Content.Find
        With findObj
            .ClearFormatting
            .Text = "{table2}"
            If .Execute Then
                attempts = 0
                Do While attempts < 3
                    On Error Resume Next
                    rngTable2.Copy
                    findObj.Parent.PasteExcelTable False, False, False
                    If Err.Number = 0 Then
                        On Error GoTo 0
                        Exit Do ' Successful paste
                    End If
                    Err.Clear
                    On Error GoTo 0
                    attempts = attempts + 1
                    DoEvents
                    Application.Wait (Now + TimeValue("0:00:01")) ' Wait 1 second and retry
                Loop
                Application.CutCopyMode = False
            End If
        End With
    Else
        MsgBox "Failed to load WordEditor. One of the emails must be modified manually.", vbWarning, "Editor Error"
    End If

    ' Configure addressing and subject
    With mailItem
        If useBCC Then
            .To = outlookApp.Session.CurrentUser.Address ' Send to self
            .BCC = recipientAddress ' Hide recipients
        Else
            .To = recipientAddress
            .BCC = ""
        End If
        
        .Subject = locationName & " prices " & todayDate
    End With

    ' Cleanup local objects
    Set mailItem = Nothing
    Set inspector = Nothing
    Set wdDoc = Nothing
    Set findObj = Nothing
End Sub


' =========================================================================================
' === HELPER FUNCTION: DYNAMICALLY LOCATES THE OUTLOOK TEMPLATE ACROSS ENVIRONMENTS     ===
' =========================================================================================
Private Function FindTemplatePath(ByVal fileName As String) As String
    Dim possiblePaths As Variant
    Dim pathVariant As Variant
    Dim username As String
    username = Environ("username")

    ' Array of potential network drives and local directories where the template might reside
    possiblePaths = Array( _
        "C:\Users\Public\TradeDept\Pricing\Templates", _
        "C:\CompanyShares\TradeDept\Pricing\Templates", _
        "C:\Users\" & username & "\TradeDept\Pricing\Templates", _
        "C:\Users\" & username & "\Documents\Company_Local\Templates" _
    )

    ' Loop through paths to find the valid existing template file
    For Each pathVariant In possiblePaths
        If Dir(pathVariant & "\" & fileName) <> "" Then
            FindTemplatePath = pathVariant & "\" & fileName
            Exit Function
        End If
    Next pathVariant
    
    ' Return empty string if not found
    FindTemplatePath = ""
End Function
