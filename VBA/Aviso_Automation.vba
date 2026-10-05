' =========================================================================================
' MULTI-USER OUTLOOK AVISO AUTOMATION MACRO
' -----------------------------------------------------------------------------------------
' Author: Nándor Magyar
' Description: 
' This VBA macro automates the generation of notification (Aviso) emails for logistics partners.
' It dynamically detects the current Windows user to resolve network/OneDrive paths, locates 
' the appropriate Outlook template (.oft) based on the partner's name, extracts specific 
' data ranges from Excel, formats them into HTML tables, and merges them into the email body.
'
' Business Impact: Saves approximately 1 hour of manual administrative work daily per user.
' =========================================================================================

Option Explicit

Sub SendEmailsByPartnerList()

    Dim OutApp As Object, OutMail As Object
    Dim wsAviso As Worksheet, wsHelperTab As Worksheet
    Dim rngTable1 As Range, rngTable2 As Range
    Dim htmlBody1 As String, htmlBody2 As String
    Dim templatePath As String
    Dim lastRowT1 As Long, lastRowT2 As Long, lastRowMain As Long
    Dim cellB7Value As Variant, cellS7Value As Variant
    Dim partnerName As String
    Dim i As Long
    Dim successList As String, errorList As String
    
    ' ==============================================================================
    ' CONFIGURATION BLOCK (Dynamic paths to support execution on any local machine)
    ' ==============================================================================
    
    ' 1. Dynamically retrieve the current Windows User ID
    Dim currentUser As String
    currentUser = Environ("USERNAME")
    
    ' 2. Define dynamic template directories (Anonymized for public repository)
    Dim BasePath1 As String
    Dim BasePath2 As String
    Dim BasePath3 As String
    
    ' Primary Logistics Templates folder
    BasePath1 = "C:\Users\" & currentUser & "\Company\Logistics_Execution\Documents\Aviso_Templates\"
    ' Alternative folder for Trade Department
    BasePath2 = "C:\Users\" & currentUser & "\Company\Trade_Dept\Price_Calc\Email_Templates\"
    ' Secondary Logistics fallback folder
    BasePath3 = "C:\Users\" & currentUser & "\Company\Logistics_Execution\Secondary_Aviso_Templates\"
    
    ' Worksheet names and boundary definitions
    Const SheetNameAviso As String = "Aviso"
    Const SheetNameHelper As String = "Segédtábla"
    Const Table1StartCell As String = "B6"
    Const Table2StartCell As String = "S6"
    Const Table1EndCol As String = "H"
    Const Table2EndCol As String = "V"
    
    ' ==============================================================================

    ' Initialize Outlook Application (Hook into active instance or create a new one)
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If Err.Number <> 0 Then
        Set OutApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    Set wsAviso = ThisWorkbook.Sheets(SheetNameAviso)
    Set wsHelperTab = ThisWorkbook.Sheets(SheetNameHelper)

    successList = ""
    errorList = ""

    ' Determine the last row of the partner list in column J
    lastRowMain = wsHelperTab.Cells(wsHelperTab.Rows.Count, "J").End(xlUp).Row

    ' Iterate through the partner list to process each pending notification
    For i = 1 To lastRowMain
        partnerName = wsHelperTab.Cells(i, "J").Value

        ' Skip empty rows and the grand total row
        If partnerName <> "Végösszeg" And Trim(partnerName) <> "" Then
            
            ' Trigger Excel calculation by inputting partner name into master cells
            wsAviso.Range("C4").Value = partnerName
            wsAviso.Range("T4").Value = partnerName

            ' --- TEMPLATE DISCOVERY (.oft) ---
            ' Fallback logic: Searches through designated folders to find the partner's template
            templatePath = FindTemplateFile(BasePath1, partnerName)
            
            If templatePath = "" Then
                templatePath = FindTemplateFile(BasePath2, partnerName)
            End If
            
            If templatePath = "" Then
                templatePath = FindTemplateFile(BasePath3, partnerName)
            End If

            ' Log error and skip iteration if no template is found
            If templatePath = "" Then
                Debug.Print "[ERROR] Template not found for partner: " & partnerName
                errorList = errorList & vbCrLf & "- " & partnerName & " (Missing Template)"
                GoTo NextPartner
            End If

            ' --- FIRST TABLE HTML EXTRACTION ---
            cellB7Value = wsAviso.Range("B7").Value
            If Trim(CStr(cellB7Value)) <> "" Then
                lastRowT1 = wsAviso.Cells(wsAviso.Rows.Count, "B").End(xlUp).Row
                If lastRowT1 < 7 Then lastRowT1 = 7
                Set rngTable1 = wsAviso.Range(Table1StartCell & ":" & Table1EndCol & lastRowT1)
                htmlBody1 = ConvertRangeToHTML(rngTable1)
            Else
                htmlBody1 = ""
            End If

            ' --- SECOND TABLE HTML EXTRACTION ---
            cellS7Value = wsAviso.Range("S7").Value
            If Trim(CStr(cellS7Value)) <> "" Then
                lastRowT2 = wsAviso.Cells(wsAviso.Rows.Count, "S").End(xlUp).Row
                If lastRowT2 < 7 Then lastRowT2 = 7
                Set rngTable2 = wsAviso.Range(Table2StartCell & ":" & Table2EndCol & lastRowT2)
                htmlBody2 = ConvertRangeToHTML(rngTable2)
            Else
                htmlBody2 = ""
            End If

            ' --- OUTLOOK EMAIL GENERATION ---
            On Error Resume Next
            Set OutMail = OutApp.CreateItemFromTemplate(templatePath)
            If Err.Number <> 0 Then
                Debug.Print "[ERROR] Failed to generate email from template: " & templatePath
                Err.Clear
                errorList = errorList & vbCrLf & "- " & partnerName & " (Corrupted Template)"
                GoTo NextPartner
            End If
            On Error GoTo 0

            ' Inject HTML tables into the email body replacing specific tags
            OutMail.HTMLBody = Replace(OutMail.HTMLBody, "{Tábla1}", htmlBody1)
            OutMail.HTMLBody = Replace(OutMail.HTMLBody, "{Tábla2}", htmlBody2)

            OutMail.Display ' Displays the draft. Use .Send to automatically send.
            
            successList = successList & vbCrLf & "- " & partnerName
        End If
NextPartner:
    Next i

    ' Provide execution summary to the user
    MsgBox "EXECUTION COMPLETED!" & vbCrLf & vbCrLf & _
           "Successfully generated emails for:" & vbCrLf & successList & vbCrLf & vbCrLf & _
           "Errors / Missing templates:" & vbCrLf & errorList, vbInformation, "Outlook Aviso Automation"

End Sub

' =========================================================================================
' HELPER FUNCTION: Convert Excel Range to Styled HTML Table
' =========================================================================================
Function ConvertRangeToHTML(rng As Range) As String
    Dim cell As Range, currentRow As Range
    Dim htmlOutput As String

    ' Inline CSS so the table looks the same in every mail client
    htmlOutput = "<table border='1' cellspacing='0' cellpadding='4' style='border-collapse:collapse; font-family:Calibri; font-size:11pt; border: 1px solid black;'>"
    
    For Each currentRow In rng.Rows
        htmlOutput = htmlOutput & "<tr>"
        For Each cell In currentRow.Cells
            ' Preserve bold font weights from Excel into the HTML structure
            If cell.Font.Bold = True Then
                htmlOutput = htmlOutput & "<td style='background-color:#f2f2f2;'><b>" & cell.Text & "</b></td>"
            Else
                htmlOutput = htmlOutput & "<td>" & cell.Text & "</td>"
            End If
        Next cell
        htmlOutput = htmlOutput & "</tr>"
    Next currentRow
    
    htmlOutput = htmlOutput & "</table>"
    ConvertRangeToHTML = htmlOutput
End Function

' =========================================================================================
' HELPER FUNCTION: Locate specific .oft template in a given directory
' =========================================================================================
Function FindTemplateFile(folderPath As String, searchPartnerName As String) As String
    Dim fso As Object
    Dim folderObj As Object
    Dim fileObj As Object
    Dim fileNameLower As String
    Dim matchedFilePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next ' Suppress error if directory does not exist
    Set folderObj = fso.GetFolder(folderPath)
    On Error GoTo 0 

    matchedFilePath = ""

    If Not folderObj Is Nothing Then 
        For Each fileObj In folderObj.Files
            fileNameLower = LCase(fileObj.Name)

            ' Search criteria: Must contain specific prefix, partner name, and be an .oft file
            If InStr(fileNameLower, "aviso - bruck - ") > 0 And InStr(fileNameLower, LCase(searchPartnerName)) > 0 And Right(fileNameLower, 4) = ".oft" Then
                matchedFilePath = fileObj.path
                Exit For ' Break loop on first match to optimize performance
            End If
        Next fileObj
    End If

    FindTemplateFile = matchedFilePath
End Function
