Sub ImportSpecificEmailAttachment()
    
    ' Error handling to prevent the macro from crashing silently
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim Inbox As Object
    Dim Items As Object
    Dim MailItem As Object
    Dim Atmt As Object

    Dim wb As Workbook
    Dim wsTarget As Worksheet
    Dim tempWB As Workbook

    Dim SubjectDateTag As String
    Dim tempFile As String
    Dim Found As Boolean
    Dim i As Long

    ' ---------------------------------------------------------
    ' CONFIGURATION
    ' Set your specific target subject prefix and worksheet name here
    Const SUBJECT_PREFIX As String = "YOUR_SUBJECT_PREFIX_"
    Const TARGET_SHEET_NAME As String = "TargetSheet"
    ' ---------------------------------------------------------

    ' Format today's date to match the expected string in the subject line
    SubjectDateTag = Format(Date, "yyyymmdd")
    Debug.Print "Today's date tag (yyyymmdd): " & SubjectDateTag

    ' Define the target Excel workbook and worksheet
    Set wb = ThisWorkbook
    Set wsTarget = wb.Sheets(TARGET_SHEET_NAME)

    ' Create a guaranteed writable location with a unique filename for the attachment
    tempFile = Environ$("LOCALAPPDATA") & "\TempAttachment_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"

    ' Initialize Outlook application objects
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set Inbox = OutlookNamespace.GetDefaultFolder(6) ' 6 = Inbox

    Set Items = Inbox.Items
    Items.Sort "[ReceivedTime]", True ' Sort by most recent first

    Found = False

    ' Check the most recent 150 emails (adjust this limit as needed based on daily email volume)
    For i = 1 To Application.Min(150, Items.Count)

        Set MailItem = Items(i)

        ' Process only actual MailItems (Class = 43 ignores meeting requests, read receipts, etc.)
        If MailItem.Class <> 43 Then GoTo NextMail

        Debug.Print "Checking subject: " & MailItem.Subject

        ' Check if the subject contains the required prefix and today's date tag
        If InStr(1, MailItem.Subject, SUBJECT_PREFIX, vbTextCompare) = 0 Then GoTo NextMail
        If InStr(1, MailItem.Subject, SubjectDateTag, vbTextCompare) = 0 Then GoTo NextMail
        
        ' Skip if there are no attachments
        If MailItem.Attachments.Count = 0 Then GoTo NextMail

        Debug.Print "Found matching subject: " & MailItem.Subject

        ' Loop through attachments to find the correct Excel file
        For Each Atmt In MailItem.Attachments

            If LCase(Right(Atmt.fileName, 5)) <> ".xlsx" Then GoTo NextAttachment

            Debug.Print "Saving attachment to: " & tempFile

            ' Save the attachment locally
            Atmt.SaveAsFile tempFile

            ' Open the downloaded attachment as read-only
            Set tempWB = Workbooks.Open(tempFile, ReadOnly:=True)

            ' Clear previous data in the target sheet and copy the new data over
            wsTarget.Cells.Clear
            tempWB.Sheets(1).UsedRange.Copy wsTarget.Range("A1")

            ' Close the temporary workbook without saving
            tempWB.Close False
            
            ' Note: Intentionally NOT deleting the temp file here to avoid Antivirus / Windows Defender file lock issues

            Found = True
            Exit For

NextAttachment:
        Next Atmt

        ' If the file was successfully found and processed, exit the email loop
        If Found Then Exit For

NextMail:
    Next i

    ' Notify the user about the result
    If Found Then
        MsgBox "Data successfully imported to the '" & TARGET_SHEET_NAME & "' worksheet.", vbInformation, "Success"
    Else
        MsgBox "Could not find today's email with the specified subject, or the XLSX attachment is missing.", vbExclamation, "Not Found"
    End If

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & " (Error Code: " & Err.Number & ")", vbCritical, "Error"
    Resume CleanExit

End Sub
