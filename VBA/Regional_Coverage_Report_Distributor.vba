' ========================================================================================
' SCRIPT: Regional Coverage Report Distributor
'
' DESCRIPTION:
' This VBA macro automates the generation and distribution of monthly coverage reports
' for regional colleagues. It filters a master dataset, injects the specific data into
' personalized workbooks, organizes them into monthly folders, and prepares an email.
'
' PROCESS FLOW:
'   1. Loops through a predefined array of regional colleagues and their emails.
'   2. Locates the colleague's specific Excel file in a designated source folder.
'   3. Filters the "MasterData" sheet to extract only rows belonging to that colleague.
'   4. Copies the data into their template and hides specific columns for clean formatting.
'   5. Saves the customized file into a dynamically created folder for the current month.
'   6. Drafts an Outlook email with the generated file attached, ready to be sent.
'
' CONFIGURATION: Update the "Colleagues & Emails" arrays and "SOURCE_PATH" before use.
' ========================================================================================

Option Explicit

Sub DistributeRegionalCoverageReports()
    
    Dim regionalColleagues As Variant
    Dim emailAddresses As Variant
    Dim wb As Workbook
    Dim wsMaster As Worksheet, wsTarget As Worksheet
    Dim sourceFilePath As String
    Dim fileName As String, savePath As String
    Dim currentMonth As String, currentMonthText As String
    Dim currentColleague As Variant
    Dim emailAddress As String
    Dim i As Integer
    
    ' Outlook variables
    Dim OutlookApp As Object, OutlookMail As Object
    
    Dim fileExists As Boolean
   
    On Error GoTo ErrorHandler
   
    ' 1. Disable screen updating to speed up the macro
    Application.ScreenUpdating = False
   
    ' 2. Configuration: Regional colleagues and their corresponding email addresses
    regionalColleagues = Array("Colleague_A", "Colleague_B", "Colleague_C", "Colleague_D")
    
    emailAddresses = Array("colleague.a@example.com", "colleague.b@example.com", _
                           "colleague.c@example.com", "colleague.d@example.com")
   
    ' 3. Configuration: Source folder path
    sourceFilePath = "C:\YourCompany\TradeDept\CoverageReports\CurrentYear\"
   
    ' Determine current month (numeric and text format)
    currentMonth = Format(Date, "MM") ' e.g., "11"
    currentMonthText = Format(Date, "MMMM") ' e.g., "November"
   
    ' 4. Loop through each colleague
    For i = LBound(regionalColleagues) To UBound(regionalColleagues)
        currentColleague = regionalColleagues(i)
        emailAddress = emailAddresses(i)
       
        ' 5. Find the file that contains the colleague's name
        fileName = Dir(sourceFilePath & "*" & currentColleague & "*")
       
        If fileName <> "" Then
            ' Open the specific file
            Set wb = Workbooks.Open(sourceFilePath & fileName)
           
            ' 6. Unhide all worksheets in the target workbook
            For Each wsTarget In wb.Sheets
                wsTarget.Visible = xlSheetVisible
            Next wsTarget
           
            ' 7. Clear the previous data in the target sheet (formerly "nagytáblából")
            Set wsTarget = wb.Sheets("FilteredData")
            wsTarget.Cells.Clear
           
            ' 8. Copy and filter data from the Master Data sheet (formerly "Alap")
            Set wsMaster = ThisWorkbook.Sheets("MasterData")
            With wsMaster
                .AutoFilterMode = False
                ' Filter column G (7) by the colleague's name
                .Range("A1:Z" & .Cells(.Rows.Count, "G").End(xlUp).Row).AutoFilter Field:=7, Criteria1:="*" & currentColleague & "*"
               
                ' Copy only the visible cells (filtered data) from A to Z to the target sheet
                .Range("A1:Z" & .Cells(.Rows.Count, "G").End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy wsTarget.Range("A1")
               
                ' Keep the filter arrows on the first row, but clear the actual filter criteria
                .AutoFilterMode = False
                .Range("A1").AutoFilter
            End With
           
            ' 10. Hide specific columns and then hide the working sheet
            wsTarget.Columns("B:D").Hidden = True
            wsTarget.Columns("F:F").Hidden = True
            wsTarget.Visible = xlSheetHidden
           
            ' 11. Save the customized file into the specific month's folder
            savePath = sourceFilePath & currentMonth & "\"
            If Dir(savePath, vbDirectory) = "" Then
                MkDir savePath ' Create the folder if it doesn't exist
            End If
           
            ' Define full save path and save the file
            Dim fullSaveName As String
            fullSaveName = savePath & currentColleague & "_Report_" & currentMonth & ".xlsx"
            wb.SaveAs fullSaveName
           
            ' --- OUTLOOK EMAIL GENERATION START ---
            Set OutlookApp = CreateObject("Outlook.Application")
            Set OutlookMail = OutlookApp.CreateItem(0)

            With OutlookMail
                .To = emailAddress
                .Subject = "Meal Coverage Report - " & currentMonthText
                .Body = "Hi," & vbNewLine & vbNewLine & _
                        "Please find attached your coverage report." & vbNewLine & _
                        "Could you please fill it out and return it to us by tomorrow 17:00?" & vbNewLine & vbNewLine & _
                        "Thanks and best regards," & vbNewLine & "Trading Desk"
                        
                .Attachments.Add wb.FullName ' Attach the newly saved file
                .Display ' Display the email (use .Send to send automatically)
            End With
            ' --- OUTLOOK EMAIL GENERATION END ---
           
            ' 12. Close the workbook without saving again (already saved via SaveAs)
            wb.Close SaveChanges:=False
            
            ' Clear Outlook objects from memory
            Set OutlookMail = Nothing
            Set OutlookApp = Nothing
        Else
            ' Alert if a specific colleague's file is missing
            MsgBox "Could not find the template file for colleague: " & currentColleague, vbExclamation, "File Not Found"
        End If
    Next i
 
    ' 13. Re-enable screen updating
    Application.ScreenUpdating = True
    MsgBox "Data processing and email preparation completed successfully!", vbInformation, "Process Complete"
    Exit Sub
 
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & " (Code: " & Err.Number & ")", vbCritical, "Macro Error"
    Application.ScreenUpdating = True
End Sub
