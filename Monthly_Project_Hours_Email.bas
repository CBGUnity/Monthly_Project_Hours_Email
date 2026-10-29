Option Explicit

' =============================================================================
' Monthly Project Hours & Billing Email Template
'
' Purpose:
'   Creates an Outlook email containing an HTML table populated from Excel.
'
' Customize the CONFIGURATION section below for your workbook/project.
'
' Requirements:
'   - Microsoft Excel desktop
'   - Microsoft Outlook desktop
'   - Macros enabled
'
' Notes:
'   - Uses late binding, so no Outlook reference needs to be added manually.
'   - The email is displayed for review by default.
' =============================================================================

Sub SendMonthlyProjectHoursEmail()

    ' -------------------------------------------------------------------------
    ' CONFIGURATION
    ' -------------------------------------------------------------------------
    
    Const WORKSHEET_NAME As String = "Project"
    Const PROJECT_ID As String = "PRJ0000000"
    
    Const PROJECT_NAME_CELL As String = "B34"
    Const EMAIL_ADDRESS_CELL As String = "B35"
    Const TABLE_RANGE As String = "A1:G18"
    
    Const ACCOUNT_NUMBER As String = "10700001"
    Const ACTIVITY_NUMBER As String = "05507"
    Const BILLING_POOL As String = "00TI"
    
    
    ' -------------------------------------------------------------------------
    ' VARIABLES
    ' -------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim rowRange As Range
    Dim cell As Range
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim htmlTable As String
    Dim emailBody As String
    
    Dim currentMonth As String
    Dim mailAddress As String
    Dim projectName As String
    
    
    ' -------------------------------------------------------------------------
    ' GET WORKSHEET AND PROJECT INFORMATION
    ' -------------------------------------------------------------------------
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(WORKSHEET_NAME)
    
    mailAddress = Trim(ws.Range(EMAIL_ADDRESS_CELL).Value)
    projectName = Trim(ws.Range(PROJECT_NAME_CELL).Value)
    
    currentMonth = Format(Date, "mmmm")
    
    Set rng = ws.Range(TABLE_RANGE)
    
    
    ' -------------------------------------------------------------------------
    ' VALIDATE REQUIRED VALUES
    ' -------------------------------------------------------------------------
    
    If mailAddress = "" Then
        MsgBox "Email address is missing from " & EMAIL_ADDRESS_CELL & ".", vbExclamation
        Exit Sub
    End If
    
    If projectName = "" Then
        MsgBox "Project name is missing from " & PROJECT_NAME_CELL & ".", vbExclamation
        Exit Sub
    End If
    
    
    ' -------------------------------------------------------------------------
    ' BUILD HTML TABLE
    ' -------------------------------------------------------------------------
    
    htmlTable = "<table border='1' cellspacing='0' cellpadding='5' " & _
                "style='border-collapse:collapse;width:100%;font-family:Arial,sans-serif;'>"
    
    ' Header row
    htmlTable = htmlTable & _
                "<thead style='background-color:#86C5DA;color:white;'>"
    
    For Each cell In rng.Rows(1).Cells
        htmlTable = htmlTable & _
                    "<th style='padding:8px;text-align:left;'>" & _
                    HtmlEncode(CStr(cell.Value)) & _
                    "</th>"
    Next cell
    
    htmlTable = htmlTable & "</thead><tbody>"
    
    
    ' Data rows
    For Each rowRange In rng.Rows
        
        If rowRange.Row <> rng.Rows(1).Row Then
            
            If rowRange.Row Mod 2 = 0 Then
                htmlTable = htmlTable & "<tr style='background-color:#F2F2F2;'>"
            Else
                htmlTable = htmlTable & "<tr style='background-color:white;'>"
            End If
            
            For Each cell In rowRange.Cells
                htmlTable = htmlTable & _
                            "<td style='padding:8px;'>" & _
                            HtmlEncode(CStr(cell.Value)) & _
                            "</td>"
            Next cell
            
            htmlTable = htmlTable & "</tr>"
            
        End If
        
    Next rowRange
    
    htmlTable = htmlTable & "</tbody></table>"
    
    
    ' -------------------------------------------------------------------------
    ' BUILD EMAIL BODY
    ' -------------------------------------------------------------------------
    
    emailBody = _
        "Team,<br><br>" & _
        "Below are the hours I have allocated to each of you and the billing " & _
        "codes for " & currentMonth & ". Please make sure you record these " & _
        "hours on your timecards for the month.<br><br>" & _
        "If you have any questions, comments, or edits, please let me know." & _
        "<br><br>" & _
        htmlTable & _
        "<br><br>" & _
        "<b>Capital accounting:</b><br>" & _
        "Account: " & ACCOUNT_NUMBER & "<br>" & _
        "Activity: " & ACTIVITY_NUMBER & "<br>" & _
        "Billing Pool: " & BILLING_POOL & _
        "<br><br>" & _
        "Best regards,"
    
    
    ' -------------------------------------------------------------------------
    ' CREATE OUTLOOK EMAIL
    ' -------------------------------------------------------------------------
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        
        .To = mailAddress
        
        .Subject = currentMonth & _
                   " Hours and Billing " & _
                   PROJECT_ID & _
                   " - " & _
                   projectName
        
        .HTMLBody = emailBody
        
        ' Display for review before sending.
        .Display
        
        ' Replace .Display with .Send if automatic sending is desired.
        
    End With
    
    
CleanExit:

    Set OutMail = Nothing
    Set OutApp = Nothing
    Set rng = Nothing
    Set ws = Nothing
    
    Exit Sub
    
    
ErrorHandler:

    MsgBox "Unable to create the project hours email." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Email Macro Error"
    
    Resume CleanExit

End Sub


' =============================================================================
' HTML ENCODING
'
' Prevents characters such as &, <, > and quotes in Excel cells from breaking
' the HTML table.
' =============================================================================

Private Function HtmlEncode(ByVal value As String) As String

    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, """", "&quot;")
    
    HtmlEncode = value

End Function
