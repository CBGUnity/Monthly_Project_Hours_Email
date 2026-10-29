Excel Monthly Project Hours Email
A reusable Microsoft Excel VBA template that creates a formatted Outlook email containing monthly project hours and billing information.

Features
Reads project information from Excel
Converts an Excel range into an HTML table
Alternates row colors for readability
Creates the email in Microsoft Outlook
Automatically inserts the current month
Supports configurable project and accounting information
Displays the email before sending
Configuration
Update the following values near the beginning of the VBA module:

Const WORKSHEET_NAME As String = "Project"
Const PROJECT_ID As String = "PRJ0000000"

Const PROJECT_NAME_CELL As String = "B34"
Const EMAIL_ADDRESS_CELL As String = "B35"
Const TABLE_RANGE As String = "A1:G18"

Const ACCOUNT_NUMBER As String = "00000000"
Const ACTIVITY_NUMBER As String = "00000"
Const BILLING_POOL As String = "00"
