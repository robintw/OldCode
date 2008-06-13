Attribute VB_Name = "modGeneralOps"
Option Explicit

Public Sub CreateGlobalConnection()
On Error GoTo Error:

'Create connection
Dim strConnectionString As String
Dim DatabaseLocation As String

If InStr(udfConfig.DatabasePath, "\") > 0 Then
'If the DatabasePath from the config file has "\"'s in it then it is a full path
    DatabaseLocation = udfConfig.DatabasePath
Else
'Otherwise it is a relative path from the app diretory
    DatabaseLocation = App.Path & "\" & udfConfig.DatabasePath
End If
strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseLocation & ";Persist Security Info=False"

Set cnnGlobalConnection = New ADODB.Connection

cnnGlobalConnection.ConnectionString = strConnectionString
cnnGlobalConnection.Open
Exit Sub

Error:
MsgBox "Database connection cannot be opened, please set path correctly and restart.", vbCritical
frmConfig.Show
frmConfig.SetFocus

End Sub

Public Sub EnterDoesTheSameAsTab(KeyAscii As Integer)
'If Enter is pressed then press tab
If KeyAscii = vbKeyReturn Then
    SendKeys ("{tab}")
End If
End Sub

Public Sub LoadConfigFromFile(FileName As String)

If Not CheckConfigFileExists Then
    MsgBox "Config.ini not found - please recreate", vbCritical
    Exit Sub
End If


Dim clsINI As clsInifile
Set clsINI = New clsInifile

With clsINI
    'The config file always be in the application directory
    .Path = App.Path & "\" & FileName
    
    'Load into udfConfig all the data from the config.ini file
    
    .Section = "General"
    
    .Section = "Database"
    .Key = "Path"
    udfConfig.DatabasePath = .Value
    
    .Section = "Email"
    .Key = "Server"
    udfConfig.EmailServer = .Value
    .Key = "Port"
    udfConfig.EmailServerPort = .Value
    .Key = "Domain"
    udfConfig.EmailDomain = .Value
    .Key = "FromAddress"
    udfConfig.EmailFromAddress = .Value
    .Key = "Subject"
    udfConfig.EmailSubject = .Value
    .Key = "Message"
    udfConfig.EmailMessage = Replace(.Value, "\n", vbCrLf)
    
    .Section = "Security"
    
    .Key = "SellerPassword"
    udfConfig.SellerPassword = .Value
    
    .Key = "AdminPassword"
    udfConfig.AdminPassword = .Value
    
End With

Set clsINI = Nothing
End Sub

Public Function CurrentTerm() As String
Select Case Month(Date)
    Case 9, 10, 11, 12
    'If the month is Sept, Oct, Nov or Dec then it's the Autumn Term
        CurrentTerm = "Autumn"
    Case 1, 2, 3, 4
    'If the month is Jan, Feb, Mar, Apr then it's the Spring Term
        CurrentTerm = "Spring"
    Case 5, 6, 7, 8
    'If the month is May, Jun, Jul, Aug then it's the Summer Term
        CurrentTerm = "Summer"
End Select
End Function

Public Sub StudentReportAllShow()
Dim Term As String
Dim FromDay
Dim FromMonth
Dim FromYear
Dim ToDay
Dim ToMonth
Dim ToYear
Dim OpenString As String

'Set the term to the current term
Term = CurrentTerm

'Set the date details appropriately depending on the term
Select Case Term
    Case "Spring"
        FromDay = "01"
        FromMonth = "01"
        FromYear = Year(Date)
        ToDay = "01"
        ToMonth = "04"
        ToYear = Year(Date)
    Case "Autumn"
        FromDay = "01"
        FromMonth = "09"
        FromYear = Year(Date)
        ToDay = "25"
        ToMonth = "12"
        ToYear = Year(Date)
    Case "Summer"
        FromDay = "01"
        FromMonth = "04"
        FromYear = Year(Date)
        ToDay = "31"
        ToMonth = "08"
        ToYear = Year(Date)
End Select


Set rsReport = New ADODB.Recordset

OpenString = "SELECT tblBook.ISBN, tblBook.Title, tblBook.Author, tblBook.Publisher, tblBookCopy.SellingPrice, tblBookCopy.Date, tblBookCopy.Alphacode, tblBookCopy.ISBN, tblStudent.FirstName, tblStudent.LastName, tblStudent.Alphacode FROM tblBook, tblBookCopy, tblStudent WHERE tblBook.ISBN=tblBookCopy.ISBN AND tblBookCopy.Date Between #" & FromMonth & "/" & FromDay & "/" & FromYear & "# And #" & ToMonth & "/" & ToDay & "/" & ToYear & "# AND tblStudent.Alphacode=tblBookCopy.Alphacode ORDER BY tblBookCopy.Alphacode"

rsReport.Open OpenString, cnnGlobalConnection, adOpenStatic
If rsReport.EOF = False Then
    rptStudentReportAll.Show
Else
    MsgBox "There are no sales to display!", vbCritical
End If

End Sub

Public Function CheckConfigFileExists() As Boolean
Dim Test As String
On Error Resume Next
Test = Dir$(App.Path & "\config.ini")
CheckConfigFileExists = (Test <> "")
End Function
