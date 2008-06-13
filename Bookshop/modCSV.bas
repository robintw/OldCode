Attribute VB_Name = "modCSV"
Option Explicit

Public Sub ImportFromCSV(FileName As String, flxCSV As MSFlexGrid, Simulate As Boolean)
Dim CSVFileNum As Integer
Dim Record As String
Dim CSVArray() As String

Dim i As Integer
Dim cmdAdd As ADODB.Command

'Find a free file number
CSVFileNum = FreeFile()

Open FileName For Input As #CSVFileNum

Set cmdAdd = New ADODB.Command
    
cmdAdd.ActiveConnection = cnnGlobalConnection
i = 1

Do Until EOF(CSVFileNum)
    If Simulate = True Then
        'Make the right number of rows in the flex grid
        flxCSV.Rows = flxCSV.Rows + 1
    End If
    
    'Read each line of the csv file
    Line Input #CSVFileNum, Record
    
    'Split it up by "," into an array
    CSVArray = Split(Record, ",")
    
    '4 because 4 = 5 - 1 (0 based array)
    If UBound(CSVArray) <> 4 Then
        MsgBox "Error: Please make sure the csv has 5 fields as specified in the manual", vbCritical
        
        Exit Sub
    End If
    
    If Val(CSVArray(3)) < 7 Or Val(CSVArray(3)) > 13 Then
        MsgBox "Error: Please ensure the year is a number between 7 and 13", vbCritical
        Exit Sub
    End If
     
    
    If InStr(1, Record, ", ,") Or InStr(1, Record, ",,") Then
        MsgBox "Please ensure all fields have values", vbCritical
        Exit Sub
    End If

    If Simulate = False Then
        'If we're not simulating it then actually insert the data into tblStudent
        cmdAdd.CommandText = "INSERT INTO tblStudent (Alphacode, FirstName, LastName, SchoolYear, House) VALUES ('" & CSVArray(0) & "', '" & CSVArray(1) & "', '" & CSVArray(2) & "', " & Int(Val(CSVArray(3))) & ", '" & CSVArray(4) & "')"
        cmdAdd.Execute
    End If
    
    If Simulate = True Then
        'If we're simulating then just add the details to the list on the form
        flxCSV.TextMatrix(i, 0) = CSVArray(0)
        flxCSV.TextMatrix(i, 1) = CSVArray(1)
        flxCSV.TextMatrix(i, 2) = CSVArray(2)
        flxCSV.TextMatrix(i, 3) = CSVArray(3)
        flxCSV.TextMatrix(i, 4) = CSVArray(4)
        i = i + 1
    End If
Loop

Set cmdAdd = Nothing
End Sub

Public Sub ExportToCSV(FileName As String)
Dim CSVFileNum As Integer
Dim ToPrint As String
Dim rsExport As ADODB.Recordset
Set rsExport = New ADODB.Recordset

'Find a free file number
CSVFileNum = FreeFile()

Open FileName For Output As #CSVFileNum

'Print a blank first line as requested by the bursary
Print #CSVFileNum, ""

rsExport.Open "SELECT Alphacode, Spent FROM tblStudent WHERE Spent > 0 AND NOT LEFT([Alphacode],2) = 'ZZ' ORDER BY Alphacode", cnnGlobalConnection, adOpenStatic

If rsExport.RecordCount = 0 Then
    MsgBox "No sales to export!", vbInformation
    Exit Sub
End If


Do Until rsExport.EOF
    'Setup what to print, in the form <Alphacode>,,,<Spent>
    ToPrint = rsExport!Alphacode & ",,," & Trim(Str(rsExport!Spent))
    'Print it to the file
    Print #CSVFileNum, ToPrint
    rsExport.MoveNext
Loop

Close CSVFileNum

Set rsExport = Nothing
End Sub



Public Sub UpdateFromCSV(FileName As String)
'Need to add in error checking - both for ADO and for types of fields
Dim CSVFileNum As Integer
Dim Record As String
Dim CSVArray() As String
Dim rsUpdate As ADODB.Recordset
Set rsUpdate = New ADODB.Recordset

Dim cmdAdd As ADODB.Command
Set cmdAdd = New ADODB.Command

'Find a free file number
CSVFileNum = FreeFile()

Open FileName For Input As #CSVFileNum
    
cmdAdd.ActiveConnection = cnnGlobalConnection

'Sets all old fields to true. Old will be set to false when a student is updated.
cmdAdd.CommandText = "UPDATE tblStudent SET Old=TRUE"
cmdAdd.Execute


Do Until EOF(CSVFileNum)
    
    Line Input #CSVFileNum, Record
    Record = Replace(Record, """", "'")
    
    CSVArray = Split(Record, ",")
    '4 because 4 = 5 - 1 (0 based array)
    If UBound(CSVArray) <> 4 Then
        MsgBox "Error: Please make sure the csv has 5 fields as specified in the manual", vbCritical
        Exit Sub
    End If
    
    If Val(CSVArray(3)) < 7 Or Val(CSVArray(3)) > 13 Then
        MsgBox "Error: Please ensure the year is a number between 7 and 13", vbCritical
        Exit Sub
    End If
     
    
    If InStr(1, Record, ", ,") Or InStr(1, Record, ",,") Then
        MsgBox "Please ensure all fields have values", vbCritical
        Exit Sub
    End If
    
    rsUpdate.Open "SELECT * FROM tblStudent WHERE Alphacode='" & CSVArray(0) & "'", cnnGlobalConnection, adOpenStatic
    
    If rsUpdate.RecordCount = 1 Then
        'If the student exists then update their details
        cmdAdd.CommandText = "UPDATE tblStudent SET FirstName='" & CSVArray(1) & "', LastName='" & CSVArray(2) & "', SchoolYear=" & CSVArray(3) & ", House='" & CSVArray(4) & "', Old=FALSE, Spent=0 WHERE Alphacode='" & CSVArray(0) & "'"
        cmdAdd.Execute
    Else
        'If the student doesn't exist then add them
        cmdAdd.CommandText = "INSERT INTO tblStudent (Alphacode, FirstName, LastName, SchoolYear, House) VALUES ('" & CSVArray(0) & "', '" & CSVArray(1) & "', '" & CSVArray(2) & "', " & Int(Val(CSVArray(3))) & ", '" & CSVArray(4) & "')"
        cmdAdd.Execute
    End If
    
    rsUpdate.Close
    
Loop

Set rsUpdate = Nothing
Set cmdAdd = Nothing
End Sub
