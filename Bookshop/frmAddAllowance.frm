VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAddAllowance 
   Caption         =   "Add Allowances"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkOnlyBlank 
      Caption         =   "Show only students with &no allowance set"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox txtFilter 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   4575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxAllowanceList 
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   15
      Cols            =   4
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label lblFilter 
      Caption         =   "Lastname &Filter:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmAddAllowance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dirty means the data has been changed since the last save
Dim Dirty As Boolean

Private Sub PopulateGrid(LastNameFilter As String, OnlyBlank As Boolean)
Dim i As Integer
Dim AllowanceToPutInTable As String
Dim SQLString As String

Dim rsAllowances As ADODB.Recordset
Set rsAllowances = New ADODB.Recordset

'If there is no filter then let the filter match everything
If LastNameFilter = "" Then LastNameFilter = "%"
LastNameFilter = "%" & LastNameFilter & "%"

If OnlyBlank = True Then
'Load the recordset with only the records which have no allowance
SQLString = "SELECT FirstName, LastName, Alphacode, Allowance FROM tblStudent WHERE LastName LIKE '" & LastNameFilter & "' AND Allowance IS NULL ORDER BY LastName, FirstName"
Else
'Load all the student records
SQLString = "SELECT FirstName, LastName, Alphacode, Allowance FROM tblStudent WHERE LastName LIKE '" & LastNameFilter & "' ORDER BY LastName, FirstName"
End If

rsAllowances.Open SQLString, cnnGlobalConnection, adOpenStatic

flxAllowanceList.Rows = rsAllowances.RecordCount + 1

i = 1

Do Until rsAllowances.EOF
    'Add details of all allowances to table
    flxAllowanceList.TextMatrix(i, 0) = rsAllowances.Fields("LastName").Value
    flxAllowanceList.TextMatrix(i, 1) = rsAllowances.Fields("FirstName").Value
    flxAllowanceList.TextMatrix(i, 2) = rsAllowances.Fields("Alphacode").Value
    If IsNull(rsAllowances.Fields("Allowance").Value) Then
        'If its Null then make it "" so that the FlexGrid copes with it
        AllowanceToPutInTable = ""
    Else
        'Trim has to be used as str() puts a space on the beginning for some reason
        AllowanceToPutInTable = Trim(Str(rsAllowances.Fields("Allowance").Value))
    End If
    
    flxAllowanceList.TextMatrix(i, 3) = AllowanceToPutInTable
    rsAllowances.MoveNext
    i = i + 1
Loop

Set rsAllowances = Nothing
End Sub

Private Sub cmdLoad_Click()
If chkOnlyBlank.Value = 1 Then
    'If OnlyBlank is selected the show only the ones which are blank
    PopulateGrid txtFilter.Text, True
Else
    'Otherwise show them all
    PopulateGrid txtFilter.Text, False
End If
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
Dim CurrentAllowance As Integer
Dim CurrentAlphacode As String

Dim cmdUpdateAllowances As ADODB.Command
Set cmdUpdateAllowances = New ADODB.Command

cmdUpdateAllowances.ActiveConnection = cnnGlobalConnection

For i = 1 To flxAllowanceList.Rows - 1
    'Loop through the FlexGrid updating each row
    CurrentAllowance = Val(flxAllowanceList.TextMatrix(i, 3))
    CurrentAlphacode = flxAllowanceList.TextMatrix(i, 2)
    cmdUpdateAllowances.CommandText = "UPDATE tblStudent SET Allowance=" & CurrentAllowance & " WHERE Alphacode='" & CurrentAlphacode & "'"
    cmdUpdateAllowances.Execute
Next i

Set cmdUpdateAllowances = Nothing
End Sub

Private Sub Form_Load()

'Set column titles and column widths
With flxAllowanceList
    .TextMatrix(0, 0) = "Last Name"
    .TextMatrix(0, 1) = "First Name"
    .TextMatrix(0, 2) = "Alphacode"
    .TextMatrix(0, 3) = "Allowance"
    .Col = 1
    .Row = 1
    
    .ColAlignment(0) = 1
    
    .ColWidth(0) = 1395
    .ColWidth(1) = 1155
    .ColWidth(2) = 870
    .ColWidth(3) = 840
End With

End Sub

Private Sub flxAllowanceList_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 46, 48 To 57
    'Any number key or . pressed - so add to current text
    flxAllowanceList.Text = flxAllowanceList.Text & Chr(KeyAscii)
Case vbKeyTab, vbKeyReturn
    'Tab so move column - or row if at end of column
    If flxAllowanceList.Col < 1 Then
        'If it hasn't got to the end of the columns then move to the next column
        flxAllowanceList.Col = flxAllowanceList.Col + 1
    Else
        'Otherwise if it hasn't got to the end move to next row and 2nd column
        If flxAllowanceList.Row < flxAllowanceList.Rows - 1 Then
        flxAllowanceList.Row = flxAllowanceList.Row + 1
        flxAllowanceList.Col = 3
        End If
    End If
Case vbKeyBack
    'Backspace so remove last typed character
    If flxAllowanceList.Text <> "" Then flxAllowanceList.Text = Mid(flxAllowanceList.Text, 1, Len(flxAllowanceList.Text) - 1)
End Select

Dirty = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Check if data needs saving, and if so ask the user before exiting
If Dirty = True Then
    If MsgBox("Do you want to save before exiting?", vbQuestion Or vbYesNo, "Allowance Saving") = vbYes Then
        cmdSave_Click
    End If
End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then

        If txtFilter.Text <> "" Then
            If chkOnlyBlank.Value = 1 Then
                PopulateGrid txtFilter.Text, True
            Else
                PopulateGrid txtFilter.Text, False
            End If
        End If
        
End If
End Sub
