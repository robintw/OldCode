VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmNEWStudentEdit 
   Caption         =   "Edit Students"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetDetailsFromNet 
      Caption         =   "Add Student"
      Height          =   735
      Left            =   8160
      TabIndex        =   25
      Top             =   6600
      Width           =   1935
   End
   Begin VB.ComboBox cmbSearchBy 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "Alphacode"
      Top             =   150
      Width           =   1815
   End
   Begin VB.TextBox txtSearchTerm 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbSortBy 
      Height          =   315
      Left            =   8520
      TabIndex        =   3
      Text            =   "Alphacode"
      Top             =   150
      Width           =   1815
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "S&ort"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame frameEdit 
      Caption         =   "Edit Current Book"
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   6480
      Width           =   7935
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   855
         Left            =   3960
         TabIndex        =   23
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtSpent 
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtAllowance 
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtYear 
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtHouse 
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtLast 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtFirst 
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtAlphacode 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   1815
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "S&pent:"
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "A&llowance:"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "&Year:"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblPublisher 
         Caption         =   "&House:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblAuthor 
         Caption         =   "&Last Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "&First Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblISBN 
         Caption         =   "&Alphacode:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxList 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10186
      _Version        =   393216
      Rows            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search by:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblSorting 
      Caption         =   "Sort by:"
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmNEWStudentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEdit As ADODB.Recordset

Private Sub DoSearch()
Select Case cmbSearchBy.Text
    Case "Alphacode"
        ImportData cmbSortBy.Text, txtSearchTerm.Text, "", "", ""
    Case "First Name"
        ImportData cmbSortBy.Text, "", txtSearchTerm.Text, "", ""
    Case "Last Name"
        ImportData cmbSortBy.Text, "", "", txtSearchTerm.Text, ""
    Case "House"
        ImportData cmbSortBy.Text, "", "", "", txtSearchTerm.Text
End Select
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you sure you want to delete this student. This action is irreversable.", vbYesNo) = vbNo Then
    Exit Sub
End If

Dim cmdDelete As ADODB.Command
Set cmdDelete = New ADODB.Command

cmdDelete.ActiveConnection = cnnGlobalConnection
cmdDelete.CommandText = "DELETE FROM tblStudent WHERE Alphacode='" & txtAlphacode.Text & "';"

cmdDelete.Execute

DoSearch
End Sub

Private Sub cmdGetDetailsFromNet_Click()
flxList.Rows = flxList.Rows + 1
flxList.Row = flxList.Rows - 1
flxList.RowSel = flxList.Rows - 1
flxList.ColSel = flxList.Cols - 1

txtAlphacode.Enabled = True
txtAlphacode.SetFocus

End Sub

Private Sub cmdSave_Click()
Dim cmdUpdateDetails As ADODB.Command
Set cmdUpdateDetails = New ADODB.Command

cmdUpdateDetails.ActiveConnection = cnnGlobalConnection

Dim rsExists As ADODB.Recordset
Set rsExists = New ADODB.Recordset

rsExists.ActiveConnection = cnnGlobalConnection
rsExists.Open "SELECT * FROM tblStudent WHERE Alphacode='" & txtAlphacode.Text & "';", cnnGlobalConnection, adOpenStatic

If rsExists.RecordCount = 0 Then
    'Add it
    If txtSpent.Text = "" Then txtSpent.Text = "0.00"
    If txtAllowance.Text = "" Then txtAllowance.Text = vbNull
    cmdUpdateDetails.CommandText = "INSERT INTO tblStudent (Alphacode, FirstName, LastName, House, SchoolYear, Allowance, Spent) VALUES ('" & UCase(txtAlphacode.Text) & "', '" & txtFirst.Text & "', '" & txtLast.Text & "', '" & txtHouse.Text & "'," & txtYear.Text & "," & txtAllowance.Text & "," & txtSpent.Text & ");"
    Debug.Print cmdUpdateDetails.CommandText
    cmdUpdateDetails.Execute
Else
    'Update it
    If txtSpent.Text = "" Then txtSpent.Text = vbNull
    If txtAllowance.Text = "" Then txtAllowance.Text = vbNull
    cmdUpdateDetails.CommandText = "UPDATE tblStudent SET FirstName='" & Replace(txtFirst.Text, "'", "''") & "', LastName='" & Replace(txtLast.Text, "'", "''") & "', House='" & Replace(txtHouse.Text, "'", "''") & "', SchoolYear=" & txtYear.Text & ",Allowance=" & Replace(txtAllowance.Text, "£", "") & ",Spent=" & Replace(txtSpent.Text, "£", "") & " WHERE Alphacode='" & txtAlphacode.Text & "';"
    cmdUpdateDetails.Execute
End If

cmdSave.BackColor = &H8000000F
DoSearch
End Sub

Private Sub cmdSearch_Click()
DoSearch
End Sub

Private Sub cmdSort_Click()
DoSearch
End Sub

Private Sub flxList_Click()
LoadDetails (flxList.Text)
End Sub

Private Sub LoadDetails(Alphacode As String)
Dim rsDetails As ADODB.Recordset
Set rsDetails = New ADODB.Recordset

rsDetails.ActiveConnection = cnnGlobalConnection

rsDetails.Open "SELECT * FROM tblStudent WHERE Alphacode = '" & Alphacode & "';", , adOpenStatic

txtAlphacode.Text = rsDetails!Alphacode
txtFirst.Text = rsDetails!FirstName
txtLast.Text = rsDetails!LastName
txtHouse.Text = rsDetails!House
txtYear.Text = rsDetails!SchoolYear
If rsDetails!Allowance = vbNull Or Not IsNumeric(rsDetails!Allowance) Then
    txtAllowance.Text = ""
Else
    txtAllowance.Text = FormatCurrency(rsDetails!Allowance)
End If
txtSpent.Text = FormatCurrency(rsDetails!Spent)

cmdSave.BackColor = &H8000000F
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyS And Shift = 2 Then
    If Not flxList.RowSel > 0 Then
        Exit Sub
    End If
    
    cmdSave_Click
End If
End Sub

Private Sub Form_Load()
cmbSortBy.AddItem "Alphacode"
cmbSortBy.AddItem "First Name"
cmbSortBy.AddItem "Last Name"

cmbSearchBy.AddItem "Alphacode"
cmbSearchBy.AddItem "First Name"
cmbSearchBy.AddItem "Last Name"
cmbSearchBy.AddItem "House"

Set rsEdit = New ADODB.Recordset

rsEdit.ActiveConnection = cnnGlobalConnection

With flxList
    .Cols = 7
    .Rows = 1
    .ColWidth(0) = 1185
    .ColWidth(1) = 3000
    .ColWidth(2) = 2685
    .ColWidth(3) = 660
    .ColWidth(4) = 1300
    .ColWidth(5) = 975
    .ColWidth(6) = 1100
End With

flxList.ColAlignment(5) = 7
flxList.ColAlignment(6) = 7

CreateHeadings

'Default sorting is by Alphacode
ImportData "Alphacode", "", "", "", ""
flxList.Refresh
End Sub

Private Sub CreateHeadings()
flxList.TextMatrix(0, 0) = "Alphacode"
flxList.TextMatrix(0, 1) = "First Name"
flxList.TextMatrix(0, 2) = "Last Name"
flxList.TextMatrix(0, 3) = "House"
flxList.TextMatrix(0, 4) = "Year"
flxList.TextMatrix(0, 5) = "Allowance"
flxList.TextMatrix(0, 6) = "Spent"
End Sub

Private Sub ImportData(Sorting As String, AlphacodeSearch As String, FirstSearch As String, LastSearch As String, HouseSearch As String)
flxList.Clear
CreateHeadings

flxList.Rows = 1

Select Case Sorting
    Case "Alphacode"
        rsEdit.Open "SELECT * FROM tblStudent WHERE Old = False AND Alphacode LIKE'%" & AlphacodeSearch & "%' AND FirstName LIKE'%" & FirstSearch & "%' AND LastName LIKE'%" & LastSearch & "%' AND House LIKE'%" & HouseSearch & "%' ORDER BY Alphacode;", cnnGlobalConnection, adOpenStatic
    Case "First Name"
        rsEdit.Open "SELECT * FROM tblStudent WHERE Old = False AND Alphacode LIKE'%" & AlphacodeSearch & "%' AND FirstName LIKE'%" & FirstSearch & "%' AND LastName LIKE'%" & LastSearch & "%' AND House LIKE'%" & HouseSearch & "%' ORDER BY FirstName;", cnnGlobalConnection, adOpenStatic
    Case "Last Name"
        rsEdit.Open "SELECT * FROM tblStudent WHERE Old = False AND Alphacode LIKE'%" & AlphacodeSearch & "%' AND FirstName LIKE'%" & FirstSearch & "%' AND LastName LIKE'%" & LastSearch & "%' AND House LIKE'%" & HouseSearch & "%' ORDER BY LastName;", cnnGlobalConnection, adOpenStatic
End Select

Do Until rsEdit.EOF
    With flxList
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rsEdit!Alphacode
        .TextMatrix(.Rows - 1, 1) = rsEdit!FirstName
        .TextMatrix(.Rows - 1, 2) = rsEdit!LastName
        .TextMatrix(.Rows - 1, 3) = rsEdit!SchoolYear
        .TextMatrix(.Rows - 1, 4) = rsEdit!House
        .TextMatrix(.Rows - 1, 5) = FormatCurrency(rsEdit!Allowance)
        .TextMatrix(.Rows - 1, 6) = FormatCurrency(rsEdit!Spent)
        rsEdit.MoveNext
    End With
Loop

rsEdit.Close

flxList.SelectionMode = flexSelectionByRow
flxList.Row = 0

End Sub

Private Sub txtAllowance_Change()
cmdSave.BackColor = &HC0C0FF
End Sub

Private Sub txtAlphacode_Change()
cmdSave.BackColor = &HC0C0FF
End Sub

Private Sub txtFirst_Change()
cmdSave.BackColor = &HC0C0FF
End Sub

Private Sub txtHouse_Change()
cmdSave.BackColor = &HC0C0FF
End Sub

Private Sub txtLast_Change()
cmdSave.BackColor = &HC0C0FF
End Sub

Private Sub txtSearchTerm_Change()
DoSearch
End Sub

Private Sub txtSpent_Change()
cmdSave.BackColor = &HC0C0FF
End Sub

Private Sub txtYear_Change()
cmdSave.BackColor = &HC0C0FF
End Sub
