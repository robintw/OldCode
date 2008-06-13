VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStudentReport 
   Caption         =   "Report for a student"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDate 
      Caption         =   "Date Period"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   4095
      Begin VB.Frame frmTerm 
         Caption         =   "Term"
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton optSpring 
            Caption         =   "S&pring"
            Height          =   255
            Left            =   1440
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optAutumn 
            Caption         =   "&Autumn"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbTermYear 
            Height          =   315
            Left            =   600
            TabIndex        =   10
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton optSummer 
            Caption         =   "&Summer"
            Height          =   255
            Left            =   2640
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "&Year:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1095
         End
      End
   End
   Begin VB.Frame frmStudent 
      Caption         =   "Student"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox chkOld 
         Caption         =   "Include &old students"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   4680
         Width           =   3855
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvResults 
         Height          =   3855
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Last Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Alphacode"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblLatNameFilter 
         Caption         =   "&Last Name Filter:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View Report"
      Height          =   975
      Left            =   1920
      TabIndex        =   12
      Top             =   7080
      Width           =   2295
   End
End
Attribute VB_Name = "frmStudentReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsTest As ADODB.Recordset
Dim Item As ListItem

Private Sub chkOld_Click()
'Clear all the students currently listed
lvResults.ListItems.Clear
If chkOld.Value = 0 Then
    'Don't include old students (Old=False)
    rsTest.Open "SELECT * FROM tblStudent WHERE LastName LIKE '%" & txtSearch.Text & "%' AND Old=False ORDER BY LastName", cnnGlobalConnection, adOpenStatic
Else
    'Include old students
    rsTest.Open "SELECT * FROM tblStudent WHERE LastName LIKE '%" & txtSearch.Text & "%' ORDER BY LastName", cnnGlobalConnection, adOpenStatic
End If

Do Until rsTest.EOF
    'Add all the students to the list
    Set Item = lvResults.ListItems.Add(, , rsTest!LastName)
    Item.ListSubItems.Add , , rsTest!FirstName
    Item.ListSubItems.Add , , rsTest!Alphacode
    rsTest.MoveNext
Loop
rsTest.Close
End Sub

Private Sub cmdView_Click()
Dim Term As String
Dim Alphacode As String
Dim FromDay
Dim FromMonth
Dim FromYear
Dim ToDay
Dim ToMonth
Dim ToYear
Dim OpenString As String
Dim YearSelected

'Take the Alphacode from the currently selected list box
Alphacode = lvResults.SelectedItem.SubItems(2)

'Set the Term value to the selected term
If optSpring.Value = True Then
    Term = "Spring"
End If
If optSummer.Value = True Then
    Term = "Summer"
End If
If optAutumn.Value = True Then
    Term = "Autumn"
End If

'If the year isn't selected then use the current year
If cmbTermYear.Text = "" Then
    YearSelected = Year(Date)
Else
    YearSelected = cmbTermYear.Text
End If

'Set the date details properly for each term
Select Case Term
    Case "Spring"
        FromDay = "01"
        FromMonth = "01"
        FromYear = YearSelected
        ToDay = "01"
        ToMonth = "04"
        ToYear = YearSelected
    Case "Autumn"
        FromDay = "01"
        FromMonth = "09"
        FromYear = YearSelected
        ToDay = "25"
        ToMonth = "12"
        ToYear = YearSelected
    Case "Summer"
        FromDay = "01"
        FromMonth = "04"
        FromYear = YearSelected
        ToDay = "31"
        ToMonth = "08"
        ToYear = YearSelected
End Select


Set rsReport = New ADODB.Recordset

OpenString = "SELECT tblBook.ISBN, tblBook.Title, tblBook.Author, tblBook.Publisher, tblBookCopy.SellingPrice, tblBookCopy.Date, tblBookCopy.Alphacode, tblBookCopy.ISBN, tblStudent.FirstName, tblStudent.LastName, tblStudent.Alphacode FROM tblBook, tblBookCopy, tblStudent WHERE tblBook.ISBN=tblBookCopy.ISBN AND tblBookCopy.Date Between #" & FromMonth & "/" & FromDay & "/" & FromYear & "# And #" & ToMonth & "/" & ToDay & "/" & ToYear & "# AND tblBookCopy.Alphacode='" & Alphacode & "' AND tblStudent.Alphacode=tblBookCopy.Alphacode ORDER BY tblBookCopy.Alphacode"

'Open the recordset for the report
rsReport.Open OpenString, cnnGlobalConnection, adOpenStatic

If rsReport.EOF = True Then
'If EOF occurs immediatly (ie. no records are returned) then give an error
    MsgBox "No books were bought by this student in this date period!", vbInformation
    Exit Sub
End If

'Show the report
rptStudentReport.Show
End Sub

Private Sub Form_Load()
Dim YearAdd As Integer
Dim CurrentYear As Integer

Set rsTest = New ADODB.Recordset

rsTest.Open "SELECT * FROM tblStudent WHERE Old=False ORDER BY LastName", cnnGlobalConnection, adOpenStatic

Do Until rsTest.EOF
    'Load all students into the list box
    Set Item = lvResults.ListItems.Add(, , rsTest!LastName)
    Item.ListSubItems.Add , , rsTest!FirstName
    Item.ListSubItems.Add , , rsTest!Alphacode
    rsTest.MoveNext
Loop
rsTest.Close

CurrentYear = Year(Date)

'Display all years from 2000 to the current year in the year combo box
For YearAdd = 2000 To CurrentYear
    cmbTermYear.AddItem YearAdd
Next YearAdd


'Set the year combo box to the current year
cmbTermYear.Text = Year(Date)

Select Case CurrentTerm
'Select the current term's option button
    Case "Autumn"
        optAutumn.Value = True
    Case "Summer"
        optSummer.Value = True
    Case "Spring"
        optSpring.Value = True
End Select
End Sub

Private Sub lvResults_DblClick()
'Do the same as if you clicked on the View button
cmdView_Click
End Sub

Private Sub txtSearch_Change()
'Clear all the students in the list box at the moment
lvResults.ListItems.Clear
If chkOld.Value = 0 Then
    'Don't include old students (Old=False)
    rsTest.Open "SELECT * FROM tblStudent WHERE LastName LIKE '%" & txtSearch.Text & "%' AND Old=False ORDER BY LastName", cnnGlobalConnection, adOpenStatic
Else
    'Include old students
    rsTest.Open "SELECT * FROM tblStudent WHERE LastName LIKE '%" & txtSearch.Text & "%' ORDER BY LastName", cnnGlobalConnection, adOpenStatic
End If

Do Until rsTest.EOF
    'Add the students to the list box
    Set Item = lvResults.ListItems.Add(, , rsTest!LastName)
    Item.ListSubItems.Add , , rsTest!FirstName
    Item.ListSubItems.Add , , rsTest!Alphacode
    rsTest.MoveNext
Loop
rsTest.Close
End Sub
