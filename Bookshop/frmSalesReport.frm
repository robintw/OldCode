VERSION 5.00
Begin VB.Form frmSalesReport 
   Caption         =   "Create Sales Report"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   735
      Left            =   2400
      TabIndex        =   14
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Frame frmDate 
      Caption         =   "Date Period"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton optMonth 
         Caption         =   "Select by &month and year"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optTerm 
         Caption         =   "Select by &term"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Frame frmTerm 
         Caption         =   "Term"
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   3735
         Begin VB.OptionButton optSpring 
            Caption         =   "S&pring"
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton optAutumn 
            Caption         =   "&Autumn"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton optSummer 
            Caption         =   "&Summer"
            Height          =   255
            Left            =   2640
            TabIndex        =   10
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cmbTermYear 
            Height          =   315
            Left            =   600
            TabIndex        =   11
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Year:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         Caption         =   "OR"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblMonth 
         Caption         =   "&Month:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "&Year:"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdView_Click()
Dim Term As String
Dim FromDay
Dim FromMonth
Dim FromYear
Dim ToDay
Dim ToMonth
Dim ToYear
Dim OpenString As String
Dim YearSelected

'If we are selecting dates by term then...
If optTerm.Item(0).Value = True Then
    'Select the correct term by which option buttons are set
    If optSpring.Value = True Then
        Term = "Spring"
    End If
    If optSummer.Value = True Then
        Term = "Summer"
    End If
    If optAutumn.Value = True Then
        Term = "Autumn"
    End If
    
    'If the year drop down has no value then use the current year
    If cmbTermYear.Text = "" Then
        YearSelected = Year(Date)
    Else
        YearSelected = cmbTermYear.Text
    End If
    
    'Set the dates appropriately depending on the term selected
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
Else
'We are selecting dates by month and year so...

    YearSelected = cmbYear.Text
    'Every period will go from the 1st of one month to the 1st of the next month
    FromDay = "01"
    ToDay = "01"
    'And from the selected year to the selected year
    FromYear = YearSelected
    ToYear = YearSelected
    'Set the month values appropriately
    Select Case cmbMonth.Text
        Case "January"
            FromMonth = "01"
            ToMonth = "02"
        Case "Febuary"
            FromMonth = "02"
            ToMonth = "03"
        Case "March"
            FromMonth = "03"
            ToMonth = "04"
        Case "April"
            FromMonth = "04"
            ToMonth = "05"
        Case "May"
            FromMonth = "05"
            ToMonth = "06"
        Case "June"
            FromMonth = "06"
            ToMonth = "07"
        Case "July"
            FromMonth = "07"
            ToMonth = "08"
        Case "August"
            FromMonth = "08"
            ToMonth = "09"
        Case "September"
            FromMonth = "09"
            ToMonth = "10"
        Case "October"
            FromMonth = "10"
            ToMonth = "11"
        Case "November"
            FromMonth = "11"
            ToMonth = "12"
        Case "December"
            FromMonth = "12"
            ToMonth = "01"
            ToYear = FromYear + 1
    End Select
End If

Set rsSales = New ADODB.Recordset

OpenString = "SELECT tblBook.ISBN, tblBook.Title, tblBook.Author, tblBook.Publisher, tblBookCopy.SellingPrice, tblBookCopy.CostPrice,  tblBookCopy.ISBN, tblBookCopy.Date FROM tblBook, tblBookCopy WHERE tblBook.ISBN=tblBookCopy.ISBN AND tblBookCopy.Date Between #" & FromMonth & "/" & FromDay & "/" & FromYear & "# And #" & ToMonth & "/" & ToDay & "/" & ToYear & "# ORDER BY tblBookCopy.Date"
'Open the recordset needed for the report
rsSales.Open OpenString, cnnGlobalConnection, adOpenStatic

If rsSales.EOF = True Then
    'If there is EOf already (ie. there are no records) then give an error
    MsgBox "No sales occured in this period!", vbInformation
    Exit Sub
End If

'Show the report
rptSales.Show
End Sub

Private Sub Form_Load()
Dim CurrentYear As Integer
Dim YearAdd As Integer

CurrentYear = Year(Date)

'Fill the combo box with years from 2000 to the current year
For YearAdd = 2000 To CurrentYear
    cmbYear.AddItem YearAdd
    cmbTermYear.AddItem YearAdd
Next YearAdd

cmbMonth.AddItem "January"
cmbMonth.AddItem "February"
cmbMonth.AddItem "March"
cmbMonth.AddItem "April"
cmbMonth.AddItem "May"
cmbMonth.AddItem "June"
cmbMonth.AddItem "July"
cmbMonth.AddItem "August"
cmbMonth.AddItem "September"
cmbMonth.AddItem "October"
cmbMonth.AddItem "November"
cmbMonth.AddItem "December"


'Set all the drop downs to the current month or year
cmbYear.Text = Year(Date)
cmbTermYear.Text = Year(Date)
cmbMonth.Text = MonthName(Month(Date))

'Enable the appropriate buttons and boxes
If optTerm.Item(0).Value = True Then
    cmbYear.Enabled = False
    cmbMonth.Enabled = False
Else
    optAutumn.Enabled = False
    optSummer.Enabled = False
    optSpring.Enabled = False
    cmbTermYear.Enabled = False
End If

If Month(Date) >= 1 And Month(Date) <= 4 Then
    optSpring.Value = True
ElseIf Month(Date) >= 9 And Month(Date) <= 12 Then
    optAutumn.Value = True
ElseIf Month(Date) >= 4 And Month(Date) <= 8 Then
    optSummer.Value = True
End If

End Sub

Private Sub optMonth_Click(Index As Integer)
If optTerm.Item(0).Value = True Then
    cmbYear.Enabled = False
    cmbMonth.Enabled = False
    optAutumn.Enabled = True
    optSummer.Enabled = True
    optSpring.Enabled = True
    cmbTermYear.Enabled = True
Else
    cmbYear.Enabled = True
    cmbMonth.Enabled = True
    optAutumn.Enabled = False
    optSummer.Enabled = False
    optSpring.Enabled = False
    cmbTermYear.Enabled = False
End If
End Sub

Private Sub optTerm_Click(Index As Integer)
'When the button is clicked enable the appropriate controls
If optTerm.Item(0).Value = True Then
    cmbYear.Enabled = False
    cmbMonth.Enabled = False
    optAutumn.Enabled = True
    optSummer.Enabled = True
    optSpring.Enabled = True
    cmbTermYear.Enabled = True
Else
    cmbYear.Enabled = True
    cmbMonth.Enabled = True
    optAutumn.Enabled = False
    optSummer.Enabled = False
    optSpring.Enabled = False
    cmbTermYear.Enabled = False
End If
End Sub
