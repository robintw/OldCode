VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBookEdit 
   Caption         =   "Edit Books"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetDetailsFromNet 
      Caption         =   "&Get Details"
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
      Text            =   "Title"
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
      Caption         =   "S&earch"
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
      Text            =   "Title"
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
      Begin VB.CommandButton Command1 
         Caption         =   "&Delete"
         Height          =   855
         Left            =   3960
         TabIndex        =   23
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtSellingPrice 
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCostPrice 
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtCopyID 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtPublisher 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtAuthor 
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtTitle 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtISBN 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   1095
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
         Caption         =   "S&elling Price:"
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "&Cost Price:"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Copy &ID:"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblPublisher 
         Caption         =   "&Publisher:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblAuthor 
         Caption         =   "&Author:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "&Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblISBN 
         Caption         =   "&ISBN:"
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
Attribute VB_Name = "frmBookEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEdit As ADODB.Recordset

Private Sub DoSearch()
Select Case cmbSearchBy.Text
    Case "Title"
        ImportData cmbSortBy.Text, txtSearchTerm.Text, "", "", ""
    Case "Author"
        ImportData cmbSortBy.Text, "", txtSearchTerm.Text, "", ""
    Case "Publisher"
        ImportData cmbSortBy.Text, "", "", txtSearchTerm.Text, ""
    Case "ISBN"
        ImportData cmbSortBy.Text, "", "", "", txtSearchTerm.Text
End Select
End Sub

Private Sub cmdGetDetailsFromNet_Click()
Dim BookInfo As udfBookDetails

BookInfo = GetDetailsFromInternet(txtISBN.Text)

txtTitle.Text = BookInfo.Title
txtAuthor.Text = BookInfo.Author
txtPublisher.Text = BookInfo.Publisher

End Sub

Private Sub cmdSave_Click()
Dim cmdUpdateDetails As ADODB.Command
Set cmdUpdateDetails = New ADODB.Command

cmdUpdateDetails.ActiveConnection = cnnGlobalConnection



cmdUpdateDetails.CommandText = "UPDATE tblBook SET Title='" & Replace(txtTitle.Text, "'", "''") & "', Author='" & Replace(txtAuthor.Text, "'", "''") & "', Publisher='" & Replace(txtPublisher.Text, "'", "''") & "' WHERE ISBN='" & txtISBN.Text & "';"
cmdUpdateDetails.Execute

cmdUpdateDetails.CommandText = "UPDATE tblBookCopy SET CostPrice=" & txtCostPrice.Text & ", SellingPrice=" & txtSellingPrice.Text & " WHERE CopyID=" & txtCopyID.Text & ";"
cmdUpdateDetails.Execute


cmdSave.BackColor = &H8000000F
DoSearch
End Sub

Private Sub cmdSearch_Click()
DoSearch
End Sub

Private Sub cmdSort_Click()
DoSearch
End Sub
Private Sub Command1_Click()
If MsgBox("Are you sure you want to delete this book. This action is irreversable.", vbYesNo) = vbNo Then
    Exit Sub
End If

Dim cmdDelete As ADODB.Command
Set cmdDelete = New ADODB.Command

cmdDelete.ActiveConnection = cnnGlobalConnection
cmdDelete.CommandText = "DELETE FROM tblBookCopy WHERE CopyID=" & txtCopyID.Text & ";"

cmdDelete.Execute

DoSearch
End Sub

Private Sub flxList_Click()
LoadDetails (Int(flxList.Text))
End Sub

Private Sub LoadDetails(BookID As Integer)
Dim rsDetails As ADODB.Recordset
Set rsDetails = New ADODB.Recordset

rsDetails.ActiveConnection = cnnGlobalConnection

rsDetails.Open "SELECT * FROM tblBook, tblBookCopy WHERE tblBook.ISBN = tblBookCopy.ISBN AND tblBookCopy.CopyID = " & BookID & ";", , adOpenStatic

txtCopyID.Text = rsDetails!CopyID
txtISBN.Text = rsDetails.Fields("tblBook.ISBN")
txtTitle.Text = rsDetails!Title
txtAuthor.Text = rsDetails!Author
txtPublisher.Text = rsDetails!Publisher
txtCostPrice.Text = rsDetails!CostPrice
txtSellingPrice.Text = rsDetails!SellingPrice

cmdSave.BackColor = &H8000000F
End Sub

Private Sub Form_Load()
cmbSortBy.AddItem "Title"
cmbSortBy.AddItem "Author"
cmbSortBy.AddItem "Publisher"

cmbSearchBy.AddItem "Title"
cmbSearchBy.AddItem "Author"
cmbSearchBy.AddItem "Publisher"
cmbSearchBy.AddItem "ISBN"

Set rsEdit = New ADODB.Recordset

rsEdit.ActiveConnection = cnnGlobalConnection

With flxList
    .Cols = 7
    .Rows = 1
    .ColWidth(0) = 405
    .ColWidth(1) = 1155
    .ColWidth(2) = 3790
    .ColWidth(3) = 2425
    .ColWidth(4) = 1700
    .ColWidth(5) = 740
    .ColWidth(6) = 740
End With

flxList.ColAlignment(5) = 7
flxList.ColAlignment(6) = 7

CreateHeadings

'Default sorting is by Title
ImportData "Title", "", "", "", ""
flxList.Refresh

End Sub

Private Sub CreateHeadings()
flxList.TextMatrix(0, 0) = "ID"
flxList.TextMatrix(0, 1) = "ISBN"
flxList.TextMatrix(0, 2) = "Title"
flxList.TextMatrix(0, 3) = "Author"
flxList.TextMatrix(0, 4) = "Publisher"
flxList.TextMatrix(0, 5) = "Cost"
flxList.TextMatrix(0, 6) = "Sell"
End Sub

Private Sub ImportData(Sorting As String, TitleSearch As String, AuthorSearch As String, PublisherSearch As String, ISBNSearch As String)
flxList.Clear
CreateHeadings

flxList.Rows = 1

Select Case Sorting
    Case "Title"
        rsEdit.Open "SELECT tblBookCopy.CostPrice, tblBookCopy.SellingPrice, tblBookCopy.CopyID, tblBookCopy.ISBN, tblBook.Title, tblBook.Author, tblBook.Publisher FROM tblBookCopy, tblBook WHERE tblBookCopy.ISBN = tblBook.ISBN AND tblBookCopy.Alphacode IS NULL AND tblBook.Title LIKE'%" & TitleSearch & "%' AND tblBook.Author LIKE'%" & AuthorSearch & "%' AND tblBook.Publisher LIKE'%" & PublisherSearch & "%' AND tblBook.ISBN LIKE'%" & ISBNSearch & "%' ORDER BY tblBook.Title;", cnnGlobalConnection, adOpenStatic
    Case "Author"
        rsEdit.Open "SELECT tblBookCopy.CostPrice, tblBookCopy.SellingPrice, tblBookCopy.CopyID, tblBookCopy.ISBN, tblBook.Title, tblBook.Author, tblBook.Publisher FROM tblBookCopy, tblBook WHERE tblBookCopy.ISBN = tblBook.ISBN AND tblBookCopy.Alphacode IS NULL AND tblBook.Title LIKE'%" & TitleSearch & "%' AND tblBook.Author LIKE'%" & AuthorSearch & "%' AND tblBook.Publisher LIKE'%" & PublisherSearch & "%' AND tblBook.ISBN LIKE'%" & ISBNSearch & "%' ORDER BY tblBook.Author;", cnnGlobalConnection, adOpenStatic
    Case "Publisher"
        rsEdit.Open "SELECT tblBookCopy.CostPrice, tblBookCopy.SellingPrice, tblBookCopy.CopyID, tblBookCopy.ISBN, tblBook.Title, tblBook.Author, tblBook.Publisher FROM tblBookCopy, tblBook WHERE tblBookCopy.ISBN = tblBook.ISBN AND tblBookCopy.Alphacode IS NULL AND tblBook.Title LIKE'%" & TitleSearch & "%' AND tblBook.Author LIKE'%" & AuthorSearch & "%' AND tblBook.Publisher LIKE'%" & PublisherSearch & "%' AND tblBook.ISBN LIKE'%" & ISBNSearch & "%' ORDER BY tblBook.Publisher;", cnnGlobalConnection, adOpenStatic
End Select

Do Until rsEdit.EOF
    With flxList
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rsEdit!CopyID
        .TextMatrix(.Rows - 1, 1) = rsEdit!ISBN
        .TextMatrix(.Rows - 1, 2) = rsEdit!Title
        .TextMatrix(.Rows - 1, 3) = rsEdit!Author
        .TextMatrix(.Rows - 1, 4) = rsEdit!Publisher
        .TextMatrix(.Rows - 1, 5) = FormatCurrency(rsEdit!CostPrice)
        .TextMatrix(.Rows - 1, 6) = FormatCurrency(rsEdit!SellingPrice)
        rsEdit.MoveNext
    End With
Loop

rsEdit.Close

flxList.SelectionMode = flexSelectionByRow
flxList.Row = 0

End Sub

Private Sub txtAuthor_Change()
cmdSave.BackColor = &H8080FF
End Sub

Private Sub txtCostPrice_Change()
cmdSave.BackColor = &H8080FF
End Sub

Private Sub txtPublisher_Change()
cmdSave.BackColor = &H8080FF
End Sub

Private Sub txtSearchTerm_Change()
DoSearch
End Sub

Private Sub txtSellingPrice_Change()
cmdSave.BackColor = &H8080FF
End Sub

Private Sub txtTitle_Change()
cmdSave.BackColor = &H8080FF
End Sub
