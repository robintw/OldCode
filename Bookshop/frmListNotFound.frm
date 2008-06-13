VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmListNotFound 
   Caption         =   "Stock Check - Not Found Books"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   855
      Left            =   6840
      TabIndex        =   3
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid flxNotFound 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10186
      _Version        =   393216
      Rows            =   5
      FixedCols       =   0
      FocusRect       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmListNotFound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim cmdDelete As ADODB.Command
Set cmdDelete = New ADODB.Command
cmdDelete.ActiveConnection = cnnGlobalConnection

'Removes the currently selected book from tblNotFound
cmdDelete.CommandText = "DELETE FROM tblNotFound WHERE NotFoundID=" & flxNotFound.TextMatrix(flxNotFound.RowSel, 0)
cmdDelete.Execute

'Sets CurrentBookDetails up ready for loading frmAddBook
CurrentBookDetails.ISBN = flxNotFound.TextMatrix(flxNotFound.RowSel, 1)
CurrentBookDetails.Title = flxNotFound.TextMatrix(flxNotFound.RowSel, 2)
CurrentBookDetails.Author = flxNotFound.TextMatrix(flxNotFound.RowSel, 3)
CurrentBookDetails.Publisher = flxNotFound.TextMatrix(flxNotFound.RowSel, 4)

'Sets the fields on frmAddBook
With frmAddBook
.txtISBN = flxNotFound.TextMatrix(flxNotFound.RowSel, 1)
.txtTitle = flxNotFound.TextMatrix(flxNotFound.RowSel, 2)
.txtAuthor = flxNotFound.TextMatrix(flxNotFound.RowSel, 3)
.txtPublisher = flxNotFound.TextMatrix(flxNotFound.RowSel, 4)
'Shows frmAddBook and sets the focus to the cost price field
.Show
.txtCostPrice.SetFocus
End With

Dim rsNotFound As ADODB.Recordset
Set rsNotFound = New ADODB.Recordset

rsNotFound.Open "SELECT ISBN FROM tblBook WHERE ISBN='" & CurrentBookDetails.ISBN & "'", cnnGlobalConnection, adOpenStatic

If rsNotFound.RecordCount = 1 Then
    'If the book details are in tblBook then don't add them again
    frmAddBook.FoundInDb = True
Else
    'Otherwise set FoundInDB to False so we add them when the addition is confirmed
    frmAddBook.FoundInDb = False
End If


'Can't seem to set the flex grid to only one row - so make the 2nd row blank instead
If flxNotFound.Rows = 2 Then
    flxNotFound.TextMatrix(1, 0) = ""
    flxNotFound.TextMatrix(1, 1) = ""
    flxNotFound.TextMatrix(1, 2) = ""
    flxNotFound.TextMatrix(1, 3) = ""
    flxNotFound.TextMatrix(1, 4) = ""
Else
    flxNotFound.RemoveItem (flxNotFound.RowSel)
End If
Set cmdDelete = Nothing
End Sub

Private Sub cmdFinish_Click()
'Close and unload form
Unload Me
End Sub

Private Sub cmdIgnore_Click()
Dim cmdDelete As ADODB.Command
Set cmdDelete = New ADODB.Command
cmdDelete.ActiveConnection = cnnGlobalConnection

'Delete selected book from tblNotFound
cmdDelete.CommandText = "DELETE FROM tblNotFound WHERE NotFoundID=" & flxNotFound.TextMatrix(flxNotFound.RowSel, 0)
cmdDelete.Execute

'Can't seem to set the flex grid to only one row - so make the 2nd row blank instead
If flxNotFound.Rows = 2 Then
    flxNotFound.TextMatrix(1, 0) = ""
    flxNotFound.TextMatrix(1, 1) = ""
    flxNotFound.TextMatrix(1, 2) = ""
    flxNotFound.TextMatrix(1, 3) = ""
    flxNotFound.TextMatrix(1, 4) = ""
Else
    flxNotFound.RemoveItem (flxNotFound.RowSel)
End If

Set cmdDelete = Nothing
End Sub

Private Sub Form_Load()
Dim rsNotFound As ADODB.Recordset
Set rsNotFound = New ADODB.Recordset

'Set up FlexGrid
With flxNotFound
    .Cols = 5
    .Rows = 1
    .TextMatrix(0, 0) = "ID"
    .TextMatrix(0, 1) = "ISBN"
    .TextMatrix(0, 2) = "Title"
    .TextMatrix(0, 3) = "Author"
    .TextMatrix(0, 4) = "Publisher"
    .ColWidth(0) = 405
    .ColWidth(1) = 1155
    .ColWidth(2) = 3990
    .ColWidth(3) = 2625
    .ColWidth(4) = 1095
End With

'Load in all records from tblNotFound
rsNotFound.Open "SELECT * FROM tblNotFound", cnnGlobalConnection, adOpenStatic

Do Until rsNotFound.EOF
    flxNotFound.Rows = flxNotFound.Rows + 1
    flxNotFound.TextMatrix(flxNotFound.Rows - 1, 0) = rsNotFound!NotFoundID
    flxNotFound.TextMatrix(flxNotFound.Rows - 1, 1) = rsNotFound!ISBN
    flxNotFound.TextMatrix(flxNotFound.Rows - 1, 2) = rsNotFound!Title
    flxNotFound.TextMatrix(flxNotFound.Rows - 1, 3) = rsNotFound!Author
    flxNotFound.TextMatrix(flxNotFound.Rows - 1, 4) = rsNotFound!Publisher
    rsNotFound.MoveNext
Loop

flxNotFound.SelectionMode = flexSelectionByRow
flxNotFound.Row = 0

Set rsNotFound = Nothing

End Sub
