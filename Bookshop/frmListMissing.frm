VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmListMissing 
   Caption         =   "Stock Check - Missing books"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore (Found)"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   855
      Left            =   6840
      TabIndex        =   3
      Top             =   6000
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid flxMissing 
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
Attribute VB_Name = "frmListMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFinish_Click()
'Close and unload form
Unload Me
End Sub

Private Sub cmdIgnore_Click()
Dim cmdDelete As ADODB.Command
Set cmdDelete = New ADODB.Command
cmdDelete.ActiveConnection = cnnGlobalConnection

'Delete the selected record from the database
cmdDelete.CommandText = "UPDATE tblBookCopy SET StockCheck = TRUE WHERE CopyID=" & flxMissing.TextMatrix(flxMissing.RowSel, 0)
cmdDelete.Execute

'Can't seem to set the flex grid to only one row - so make the 2nd row blank instead
If flxMissing.Rows = 2 Then
    flxMissing.TextMatrix(1, 0) = ""
    flxMissing.TextMatrix(1, 1) = ""
    flxMissing.TextMatrix(1, 2) = ""
    flxMissing.TextMatrix(1, 3) = ""
    flxMissing.TextMatrix(1, 4) = ""
    cmdRemove.Enabled = False
    cmdIgnore.Enabled = False
    cmdFinish.Default = True
Else
    'Just remove the row of the flex grid - as there are enough left for it not to cause an error
    flxMissing.RemoveItem (flxMissing.RowSel)
End If

Set cmdDelete = Nothing
End Sub

Private Sub cmdRemove_Click()
Dim cmdDelete As ADODB.Command
Set cmdDelete = New ADODB.Command
cmdDelete.ActiveConnection = cnnGlobalConnection

cmdDelete.CommandText = "DELETE FROM tblBookCopy WHERE CopyID=" & flxMissing.TextMatrix(flxMissing.RowSel, 0)
cmdDelete.Execute


If flxMissing.Rows = 2 Then
'Can't seem to set the flex grid to only one row - so make the 2nd row blank instead
    flxMissing.TextMatrix(1, 0) = ""
    flxMissing.TextMatrix(1, 1) = ""
    flxMissing.TextMatrix(1, 2) = ""
    flxMissing.TextMatrix(1, 3) = ""
    flxMissing.TextMatrix(1, 4) = ""
    cmdRemove.Enabled = False
    cmdIgnore.Enabled = False
    cmdFinish.Default = True
Else
    'Just remove the row of the flex grid - as there are enough left for it not to cause an error
    flxMissing.RemoveItem (flxMissing.RowSel)
End If

Set cmdDelete = Nothing
End Sub

Private Sub Form_Load()
Dim rsMissing As ADODB.Recordset
Set rsMissing = New ADODB.Recordset

'Set up FlexGrid
With flxMissing
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
rsMissing.Open "SELECT tblBookCopy.CopyID, tblBookCopy.ISBN, tblBook.Title, tblBook.Author, tblBook.Publisher FROM tblBookCopy, tblBook WHERE StockCheck=FALSE AND tblBookCopy.ISBN = tblBook.ISBN", cnnGlobalConnection, adOpenStatic

Do Until rsMissing.EOF
    With flxMissing
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rsMissing!CopyID
        .TextMatrix(.Rows - 1, 1) = rsMissing!ISBN
        .TextMatrix(.Rows - 1, 2) = rsMissing!Title
        .TextMatrix(.Rows - 1, 3) = rsMissing!Author
        .TextMatrix(.Rows - 1, 4) = rsMissing!Publisher
        rsMissing.MoveNext
    End With
Loop

flxMissing.SelectionMode = flexSelectionByRow
flxMissing.Row = 0

Set rsMissing = Nothing

End Sub
