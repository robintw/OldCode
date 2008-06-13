VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmStockCheck 
   Caption         =   "Stock Check"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish Stock Check"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   7680
      Width           =   12375
   End
   Begin VB.ComboBox cmbPrice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid flxScanned 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10186
      _Version        =   393216
      FixedCols       =   0
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.TextBox txtISBN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label lblQmark 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   72
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   8280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblCross 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   99.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   8280
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblTick 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   99.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblPrice 
      Caption         =   "&Price:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblISBN 
      Caption         =   "&ISBN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmStockCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBooks As ADODB.Recordset

Private Sub cmbPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'When enter is pressed...
        Dim rsDifferentPriceBooks As ADODB.Recordset
        Set rsDifferentPriceBooks = New ADODB.Recordset
        Dim cmdUpdate As ADODB.Command
        Set cmdUpdate = New ADODB.Command
        cmdUpdate.ActiveConnection = cnnGlobalConnection
        
        rsDifferentPriceBooks.Open _
            "SELECT tblBookCopy.CopyID, tblBookCopy.ISBN, tblBookCopy.CostPrice, tblBookCopy.SellingPrice, tblBook.Title, tblBook.Author, tblBook.Publisher from tblBookCopy, tblBook WHERE tblBookCopy.ISBN='" _
            & txtISBN.Text & "' AND tblBook.ISBN = tblBookCopy.ISBN AND tblBookCopy.Alphacode IS NULL AND tblBookCopy.StockCheck = FALSE AND SellingPrice =" & cmbPrice.Text, cnnGlobalConnection, adOpenStatic
        
        If rsDifferentPriceBooks.RecordCount = 1 Then
            'If there is only one book with the selected price then set the StockCheck field to true
            cmdUpdate.CommandText = "UPDATE tblBookCopy SET StockCheck = TRUE WHERE ISBN='" & txtISBN.Text & "' AND SellingPrice =" & cmbPrice.Text
            cmdUpdate.Execute
            'Add the book to the list on the form
            flxScanned.Rows = flxScanned.Rows + 1
            flxScanned.TextMatrix(flxScanned.Rows - 1, 0) = rsDifferentPriceBooks!CopyID
            flxScanned.TextMatrix(flxScanned.Rows - 1, 1) = rsDifferentPriceBooks!ISBN
            flxScanned.TextMatrix(flxScanned.Rows - 1, 2) = rsDifferentPriceBooks!Title
            flxScanned.TextMatrix(flxScanned.Rows - 1, 3) = rsDifferentPriceBooks!Author
            flxScanned.TextMatrix(flxScanned.Rows - 1, 4) = rsDifferentPriceBooks!Publisher
            flxScanned.TextMatrix(flxScanned.Rows - 1, 5) = FormatCurrency(rsDifferentPriceBooks!SellingPrice)
            'Clear all the boxes on the form and set the focus back
            ClearBoxes
        Else
        'There is more than one book with that price - so set the StockCheck field to true on the selected CopyID
            'Set the StockCheck field to true
            cmdUpdate.CommandText = "UPDATE tblBookCopy SET StockCheck = TRUE WHERE ISBN='" & txtISBN.Text & "' AND SellingPrice =" & cmbPrice.Text & " AND CopyID =" & rsDifferentPriceBooks!CopyID
            cmdUpdate.Execute
            'Add the book to the list on the form
            flxScanned.Rows = flxScanned.Rows + 1
            flxScanned.TextMatrix(flxScanned.Rows - 1, 0) = rsDifferentPriceBooks!CopyID
            flxScanned.TextMatrix(flxScanned.Rows - 1, 1) = rsDifferentPriceBooks!ISBN
            flxScanned.TextMatrix(flxScanned.Rows - 1, 2) = rsDifferentPriceBooks!Title
            flxScanned.TextMatrix(flxScanned.Rows - 1, 3) = rsDifferentPriceBooks!Author
            flxScanned.TextMatrix(flxScanned.Rows - 1, 4) = rsDifferentPriceBooks!Publisher
            flxScanned.TextMatrix(flxScanned.Rows - 1, 5) = FormatCurrency(rsDifferentPriceBooks!SellingPrice)
            'Clear all the boxes on the form and set the focus back
            ClearBoxes
        End If
        
End If
Set rsDifferentPriceBooks = Nothing
Set cmdUpdate = Nothing
End Sub


Private Sub cmdFinish_Click()
If MsgBox("Are you sure you want to finish the stock check?", vbYesNo Or vbQuestion) = vbNo Then
    Exit Sub
End If
'If the user doesn't say no (ie. says yes) then frmFetchDetails is shown to fetch details
'of all the books that aren't found
frmFetchDetails.Show
End Sub

Private Sub Form_Load()
'Initialise the flex grid - rows, columns, column titles and column widths
flxScanned.Rows = 1
flxScanned.Cols = 6

flxScanned.TextMatrix(0, 0) = "CopyID"
flxScanned.TextMatrix(0, 1) = "ISBN"
flxScanned.TextMatrix(0, 2) = "Title"
flxScanned.TextMatrix(0, 3) = "Author"
flxScanned.TextMatrix(0, 4) = "Publisher"
flxScanned.TextMatrix(0, 5) = "SellingPrice"

flxScanned.ColWidth(0) = 645
flxScanned.ColWidth(1) = 1005
flxScanned.ColWidth(2) = 5745
flxScanned.ColWidth(3) = 2385
flxScanned.ColWidth(4) = 1530
flxScanned.ColWidth(5) = 960

'Allows the user to continue with a previous stock check if it takes more than one session to check all the books
If MsgBox("Do you want to start a new stock check", vbYesNo Or vbQuestion) = vbYes Then
    'If a new stock check has been started then set all StockCheck fields to false, and delete everything in tblNotFound
    Dim cmdClear As ADODB.Command
    Set cmdClear = New ADODB.Command
    
    cmdClear.ActiveConnection = cnnGlobalConnection
    
    cmdClear.CommandText = "UPDATE tblBookCopy SET StockCheck = FALSE"
    cmdClear.Execute
    cmdClear.CommandText = "DELETE FROM tblNotFound"
    cmdClear.Execute
    
Else
'We're continuing with an old stock check
    Dim rsOldCheck As ADODB.Recordset
    Set rsOldCheck = New ADODB.Recordset
    rsOldCheck.Open "SELECT tblBookCopy.CopyID, tblBookCopy.ISBN, tblBook.Title, tblBook.Author, tblBook.Publisher, tblBookCopy.SellingPrice FROM tblBookCopy, tblBook WHERE StockCheck=True AND tblBook.ISBN = tblBookCopy.ISBN", cnnGlobalConnection, adOpenStatic
    
    Do Until rsOldCheck.EOF
        'For each record in the recordset add the record to the list on the form
        'so it looks just like it did when the user closed the form last time
        flxScanned.Rows = flxScanned.Rows + 1
        flxScanned.TextMatrix(flxScanned.Rows - 1, 0) = rsOldCheck!CopyID
        flxScanned.TextMatrix(flxScanned.Rows - 1, 1) = rsOldCheck!ISBN
        flxScanned.TextMatrix(flxScanned.Rows - 1, 2) = rsOldCheck!Title
        flxScanned.TextMatrix(flxScanned.Rows - 1, 3) = rsOldCheck!Author
        flxScanned.TextMatrix(flxScanned.Rows - 1, 4) = rsOldCheck!Publisher
        flxScanned.TextMatrix(flxScanned.Rows - 1, 5) = FormatCurrency(rsOldCheck!SellingPrice)
        
        rsOldCheck.MoveNext
    Loop

End If

Set cmdClear = Nothing
End Sub

Private Sub txtISBN_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then

Select Case Len(txtISBN.Text)
    Case 13
        'If its 13 Chars long then it is an EAN - so convert to ISBN
        txtISBN.Text = ConvertEANToISBN(txtISBN.Text)
    Case 10
        'If its 10 then it is an ISBN - its fine but validate it
        If ValidateISBN(txtISBN.Text) = False Then
            MsgBox "Please enter a valid ISBN"
            Exit Sub
        End If
    Case Else
        'Otherwise make it blank so that it is caught below
        txtISBN.Text = ""
End Select

If txtISBN.Text = "" Then
    MsgBox "Error: The EAN/ISBN entered was not valid, please re-scan the book", vbCritical
    txtISBN.SetFocus
    Exit Sub
End If

Set rsBooks = New ADODB.Recordset

Dim FirstPrice As Single
Dim PricesDifferent As Boolean
Dim CurrentCopyID As Integer


Dim cmdUpdate As ADODB.Command
Set cmdUpdate = New ADODB.Command
cmdUpdate.ActiveConnection = cnnGlobalConnection

PricesDifferent = False

    rsBooks.Open _
        "SELECT tblBookCopy.CopyID, tblBookCopy.ISBN, tblBookCopy.SellingPrice, tblBook.Title, tblBook.Author, tblBook.Publisher from tblBookCopy, tblBook WHERE tblBookCopy.ISBN='" _
        & txtISBN.Text & _
        "' AND tblBook.ISBN = tblBookCopy.ISBN AND tblBookCopy.Alphacode IS NULL AND tblBookCopy.StockCheck = FALSE ORDER BY SellingPrice", _
        cnnGlobalConnection, adOpenStatic
    
    If rsBooks.RecordCount = 0 Then
    'No books that match the ISBN entered
    'Show the red cross, and give an error message
        lblCross.Visible = True
        lblTick.Visible = False
        lblQmark.Visible = False
        MsgBox "The ISBN entered does not match a book in the database." & _
            vbCrLf & vbCrLf & _
            "Please put this book in a separate pile, it will needed once the stock check is finished.", _
            vbCritical, "Book not found"
        'Add a record to tblNotFound with the books ISBN - so it can be listed easily later
        Dim cmdInsert As ADODB.Command
        Set cmdInsert = New ADODB.Command
        cmdInsert.ActiveConnection = cnnGlobalConnection
        cmdInsert.CommandText = "INSERT INTO tblNotFound (ISBN) VALUES ('" & txtISBN.Text & "')"
        cmdInsert.Execute
        'Clear all the boxes on the form and set the focus back
        ClearBoxes
        Exit Sub
    ElseIf rsBooks.RecordCount = 1 Then
        'Only one book so can't be different prices
        PricesDifferent = False
        cmbPrice.Clear
        cmbPrice.Enabled = False
        'Set the StockCheck field to true for the current book
        cmdUpdate.CommandText = _
            "UPDATE tblBookCopy SET StockCheck = TRUE WHERE ISBN='" & txtISBN.Text _
            & "'"
        cmdUpdate.Execute
        'Add the book to the list on the form
        flxScanned.Rows = flxScanned.Rows + 1
        flxScanned.TextMatrix(flxScanned.Rows - 1, 0) = rsBooks!CopyID
        flxScanned.TextMatrix(flxScanned.Rows - 1, 1) = rsBooks!ISBN
        flxScanned.TextMatrix(flxScanned.Rows - 1, 2) = rsBooks!Title
        flxScanned.TextMatrix(flxScanned.Rows - 1, 3) = rsBooks!Author
        flxScanned.TextMatrix(flxScanned.Rows - 1, 4) = rsBooks!Publisher
        flxScanned.TextMatrix(flxScanned.Rows - 1, 5) = FormatCurrency(rsBooks!SellingPrice)
        'Show a tick to say the book has been found
        lblCross.Visible = False
        lblTick.Visible = True
        lblQmark.Visible = False
        'Clear all the boxes on the form and set the focus back
        ClearBoxes
    Else
        'There are multiple books - could be different prices
        FirstPrice = rsBooks.Fields("SellingPrice").Value
        
        PricesDifferent = False
        
        'Step through every book found, checking each price, to see if there are different prices
        Do Until rsBooks.EOF
            If rsBooks.Fields("SellingPrice").Value = FirstPrice Then
                rsBooks.MoveNext
            Else
                'If there are set PricesDifferent to true so we know about it
                PricesDifferent = True
                Exit Do
            End If
        Loop
        
        'Move back to the beginning so that the first price in the combo box is the lowest
        rsBooks.MoveFirst
        
        CurrentCopyID = rsBooks!CopyID
        
        If PricesDifferent = False Then
        'If there are multiple copies but the prices are not different then...
            cmbPrice.Clear
            cmbPrice.Enabled = False
            'Set the StockCheck field to true
            cmdUpdate.CommandText = _
                "UPDATE tblBookCopy SET StockCheck = TRUE WHERE ISBN='" & _
                txtISBN.Text & "' AND CopyID = " & CurrentCopyID
            cmdUpdate.Execute
            'Add the book to the list on the form
            flxScanned.Rows = flxScanned.Rows + 1
            flxScanned.TextMatrix(flxScanned.Rows - 1, 0) = rsBooks!CopyID
            flxScanned.TextMatrix(flxScanned.Rows - 1, 1) = rsBooks!ISBN
            flxScanned.TextMatrix(flxScanned.Rows - 1, 2) = rsBooks!Title
            flxScanned.TextMatrix(flxScanned.Rows - 1, 3) = rsBooks!Author
            flxScanned.TextMatrix(flxScanned.Rows - 1, 4) = rsBooks!Publisher
            flxScanned.TextMatrix(flxScanned.Rows - 1, 5) = FormatCurrency(rsBooks!SellingPrice)
            
            'Show a tick to say the book has been found
            lblCross.Visible = False
            lblTick.Visible = True
            lblQmark.Visible = False
            'Clear all the boxes on the form and set the focus back
            ClearBoxes
            
        Else
        'There are multiple copies and the prices are different
            rsBooks.MoveFirst
            'Highlight the combo box - so the user sees it
            cmbPrice.BackColor = &HC0C0FF
            cmbPrice.Enabled = True
            
            lblPrice.ForeColor = &H80000012
            cmbPrice.SetFocus
            
            Do Until rsBooks.EOF
                'Add all the prices to the combo box
                cmbPrice.AddItem (rsBooks.Fields("SellingPrice").Value)
                rsBooks.MoveNext
            Loop
            'Show a question mark to tell the user that more information is needed
            lblCross.Visible = False
            lblTick.Visible = False
            lblQmark.Visible = True
        End If
        
    End If
Else
            'If the key pressed is not enter then set all the pictures to not be visible
            'so that as soon as you start typing after adding a book the pictures reset ready
            'for the next book
            lblCross.Visible = False
            lblTick.Visible = False
            lblQmark.Visible = False
            
End If

Set cmdUpdate = Nothing
End Sub

Private Sub ClearBoxes()
'Clear the text boxes on the form, reset the highlight, and set the focus back to the ISBN text box
txtISBN.Text = ""
cmbPrice.Clear
cmbPrice.Enabled = False
cmbPrice.BackColor = &HFFFFFF
lblPrice.ForeColor = &H80000011
txtISBN.SetFocus
End Sub
