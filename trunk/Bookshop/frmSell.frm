VERSION 5.00
Begin VB.Form frmSell 
   Caption         =   "Sell a book"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSellPrice 
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtISBN 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtTitle 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtAuthor 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtPublisher 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtSellPrice 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtAlphacode 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtLastName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   20
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtFirstName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtHouse 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   14
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtYear 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   21
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtAllowance 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   22
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtSpent 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   25
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3000
      TabIndex        =   27
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&Confirm"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   1695
      Left            =   0
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label lblISBN 
      Caption         =   "&ISBN:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      Caption         =   "Sell a book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label lblTitle 
      Caption         =   "&Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblAuthor 
      Caption         =   "&Author:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblPublisher 
      Caption         =   "&Publisher:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblSellPrice 
      Caption         =   "&Sell Price:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4320
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   0
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblAlphacodeDet 
      Caption         =   "Al&phacode:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblLastNameDet 
      Caption         =   "&Last Name:"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblFirstNameDet 
      Caption         =   "&First Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblHouse 
      Caption         =   "&House:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblYear 
      Caption         =   "&Year:"
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label lblAllowance 
      Caption         =   "A&llowance:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblSpent 
      Caption         =   "Sp&ent:"
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   4680
      Width           =   975
   End
End
Attribute VB_Name = "frmSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSellPrice_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub cmdCancel_Click()
'Close and unload form
Unload Me
End Sub

Private Sub cmdConfirm_Click()
Dim StudentAlphacode As String
Dim BookISBN As String
Dim SellingPrice As Single
Dim NewSpentValue As Single

'If either the name is blank or the ISBN is blank then give an error
If Trim(txtFirstName.Text) = "" Or Trim(txtISBN.Text) = "" Then
    MsgBox "Error: You must select both a book to sell and a person to sell it to.", vbCritical
    txtAlphacode.SetFocus
    Exit Sub
End If

'If the allowance is blank then give an error
If Trim(txtAllowance.Text) = "" Then
    MsgBox "This student has no allowance value set. Please set one before trying to sell them a book!", vbCritical
    Exit Sub
End If

StudentAlphacode = txtAlphacode.Text
BookISBN = txtISBN.Text

'Get the correct selling price from either the text box or combo box
If txtSellPrice.Text = "" Then
    SellingPrice = cmbSellPrice.Text
Else
    SellingPrice = txtSellPrice.Text
End If

'Calculate the new spent value once this book has been bought
NewSpentValue = txtSpent.Text + SellingPrice

'If the new spent value will go over their allowance then give a message but allow the user to continue
If NewSpentValue > txtAllowance.Text Then
    MsgboxAnswer = MsgBox("This will go over the students allowance. Do you want to continue?", vbYesNo Or vbQuestion)
    If MsgboxAnswer = vbNo Then
    'If they say no then unload the form and stop the sale
    Unload Me
    Exit Sub
    End If
End If
  

Dim rsSell As ADODB.Recordset
Set rsSell = New ADODB.Recordset

rsSell.ActiveConnection = cnnGlobalConnection

rsSell.Open "SELECT CopyID, ISBN, SellingPrice FROM tblBookCopy WHERE ISBN='" & BookISBN & "' AND SellingPrice=" & SellingPrice

Dim cmdSell As ADODB.Command
Set cmdSell = New ADODB.Command

'Update tblBookCopy with the appropriate data
cmdSell.ActiveConnection = cnnGlobalConnection
cmdSell.CommandText = "UPDATE tblBookCopy SET tblBookCopy.Alphacode='" & _
    StudentAlphacode & "', tblBookCopy.Date='" & Date & "' WHERE tblBookCopy.ISBN='" & BookISBN _
    & "' AND tblBookCopy.SellingPrice=" & SellingPrice & " AND tblBookCopy.CopyID=" & rsSell!CopyID

cmdSell.Execute

'Update the spent data as well
cmdSell.CommandText = "UPDATE tblStudent SET tblStudent.Spent=" & NewSpentValue & " WHERE tblStudent.Alphacode='" & StudentAlphacode & "'"

cmdSell.Execute

Set cmdSell = Nothing

'Close and unload form
Unload Me
End Sub



Private Sub txtAlphacode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    
    'All Alphacodes should be in upper case
    txtAlphacode.Text = UCase(txtAlphacode.Text)

    Dim rsStudents As ADODB.Recordset
    Set rsStudents = New ADODB.Recordset

    rsStudents.Open "SELECT * from tblStudent WHERE Alphacode='" & txtAlphacode.Text & "' AND Old=False", cnnGlobalConnection, adOpenStatic
    
    If rsStudents.RecordCount <> 0 Then
        'Show the details of the student with that alphacode
        txtFirstName.Text = rsStudents!FirstName
        txtLastName.Text = rsStudents!LastName
        txtYear.Text = rsStudents!SchoolYear
        txtHouse.Text = rsStudents!House
        txtAllowance.Text = FormatCurrency(rsStudents!Allowance)
        txtSpent.Text = FormatCurrency(rsStudents!Spent)
        cmdConfirm.SetFocus
    Else
        'If the student can't be found then give an error
        MsgBox "Error: Student does not exist", vbCritical
        txtAlphacode.SetFocus
    End If
    
End If

Set rsStudents = Nothing
End Sub

Private Sub txtISBN_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Dim rsBooks As ADODB.Recordset
Set rsBooks = New ADODB.Recordset

    Dim FirstPrice As Single
    Dim PricesDifferent As Boolean
    
    
Select Case Len(txtISBN.Text)
    Case 13
        'If its 13 Chars long then it is an EAN - so convert to ISBN
        txtISBN.Text = ConvertEANToISBN(txtISBN.Text)
    Case 10
        'If its 10 then it is an ISBN - its fine but validate it
        If ValidateISBN(txtISBN.Text) = False Then
            MsgBox "Error: Invalid ISBN"
            Exit Sub
        End If
    Case Else
        'Otherwise we've no idea what it is!
        MsgBox "Error: Please enter an ISBN or EAN"
        txtISBN.Text = ""
        txtISBN.SetFocus
        
        Exit Sub
End Select
    
    rsBooks.Open _
        "SELECT tblBookCopy.CostPrice, tblBookCopy.SellingPrice, tblBook.Title, tblBook.Author, tblBook.Publisher from tblBookCopy, tblBook WHERE tblBookCopy.ISBN='" _
        & txtISBN.Text & "' AND tblBook.ISBN = tblBookCopy.ISBN AND tblBookCopy.Alphacode IS NULL ORDER BY SellingPrice", _
        cnnGlobalConnection, adOpenStatic
    
    'If no records are returned then give an error
    If rsBooks.RecordCount = 0 Then
        MsgBox "Error: Book not found", vbCritical
        Exit Sub
    End If
        
        
    FirstPrice = rsBooks!SellingPrice
    PricesDifferent = False
    
    'Step through all records returned, checking each price to see if there are
    'any different prices. If so, set PricesDifferent to True.
    Do Until rsBooks.EOF
        If rsBooks!SellingPrice = FirstPrice Then
            rsBooks.MoveNext
        Else
            PricesDifferent = True
            Exit Do
        End If
    Loop
    'Move back to the beginning so that the first price in the combo box is the lowest
    'price which is the one the student is likely to want!
    rsBooks.MoveFirst
    
    If PricesDifferent = True Then
    'If the prices are different then highlight the combo box and set the focus to it
        cmbSellPrice.Visible = True
        cmbSellPrice.BackColor = &HC0C0FF
        cmbSellPrice.SetFocus
        'Make sure the text box isn't visible underneath the combo box
        txtSellPrice.Visible = False
        'Show the book details on the form
        txtTitle.Text = rsBooks!Title
        txtAuthor.Text = rsBooks!Author
        txtPublisher.Text = rsBooks!Publisher
                
        Do Until rsBooks.EOF
            'Add each price to the combo box
            cmbSellPrice.AddItem (FormatCurrency(rsBooks!SellingPrice))
            rsBooks.MoveNext
        Loop
    Else
    'If the prices aren't different then show the text box with the price in it
        cmbSellPrice.Visible = False
        txtSellPrice.Visible = True
        txtAlphacode.SetFocus
        'Show the book details on the form
        txtTitle.Text = rsBooks!Title
        txtAuthor.Text = rsBooks!Author
        txtPublisher.Text = rsBooks!Publisher
        txtSellPrice.Text = FormatCurrency(rsBooks!SellingPrice)
    End If

End If

Set rsBooks = Nothing
End Sub
