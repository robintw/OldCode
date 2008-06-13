VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmAddBook 
   Caption         =   "Add Book"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPublisher 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   735
      Left            =   2880
      TabIndex        =   15
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CheckBox chkSecondHand 
      Caption         =   "Second &Hand"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtSellingPrice 
      Height          =   285
      Left            =   3240
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtCostPrice 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtISBN 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "C&onfirm"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock smtp 
      Left            =   2040
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   0
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblPublisher 
      Caption         =   "&Publisher:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblAuthor 
      Caption         =   "&Author:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Caption         =   "&Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "&Selling Price:"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "&Cost Price:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      Caption         =   "Add a book"
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
   Begin VB.Label lblISBN 
      Caption         =   "&ISBN:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmAddBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FoundInDb As Boolean
Dim WithEvents EmailSend As clsEmailSend
Attribute EmailSend.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
'Close and unload form
Unload Me
End Sub

Private Sub cmdConfirm_Click()
Dim cmdAddBook As ADODB.Command
Set cmdAddBook = New ADODB.Command
Dim MessageText As String
Dim SubjectText As String
cmdAddBook.ActiveConnection = cnnGlobalConnection

'In all validation checks Trim() is used so the user cannot enter just spaces

If Trim(txtISBN.Text) = "" Then
MsgBox "Error: Please enter an ISBN"
Exit Sub
End If

If Trim(txtTitle.Text) = "" Then
MsgBox "Error: Please enter a Title"
Exit Sub
End If

If Trim(txtAuthor.Text) = "" Then
MsgBox "Error: Please enter an Author"
Exit Sub
End If

If Trim(txtPublisher.Text) = "" Then
MsgBox "Error: Please enter a Publisher"
Exit Sub
End If

If Trim(txtSellingPrice.Text) = "" Then
    MsgBox "Error: Please make sure you have specified a selling price", vbCritical
    txtSellingPrice.SetFocus
    Exit Sub
End If

If IsNumeric(txtSellingPrice.Text) = False Then
    MsgBox "Error: Please ensure you have entered a numeric value for the selling price!"
    txtSellingPrice.SetFocus
    Exit Sub
End If

If IsNumeric(txtCostPrice.Text) = False And txtCostPrice.Text <> "" Then
    MsgBox "Error: Please ensure you have entered a numeric value for the cost price!"
    txtCostPrice.SetFocus
    Exit Sub
End If

If Trim(txtCostPrice.Text) = "" Then
    txtCostPrice.Text = vbNull
End If

    
If FoundInDb = True Then
        'Update details in CurrentBookDetails from text boxes - allows user editing
        CurrentBookDetails.Title = txtTitle.Text
        CurrentBookDetails.Author = txtAuthor.Text
        CurrentBookDetails.Publisher = txtPublisher.Text
    
        'Escapes any ' in titles
        CurrentBookDetails.Title = Replace(CurrentBookDetails.Title, "'", "''")
        CurrentBookDetails.Author = Replace(CurrentBookDetails.Author, "'", "''")
        CurrentBookDetails.Publisher = Replace(CurrentBookDetails.Publisher, "'", "''")
        
        cmdAddBook.CommandText = "INSERT INTO tblBookCopy (ISBN, CostPrice, SellingPrice) VALUES ('" & CurrentBookDetails.ISBN & "'," & txtCostPrice.Text & "," & txtSellingPrice.Text & ")"
        cmdAddBook.Execute
Else
        'Update details in CurrentBookDetails from text boxes - allows user editing
        CurrentBookDetails.Title = txtTitle.Text
        CurrentBookDetails.Author = txtAuthor.Text
        CurrentBookDetails.Publisher = txtPublisher.Text
        
        'Escapes any ' in titles
        CurrentBookDetails.Title = Replace(CurrentBookDetails.Title, "'", "''")
        CurrentBookDetails.Author = Replace(CurrentBookDetails.Author, "'", "''")
        CurrentBookDetails.Publisher = Replace(CurrentBookDetails.Publisher, "'", "''")
        
        'Adds details to Book table
        cmdAddBook.CommandText = "INSERT INTO tblBook (ISBN, Title, Author, Publisher) VALUES ('" & CurrentBookDetails.ISBN & "', '" & CurrentBookDetails.Title & "', '" & CurrentBookDetails.Author & "', '" & CurrentBookDetails.Publisher & "')"
        cmdAddBook.Execute
        
        'Adds details to BookCopy table
        cmdAddBook.CommandText = "INSERT INTO tblBookCopy (ISBN, CostPrice, SellingPrice) VALUES ('" & CurrentBookDetails.ISBN & "'," & Val(txtCostPrice.Text) & "," & Val(txtSellingPrice.Text) & ")"
        cmdAddBook.Execute
    End If
    
    'Emails students or displays message for staff
    Dim rsOrder As ADODB.Recordset
    Set rsOrder = New ADODB.Recordset
   
    rsOrder.Open "SELECT tblOrder.OrderID, tblOrder.Title, tblOrder.Author, tblOrder.Publisher, tblOrder.DateOrdered, tblOrder.Alphacode, tblStudent.Alphacode, tblStudent.FirstName, tblStudent.LastName FROM tblOrder, tblStudent WHERE ISBN='" & CurrentBookDetails.ISBN & "' AND tblOrder.Alphacode = tblStudent.Alphacode ORDER BY DateOrdered", cnnGlobalConnection, adOpenStatic
    
    'If it has been ordered
    If rsOrder.RecordCount > 0 Then
        
        If Left(rsOrder.Fields("tblOrder.Alphacode").Value, 2) = "ZZ" Then
        'Don't bother emailing staff - they don't check their email regularly
            MsgBox "This book has been ordered by a member of staff. Please notify them that their book has arrived."
        Else
        
        MsgBox "This book has been ordered by " & rsOrder!FirstName & " " & rsOrder!LastName & vbCrLf & "When you press OK, he/she will be emailed.", vbInformation
        
        Set EmailSend = New clsEmailSend
        'Send an email using the details stored in the config file

        With EmailSend
          
          .Host = udfConfig.EmailServer
          
          .Sender = udfConfig.EmailFromAddress
          
          'Change <FirstName> etc to the actual details
          SubjectText = Replace(udfConfig.EmailSubject, "<FirstName>", rsOrder!FirstName)
          SubjectText = Replace(SubjectText, "<LastName>", rsOrder!LastName)
          SubjectText = Replace(SubjectText, "<Alphacode>", rsOrder.Fields("tblStudent.Alphacode").Value)
          SubjectText = Replace(SubjectText, "<Title>", CurrentBookDetails.Title)
          SubjectText = Replace(SubjectText, "<Author>", CurrentBookDetails.Author)
          SubjectText = Replace(SubjectText, "<Publisher>", CurrentBookDetails.Publisher)
          
          .Subject = SubjectText
    
          .Port = udfConfig.EmailServerPort
          
          'Change <FirstName> etc to the actual details
          MessageText = Replace(udfConfig.EmailMessage, "<FirstName>", rsOrder!FirstName)
          MessageText = Replace(MessageText, "<LastName>", rsOrder!LastName)
          MessageText = Replace(MessageText, "<Alphacode>", rsOrder.Fields("tblStudent.Alphacode").Value)
          MessageText = Replace(MessageText, "<Title>", CurrentBookDetails.Title)
          MessageText = Replace(MessageText, "<Author>", CurrentBookDetails.Author)
          MessageText = Replace(MessageText, "<Publisher>", CurrentBookDetails.Publisher)

          .Message = MessageText
          
          'Possibly need <>'s around email address
          'Names work with Name <email.address>
          .FromName = udfConfig.EmailFromAddress
          
          
          Set .WinsockControl = smtp
          .Recipient = rsOrder.Fields("tblStudent.Alphacode").Value & udfConfig.EmailDomain
          
          
          'Names work with Name <email.address>
          .ToName = rsOrder.Fields("tblStudent.Alphacode").Value & udfConfig.EmailDomain
          
          .Form = Me
          
          .SendMail
          
        End With
        End If
        'Remove order from list as it has been delt with now
        cmdAddBook.CommandText = "DELETE FROM tblOrder WHERE OrderID=" & rsOrder!OrderID
        cmdAddBook.Execute
                
    End If
Set cmdAddBook = Nothing
ResetForm

End Sub

Private Sub ResetForm()
'Set all fields to "", set all of CurrentBookDetails to Null and set focus to txtISBN
'Make it as if we've just loaded the form again
txtISBN.Text = ""
txtTitle.Text = ""
txtAuthor.Text = ""
txtPublisher.Text = ""
txtCostPrice.Text = ""
txtSellingPrice.Text = ""

CurrentBookDetails.Author = vbNull
CurrentBookDetails.Title = vbNull
CurrentBookDetails.Publisher = vbNull
CurrentBookDetails.ISBN = vbNull

txtISBN.SetFocus
End Sub

Private Sub txtAuthor_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtCostPrice_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
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

    Dim rsLookup As ADODB.Recordset
    Set rsLookup = New ADODB.Recordset

    CurrentBookDetails.ISBN = txtISBN.Text
    
    rsLookup.Open "SELECT Title, Author, Publisher FROM tblBook WHERE ISBN='" & _
        CurrentBookDetails.ISBN & "'", cnnGlobalConnection, adOpenStatic
    
    If rsLookup.RecordCount = 1 Then
        'If it is found in tblBook then take the details from there - saves looking it up
        'on the internet
        CurrentBookDetails.Title = rsLookup!Title
        CurrentBookDetails.Author = rsLookup!Author
        CurrentBookDetails.Publisher = rsLookup!Publisher
        
        txtTitle.Text = CurrentBookDetails.Title
        txtAuthor.Text = CurrentBookDetails.Author
        txtPublisher.Text = CurrentBookDetails.Publisher
        
        'So we know later on not to add a record to tblBook
        FoundInDb = True
        
        txtCostPrice.SetFocus
    Else
        'If it's not found in tblBook then get the details from Amazon's website
        CurrentBookDetails = GetDetailsFromInternet(CurrentBookDetails.ISBN)
        
        'If the book isn't found the title will be set to "Error"
        If CurrentBookDetails.Title <> "Error" Then
            txtTitle.Text = CurrentBookDetails.Title
            txtAuthor.Text = CurrentBookDetails.Author
            txtPublisher.Text = CurrentBookDetails.Publisher
            txtCostPrice.SetFocus
        Else
            'There is an error so let the user manually input the title
            txtISBN.Text = CurrentBookDetails.ISBN
            txtTitle.SetFocus
        End If
        
    End If

End If
Set rsLookup = Nothing
End Sub

Private Sub txtPublisher_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtSellingPrice_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub
