VERSION 5.00
Begin VB.Form frmAddOrder 
   Caption         =   "Order a book"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAlphacode 
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "C&onfirm"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtISBN 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   735
      Left            =   3120
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtPublisher 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      Caption         =   "Order a book"
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
   Begin VB.Label Label1 
      Caption         =   "Please enter ISBN, Alphacode and at least one other field."
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblAlphacodeDet 
      Caption         =   "&Alphacode:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblISBN 
      Caption         =   "&ISBN:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
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
   Begin VB.Label lblAuthor 
      Caption         =   "A&uthor:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblPublisher 
      Caption         =   "&Publisher:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
End
Attribute VB_Name = "frmAddOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
'Close and unload form
Unload Me
End Sub

Private Sub cmdConfirm_Click()
Dim cmdAdd As ADODB.Command
Set cmdAdd = New ADODB.Command
cmdAdd.ActiveConnection = cnnGlobalConnection

'Check they have entered an ISBN
If Trim(txtISBN.Text) = "" Then
    MsgBox "Error: Please enter ISBN", vbCritical
    Exit Sub
End If

'Check it is valid
If ValidateISBN(txtISBN.Text) = False Then
    MsgBox "Please enter a valid ISBN"
    Exit Sub
End If

'Add details to tblOrder
cmdAdd.CommandText = "INSERT INTO tblOrder (ISBN, Title, Author, Publisher, Alphacode, DateOrdered) VALUES ('" & txtISBN.Text & "','" & txtTitle.Text & "','" & txtAuthor.Text & "','" & txtPublisher.Text & "','" & txtAlphacode.Text & "','" & Date & "')"

cmdAdd.Execute

Set cmdAdd = Nothing
'Close form
Unload Me
End Sub

Private Sub txtAlphacode_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub


Private Sub txtAuthor_KeyPress(KeyAscii As Integer)
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
End If

EnterDoesTheSameAsTab (KeyAscii)

End Sub

Private Sub txtPublisher_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub
