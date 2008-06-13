VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Bookshop Management Login"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   360
      Picture         =   "frmLogin.frx":0000
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Frame famUser 
      Caption         =   "User"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   2295
      Begin VB.OptionButton optAdmin 
         Caption         =   "&Admin"
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSeller 
         Caption         =   "&Seller"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
'Close and unload form
Unload Me
End Sub

Private Sub cmdLogin_Click()
Dim SellerPassword As String
Dim AdminPassword As String

Dim clsPasswordINI As clsInifile
Set clsPasswordINI = New clsInifile

'Load passwords from ini file. Cannot load them from udfConfig as whole config file has
'not been read yet
With clsPasswordINI
    .Path = App.Path & "\config.ini"
    
    .Section = "Security"
    
    .Key = "SellerPassword"
    SellerPassword = .Value
    
    .Key = "AdminPassword"
    AdminPassword = .Value
End With

If optSeller.Value = True Then
    'If Seller is selected then check for correct seller password
    If txtPassword.Text = SellerPassword Then
        frmMenuSeller.Show
        Unload Me
    Else
        'If the password is incorrect then display an error message
        MsgBox "Incorrect password. Please try again", vbCritical
        'Set txtPassword to blank, and put the focus back to it ready for a 2nd try
        txtPassword.Text = ""
        txtPassword.SetFocus
    End If
Else
    If txtPassword.Text = AdminPassword Then
        'If Admin is selected then check for correct admin password
        frmMenuAdmin.Show
        Unload Me
    Else
        'If the password is incorrect then display an error message
        MsgBox "Incorrect password. Please try again", vbCritical
        'Set txtPassword to blank, and put the focus back to it ready for a 2nd try
        txtPassword.Text = ""
        txtPassword.SetFocus
    End If
End If


Set clsPasswordINI = Nothing
End Sub

Private Sub Form_Activate()
'Set the focus to the option buttons so it is easy to select the user
optSeller.SetFocus
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'If Ctrl-A is pressed then select the optAdmin
If KeyCode = 65 And Shift = 2 Then
    optAdmin.Value = True
    txtPassword.SetFocus
End If

'If Ctrl-S is pressed then select the optSeller
If KeyCode = 83 And Shift = 2 Then
    optSeller.Value = True
    txtPassword.SetFocus
End If
End Sub

Private Sub optAdmin_KeyPress(KeyAscii As Integer)
'If return is pressed then jump to the password box
If KeyAscii = vbKeyReturn Then
    txtPassword.Text = ""
    txtPassword.SetFocus
End If
End Sub

Private Sub optSeller_KeyPress(KeyAscii As Integer)
'If return is pressed then jump to the password box
If KeyAscii = vbKeyReturn Then
    txtPassword.Text = ""
    txtPassword.SetFocus
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
'If return is pressed then run cmdLogin_Click
If KeyAscii = vbKeyReturn Then
    cmdLogin_Click
End If
End Sub
