VERSION 5.00
Begin VB.Form frmMenuSeller 
   Caption         =   "Bookshop Management - Seller Menu"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOrder 
      Caption         =   "&Order a book"
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "&Sell a book"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   2415
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   840
      Picture         =   "frmMenuSeller.frx":0000
      ScaleHeight     =   3975
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmMenuSeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOrder_Click()
frmAddOrder.Show
End Sub

Private Sub cmdSell_Click()
frmSell.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If S is pressed then do the same as clicking the cmdSell button
If KeyAscii = 115 Then
    cmdSell_Click
End If
'If O is pressed then do the same as clicking the cmdOrder button
If KeyAscii = 111 Then
    cmdOrder_Click
End If
End Sub

Private Sub Form_Load()
If CheckConfigFileExists Then
    LoadConfigFromFile ("config.ini")
Else
    MsgBox "Cannot find config.ini file. Please contact Bookshop Manager.", vbCritical
    Unload Me
    End
End If

CreateGlobalConnection
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Closes the global connection
cnnGlobalConnection.Close
'Closes all open files
Reset
End Sub
