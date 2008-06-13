VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MGC Bookshop Management"
   ClientHeight    =   3585
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6510
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2474.431
   ScaleMode       =   0  'User
   ScaleWidth      =   6113.227
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStats 
      Caption         =   "&DB Stats"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   2640
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1065
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   747.985
      ScaleMode       =   0  'User
      ScaleWidth      =   958.685
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5160
      TabIndex        =   7
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Label lblHelp 
      Caption         =   "For help contact Robin Wilson: r.t.wilson@rmplc.co.uk or via Martyn Wilson in the IT Dept."
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   6085.056
      Y1              =   1573.697
      Y2              =   1573.697
   End
   Begin VB.Label lblDescription 
      Caption         =   "Bespoke software developed for Malvern Girls' College to manage the college bookshop."
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1650
      TabIndex        =   3
      Top             =   1005
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1650
      TabIndex        =   1
      Top             =   120
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   6085.056
      Y1              =   1490.87
      Y2              =   1490.87
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1650
      TabIndex        =   2
      Top             =   660
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: This program is protected under international copyright law."
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   4815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub cmdStats_Click()
Dim rsStats As ADODB.Recordset
Set rsStats = New ADODB.Recordset

Dim NumOfBooks As Integer
Dim SizeOfDB As Long

'Selects all the books that haven't been sold
rsStats.Open "SELECT * FROM tblBookCopy WHERE Alphacode IS NULL", cnnGlobalConnection, adOpenStatic

NumOfBooks = rsStats.RecordCount
SizeOfDB = FileLen(udfConfig.DatabasePath)
'FileLen gives the result in Bytes, I want it in Kb
SizeOfDB = SizeOfDB / 1024

MsgBox "Number of books in database: " & NumOfBooks & vbCrLf & "Size of bookshop.mdb file: " & SizeOfDB & "Kb", vbInformation Or vbOKOnly, "Database Statistics"

Set rsStats = Nothing
End Sub

Private Sub Form_Load()
    'Set Labels with version and other information
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = "MGC Bookshop Management Application"
    lblDisclaimer.Caption = "Warning: This program is protected under international copyright law. © Robin Wilson 2006" & vbCrLf & vbCrLf & "This product include software developed by vbAccelerator (http://vbaccelerator.com/)."
End Sub

