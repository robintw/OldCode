VERSION 5.00
Begin VB.Form frmMenuAdmin 
   Caption         =   "Bookshop Management - Admin Menu"
   ClientHeight    =   10695
   ClientLeft      =   6225
   ClientTop       =   1575
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLOL 
      Caption         =   "Edit &Book Details"
      Height          =   855
      Index           =   0
      Left            =   3240
      TabIndex        =   12
      Top             =   8760
      Width           =   2415
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Report Menu"
      Height          =   855
      Left            =   360
      TabIndex        =   9
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "&Sell a book"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "&Order a book"
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton cmdEndOf 
      Caption         =   "End of term and &year activities"
      Height          =   855
      Left            =   360
      TabIndex        =   10
      Top             =   8760
      Width           =   2415
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "Co&nfiguration"
      Height          =   855
      Left            =   3240
      TabIndex        =   11
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton cmdAddEditStudents 
      Caption         =   "Add/Edit S&tudents"
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CommandButton cmdTestCSV 
      Caption         =   "&Export to bursary (as CSV)"
      Height          =   855
      Left            =   3240
      TabIndex        =   8
      Top             =   7800
      Width           =   2415
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   1080
      Picture         =   "frmMenuAdmin.frx":0000
      ScaleHeight     =   4215
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label lblAbout 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Click for 'About' information"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   3960
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdStockCheck 
      Caption         =   "S&tock Check"
      Height          =   855
      Left            =   3240
      TabIndex        =   6
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CommandButton cmdCSV 
      Caption         =   "&Import students from CSV"
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdAddBook 
      Caption         =   "&Add Book"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add A&llowances"
      Height          =   855
      Left            =   3240
      TabIndex        =   4
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5880
      Y1              =   5640
      Y2              =   5640
   End
End
Attribute VB_Name = "frmMenuAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddBook_Click()
frmAddBook.Show
End Sub

Private Sub cmdAddEditStudents_Click()
frmNEWStudentEdit.Show
End Sub

Private Sub cmdConfig_Click()
frmConfig.Show
End Sub

Private Sub cmdCSV_Click()
frmStudentImport.Show
End Sub

Private Sub cmdEndOf_Click()
frmEndOf.Show
End Sub

Private Sub cmdLOL_Click(Index As Integer)
frmBookEdit.Show
End Sub

Private Sub cmdReport_Click()
frmReportMenu.Show
End Sub

Private Sub cmdROFL_Click()
frmNEWStudentEdit.Show
End Sub

Private Sub cmdStockCheck_Click()
frmStockCheck.Show
End Sub


Private Sub cmdTestCSV_Click()
frmExport.Show
End Sub


Private Sub Command2_Click()
frmAddAllowance.Show
End Sub

Private Sub cmdOrder_Click()
frmAddOrder.Show
End Sub

Private Sub cmdSell_Click()
frmSell.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If S is pressed then do the same as if cmdSell was clicked
If KeyAscii = vbKeyS Then
    cmdSell_Click
End If

'If O is pressed then do the same as if cmdOrder was clicked
If KeyAscii = vbKeyO Then
    cmdOrder_Click
End If
End Sub

Private Sub Form_Load()
'Create the connection, and load the config info
If CheckConfigFileExists Then
    LoadConfigFromFile ("config.ini")
Else
    MsgBox "Cannot find config.ini file please create and then restart.", vbCritical
    Unload Me
    End
End If

CreateGlobalConnection
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'Closes the global connection
cnnGlobalConnection.Close
'Closes all open files
Reset
End Sub

Private Sub lblAbout_Click()
picLogo_Click
End Sub

Private Sub picLogo_Click()
frmAbout.Show
End Sub
