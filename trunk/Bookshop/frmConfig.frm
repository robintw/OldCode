VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmConfig 
   Caption         =   "Configuration"
   ClientHeight    =   5745
   ClientLeft      =   13095
   ClientTop       =   1920
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   240
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3360
      TabIndex        =   21
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      Top             =   5160
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8705
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "&Database"
      TabPicture(0)   =   "frmConfig.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtPath"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBrowse"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Security"
      TabPicture(1)   =   "frmConfig.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtAdminPassword"
      Tab(1).Control(1)=   "txtSellerPassword"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label5"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Email"
      TabPicture(2)   =   "frmConfig.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtSubject"
      Tab(2).Control(1)=   "txtFromAddress"
      Tab(2).Control(2)=   "txtDomain"
      Tab(2).Control(3)=   "txtPort"
      Tab(2).Control(4)=   "txtServer"
      Tab(2).Control(5)=   "txtBody"
      Tab(2).Control(6)=   "Label3"
      Tab(2).Control(7)=   "Label2"
      Tab(2).Control(8)=   "Label1"
      Tab(2).Control(9)=   "lblKeyForBody"
      Tab(2).Control(10)=   "lblPort"
      Tab(2).Control(11)=   "lblServer"
      Tab(2).Control(12)=   "lblBody"
      Tab(2).ControlCount=   13
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtAdminPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73440
         TabIndex        =   9
         Text            =   "drowssap"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtSellerPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73440
         TabIndex        =   12
         Text            =   "password"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   -74160
         TabIndex        =   17
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtFromAddress 
         Height          =   285
         Left            =   -74160
         TabIndex        =   15
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtDomain 
         Height          =   285
         Left            =   -74160
         TabIndex        =   11
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   -71640
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   -74160
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtBody 
         Height          =   2445
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "&Path:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "&Admin Password:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "&Seller Password:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "&Subject:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "&From Address:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Do&main:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblKeyForBody 
         Caption         =   "Instructions go here"
         Height          =   2415
         Left            =   -71760
         TabIndex        =   20
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblPort 
         Caption         =   "Po&rt:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblServer 
         Caption         =   "Serve&r:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblBody 
         Caption         =   "&Body:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   2160
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
'Restrict dialog box to access databases
cmndlg.Filter = "Access Databases | *.mdb"
'Don't raise an error on cancel
cmndlg.CancelError = False
cmndlg.ShowOpen
txtPath.Text = cmndlg.FileName
End Sub

Private Sub cmdCancel_Click()
'Close and unload form
Unload Me
End Sub

Private Sub cmdOk_Click()
'Deletes current config file - all details held in memory so thats ok
Kill "config.ini"

Dim clsINI As clsInifile
Set clsINI = New clsInifile
'Write config details to appropriate sections of config file
With clsINI
    .Path = App.Path & "\config.ini"
       
    .Section = "Database"
    .Key = "Path"
    .Value = txtPath.Text
        
    .Section = "Email"
    .Key = "Server"
    .Value = txtServer.Text
    .Key = "Port"
    .Value = txtPort.Text
    .Key = "Domain"
    .Value = txtDomain.Text
    .Key = "FromAddress"
    .Value = txtFromAddress.Text
    .Key = "Subject"
    .Value = txtSubject.Text
    .Key = "Message"
    .Value = Replace(txtBody.Text, vbCrLf, "\n")
    
    .Section = "Security"
    .Key = "SellerPassword"
    .Value = txtSellerPassword.Text
    
    .Key = "AdminPassword"
    .Value = txtAdminPassword.Text
    
End With

'Reload the config from the file - to update udfConfig
LoadConfigFromFile ("config.ini")
Set clsINI = Nothing

'Close and unload form
Unload Me
End Sub

Private Sub Form_Load()
'Load config in case it has changed since program startup
LoadConfigFromFile ("config.ini")

'Add instructions
lblKeyForBody.Caption = "You may use the following fields in the body and subject of the email:" & vbCrLf & vbCrLf & "<FirstName>" & vbCrLf & "<LastName>" & vbCrLf & "<Alphacode>" & vbCrLf & "<Title>" & vbCrLf & "<Author>" & vbCrLf & "<Publisher>" & vbCrLf & "<DateOrdered>"


'Load all config data into form
txtPath.Text = udfConfig.DatabasePath

txtAdminPassword.Text = udfConfig.AdminPassword
txtSellerPassword.Text = udfConfig.SellerPassword

txtServer.Text = udfConfig.EmailServer
txtPort.Text = udfConfig.EmailServerPort
txtDomain.Text = udfConfig.EmailDomain
txtFromAddress.Text = udfConfig.EmailFromAddress
txtSubject.Text = udfConfig.EmailSubject
txtBody.Text = udfConfig.EmailMessage
End Sub

