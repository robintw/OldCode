VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEndOf 
   Caption         =   "End of Year activities"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "End of Year"
      Height          =   5415
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   3375
      Begin MSComDlg.CommonDialog cmndlg 
         Left            =   120
         Top             =   4800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Import CSV"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "NB: The update process may take some time."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label lblYearInstructions 
         Caption         =   "Instructions go here!"
         Height          =   3015
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "End of Term"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdPrintStudentReports 
         Caption         =   "&Print Student Reports"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear Spent"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label lblTermInstructions 
         Caption         =   "Instructions go here!"
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEndOf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
Dim cmdClear As ADODB.Command
Set cmdClear = New ADODB.Command
cmdClear.ActiveConnection = cnnGlobalConnection
'Clear all spent details in tblStudent
cmdClear.CommandText = "UPDATE tblStudent SET Spent=0"
cmdClear.Execute

Set cmdClear = Nothing
End Sub

Private Sub cmdImport_Click()
'Only show csv files in dialog
cmndlg.Filter = "Comma Separated Values file (*.csv) | *.csv"
'Don't raise an error on cancel
cmndlg.CancelError = False
cmndlg.ShowOpen
Screen.MousePointer = vbHourglass
UpdateFromCSV (cmndlg.FileName)
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrintStudentReports_Click()
StudentReportAllShow
End Sub

Private Sub Form_Load()
'Add instructions
lblYearInstructions.Caption = "At the end of each year all new students need adding, and all current students need to be updated with regard to year and house." & vbCrLf & "To do this a csv file export from CMIS of all students is required, with the following fields, in order:" & vbCrLf & "Alphacode, FirstName, LastName, SchoolYear, House." & vbCrLf & vbCrLf & "Once this file has been exported, click the button below and browse to select the file. The update process will then start automatically."
lblTermInstructions.Caption = "At the end of each term the spent amounts for each student need reseting, as their allowance is termly and not able to be carried forward" & vbCrLf & "To do this click the button below, which will clear the spent information for all students." & vbCrLf & vbCrLf & "Normally at the end of each term reports are printed listing every book that a girl has bought. To do this click the button below."
End Sub

