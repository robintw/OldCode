VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmStudentImport 
   Caption         =   "Import Student Data from CSV"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddToDB 
      Caption         =   "&Confirm and add to database"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   8655
   End
   Begin MSFlexGridLib.MSFlexGrid flxImport 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   8280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblPath 
      Caption         =   "&Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmStudentImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileName

Private Sub cmdAddToDB_Click()
'Set the pointer to an hourglass to let the user know the operation is being executed and
'might take some time
Screen.MousePointer = vbHourglass

'Import from the CSV file specified
ImportFromCSV txtPath.Text, flxImport, False

'Set the pointer back to normal
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBrowse_Click()
'Only show CSV files in the dialog box
cmndlg.Filter = "Comma Separated Values file (*.csv) | *.csv"
'Don't raise an error if the user presses cancel
cmndlg.CancelError = False
cmndlg.ShowOpen
FileName = cmndlg.FileName

txtPath.Text = FileName
flxImport.Rows = 1

'Import, simulating into the flexgrid
ImportFromCSV txtPath.Text, flxImport, True

End Sub

Private Sub Form_Load()
'Setup the flex grid titles
flxImport.TextMatrix(0, 0) = "Alphacode"
flxImport.TextMatrix(0, 1) = "First Name"
flxImport.TextMatrix(0, 2) = "Last Name"
flxImport.TextMatrix(0, 3) = "Year"
flxImport.TextMatrix(0, 4) = "House"
End Sub
