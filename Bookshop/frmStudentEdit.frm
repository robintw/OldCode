VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmStudentEdit 
   Caption         =   "Add/Edit Students"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Del"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtAllowance 
      DataField       =   "Allowance"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """£""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   2
      EndProperty
      DataSource      =   "adodcStudent"
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtHouse 
      DataField       =   "House"
      DataSource      =   "adodcStudent"
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtYear 
      DataField       =   "SchoolYear"
      DataSource      =   "adodcStudent"
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LastName"
      DataSource      =   "adodcStudent"
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "FirstName"
      DataSource      =   "adodcStudent"
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtAlphacode 
      DataField       =   "Alphacode"
      DataSource      =   "adodcStudent"
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc adodcStudent 
      Height          =   375
      Left            =   840
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   2
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=bookshop.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=bookshop.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblStudent WHERE Old=FALSE ORDER BY LastName"
      Caption         =   "Students"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      Caption         =   "Add/Edit Students"
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
   Begin VB.Label Label5 
      Caption         =   "Allowance:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "House:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Year:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblAlphacode 
      Caption         =   "Alphacode:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmStudentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
'Add a new record
adodcStudent.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
Dim cmdDelete As ADODB.Command
Set cmdDelete = New ADODB.Command
cmdDelete.ActiveConnection = cnnGlobalConnection

'Delete the student from tblStudent
cmdDelete.CommandText = "DELETE FROM tblStudent WHERE Alphacode='" & adodcStudent.Recordset.Fields("Alphacode") & "'"
cmdDelete.Execute

If MsgBox("Do you want to remove the details for all the books this student bought?", vbYesNo Or vbQuestion) = vbYes Then
    'Delete all the rceords from tblBookCopy that this student is mentioned in
    'ie. all the books the student bought
    cmdDelete.CommandText = "DELETE FROM tblBookCopy WHERE Alphacode='" & adodcStudent.Recordset.Fields("Alphacode") & "'"
    cmdDelete.Execute
    Set cmdDelete = Nothing
End If
adodcStudent.Recordset.Delete
End Sub

Private Sub txtAllowance_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtAlphacode_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtHouse_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
EnterDoesTheSameAsTab (KeyAscii)
End Sub

