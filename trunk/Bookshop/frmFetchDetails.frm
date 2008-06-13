VERSION 5.00
Begin VB.Form frmFetchDetails 
   Caption         =   "Processing . . ."
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   1095
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdMissing 
      Caption         =   "Books &in the database but not scanned"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdNotFound 
      Caption         =   "Books scanned but &not found in the database"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblFinished 
      Caption         =   "finished"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCaption 
      Caption         =   "Please wait while the details are fetched for all books that could not be found..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmFetchDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFinish_Click()
'Close and unload form
Unload Me
End Sub

Private Sub cmdMissing_Click()
frmListMissing.Show
End Sub

Private Sub cmdNotFound_Click()
frmListNotFound.Show
End Sub

Private Sub Form_Activate()
'Set pointer to hourglass - gives the user visual feedback that the system is going something
Screen.MousePointer = vbHourglass

'Allow the form to complete loading - stops any redraw problems
DoEvents

Dim ReturnedDetails As udfBookDetails

Dim rsNotFound As ADODB.Recordset
Set rsNotFound = New ADODB.Recordset
Dim cmdUpdate As ADODB.Command
Set cmdUpdate = New ADODB.Command
cmdUpdate.ActiveConnection = cnnGlobalConnection


rsNotFound.Open "SELECT * FROM tblNotFound", cnnGlobalConnection, adOpenStatic

'For all books in tblNotFound
Do Until rsNotFound.EOF
    'Find details from the internet
    ReturnedDetails = GetDetailsFromInternet(rsNotFound!ISBN)
    ReturnedDetails.Title = Replace(ReturnedDetails.Title, "'", "''")
    ReturnedDetails.Author = Replace(ReturnedDetails.Author, "'", "''")
    ReturnedDetails.Publisher = Replace(ReturnedDetails.Publisher, "'", "''")
    'And update the table with those details
    cmdUpdate.CommandText = "UPDATE tblNotFound SET Title='" & ReturnedDetails.Title & "', Author='" & ReturnedDetails.Author & "', Publisher='" & ReturnedDetails.Publisher & "' WHERE ISBN='" & rsNotFound!ISBN & "'"
    cmdUpdate.Execute
    rsNotFound.MoveNext
Loop
'Change mouse pointer back to normal
Screen.MousePointer = vbDefault

'Allow user to click buttons - and display 'finished' message
lblFinished.Visible = True
cmdNotFound.Enabled = True
cmdMissing.Enabled = True

Set rsNotFound = Nothing
Set cmdUpdate = Nothing
End Sub

