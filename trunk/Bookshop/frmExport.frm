VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmExport 
   Caption         =   "Export to bursary"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock smtp 
      Left            =   1560
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtEmailAddress 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.TextBox txtMessage 
      Height          =   2445
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "To &Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblBody 
      Caption         =   "&Body:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblKeyForBody 
      Caption         =   "Instructions go here"
      Height          =   2415
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "S&ubject:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents EmailSend As clsEmailSend
Attribute EmailSend.VB_VarHelpID = -1
Dim FileName As String
Dim FilePath As String

Private Sub cmdSend_Click()

'Check all fields are filled in (with Trim())
If Trim(txtSubject.Text) = "" Then
    MsgBox "Error: Please enter a subject"
    Exit Sub
End If

If Trim(txtEmailAddress.Text) = "" Then
    MsgBox "Error: Please enter an email address"
    Exit Sub
End If

If Trim(txtMessage.Text) = "" Then
    MsgBox "Error: Please enter a message body"
    Exit Sub
End If

Set EmailSend = New clsEmailSend
'Send email to bursary with details specified in config.ini and on the form
        With EmailSend
          
          .Host = udfConfig.EmailServer
          .Sender = udfConfig.EmailFromAddress
        
          .Subject = txtSubject.Text
    
          .Port = udfConfig.EmailServerPort
          
          .Message = txtMessage.Text
          
          'Possibly need <>'s around email address
          'Names work with Name <email.address>
          .FromName = udfConfig.EmailFromAddress
          
          
          Set .WinsockControl = smtp
          
          .Recipient = txtEmailAddress.Text
          
          
          'Names work with Name <email.address>
          .ToName = txtEmailAddress.Text
          
          .Form = Me
            'Attach the file we created when the form was loaded
          .AttachmentPath = FilePath
          
          .SendMail

            

        End With
End Sub

Private Sub Form_Load()
On Error Resume Next
'Create a filename from the current term and date
FileName = "Bursary-" & CurrentTerm & "-" & Replace(Date, "/", "-") & ".csv"
FilePath = ".\BursaryData\" & FileName
'Export the data to the path we've just created
ExportToCSV (FilePath)
'Show instructions
lblKeyForBody.Caption = "The data has been exported and stored in the BursaryData subfolder. Fill in the details here to send an email to the bursary with the file automatically attached."
End Sub
