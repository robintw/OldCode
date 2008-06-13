VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmailSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author: Alan Pucill - altered from open source code on the internet (arp@hereford.ac.uk)
'         Altered by Robin Wilson, Jan-Mar 06 for use in Computing A-Level Coursework
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Enum SMTP_State              'The smtp states enum
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

'Private sAdressData       As AdressData   'the Adress data buffer
Private smtpState         As SMTP_State   'the smtp states
Private EncodedFile       As String       'the encoded files buffer (attachments)
Dim frm                  As Form         'the variable for new forms
Private MessageBuffer(50) As String       'some Buffer variables
Private FromBuffer(50)    As String
Dim SubjectBuffer(50)    As String


Dim txtHost As String
Dim txtAttachment As String
Dim txtSender As String
Dim txtRecipient As String
Dim txtSubject As String
Dim txtMessage As String
Dim txtTo As String
Dim txtPort As String
Dim txtFrom As String
Dim WinsockError As String

Private WithEvents mSMTP As Winsock
Attribute mSMTP.VB_VarHelpID = -1
Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1

Public Event MessageSent()
Public Event SendError()

Public Property Let AttachmentPath(sAttachment As String)
txtAttachment = sAttachment
End Property
Public Property Let Port(sPort As String)
txtPort = sPort
End Property
Public Property Let Host(sHost As String)
txtHost = sHost
End Property
Public Property Let Sender(sSender As String)
txtSender = sSender
End Property
Public Property Let Recipient(sRecipient As String)
txtRecipient = sRecipient
End Property
Public Property Let Message(sMessage As String)
txtMessage = sMessage
End Property
Public Property Let ToName(sToName As String)
txtTo = sToName
End Property
Public Property Let FromName(sFromName As String)
txtFrom = sFromName
End Property
Public Property Let Subject(sSubject As String)
txtSubject = sSubject
End Property

Public Property Let Form(ctlForm As Form)
Set mForm = ctlForm
End Property

Public Property Set WinsockControl(ByRef ctlWinsock As Winsock)
Set mSMTP = ctlWinsock
End Property

Private Sub msmtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Tell the user that an error occured
    WinsockError = "Email Sending Error" & "Winsock Error number " & Number & vbCrLf & Description
    mSMTP.Close
    MsgBox WinsockError, vbCritical, "Winsock Error"
End Sub

Private Sub mSMTP_DataArrival(ByVal bytesTotal As Long)
    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    'Retrive data from winsock buffer
    mSMTP.GetData strServerResponse
       
    
    'Get server response code (first three symbols)
    strResponseCode = Left(strServerResponse, 3)
    'Only these three codes from the server tell us
    'that the command was accepted
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
        Select Case smtpState
            Case MAIL_CONNECT
                smtpState = MAIL_HELO
                'Remove blank spaces
                strDataToSend = Trim$(txtSender)
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                mSMTP.SendData "HELO " & strDataToSend & vbCrLf
            Case MAIL_HELO
                smtpState = MAIL_FROM
                'Send MAIL FROM command to the server
                'so it knows from who the message comes
                mSMTP.SendData "MAIL FROM:" & Trim$(txtSender) & vbCrLf
                
            Case MAIL_FROM
                smtpState = MAIL_RCPTTO
                'Send RCPT TO command to the server
                'so it knows where to send the message
                mSMTP.SendData "RCPT TO:" & Trim$(txtRecipient) & vbCrLf
                
            Case MAIL_RCPTTO
                smtpState = MAIL_DATA
                'Send DATA command to the server
                'so it knows that we want to send the message
                mSMTP.SendData "DATA" & vbCrLf
                
            Case MAIL_DATA
                smtpState = MAIL_DOT
                
                'Send Subject
                mSMTP.SendData "SUBJECT:" & txtSubject & vbCrLf
                mSMTP.SendData "FROM: " & txtFrom & vbCrLf
                mSMTP.SendData "TO:" & txtTo & vbCrLf
                                
                Dim varLines    As Variant
                Dim varLine     As Variant
                Dim strMessage  As String
                
                'Add atacchments
                strMessage = txtMessage & vbCrLf & vbCrLf & EncodedFile
                'clear the buffer for the encoded files
                EncodedFile = ""
                'Parse message to get lines
                varLines = Split(strMessage, vbCrLf)
                'clear message buffer
                strMessage = ""
                'Send each line of the message
                
                For Each varLine In varLines
                    mSMTP.SendData CStr(varLine) & vbLf
                    
                Next
                
                'Send a dot symbol so the server knows
                'that the end of the message is reached
                mSMTP.SendData vbCrLf & "." & vbCrLf
                
                
            Case MAIL_DOT
                smtpState = MAIL_QUIT
                'Send QUIT command
                mSMTP.SendData "QUIT" & vbCrLf
                
            Case MAIL_QUIT
                'Close the connection to the smtp server
                mSMTP.Close
                
        End Select
    Else
        'Check if an error occured
        mSMTP.Close
        If Not smtpState = MAIL_QUIT Then
            'If yes then print the error
            
            MsgBox "Email Sending Error: " & strServerResponse, vbCritical, "Error"
            RaiseEvent SendError
        Else
            'if the message sent successfully, print it
            RaiseEvent MessageSent
            End If
    End If
End Sub

Public Sub SendMail()
'Connect to the smtp server.
    If Not txtAttachment = "" Then
        EncodedFile = UUEncodeFile(txtAttachment)
    End If
    mSMTP.Connect txtHost, txtPort
    'reset the state
    smtpState = MAIL_CONNECT
End Sub


Public Function UUDecodeToFile(strUUCodeData As String, strFilePath As String)
    On Error Resume Next
    Dim vDataLine   As Variant      'some variables needed for decoding
    Dim vDataLines  As Variant
    Dim strDataLine As String
    Dim intSymbols  As Integer
    Dim intFile     As Integer
    Dim strTemp     As String
    Dim i           As Integer
    If Left$(strUUCodeData, 6) = "begin " Then  'check if it is a encoded file
        strUUCodeData = Mid$(strUUCodeData, InStr(1, strUUCodeData, vbLf) + 1)
    End If
    If Right$(strUUCodeData, 4) = "end" + vbLf Then 'check if "end" is available
        strUUCodeData = Left$(strUUCodeData, Len(strUUCodeData) - 7)
    End If
    intFile = FreeFile
    Open strFilePath For Binary As intFile  'open output file
        vDataLines = Split(strUUCodeData, vbLf)
        For Each vDataLine In vDataLines    'get every line
                strDataLine = CStr(vDataLine)
                intSymbols = Asc(Left$(strDataLine, 1)) 'get number of chars in
                                                        'one line. This is important
                                                        'for decoding
                strDataLine = Mid$(strDataLine, 2, intSymbols)
                For i = 1 To Len(strDataLine) Step 4
                    'now some decoding
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i, 1)) - 32) * 4 + _
                              (Asc(Mid(strDataLine, i + 1, 1)) - 32) \ 16)
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i + 1, 1)) Mod 16) * 16 + _
                              (Asc(Mid(strDataLine, i + 2, 1)) - 32) \ 4)
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i + 2, 1)) Mod 4) * 64 + _
                              Asc(Mid(strDataLine, i + 3, 1)) - 32)
                Next i
                'put the decoded data in the file
                Put intFile, , strTemp
                strTemp = ""
        Next
    'close the file
    Close intFile
End Function

Public Function UUEncodeFile(strFilePath As String) As String
    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim lFileSize       As Long         'size of the file
    Dim strFileName     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    Dim strResult       As String
    'Get file name
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    'This important: "begin 664"
    strResult = "begin 664 " + strFileName + vbLf
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize / 45 + 1
    'you need to encode every 45 bytes
    strFileData = Space(45)
    intFile = FreeFile
    'open the output file
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            If i = lEncodedLines Then
                strFileData = Space(lFileSize Mod 45)
            End If
            'get data
            Get intFile, , strFileData
            'the first byte in a line is a char, which number describes
            'how many bytes are in the line
            strTempLine = Chr(Len(strFileData) + 32)
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then
                strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            End If
            'now some encoding
            For j = 1 To Len(strFileData) Step 3
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            strResult = strResult + strTempLine + vbLf
            strTempLine = ""
            'get next line
        Next i
        'close the file
    Close intFile
    'add the "end" string
    strResult = strResult & "'" & vbLf + "end" + vbLf
    'return the encoded string
    UUEncodeFile = strResult
End Function



