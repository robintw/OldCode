VERSION 5.00
Begin VB.Form frmStockReport 
   Caption         =   "Stock Report"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewReport 
      Caption         =   "View Report"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Frame frameSorting 
      Caption         =   "Sort By:"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   3135
      Begin VB.OptionButton optISBN 
         Caption         =   "ISBN"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optTitle 
         Caption         =   "Title"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label lblWarning 
      Caption         =   "Warning"
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmStockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdViewReport_Click()
Dim OpenString As String

Set rsReport = New ADODB.Recordset

If optISBN.Value = True Then
OpenString = "SELECT tblBookCopy.ISBN, tblBook.ISBN, Title, Author, Publisher, SellingPrice, CopyID FROM tblBook, tblBookCopy WHERE tblBookCopy.ISBN = tblBook.ISBN ORDER BY tblBookCopy.ISBN, CopyID;"
Else
OpenString = "SELECT tblBookCopy.ISBN, tblBook.ISBN, Title, Author, Publisher, SellingPrice, CopyID FROM tblBook, tblBookCopy WHERE tblBookCopy.ISBN = tblBook.ISBN ORDER BY Title, CopyID;"
End If

'Open the recordset for the report
rsReport.Open OpenString, cnnGlobalConnection, adOpenStatic

If rsReport.EOF = True Then
'If EOF occurs immediatly (ie. no records are returned) then give an error
    MsgBox "No books were bought by this student in this date period!", vbInformation
    Exit Sub
End If

'Show the report
rptStockReport.Show
End Sub

Private Sub Form_Load()
lblWarning.Caption = "WARNING:" & vbCrLf & vbCrLf & "This report will be list all the books in stock at the moment, and will therefore be very LARGE, probably running to many ten's of pages. Be sure to check how many pages it is before clicking print!"

End Sub
