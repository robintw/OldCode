VERSION 5.00
Begin VB.Form frmReportMenu 
   Caption         =   "Report Menu"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStockReport 
      Caption         =   "Stock"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdOrderReport 
      Caption         =   "&Orders"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton cmdStudentReport 
      Caption         =   "S&elected student reports"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdSales 
      Caption         =   "&Sales"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdStudentReportAll 
      Caption         =   "&All student reports"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmReportMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOrderReport_Click()
Dim rsAnyThere As ADODB.Recordset
Set rsAnyThere = New ADODB.Recordset

rsAnyThere.ActiveConnection = cnnGlobalConnection
rsAnyThere.Open "SELECT * from tblOrder", cnnGlobalConnection, adOpenStatic

If rsAnyThere.RecordCount = 0 Then
    MsgBox "Sorry, there are no orders to view at the moment.", vbInformation
    Exit Sub
End If

rptOrderForm.Show
End Sub

Private Sub cmdSales_Click()
frmSalesReport.Show
End Sub

Private Sub cmdStockReport_Click()
frmStockReport.Show
End Sub

Private Sub cmdStudentReport_Click()
frmStudentReport.Show
End Sub

Private Sub cmdStudentReportAll_Click()
StudentReportAllShow
End Sub
