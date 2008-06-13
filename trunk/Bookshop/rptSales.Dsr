VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSales 
   Caption         =   "Sales"
   ClientHeight    =   13995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   24686
   SectionData     =   "rptSales.dsx":0000
End
Attribute VB_Name = "rptSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Sales"
Option Explicit

Private Sub ActiveReport_DataInitialize()
Fields.Add "ISBN"
Fields.Add "Title"
Fields.Add "Author"
Fields.Add "Publisher"
Fields.Add "CostPrice"
Fields.Add "SellingPrice"
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
If rsSales.EOF = True Then
rsSales.Close
Exit Sub
End If

EOF = False

Fields("ISBN").Value = rsSales.Fields("tblBookCopy.ISBN").Value
Fields("Title").Value = rsSales!Title
Fields("Author").Value = rsSales!Author
Fields("Publisher").Value = rsSales!Publisher
Fields("CostPrice").Value = rsSales!CostPrice
Fields("SellingPrice").Value = rsSales!SellingPrice

rsSales.MoveNext

End Sub

