Attribute VB_Name = "modTypeDefinitions"
Option Explicit

Global CurrentBookDetails As udfBookDetails
'Global OrderDetails As udfOrderDetails
Global cnnGlobalConnection As ADODB.Connection
Global udfConfig As udfConfig
Global rsReport As ADODB.Recordset
Global rsSales As ADODB.Recordset

Public Type udfBookDetails
'Type to hold book details returned from Internet
    ISBN As String
    Title As String
    Author As String
    Publisher As String
End Type

Public Type udfOrderDetails
'Type to hold order details for passing between proc's
    BookDetails As udfBookDetails
    Alphacode As String
    DateOrdered As Date
End Type


Public Type udfConfig
'Type to hold all config info
    DatabasePath As String
    AllowanceSearchAsYouType As Boolean
    EmailSubject As String
    EmailMessage As String
    EmailDomain As String 'eg: @mgc.worcs.sch.uk
    EmailServer As String
    EmailServerPort As String
    EmailFromAddress As String
    TemporaryFolder As String
    AdminPassword As String
    SellerPassword As String
    'Add more here if needed
End Type

