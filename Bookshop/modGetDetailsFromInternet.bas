Attribute VB_Name = "modGetDetailsFromInternet"
'---------------------------------------------------------------------------------------
' Module    : modGetDetailsFromInternet
' Author    : Robin Wilson
' Purpose   : Get details for a particular book from Amazon's website
'---------------------------------------------------------------------------------------

Option Explicit


Public Function GetDetailsFromInternet(strISBN As String) As udfBookDetails
'---------------------------------------------------------------------------------------
' Procedure : GetDetailsFromInternet in modGetDetailsFromInternet
' Author    : Robin Wilson
' Purpose   : Get the details for the book with the ISBN passed, and return a
'             UDF with Title, Author and Publisher in it
'---------------------------------------------------------------------------------------
'
    Dim xmlDoc As New DOMDocument
    Dim root As IXMLDOMElement
    Dim Success As Boolean
    
    Dim nodTitle As IXMLDOMNode
    Dim nodAuthor As IXMLDOMNode
    Dim nodPublisher As IXMLDOMNode
    
    ' Allow the document to complete loading
    xmlDoc.async = False
    'Uses amazon.co.uk at the moment - can replace with amazon.com
    Success = xmlDoc.Load("http://ecs.amazonaws.com/onca/xml?Service=AWSECommerceService&AWSAccessKeyId=1CE7SK4ZPTNDQZCWBP82&Operation=ItemLookup&AssociateTag=devconn-20&Version=2007-01-15&ItemId=" & strISBN & "&ResponseGroup=ItemAttributes&type=lite&f=xml&locale=uk")

    'If the document succeeded loading
    If Success = True Then
        
        'Get the details from the relevant parts of the XML
        GetDetailsFromInternet.ISBN = strISBN
        
        Set nodTitle = xmlDoc.selectSingleNode("//Title")
        
        If nodTitle Is Nothing = False Then
            GetDetailsFromInternet.Title = nodTitle.Text
        End If
        
        Set nodAuthor = xmlDoc.selectSingleNode("//Author")
        
        If nodAuthor Is Nothing = False Then
            GetDetailsFromInternet.Author = nodAuthor.Text
        End If
        
        Set nodPublisher = xmlDoc.selectSingleNode("//Publisher")
        
        If nodPublisher Is Nothing = False Then
            GetDetailsFromInternet.Publisher = nodPublisher.Text
        End If
        
        
    Else
            'Otherwise fill all data with Null's
            GetDetailsFromInternet.Title = vbNull
            GetDetailsFromInternet.Author = vbNull
            GetDetailsFromInternet.Publisher = vbNull
    End If
    
    Set root = Nothing
End Function
