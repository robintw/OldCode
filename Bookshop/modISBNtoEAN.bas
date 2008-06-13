Attribute VB_Name = "modISBNtoEAN"
'---------------------------------------------------------------------------------------
' Module    : modISBNtoEAN
' Author    : Robin Wilson
' Purpose   : Contains functions used for converting the scanned in EAN to an ISBN
'---------------------------------------------------------------------------------------

Option Explicit

Public Function ValidateISBN(ByVal ISBN As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : ValidateISBN in modISBNtoEAN
' Author    : Robin Wilson
' Purpose   : Check the Check Digit in a 10 character ISBN
'---------------------------------------------------------------------------------------
'
Dim i As Integer
Dim RunningTotal As Integer
Dim ProductForChar As Integer
Dim ModResult As String

RunningTotal = 0

'All ISBNs must be 10 characters
If Len(ISBN) <> 10 Then
    ValidateISBN = False
    Exit Function
End If

For i = 1 To Len(ISBN) - 1
    ProductForChar = Mid(ISBN, i, 1) * i
    RunningTotal = RunningTotal + ProductForChar
Next i

ModResult = Trim(Str(RunningTotal Mod 11))

'Deals with a CD of 10 being written as X
If ModResult = "10" Then ModResult = "X"

If ModResult = Mid(ISBN, Len(ISBN), 1) Then
    ValidateISBN = True
Else
    ValidateISBN = False
End If
End Function

Private Function ValidateEAN(ByVal EAN As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : ValidateEAN in modISBNtoEAN
' Author    : Robin Wilson
' Purpose   : Checks the Check Digit in a 13 character EAN
'---------------------------------------------------------------------------------------
'
Dim i As Integer
Dim CharTotal As Integer
Dim RunningTotal As Integer
Dim ModdedValue As Integer
Dim Weighting As String

'Odd digits (counting from the right) are multiplied by 3, even digits by 2
Weighting = "1313131313131"

'All EANs must be 13 characters long
If Len(EAN) <> 13 Then
    ValidateEAN = False
    Exit Function
End If

RunningTotal = 0

For i = 1 To 13
    CharTotal = Mid(EAN, i, 1) * Mid(Weighting, i, 1)
    RunningTotal = RunningTotal + CharTotal
Next i

ModdedValue = RunningTotal Mod 10

If ModdedValue = 0 Then
    ValidateEAN = True
Else
    ValidateEAN = False
End If
End Function


Private Function CreateISBNCheckDigit(ISBN As String) As String
'---------------------------------------------------------------------------------------
' Procedure : CreateISBNCheckDigit in modISBNtoEAN
' Author    : Robin Wilson
' Purpose   : Creates a Check Digit for the given 9 character ISBN without a Check Digit
'---------------------------------------------------------------------------------------
'
Dim i As Integer
Dim RunningTotal As Integer
Dim ProductForChar As Integer
Dim ModResult As Integer
Dim ISBNWithCD As String
Dim NewCheckDigit As String

RunningTotal = 0

'An ISBN without a check digit will be 9 chars
If Len(ISBN) <> 9 Then
    Err.Raise 1, , "The ISBN did not have the required number of digits"
    CreateISBNCheckDigit = vbNull
    Exit Function
End If


For i = 1 To Len(ISBN)
    ProductForChar = Mid(ISBN, i, 1) * i
    RunningTotal = RunningTotal + ProductForChar
Next i

ModResult = RunningTotal Mod 11

'Deals with VB adding a space to the left of the Str(ModResult) result
If ModResult = 10 Then
NewCheckDigit = "X"
Else
NewCheckDigit = Right(Str(ModResult), 1)
End If

ISBNWithCD = ISBN & NewCheckDigit

CreateISBNCheckDigit = ISBNWithCD
End Function

Private Function StringOpsOnEAN(EAN As String)
'---------------------------------------------------------------------------------------
' Procedure : StringOpsOnEAN in modISBNtoEAN
' Author    : Robin Wilson
' Purpose   : Removes 978 from the beginning, and the Check Digit from the end of an EAN
'---------------------------------------------------------------------------------------
'
Dim ISBN As String
'Remove 978
ISBN = Right(EAN, 10)

'Remove the EAN13 Check Digit
ISBN = Left(ISBN, 9)

StringOpsOnEAN = ISBN
End Function

Public Function ConvertEANToISBN(EAN As String) As String
'---------------------------------------------------------------------------------------
' Procedure : ConvertEANToISBN in modISBNtoEAN
' Author    : Robin Wilson
' Purpose   : Converts the given EAN to an ISBN - and returns it
'---------------------------------------------------------------------------------------
'
Dim ISBN As String

If ValidateEAN(EAN) Then
        ISBN = CreateISBNCheckDigit(StringOpsOnEAN(EAN))
End If

ConvertEANToISBN = ISBN
End Function


