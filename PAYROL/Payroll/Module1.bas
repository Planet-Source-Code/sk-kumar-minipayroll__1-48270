Attribute VB_Name = "Module1"
Public strBranch As String
Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public CompanyName As String
Public SubName As String
Public rs2 As New ADODB.Recordset

Public Sub clear()
    On Error Resume Next
        For i = 0 To Screen.ActiveForm.Count - 1
            If TypeOf Screen.ActiveForm.Controls(i) Is TextBox Then
                Screen.ActiveForm.Controls(i).Text = ""
            ElseIf TypeOf Screen.ActiveForm.Controls(i) Is MaskEdBox Then
                Screen.ActiveForm.Controls(i).Text = "__/__/____"
            ElseIf TypeOf Screen.ActiveForm.Controls(i) Is ComboBox Then
                Screen.ActiveForm.Controls(i).Text = ""
            ElseIf TypeOf Screen.ActiveForm.Controls(i) Is ListBox Then
                Screen.ActiveForm.Controls(i).clear
            End If
        Next
End Sub

Public Sub clear1()
    On Error Resume Next
        For i = 0 To Screen.ActiveForm.Count - 1
            If TypeOf Screen.ActiveForm.Controls(i) Is TextBox Then
                Screen.ActiveForm.Controls(i).Text = ""
            ElseIf TypeOf Screen.ActiveForm.Controls(i) Is MaskEdBox Then
                Screen.ActiveForm.Controls(i).Text = "__/__/____"
            ElseIf TypeOf Screen.ActiveForm.Controls(i) Is ListBox Then
                Screen.ActiveForm.Controls(i).clear
            End If
        Next
End Sub

Public Sub gspNumeric(s As Control, value As Integer)
    If (value >= 65 And value <= 90) Or (value >= 97 And value <= 122) Or (value = 39) Then
        's.Locked = True
        value = 0
        MsgBox "Only Numeric Values Accepted", vbExclamation, "Payroll"
        Exit Sub
    End If
    If (value >= 33 And value <= 47) Or (value = 32) Or (value = 64) Or (value = 124) Or (value = 92) Or (value = 94) Or (value = 96) Or (value = 126) Then
        value = 0
        MsgBox "Only Numeric Values Accepted", vbExclamation, "Payroll"
        Exit Sub
    End If
End Sub
Public Sub gspAlphaNumeric(s As Control, value As Integer)
    If (value >= 49 And value <= 57) Or (value = 48) Then
        's.Locked = True
        value = 0
        MsgBox "Enter Character  Values", vbExclamation, "Payroll"
        Exit Sub
    End If
    If (value >= 33 And value <= 47) Or (value = 64) Or (value = 124) Or (value = 92) Or (value = 94) Or (value = 96) Or (value = 126) Then
        value = 0
        MsgBox "Only String Values Accepted", vbExclamation, "Payroll"
        Exit Sub
    End If
End Sub
Public Function CheckMonth(Text As String) As Boolean
CheckMonth = True
 If Text <> "" Then
 If Val(Mid(Text, 1, InStr(1, Text, "/") - 1)) > 12 Or Val(Text) = 0 Then
    CheckMonth = False
 End If
 End If
 
End Function
Public Function CheckDate(Text As String)
CheckDate = True
If Text <> "" Then
If Not IsDate(Text) Then
    CheckDate = False
End If
End If
End Function
Public Function DateDiffer(Text1 As String, Text2 As String)
If Text1 <> "" And Text2 <> "" Then
  DateDiffer = DateDiff("d", Text1, Text2)

  
End If
    
End Function
Public Sub CheckSpecialChar(s As Control, value As Integer)

If (value >= 33 And value <= 47) Or (value = 32) Or (value = 64) Or (value = 124) Or (value = 92) Or (value = 94) Or (value = 96) Or (value = 126) Then
        value = 0
        MsgBox "Not Accepted Special Character", vbExclamation, "Payroll"
        Exit Sub
    End If
End Sub
Public Function NumericCheck(TextValue As String, Values As Byte) As Byte
If Values > 47 And Values < 58 Or Values = 8 Or Values = 46 Then
    NumericCheck = Values
Else
    NumericCheck = 0
End If
Mycheck = TextValue Like "*.*"
If Mycheck = True Then
If NumericCheck = 46 Then
  NumericCheck = 0
End If

End If
End Function
Public Function NumericCheck1(TextValue As String, Values As Byte) As Byte

If Values > 47 And Values < 58 Or Values = 8 Or Values = 46 Then
    NumericCheck1 = Values
Else
    NumericCheck1 = 0
End If

End Function

' Return words for this value between 1 and 999.
Public Function Words_1_999(ByVal num As Integer) As String
Dim hundreds As Integer
Dim remainder As Integer
Dim result As String

    hundreds = num \ 100
    remainder = num - hundreds * 100

    If hundreds > 0 Then
        result = Words_1_19(hundreds) & " Hundred "
    End If

    If remainder > 0 Then
        result = result & Words_1_99(remainder)
    End If

    Words_1_999 = Trim$(result)
End Function
' Return a word for this value between 1 and 99.
Public Function Words_1_99(ByVal num As Integer) As String
Dim result As String
Dim tens As Integer

    tens = num \ 10

    If tens <= 1 Then
        ' 1 <= num <= 19
        result = result & " " & Words_1_19(num)
    Else
        ' 20 <= num
        ' Get the tens digit word.
        Select Case tens
            Case 2
                result = "Twenty"
            Case 3
                result = "Thirty"
            Case 4
                result = "Fourty"
            Case 5
                result = "Fifty"
            Case 6
                result = "Sixty"
            Case 7
                result = "Seventy"
            Case 8
                result = "Eighty"
            Case 9
                result = "Ninety"
        End Select

        ' Add the ones digit number.
        result = result & " " & Words_1_19(num - tens * 10)
    End If

    Words_1_99 = Trim$(result)
End Function
' Return a word for this value between 1 and 19.
Public Function Words_1_19(ByVal num As Integer) As String
    Select Case num
        Case 1
            Words_1_19 = "One"
        Case 2
            Words_1_19 = "Two"
        Case 3
            Words_1_19 = "Three"
        Case 4
            Words_1_19 = "Four"
        Case 5
            Words_1_19 = "Five"
        Case 6
            Words_1_19 = "Six"
        Case 7
            Words_1_19 = "Seven"
        Case 8
            Words_1_19 = "Eight"
        Case 9
            Words_1_19 = "Nine"
        Case 10
            Words_1_19 = "Ten"
        Case 11
            Words_1_19 = "Eleven"
        Case 12
            Words_1_19 = "Twelve"
        Case 13
            Words_1_19 = "Thirteen"
        Case 14
            Words_1_19 = "Fourteen"
        Case 15
            Words_1_19 = "Fifteen"
        Case 16
            Words_1_19 = "Sixteen"
        Case 17
            Words_1_19 = "Seventeen"
        Case 18
            Words_1_19 = "Eightteen"
        Case 19
            Words_1_19 = "Nineteen"
    End Select
End Function
' Return a string of words to represent the
' integer part of this value.
Public Function Words_1_all(ByVal num As Currency) As String
Dim power_value(1 To 5) As Currency
Dim power_name(1 To 5) As String
Dim digits As Integer
Dim result As String
Dim i As Integer

    ' Initialize the power names and values.
    power_name(1) = "Billion": power_value(1) = 1000000000
   ' power_name(2) = "billion":  power_value(2) = 1000000000
    power_name(2) = "Crores":  power_value(2) = 10000000
    'power_name(3) = "million":  power_value(3) = 1000000
    power_name(3) = "Lakhs":  power_value(3) = 100000
    power_name(4) = "Thousand": power_value(4) = 1000
    power_name(5) = "":         power_value(5) = 1

    For i = 1 To 5
        ' See if we have digits in this range.
        If num >= power_value(i) Then
            ' Get the digits.
            digits = Int(num / power_value(i))

            ' Add the digits to the result.
            If Len(result) > 0 Then result = result & " "
            result = result & _
                Words_1_999(digits) & _
                " " & power_name(i)

            ' Get the number without these digits.
            num = num - digits * power_value(i)
        End If
    Next i

    Words_1_all = Trim$(result)
End Function
' Return a string of words to represent this
' currency value in dollars and cents.
Public Function Words_Money(ByVal num As Currency) As String
Dim dollars As Currency
Dim cents As Integer
Dim dollars_result As String
Dim cents_result As String

    ' Separate the dollars and cents.
    dollars = Int(num)
    cents = CInt((num - dollars) * 100#)

    dollars_result = Words_1_all(dollars)
    If Len(dollars_result) = 0 Then dollars_result = "zero"

    cents_result = Words_1_all(cents)
    If Len(cents_result) = 0 Then cents_result = "zero"

    Words_Money = "Rupees  " & dollars_result & _
        " And Paise " & cents_result & " Only"
End Function

Public Function NumToWords(num As Double)
    NumToWords = Words_Money(num)
End Function

