Option Explicit

Sub Main
   Const PROGRAM_NAME = "Лаб. роб. Libre Basic 2"
   
   Const NUMBER_FORMAT = "0,000E+#"
   Const ERR_MSG_PROGRAM_STOP = "Програма призупинена"
   Const DEFAULT_X = 1
   Const DEFAULT_A = 3
   Const MSG_X = "Введіть дійсне число"
   Const MSG_A = "Введіть ціле число від 1 до 3"
   Const MIN_A = 1
   Const MAX_A = 3
   
   Dim symbolEnter As String
   Dim numberAsTextValue As String
   Dim x As Single
   Dim a As Single
   Dim f As Single
   Dim errMsg As String
   Dim answer As String
   
   SymbolEnter = Chr(13)
   NumberAsTextValue = InputBox(MSG_A, PROGRAM_NAME, DEFAULT_A) 
   x = InputBox(MSG_X, PROGRAM_NAME, DEFAULT_X)
   
   If Not IsNumeric(NumberAsTextValue) Then
      errMsg = "Введене значення некоректне" & SymbolEnter & ERR_MSG_PROGRAM_STOP
      msgBox errMsg, mb_IconExclamation, PROGRAM_NAME
      Exit Sub
   End If
   
   a = CSng(NumberAsTextValue)
   
   If a < MIN_A Or a > MAX_A Then
      errMsg = "Введене значення некоректне" & SymbolEnter & ERR_MSG_PROGRAM_STOP
      msgBox errMsg, mb_IconExclamation, PROGRAM_NAME
      Exit Sub
   End If
   
   If x >= 2 * a Then
      f = Log(x) / Log(a + 1) + Log(x) / Log(10)
   ElseIf 2 * a > x And x > a / 3 Then 
      f = (Sin(x) / x) + Tanh(x)
   ElseIf x = a / 3 Then
      f = 1 + x ^ (a + 1)
   ElseIf x < a / 3 Then
      f = Sin(x + a * Pi) + 2
   Else
      ' Add the appropriate action or default value for 'f' here
   End If

   answer = "f=" & Format(f, "scientific")
   msgBox answer, mb_IconInformation, PROGRAM_NAME
End Sub
