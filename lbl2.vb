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
   
   dim callFunctionResult as Object
   
   Dim symbolEnter As String
   Dim numberAsTextValue As String
   Dim x As Single
   Dim a As Single
   Dim f As Single
   Dim errMsg As String
   Dim answer As String
   
   SymbolEnter = Chr(13)
   NumberAsTextValue = InputBox(MSG_A, PROGRAM_NAME, DEFAULT_A) 
   
   If Not IsNumeric(NumberAsTextValue) Then
      errMsg = "Введене значення не є числовим" & SymbolEnter & ERR_MSG_PROGRAM_STOP
      msgBox errMsg, mb_IconExclamation, PROGRAM_NAME
      Exit Sub
   End If
   
   a = CSng(NumberAsTextValue)
   
   If a < MIN_A Or a > MAX_A Then
      errMsg = "Введене значення виходить за межі від 1 - 3" & SymbolEnter & ERR_MSG_PROGRAM_STOP
      msgBox errMsg, mb_IconExclamation, PROGRAM_NAME
      Exit Sub
   End If
   
   NumberAsTextValue = InputBox(MSG_X, PROGRAM_NAME, DEFAULT_X) 
   
   If Not IsNumeric(NumberAsTextValue) Then
      errMsg = "Введене значення не є числовим" & SymbolEnter & ERR_MSG_PROGRAM_STOP
      msgBox errMsg, mb_IconExclamation, PROGRAM_NAME
      Exit Sub
   End If
   
   dim example1 as single
   dim example2 as single
   example1 = 2 * a
   example2 = a / 3
   
   callFunctionResult = createUnoService("com.sun.star.sheet.FunctionAccess")
   
   If x >= example1 Then
      f = callFunctionResult.callFunction("LOG", Array(x + 1))
      f = f /  callFunctionResult.callFunction("LOG10", Array(x))
   ElseIf example1 > x And x > example2 Then
      f = callFunctionResult.callFunction("SIN", Array(x))
      f = f / x
      f = f + callFunctionResult.callFunction("TAN", Array(x, ))
   ElseIf x = example2 Then
      f = 1
      f = f + callFunctionResult.callFunction("POWER", Array(a + 1))
   ElseIf x < example2 Then
      f = callFunctionResult.callFunction("SIN", Array(x + a * Pi))
      f = f + 2
   End If

   answer = "f=" & Format(f, "scientific")
   msgBox answer, mb_IconInformation, PROGRAM_NAME
End Sub
