option explicit

Sub Main
   const PROGRAM_NAME = "Лаб. роб. Libre Basic 2"
   const ERR_MSG_PROGRAM_STOP= "Програма призупинена"
   const MSG_X = "Введіть дійсне число"
   const MSG_A = "Введіть ціле число від 1 до 3"
   const DEFAULT_X = 1
   const DEFAULT_A = 4

   dim x as single
   dim a as single
   dim f as single
   dim answer as string

   x = inputBox(MSG_X, PROGRAM_NAME, DEFAULT_X) 
   a = inputBox(MSG_A, PROGRAM_NAME, DEFAULT_A)

   if x >= 2*a then
      f = log(x)/log(a+1) + log(x)/log(10)
   elseif 2*a > x > a/3 then 
      f = (sin(x)/x) + tanh(x)
   elseif x = a/3 then
      f = 1 + x^a+1
   elseif x < a/3 then
      f = sin(x + a*Pi) + 2 else
      ERR MSG_PROGRAM_STOP
   end if
   
   answer = "f=" & format (f, "scientific") 
   msgBox answer, mb_IconInformation, PROGRAM_NAME
End Sub

