option explicit

Sub lab1Case1
    const PROGRAM_NAME = "Лаб. роб. 1 - спосіб 1"
    dim f as single, x as single, y as single, n as single, a as single, b as single
    dim answer as string
    dim HALF_PI as single
    
    x = inputBox("x = ", PROGRAM_NAME, 3.16e-3)
    y = inputBox("y = ", PROGRAM_NAME, 4.7e1)
    n = inputBox("n = ", PROGRAM_NAME, 4)
    a = inputBox("a = ", PROGRAM_NAME, 2.6)
    b = inputBox("b = ", PROGRAM_NAME, 7.8)
    
    HALF_PI = PI/2
    
    f = abs(sin(3*x + HALF_PI) + cos(3*x - 3*HALF_PI))
    f = f / (sqr(abs(x^(n + 2))) - y*log(a * x + b)/log(10))
    
    answer = "f = " & format(f, "scientific")
    msgBox answer, mb_IconInformation, PROGRAM_NAME
End Sub

Sub lab1Case2
    const PROGRAM_NAME = "Лаб. роб. 1 - спосіб 2"
    dim x as single, y as single, n as single, a as single, b as single
    dim f as single, f1 as single, f2 as single, f3 as single, f4 as single
    dim f5 as single, f6 as single 
    dim answer as string
    dim HALF_PI as single
    
    x = inputBox("x = ", PROGRAM_NAME, 3.16e-3)
    y = inputBox("y = ", PROGRAM_NAME, 4.7e1)
    n = inputBox("n = ", PROGRAM_NAME, 4)
    a = inputBox("a = ", PROGRAM_NAME, 2.6)
    b = inputBox("b = ", PROGRAM_NAME, 7.8)

    HALF_PI = PI/2    
    
    f6 = sin(3*x + HALF_PI)
    f5 = cos(3*x - 3*HALF_PI)
    f4 = sqr(abs(x^(n + 2)))
    f3 = y*log(a * x + b)/log(10)
    f2 = abs(f6 + f5)
    f1 = f4 - f3
    f = f2 / f1
    
    answer = "f = " & format(f, "scientific")
    msgBox answer, mb_IconInformation, PROGRAM_NAME
End Sub

Sub lab1Case3
    const PROGRAM_NAME = "Лаб. роб. 1 - спосіб 3"
    dim f as single, x as single, y as single, n as single, a as single, b as single
    dim answer as string
    dim HALF_PI as single
    
    x = inputBox("x = ", PROGRAM_NAME, 3.16e-3)
    y = inputBox("y = ", PROGRAM_NAME, 4.7e1)
    n = inputBox("n = ", PROGRAM_NAME, 4)
    a = inputBox("a = ", PROGRAM_NAME, 2.6)
    b = inputBox("b = ", PROGRAM_NAME, 7.8)
    
    HALF_PI = PI/2
    
    f = sin(3*x + HALF_PI)
    f = f + cos(3*x - 3*HALF_PI)
    f = abs(f)
    f = f / (sqr(abs(x^(n + 2))) - y*log(a * x + b)/log(10))
    
    answer = "f = " & format(f, "scientific")
    msgBox answer, mb_IconInformation, PROGRAM_NAME
End Sub