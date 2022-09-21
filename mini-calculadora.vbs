dim numeroUno
dim numeroDos
dim resultado
dim opcion

opcion = inputbox("Elija una operacion: " & vbCrLf & vbCrLf & "1. Suma" & vbCrLf &_
    "2. Resta" & vbCrLf &_
    "3. Multiplicacion" & vbCrLf &_
    "4. Divivision" & vbCrLf &_
    "5. Salir", "Visual Basic Script")

if (opcion = 5) then
    msgbox("Gracias por tu visita!!")
end if

if (opcion >= 6 or opcion <= 0) then
    msgbox("Error: Opcion no disponible!!")
end if



if (opcion = 1) then
    numeroUno = CInt(inputbox("Ingrese primer valor","Visual Basic Scritp"))
    numeroDos = CInt(inputbox("Ingrese segundo valor","Visual Basic Script"))

    resultado = numeroUno + numeroDos
    msgbox("Resultado de " & numeroUno & "+" & numeroDos & " es: " & resultado)
end if


if (opcion = 2) then
    numeroUno = CInt(inputbox("Ingrese primer valor","Visual Basic Scritp"))
    numeroDos = CInt(inputbox("Ingrese segundo valor","Visual Basic Script"))

    resultado = numeroUno - numeroDos
    msgbox("Resultado de " & numeroUno & "-" & numeroDos & " es: " & resultado)
end if

if (opcion = 3) then
    numeroUno = CInt(inputbox("Ingrese primer valor","Visual Basic Scritp"))
    numeroDos = CInt(inputbox("Ingrese segundo valor","Visual Basic Script"))

    resultado = numeroUno * numeroDos
    msgbox("Resultado de " & numeroUno & "x" & numeroDos & " es: " & resultado)
end if

if (opcion = 4) then
    numeroUno = CInt(inputbox("Ingrese primer valor","Visual Basic Scritp"))
    numeroDos = CInt(inputbox("Ingrese segundo valor","Visual Basic Script"))

    if (numeroDos = 0) then
        msgbox("No se puede dividir entre cero")
    else
        resultado = numeroUno / numeroDos
        msgbox("Resultado de " & numeroUno & "/" & numeroDos & " es: " & resultado)
    end if
end if

