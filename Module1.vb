﻿Module Module1

    Sub Main()
        'FunNow()
        'FunDateAdd()
        FunDateDiff()

        Console.Read()
    End Sub

    Sub FunNow()
        Dim hoy
        Dim hoy2 As Date = Now

        hoy = Now
        Console.WriteLine("La fecha y hora es: " & hoy)
        Console.WriteLine("La fecha y hora es: " & hoy2)

    End Sub

    Sub FunDateAdd()
        'Dim fecha = DateAdd("m", 1, "31-enero-2024")
        'Console.WriteLine(fecha)

        Dim fecha1 As Date
        Dim intervalo As String
        Dim numero As Integer
        Dim mensaje As String

        intervalo = "m" 'Console.ReadLine("Ingresa el intervalo a utilizar: ")
        intervalo = InputBox("Ingrese el intervalo a utilizar")

        fecha1 = InputBox("Ingrese fecha")
        numero = InputBox("Ingrese un numero de meses a añadir")
        mensaje = "La nueva fecha es: " & DateAdd(intervalo, numero, fecha1)
        MsgBox(mensaje)


    End Sub

    Sub FunDateDiff()
        Dim fecha As Date
        Console.WriteLine("Ingresa una Fecha")
        fecha = Console.ReadLine()
        Console.WriteLine("Los dias desde hoy: " & DateDiff("d", "1-1-2023", fecha))

    End Sub



End Module
