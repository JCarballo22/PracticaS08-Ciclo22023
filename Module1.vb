Module Module1

    Sub Main()
        'FunNow()
        'FunDateAdd()
        'FunDateDiff()
        'FunDatePart()
        'FunFormat()
        Excepciones()


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

    Sub FunDatePart()
        Dim fechaString, mensaje As String
        Dim fechaActual As Date

        'Formas de ingresar fechas: Septiembre 21, 2023. 21/Septiembre/2023. 21/09/2023.
        fechaString = InputBox("Ingrese una Fecha:")
        fechaActual = CDate(fechaString)

        mensaje = "Trimestral: " & DatePart(DateInterval.Quarter, fechaActual)
        MsgBox(mensaje)

        mensaje = "Día del mes: " & DatePart(DateInterval.Day, fechaActual)
        MsgBox(mensaje)

        mensaje = "Las semanas del año: " & DatePart("ww", fechaActual)
        MsgBox(mensaje)

        mensaje = "Día de la semana: " & DatePart("w", fechaActual)
        MsgBox(mensaje)


    End Sub

    Sub FunFormat()
        Dim miHora, miFecha, miString
        miHora = #17:05:05#
        miFecha = "Septiembre 21, 2023"

        miString = Format(TimeOfDay, "Long time")
        Console.WriteLine(miString)

        miString = Format(miFecha, "Long date")
        Console.WriteLine(miString)

        miString = Format(miHora, "h:m:s")
        Console.WriteLine(miString)

        miString = Format(miHora, "hh:mm:ss")
        Console.WriteLine(miString)

        miString = Format(miFecha, "Long date")
        Console.WriteLine(miString)



    End Sub

    Sub Excepciones()
        Try
            Console.WriteLine("Por favor ingresa un numero: ")
            Dim numS = Console.ReadLine()
            Dim numI = Integer.Parse(numS)
            Dim cuadrado = numI * numI
            Console.WriteLine("Su numero " & numI & " al cuadrado es: " & cuadrado)
        Catch ex As Exception
            Console.WriteLine("Usuaro el numero no deber ir en letras " & ex.ToString())
        Finally
            Console.WriteLine("Se ejecuto Finally")
        End Try




    End Sub

End Module
