Option Explicit On

Module Web
    Public Function PDF() As String
        Dim MyWs As New WebPrinter.PrinterDocument

        PDF = ""

        Try
            MyWs.Url = Datos.URL
            MyWs.Credentials = New System.Net.NetworkCredential(Datos.User, Datos.Password, Datos.Domain)
            PDF = MyWs.Pendiente_PDF
        Catch ex As Exception
            Log(True, "Error en los servicios WEB Pendiente PDF: " + ex.Message)
            Log(False, "URL: " + Datos.URL)
        End Try

        MyWs.Dispose()
    End Function

    Public Function Etiquetas_Kit() As String
        Dim MyWs As New WebPrinter.PrinterDocument

        Etiquetas_Kit = ""

        Try
            MyWs.Url = Datos.URL
            MyWs.Credentials = New System.Net.NetworkCredential(Datos.User, Datos.Password, Datos.Domain)
            Etiquetas_Kit = MyWs.Etiquetas_Kits()
        Catch ex As Exception
            Log(True, "Error en los servicios WEB Etiquetas Kit: " + ex.Message)
            Log(False, "URL: " + Datos.URL)
        End Try

        MyWs.Dispose()
    End Function

    Public Function Etiquetas_Kit_GS1() As String
        Dim MyWs As New WebPrinter.PrinterDocument

        Etiquetas_Kit_GS1 = ""

        Try
            MyWs.Url = Datos.URL
            MyWs.Credentials = New System.Net.NetworkCredential(Datos.User, Datos.Password, Datos.Domain)
            Etiquetas_Kit_GS1 = MyWs.Etiquetas_Kits_GS1()
        Catch ex As Exception
            Log(True, "Error en los servicios WEB Etiquetas Kit: " + ex.Message)
            Log(False, "URL: " + Datos.URL)
        End Try

        MyWs.Dispose()
    End Function

    Public Function Etiquetas_Lote() As String
        Dim MyWs As New WebPrinter.PrinterDocument

        Etiquetas_Lote = ""

        Try
            MyWs.Url = Datos.URL
            MyWs.Credentials = New System.Net.NetworkCredential(Datos.User, Datos.Password, Datos.Domain)
            Etiquetas_Lote = MyWs.Etiquetas_GS1()
        Catch ex As Exception
            Log(True, "Error en los servicios WEB Etiquetas Lote: " + ex.Message)
            Log(False, "URL: " + Datos.URL)
        End Try

        MyWs.Dispose()
    End Function

    Public Function Etiquetas_Clientes() As String
        Dim MyWs As New WebPrinter.PrinterDocument

        Etiquetas_Clientes = ""

        Try
            MyWs.Url = Datos.URL
            MyWs.Credentials = New System.Net.NetworkCredential(Datos.User, Datos.Password, Datos.Domain)
            Etiquetas_Clientes = MyWs.Etiquetas_Envios()
        Catch ex As Exception
            Log(True, "Error en los servicios WEB Etiquetas Clientes: " + ex.Message)
            Log(False, "URL: " + Datos.URL)
        End Try

        MyWs.Dispose()
    End Function


    Public Sub Resolver(xId As Integer, xError As String)
        Dim MyWs As New WebPrinter.PrinterDocument

        Try
            MyWs.Url = Datos.URL
            MyWs.Credentials = New System.Net.NetworkCredential(Datos.User, Datos.Password, Datos.Domain)
            MyWs.resolver(xId, xError)
        Catch ex As Exception
            Log(True, "Error en los servicios WEB Resolver: " + ex.Message)
            Log(False, "URL: " + Datos.URL)
        End Try

        MyWs.Dispose()
    End Sub

End Module
