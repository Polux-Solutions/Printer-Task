Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Drawing.Printing
Imports iTextSharp.text
Imports iTextSharp.text.pdf


Public Structure stDatos
    Public URL As String
    Public Delay As Integer
    Public Folder As String
    Public Log As String
    Public Log_Activity As String
    Public Version As String
    Public User As String
    Public Password As String
    Public Domain As String
End Structure

Module Principal
    Public Datos As stDatos
    Sub Main()

        KillAll()

        If Leer_Parametros() Then bucle_Infinito()
    End Sub

    Private Sub bucle_Infinito()
        Log(True, "Inicio Proceso")
        Proceso()
        'Exit Sub

        Do While True
            Threading.Thread.Sleep(Datos.Delay)
            Proceso()
        Loop
    End Sub
    Public Sub Proceso()
        Dim xml As String = ""

        For n = 1 To 4
            Threading.Thread.Sleep(Datos.Delay)
            xml = ""

            Select Case n
                Case 1 : xml = Web.PDF()
                    If xml <> "" Then Enviar_PDF(xml)
                Case 2 : xml = Web.Etiquetas_Kit()
                Case 3 : xml = Web.Etiquetas_Clientes()
                Case 4 : xml = Web.Etiquetas_Lote()
                Case 5 : xml = Web.Etiquetas_Kit_GS1()
            End Select

            If n <> 1 Then
                If xml <> "" Then Imprimir_Etiqueta(xml)
            End If

        Next
    End Sub

    Private Sub KillAll()
        Dim MiProceso As String = System.Diagnostics.Process.GetCurrentProcess.ProcessName
        Dim i As Integer

        i = InStr(MiProceso, ".")
        If i > 0 Then MiProceso = MiProceso.Substring(0, i - 1)
        For Each pr As Process In Diagnostics.Process.GetProcessesByName(MiProceso)
            If pr.Id <> Diagnostics.Process.GetCurrentProcess.Id Then
                pr.Kill()
            End If
        Next
    End Sub

    Private Function Leer_Parametros() As Boolean
        Dim Config As New System.Configuration.AppSettingsReader
        Dim t As String = ""

        Leer_Parametros = True

        Try
            Datos.URL = Config.GetValue("URL", GetType(System.String)).ToString
            Datos.Folder = Config.GetValue("FOLDER", GetType(System.String)).ToString
            Datos.Version = Config.GetValue("VERSION", GetType(System.String)).ToString
            Datos.User = Config.GetValue("USER", GetType(System.String)).ToString
            Datos.Password = Config.GetValue("PASSWORD", GetType(System.String)).ToString
            Datos.Domain = Config.GetValue("DOMAIN", GetType(System.String)).ToString
            Datos.Delay = 5000
            t = Config.GetValue("DELAY", GetType(System.String)).ToString
            If IsNumeric(t) Then Datos.Delay = CInt(t)

            Datos.Log = Datos.Folder + "\Printer-Task.log"
            Datos.Log_Activity = Datos.Folder + "\Printer-Task-Activity.log"
        Catch ex As Exception
            MsgBox("Error al leer parámetros " & ex.Message, MsgBoxStyle.Critical)
            Leer_Parametros = False
            Exit Function
        End Try

        Try
            Datos.User = Decrypt("_venus13", Datos.User)
            Datos.Password = Decrypt("_venus13", Datos.Password)
            Datos.Domain = Decrypt("_venus13", Datos.Domain)
        Catch ex As Exception
            Log(True, "Error al leer parámetros, desencriptar usuario " & ex.Message)
            Leer_Parametros = False
            Exit Function
        End Try


    End Function
    Public Sub Log(DatosFecha As Boolean, ByVal texto As String)
        Dim sr As System.IO.StreamWriter

        Try
            sr = New System.IO.StreamWriter(Datos.Log, True)
            If DatosFecha Then
                sr.WriteLine(Datos.Version + " <" + Format(Now, "dd.MM.yy HH:mm:ss") & ">   " & texto)
            Else
                sr.WriteLine(Space(35) + texto)
            End If

            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Log_Activity(ByVal texto As String)
        Dim sr As System.IO.StreamWriter

        Try
            sr = New System.IO.StreamWriter(Datos.Log_Activity, True)
            sr.WriteLine(Datos.Version + " <" + Format(Now, "dd.MM.yy HH:mm:ss") & ">   " & texto)
            sr.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Private Function Buscar_Impresora(PrinterName As String) As Boolean
        Buscar_Impresora = False

        For Each Impresoras In PrinterSettings.InstalledPrinters
            If Impresoras.ToString = PrinterName Then
                Buscar_Impresora = True
                Exit For
            End If
        Next
    End Function


    Public Function Cargar_XML(ByRef ds As DataSet, xml As String) As Boolean
        Cargar_XML = True

        Try
            ds = New DataSet
            ds.ReadXml(New System.Xml.XmlTextReader(New StringReader(xml)))
        Catch ex As Exception
            Cargar_XML = False
            Log(True, "Error cargar XML: " + ex.Message)
        End Try
    End Function


    Public Function Decrypt(clave As String, cipherText As String) As String
        Dim EncryptionKey As String = clave
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
             &H65, &H64, &H76, &H65, &H64, &H65,
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function

    Private Function Cargar_Dataset(xml As String, ByRef ds As DataSet) As Boolean
        Cargar_Dataset = True

        If xml = "" Then
            Log(True, "Se ha recibido un xml vacío, Función enviar_PDF")
            Cargar_Dataset = False
        End If

        If Cargar_Dataset Then
            ds = Nothing
            Cargar_Dataset = Cargar_XML(ds, xml)

            If Not Cargar_Dataset Then
                Log(True, "El xml recibido no puede convertirse en un dataset")
                Log(False, xml)
            End If
        End If
    End Function

    Private Sub Enviar_PDF(xml As String)
        Dim TextoError As String = ""
        Dim ds As DataSet = Nothing

        If xml = "" Then Exit Sub
        If Not Cargar_Dataset(xml, ds) Then Exit Sub

        For Each dt As DataRow In ds.Tables(0).Rows
            If Not Buscar_Impresora(dt.Item("Printer")) Then
                Log(True, "La impresora recibida no existe: " + dt.Item("Printer"))
                Log(False, "Función Imprimir DOC")
                Log(False, "Id: " + dt.Item("Id").ToString)
                TextoError = "No se ha encontrado la impresora"
            Else
                Log_Activity("DOC: " + dt.Item("Origen").ToString)

                Dim PDFDoc As PDF = New PDF()
                PDFDoc.dt = dt
                PDFDoc.Ghostscript_PDF()
            End If

            Web.Resolver(dt.Item("Id"), TextoError)
        Next
    End Sub

    Private Sub Imprimir_Etiqueta(xml As String)
        Dim TextoError As String = ""
        Dim ds As DataSet = Nothing

        If xml = "" Then Exit Sub
        If Not Cargar_Dataset(xml, ds) Then Exit Sub

        For Each dt As DataRow In ds.Tables(0).Rows
            If Not Buscar_Impresora(dt.Item("Printer")) Then
                Log(True, "La impresora recibida no existe: " + dt.Item("Printer"))
                Log(False, "Función Etiquetas")
                Log(False, "Id: " + dt.Item("Id").ToString)
                TextoError = "No se ha encontrado la impresora"
            Else
                Dim Printer As New Impresora

                Printer.ds = ds
                Printer.Imprimir_etiquetas()
            End If
            ' REVISAR
            Web.Resolver(dt.Item("Id"), TextoError)
        Next
    End Sub

    Private Sub SaveFile(xml As String)
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("C:\Task-Printer\TMP\test.xml", True)
        file.Write(xml)
        file.Close()
    End Sub



    Private Function Pdf_Directo(dt As DataRow) As Boolean
        Dim Fichero As String = ""
        Dim Comillas As String = Chr(34)

        Pdf_Directo = True

        Fichero = Datos.Folder + "\TMP\" + dt.Item("PDFID") + ".pdf"

        Try
            Dim PDFProcess As New Process()

            PDFProcess.StartInfo.FileName = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
            PDFProcess.StartInfo.Arguments = "/N /T " + Comillas + Fichero + Comillas + " " + Comillas + dt.Item("Printer") + Comillas
            PDFProcess.StartInfo.UseShellExecute = True
            PDFProcess.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
            PDFProcess.StartInfo.Verb = "Print"
            PDFProcess.Start()
            PDFProcess.Close()

            Threading.Thread.Sleep(4000)
            For Each item As Process In Process.GetProcesses
                If item.ProcessName = "AcroRd32" Then
                    item.Kill()
                End If
            Next
        Catch ex As Exception
            Log(True, "Error Envío PDF: " + ex.Message)
            Log(False, "Fichero: " + Fichero)
            Log(False, "Impresora: " + dt.Item("Printer"))
            Pdf_Directo = False
        End Try

    End Function
End Module
