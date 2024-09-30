Option Explicit On

Imports ThoughtWorks.QRCode
Imports ThoughtWorks.QRCode.Codec
Imports ThoughtWorks.QRCode.Codec.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports RawPrint
Imports DataMatrix
Imports DataMatrix.net

Class Impresora
    Public ds As DataSet

    Private PrintDoc As New PrintDocument
    Private dt As DataRow

    Sub Imprimir_etiquetas()
        Dim NewPaperSize As Printing.PaperSize = Nothing
        Dim LandScape As Boolean = False

        For n = 0 To ds.Tables(0).Rows.Count - 1
            NewPaperSize = Nothing
            dt = ds.Tables(0).Rows(n)

            LandScape = False
            Select Case dt.Item("Tipo").ToString.ToUpper
                Case "KIT"
                    Log_Activity("Etiqueta Kit: " + dt.Item("Origen").ToString)
                    NewPaperSize = New Printing.PaperSize("ZPL", 393, 595)
                    AddHandler PrintDoc.PrintPage, AddressOf Kit_PrintPage
                    LandScape = True
                Case "KIT GS1"
                    Log_Activity("Etiqueta Kit: " + dt.Item("Origen").ToString)
                    NewPaperSize = New Printing.PaperSize("ZPL", 393, 595)
                    AddHandler PrintDoc.PrintPage, AddressOf Kit_Gs1_PrintPage
                    LandScape = True
                Case "CLIENTES"
                    Log_Activity("Etiqueta Cliente: " + dt.Item("Origen").ToString)
                    NewPaperSize = New Printing.PaperSize("ZPL", 393, 196)
                    AddHandler PrintDoc.PrintPage, AddressOf Clientes_PrintPage
                    LandScape = False
                Case "LOTE"
                    Log_Activity("Etiqueta Lote: " + dt.Item("Origen").ToString)
                    NewPaperSize = New Printing.PaperSize("ZPL", 393, 196)
                    AddHandler PrintDoc.PrintPage, AddressOf Lote_PrintPage
                    LandScape = False
            End Select

            If Not IsNothing(NewPaperSize) Then
                PrintDoc.PrinterSettings.PrinterName = dt.Item("Printer")
                PrintDoc.DefaultPageSettings.Landscape = LandScape
                PrintDoc.DefaultPageSettings.Margins.Left = 0
                PrintDoc.DefaultPageSettings.Margins.Top = 0
                PrintDoc.DefaultPageSettings.PaperSize = NewPaperSize

                PrintDoc.PrinterSettings.Copies = dt.Item("Copias")

                PrintDoc.Print()
            End If
        Next
    End Sub


    Private Sub Kit_PrintPage(ByVal sender As System.Object, e As PrintPageEventArgs)
        Dim Fuente As System.Drawing.Font
        Dim Pt As New System.Drawing.Point
        Dim Talla As New SizeF
        Dim Raya As Pen = New Pen(Color.Black, 0.5)
        Dim Brocha As Brush = New SolidBrush(Color.GreenYellow)
        Dim Paginas As Integer = 1

        For Each dt2 As DataRow In dt.GetChildRows(ds.Tables(0).ChildRelations(0))
            e.Graphics.PageUnit = GraphicsUnit.Millimeter

            ' DATOS
            '--------

            Fuente = New System.Drawing.Font("Arial", 10, FontStyle.Regular)

            Pt.X = 5
            Pt.Y = 3
            e.Graphics.DrawString("Nombre:", Fuente, Brushes.Black, Pt)

            Pt.X = 22
            e.Graphics.DrawString(dt2.Item("Description"), Fuente, Brushes.Black, Pt)

            Pt.X = 3
            Pt.Y = 11
            e.Graphics.DrawString("Referencia:", Fuente, Brushes.Black, Pt)

            Pt.X = 22
            e.Graphics.DrawString(dt2.Item("ItemNo"), Fuente, Brushes.Black, Pt)

            Pt.X = 60
            e.Graphics.DrawString("Caducidad:", Fuente, Brushes.Black, Pt)

            e.Graphics.DrawRectangle(Raya, 85, 10, 62, 7)

            Pt.Y = 11
            Pt.X = 90 + Centrar_Texto(e, dt2.Item("ExpirationDate"), Fuente, 62)
            e.Graphics.DrawString(dt2.Item("ExpirationDate"), Fuente, Brushes.Black, Pt)

            Pt.Y = 18
            Pt.X = 60
            e.Graphics.DrawString("Lote:", Fuente, Brushes.Black, Pt)

            e.Graphics.DrawRectangle(Raya, 85, 17, 62, 7)
            Pt.Y = 18
            Pt.X = 90 + (Centrar_Texto(e, dt2.Item("LotNo"), Fuente, 62))
            e.Graphics.DrawString(dt2.Item("LotNo"), Fuente, Brushes.Black, Pt)

            ' Cabecera Tabla

            e.Graphics.DrawRectangle(Raya, 5, 24, 142, 7)
            e.Graphics.FillRectangle(Brocha, 5, 24, 142, 7)

            Fuente = New System.Drawing.Font("Arial", 10, FontStyle.Bold)
            Pt.Y = 26
            Pt.X = 7 + Centrar_Texto(e, "PRODUCTOS SANITARIOS INCLUIDOS", Fuente, 142)
            e.Graphics.DrawString("PRODUCTOS SANITARIOS INCLUIDOS", Fuente, Brushes.Black, Pt)

            e.Graphics.DrawRectangle(Raya, 5, 31, 17, 7)
            e.Graphics.DrawRectangle(Raya, 22, 31, 42, 7)
            e.Graphics.DrawRectangle(Raya, 64, 31, 30, 7)
            e.Graphics.DrawRectangle(Raya, 94, 31, 35, 7)
            e.Graphics.DrawRectangle(Raya, 129, 31, 18, 7)

            Fuente = New System.Drawing.Font("Arial", 7, FontStyle.Regular)
            Pt.Y = 33
            Pt.X = 5 + Centrar_Texto(e, "Referencia", Fuente, 16)
            e.Graphics.DrawString("Referencia", Fuente, Brushes.Black, Pt)
            Pt.X = 23 + Centrar_Texto(e, "Nombre del Producto", Fuente, 42)
            e.Graphics.DrawString("Nombre del Producto", Fuente, Brushes.Black, Pt)
            Pt.X = 64 + Centrar_Texto(e, "Fabricante", Fuente, 30)
            e.Graphics.DrawString("Fabricante", Fuente, Brushes.Black, Pt)
            Pt.X = 94 + Centrar_Texto(e, "Dirección", Fuente, 35)
            e.Graphics.DrawString("Dirección", Fuente, Brushes.Black, Pt)
            Pt.X = 129 + Centrar_Texto(e, "Marcado UE", Fuente, 18)
            e.Graphics.DrawString("Marcado UE", Fuente, Brushes.Black, Pt)

            ' Lineas
            Dim Linea As Integer = 31

            For Each dt3 As DataRow In dt2.GetChildRows(ds.Tables(1).ChildRelations(0))
                Kit_Linea(e, dt3, Linea)
            Next


            ' Pie
            Pt.Y = Linea + 7
            Pt.X = 6 + Centrar_Texto(e, "Agrupado por: CRIVEL, S.A. Calle Argualas s/n Polígono Industrial Argualas, nave 31 Zaragoza 50012", Fuente, 140)
            e.Graphics.DrawRectangle(Raya, 5, Pt.Y, 142, 7)
            Pt.Y += 2
            e.Graphics.DrawString("Agrupado por: CRIVEL, S.A. Calle Argualas s/n Polígono Industrial Argualas, nave 31 Zaragoza 50012", Fuente, Brushes.Black, Pt)


            e.HasMorePages = False
        Next

    End Sub


    Private Sub Kit_Gs1_PrintPage(ByVal sender As System.Object, e As PrintPageEventArgs)
        Dim Fuente As System.Drawing.Font
        Dim Pt As New System.Drawing.Point
        Dim Talla As New SizeF
        Dim Raya As Pen = New Pen(Color.Black, 0.5)
        Dim Brocha As Brush = New SolidBrush(Color.GreenYellow)
        Dim Paginas As Integer = 1

        For Each dt2 As DataRow In dt.GetChildRows(ds.Tables(0).ChildRelations(0))
            e.Graphics.PageUnit = GraphicsUnit.Millimeter

            ' DATOS
            '--------

            Fuente = New System.Drawing.Font("Arial", 10, FontStyle.Regular)

            Pt.X = 5
            Pt.Y = 3
            e.Graphics.DrawString("Nombre:", Fuente, Brushes.Black, Pt)

            Pt.X = 22
            e.Graphics.DrawString(dt2.Item("Description"), Fuente, Brushes.Black, Pt)

            Pt.X = 3
            Pt.Y = 11
            e.Graphics.DrawString("Referencia:", Fuente, Brushes.Black, Pt)

            Pt.X = 22
            e.Graphics.DrawString(dt2.Item("ItemNo"), Fuente, Brushes.Black, Pt)

            Pt.X = 60
            e.Graphics.DrawString("Caducidad:", Fuente, Brushes.Black, Pt)

            e.Graphics.DrawRectangle(Raya, 85, 10, 62, 7)

            Pt.Y = 11
            Pt.X = 90 + Centrar_Texto(e, dt2.Item("ExpirationDate"), Fuente, 62)
            e.Graphics.DrawString(dt2.Item("ExpirationDate"), Fuente, Brushes.Black, Pt)

            Pt.Y = 18
            Pt.X = 60
            e.Graphics.DrawString("Lote:", Fuente, Brushes.Black, Pt)

            e.Graphics.DrawRectangle(Raya, 85, 17, 62, 7)
            Pt.Y = 18
            Pt.X = 90 + (Centrar_Texto(e, dt2.Item("LotNo"), Fuente, 62))
            e.Graphics.DrawString(dt2.Item("LotNo"), Fuente, Brushes.Black, Pt)

            ' Cabecera Tabla

            e.Graphics.DrawRectangle(Raya, 5, 24, 142, 7)
            e.Graphics.FillRectangle(Brocha, 5, 24, 142, 7)

            Fuente = New System.Drawing.Font("Arial", 10, FontStyle.Bold)
            Pt.Y = 26
            Pt.X = 7 + Centrar_Texto(e, "PRODUCTOS SANITARIOS INCLUIDOS", Fuente, 142)
            e.Graphics.DrawString("PRODUCTOS SANITARIOS INCLUIDOS", Fuente, Brushes.Black, Pt)

            e.Graphics.DrawRectangle(Raya, 5, 31, 17, 7)
            e.Graphics.DrawRectangle(Raya, 22, 31, 42, 7)
            e.Graphics.DrawRectangle(Raya, 64, 31, 30, 7)
            e.Graphics.DrawRectangle(Raya, 94, 31, 35, 7)
            e.Graphics.DrawRectangle(Raya, 129, 31, 18, 7)

            Fuente = New System.Drawing.Font("Arial", 7, FontStyle.Regular)
            Pt.Y = 33
            Pt.X = 5 + Centrar_Texto(e, "Referencia", Fuente, 16)
            e.Graphics.DrawString("Referencia", Fuente, Brushes.Black, Pt)
            Pt.X = 23 + Centrar_Texto(e, "Nombre del Producto", Fuente, 42)
            e.Graphics.DrawString("Nombre del Producto", Fuente, Brushes.Black, Pt)
            Pt.X = 64 + Centrar_Texto(e, "Fabricante", Fuente, 30)
            e.Graphics.DrawString("Fabricante", Fuente, Brushes.Black, Pt)
            Pt.X = 94 + Centrar_Texto(e, "Dirección", Fuente, 35)
            e.Graphics.DrawString("Dirección", Fuente, Brushes.Black, Pt)
            Pt.X = 129 + Centrar_Texto(e, "Marcado UE", Fuente, 18)
            e.Graphics.DrawString("Marcado UE", Fuente, Brushes.Black, Pt)

            ' Lineas
            Dim Linea As Integer = 31

            For Each dt3 As DataRow In dt2.GetChildRows(ds.Tables(1).ChildRelations(0))
                Kit_Linea(e, dt3, Linea)
            Next


            ' Pie
            Pt.Y = Linea + 7
            Pt.X = 6 + Centrar_Texto(e, "Agrupado por: CRIVEL, S.A. Calle Argualas s/n Polígono Industrial Argualas, nave 31 Zaragoza 50012", Fuente, 140)
            e.Graphics.DrawRectangle(Raya, 5, Pt.Y, 142, 7)
            Pt.Y += 2
            e.Graphics.DrawString("Agrupado por: CRIVEL, S.A. Calle Argualas s/n Polígono Industrial Argualas, nave 31 Zaragoza 50012", Fuente, Brushes.Black, Pt)

            Dim Im128 As System.Drawing.Image
            Dim t2 As String = ""
            Dim FNC1 As String = Chr(29)
            Dim Fecha As String = dt2.Item("ExpirationDate")

            If Fecha.Length = 10 Then Fecha = Fecha.Substring(0, 2) + Fecha.Substring(3, 2) + Fecha.Substring(8, 2)

            t2 = "01" + dt2.Item("AECOC") + "17" + Fecha + "11" + Format(Now, "ddMMyy") + "10" + dt2.Item("LotNo")
            Im128 = Imagen_Datamatrix(t2)

            Try
                Pt.X = 6
                Pt.Y = Linea + 30
                e.Graphics.DrawImage(Im128, Pt.X, Pt.Y, 8, 8)
            Catch ex As Exception
                Log(True, "Error Imagen Logo: " + ex.Message)
            End Try

            Fuente = New System.Drawing.Font("Arial", 5, FontStyle.Regular)
            Pt.X = 16
            e.Graphics.DrawString($"(01) {dt2.Item("AECOC")}", Fuente, Brushes.Black, Pt)
            Pt.Y += 2
            e.Graphics.DrawString($"(17) {Fecha}", Fuente, Brushes.Black, Pt)
            Pt.Y += 2
            e.Graphics.DrawString($"(11) {Format(Now, "ddMMyy")}", Fuente, Brushes.Black, Pt)
            Pt.Y += 2
            e.Graphics.DrawString($"(10) {dt2.Item("LotNo")}", Fuente, Brushes.Black, Pt)

            e.HasMorePages = False
        Next

    End Sub

    Private Sub Kit_Linea(e As PrintPageEventArgs, dt As DataRow, ByRef Linea As Integer)
        Dim Fuente As System.Drawing.Font
        Dim Pt As New System.Drawing.Point
        Dim Talla As New SizeF
        Dim Raya As Pen = New Pen(Color.Black, 0.5)
        Dim Texto() As String = Nothing

        Linea += 7
        e.Graphics.DrawRectangle(Raya, 5, Linea, 17, 7)
        e.Graphics.DrawRectangle(Raya, 22, Linea, 42, 7)
        e.Graphics.DrawRectangle(Raya, 64, Linea, 30, 7)
        e.Graphics.DrawRectangle(Raya, 94, Linea, 35, 7)
        e.Graphics.DrawRectangle(Raya, 129, Linea, 18, 7)

        Fuente = New System.Drawing.Font("Arial", 7, FontStyle.Regular)
        Pt.Y = Linea + 2
        Pt.X = 6
        e.Graphics.DrawString(dt.Item("Sub_ItemNo"), Fuente, Brushes.Black, Pt)
        Fuente = New System.Drawing.Font("Arial", 7, FontStyle.Regular)

        Trocear_Texto(e, dt.Item("Sub_Description"), Fuente, 40, Texto)
        Pt.Y = Linea + 1
        Pt.X = 23
        e.Graphics.DrawString(Texto(0), Fuente, Brushes.Black, Pt)
        If Texto.Length > 1 Then
            Pt.Y = Pt.Y + 3
            e.Graphics.DrawString(Texto(1), Fuente, Brushes.Black, Pt)
        End If

        Trocear_Texto(e, dt.Item("Sub_Fabricante"), Fuente, 30, Texto)
        Pt.Y = Linea + 1
        Pt.X = 65
        e.Graphics.DrawString(Texto(0), Fuente, Brushes.Black, Pt)
        If Texto.Length > 1 Then
            Pt.Y = Pt.Y + 3
            e.Graphics.DrawString(Texto(1), Fuente, Brushes.Black, Pt)
        End If

        Trocear_Texto(e, dt.Item("Sub_Fabricante_Dir"), Fuente, 35, Texto)
        Pt.Y = Linea + 1
        Pt.X = 95
        e.Graphics.DrawString(Texto(0), Fuente, Brushes.Black, Pt)
        If Texto.Length > 1 Then
            Pt.Y = Pt.Y + 3
            e.Graphics.DrawString(Texto(1), Fuente, Brushes.Black, Pt)
        End If

        Pt.Y = Linea + 1
        Pt.X = 130
        e.Graphics.DrawString("Sí", Fuente, Brushes.Black, Pt)
        Pt.Y = Pt.Y + 3
        e.Graphics.DrawString("Automarcado", Fuente, Brushes.Black, Pt)

        e.HasMorePages = False
    End Sub


    Private Sub Clientes_PrintPage(ByVal sender As System.Object, e As PrintPageEventArgs)
        Dim Fuente As System.Drawing.Font
        Dim Pt As New System.Drawing.Point
        Dim Brocha As Brush = Brushes.Black
        Dim ImagenLogo As System.Drawing.Image

        e.Graphics.PageUnit = GraphicsUnit.Millimeter

        Try
            ImagenLogo = System.Drawing.Image.FromFile(Datos.Folder + "\Logos\Crivelsa.jpg")
        Catch ex As Exception
            Log(True, "Error Carga Imagen Logo: " + ex.Message)
        End Try

        Try
            Pt.X = 2
            Pt.Y = 2
            e.Graphics.DrawImage(ImagenLogo, Pt.X, Pt.Y, 32, 12)
        Catch ex As Exception
            Log(True, "Error Imagen Logo: " + ex.Message)
        End Try

        For Each dt2 As DataRow In dt.GetChildRows(ds.Tables(0).ChildRelations(0))
            Fuente = New System.Drawing.Font("Arial", 9, FontStyle.Regular)

            Pt.X = 23 + Centrar_Texto(e, dt2.Item("Address"), Fuente, 70)
            Pt.Y = 2

            e.Graphics.DrawString(dt2.Item("Address"), Fuente, Brocha, Pt)
            Pt.X = 23 + Centrar_Texto(e, dt2.Item("Phone") + "  " + dt2.Item("PostCode"), Fuente, 70)
            Pt.Y = 7
            e.Graphics.DrawString(dt2.Item("Phone") + "  " + dt2.Item("PostCode"), Fuente, Brocha, Pt)
            Pt.X = 23 + Centrar_Texto(e, dt2.Item("City"), Fuente, 70)
            Pt.Y = 11
            e.Graphics.DrawString(dt2.Item("City"), Fuente, Brocha, Pt)

            For Each dt3 As DataRow In dt.GetChildRows(ds.Tables(0).ChildRelations(1))

                Pt.X = 2
                Pt.Y = 15
                e.Graphics.DrawString("ATT:", Fuente, Brocha, Pt)

                Pt.X = 2
                Pt.Y = 19
                e.Graphics.DrawString("Destinatario: ", Fuente, Brocha, Pt)

                Dim Texto() As String = Nothing

                Trocear_Texto(e, dt3.Item("CustomerName"), Fuente, 70, Texto)
                If Not IsNothing(Texto) Then
                    Pt.Y = 19
                    Pt.X = 22
                    e.Graphics.DrawString(Texto(0), Fuente, Brocha, Pt)

                    If Texto.Length > 1 Then
                        Pt.Y += 4
                        e.Graphics.DrawString(Texto(1), Fuente, Brocha, Pt)
                    End If
                End If

                Pt.X = 2
                Pt.Y += 4
                e.Graphics.DrawString("Domicilio: ", Fuente, Brocha, Pt)

                Trocear_Texto(e, dt3.Item("CustomerAddress"), Fuente, 70, Texto)
                If Not IsNothing(Texto) Then
                    Pt.X = 22
                    e.Graphics.DrawString(Texto(0), Fuente, Brocha, Pt)

                    If Texto.Length > 1 Then
                        Pt.Y += 4
                        e.Graphics.DrawString(Texto(1), Fuente, Brocha, Pt)
                    End If
                End If

                Pt.X = 2
                Pt.Y += 4
                e.Graphics.DrawString("Localidad: ", Fuente, Brocha, Pt)

                Pt.X = 22
                e.Graphics.DrawString(dt3.Item("CustomerPostCode") + "   " + dt3.Item("CustomerCity"), Fuente, Brocha, Pt)

                Pt.X = 2
                Pt.Y += 4
                e.Graphics.DrawString("Provincia: ", Fuente, Brocha, Pt)

                Pt.X = 22
                e.Graphics.DrawString(dt3.Item("CustomerCounty"), Fuente, Brocha, Pt)
            Next

            Pt.X = 2
            Pt.Y = 43
            e.Graphics.DrawString("Nº de Bultos: " + dt.Item("Copias").ToString + "    Mercancía Material Médico", Fuente, Brocha, Pt)


            e.HasMorePages = False
        Next

    End Sub


    Private Sub Lote_PrintPage(ByVal sender As System.Object, e As PrintPageEventArgs)
        Dim Fuente As System.Drawing.Font
        Dim Pt As New System.Drawing.Point
        Dim Brocha As Brush = Brushes.Black

        e.Graphics.PageUnit = GraphicsUnit.Millimeter

        Fuente = New System.Drawing.Font("Arial", 9, FontStyle.Regular)

        Pt.X = 2
        Pt.Y = 2
        e.Graphics.DrawString("Referencia:", Fuente, Brocha, Pt)

        Pt.X = 23
        e.Graphics.DrawString(dt.Item("Referencia"), Fuente, Brocha, Pt)

        Pt.X = 2
        Pt.Y = 6
        e.Graphics.DrawString("Lote/Serie:", Fuente, Brocha, Pt)

        Pt.X = 23
        e.Graphics.DrawString(dt.Item("Lote"), Fuente, Brocha, Pt)

        Pt.X = 2
        Pt.Y = 10
        e.Graphics.DrawString("F.Caducidad:", Fuente, Brocha, Pt)

        Dim t As String = dt.Item("Caducidad")
        If t.Length = 8 Then t = t.Substring(0, 2) + "-" + t.Substring(2, 2) + "-" + t.Substring(4, 4)

        Pt.X = 23
        e.Graphics.DrawString(t, Fuente, Brocha, Pt)

        Dim Im128 As System.Drawing.Image

        Dim t2 As String = "(01)" + dt.Item("Referencia") + "(10)" + dt.Item("Lote") + "(17)" + dt.Item("Caducidad")
        Im128 = Imagen_QR(t2)

        Try
            Pt.X = 46
            Pt.Y = 2
            e.Graphics.DrawImage(Im128, Pt.X, Pt.Y, 15, 15)
        Catch ex As Exception
            Log(True, "Error Imagen Logo: " + ex.Message)
        End Try

        e.HasMorePages = False

    End Sub

#Region "Imágenes"
    Private Function Imagen_Barcode128(valor As String) As System.Drawing.Image
        Dim Bc128 As New Barcode128

        Imagen_Barcode128 = Nothing
        If valor = "" Then Exit Function

        Try
            Bc128.Code = valor
            Imagen_Barcode128 = Bc128.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White)
        Catch ex As Exception
            Imagen_Barcode128 = Nothing
        End Try
    End Function
    Private Function Imagen_Barcode39(valor As String) As System.Drawing.Image
        Dim Bc39 As New Barcode39

        Imagen_Barcode39 = Nothing
        If valor = "" Then Exit Function

        Try
            Bc39.Code = valor
            Imagen_Barcode39 = Bc39.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White)
        Catch ex As Exception
            Imagen_Barcode39 = Nothing
        End Try
    End Function

    Private Function Imagen_QR(valor As String) As System.Drawing.Image
        Try
            Dim generarCodigoQR As QRCodeEncoder = New QRCodeEncoder

            generarCodigoQR.QRCodeEncodeMode = Codec.QRCodeEncoder.ENCODE_MODE.BYTE
            generarCodigoQR.QRCodeScale = 6
            generarCodigoQR.QRCodeErrorCorrect = Codec.QRCodeEncoder.ERROR_CORRECTION.H
            generarCodigoQR.QRCodeVersion = 0

            Imagen_QR = generarCodigoQR.Encode(valor, System.Text.Encoding.UTF8)
        Catch ex As Exception
            Log(True, "Error Generar QR Lote: " + ex.Message)
        End Try
    End Function

    Private Function Imagen_Datamatrix(valor As String) As System.Drawing.Image
        Try
            Dim Encoder As DataMatrix.net.DmtxImageEncoder = New DataMatrix.net.DmtxImageEncoder
            Dim Options As DataMatrix.net.DmtxImageEncoderOptions = New DataMatrix.net.DmtxImageEncoderOptions

            Options.ModuleSize = 5
            Options.MarginSize = 2
            Options.BackColor = Color.White
            Options.ForeColor = Color.Black
            Options.Scheme = DmtxScheme.DmtxSchemeAsciiGS1
            Imagen_Datamatrix = Encoder.EncodeImage(valor, Options)
        Catch ex As Exception
            Log(True, "Error Generar Imagen Datamatrix: " + ex.Message)
        End Try
    End Function



    Private Function Base64StringToBitmap(base64string As String) As Bitmap
        Base64StringToBitmap = Nothing

        If base64string = "" Then Exit Function
        Dim bytebuffer() As Byte = Convert.FromBase64String(base64string)
        Dim MemorySt As MemoryStream = New MemoryStream(bytebuffer)
        MemorySt.Position = 0

        Base64StringToBitmap = Bitmap.FromStream(MemorySt)
        MemorySt.Close()
        MemorySt = Nothing
        bytebuffer = Nothing
    End Function


    Public Function ResizeImage(img As System.Drawing.Image, width As Integer, height As Integer) As System.Drawing.Image
        Dim newImage = New Bitmap(width, height)
        Using gr = Graphics.FromImage(newImage)
            gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
            gr.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            gr.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
            gr.DrawImage(img, New System.Drawing.Rectangle(0, 0, width, height))
        End Using
        Return newImage
    End Function

    Public Function ResizeImage(img As System.Drawing.Image, size As Size) As System.Drawing.Image
        Return ResizeImage(img, size.Width, size.Height)
    End Function

    Public Function ResizeImage(bmp As Bitmap, width As Integer, height As Integer) As System.Drawing.Image
        Return ResizeImage(DirectCast(bmp, System.Drawing.Image), width, height)
    End Function

    Public Function ResizeImage(bmp As Bitmap, size As Size) As System.Drawing.Image
        Return ResizeImage(DirectCast(bmp, System.Drawing.Image), size.Width, size.Height)
    End Function
#End Region

#Region "Formatos"

    Private Sub Trocear_Texto(e As PrintPageEventArgs, tt As String, fuente As System.Drawing.Font, ancho As Integer, ByRef Lineas() As String)
        Dim Talla As SizeF
        Dim palabras() As String

        ReDim Lineas(0)

        palabras = tt.Split(" ")
        Lineas(0) = palabras(0)
        If palabras.Length <= 1 Then Exit Sub

        For i = 1 To palabras.Length - 1
            Talla = e.Graphics.MeasureString(Lineas(Lineas.Length - 1) + " " + palabras(i), fuente)
            If Talla.Width <= ancho Then
                Lineas(Lineas.Length - 1) += " " + palabras(i)
            Else
                ReDim Preserve Lineas(Lineas.Length)
                Lineas(Lineas.Length - 1) = palabras(i)
            End If
        Next
    End Sub

    Private Function Centrar_Texto(e As PrintPageEventArgs, tt As String, fuente As System.Drawing.Font, ancho As Integer) As Integer
        Dim Talla As SizeF


        Talla = e.Graphics.MeasureString(tt, fuente)

        Centrar_Texto = (ancho - Talla.Width) / 2
        If Centrar_Texto < 0 Then Centrar_Texto = 0
    End Function
#End Region


End Class
