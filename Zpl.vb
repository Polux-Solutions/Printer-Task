Option Explicit On

Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing.Printing
Imports System.ComponentModel
Imports RawPrint


Public Class Zpl
    Dim Imagen As Image
    Dim dt As DataRow

    Public Sub Imprimir_Etiqueta()
        Dim Ruta As String = Datos.Ruta + "\output"


        If Not Directory.Exists(Ruta) Then
            Log("No Existe la carpeta: " + Ruta)
        Else
            For Each file As String In System.IO.Directory.GetFiles(Ruta, "*.xml", System.IO.SearchOption.TopDirectoryOnly)
                Me.Text = file
                Me.Refresh()

                Dim ds As DataSet = New DataSet

                If Cargar_XML(ds, file) Then
                    Try
                        If ds.Tables(0).Rows.Count >= 1 Then
                            Dim n As Integer = 0
                            For n = 0 To ds.Tables(0).Rows.Count - 1
                                dt = ds.Tables(0).Rows(n)

                                Log("Tipo Etiqueta: " + dt.Item("Tipo").ToString)

                                Etiqueta_Producto()
                            Next
                        End If
                End If
                Catch ex As Exception
                Log("Error lectura xml Fichero: " + ex.Message)
                End Try

                Dim destino As String = Ruta + "\Procesados\" + Path.GetFileName(file)
                Try
                    If System.IO.File.Exists(destino) Then System.IO.File.Delete(destino)
                    System.IO.File.Move(file, destino)
                Catch ex As Exception
                    Log("Error al mover a procesados: " + vbCrLf + "    Origen: " + file + vbCrLf + "   Destino: " + destino)
                End Try
            Next
        End If
    End Sub

    Private Sub Etiqueta_Producto()
        Dim NewPaperSize As Printing.PaperSize = New Printing.PaperSize("ZPL", 402, 295)

        PrintDoc.PrinterSettings.PrinterName = Datos.Impresora
        PrintDoc.DefaultPageSettings.Landscape = False
        PrintDoc.DefaultPageSettings.Margins.Left = 0
        PrintDoc.DefaultPageSettings.Margins.Top = 0
        PrintDoc.DefaultPageSettings.PaperSize = NewPaperSize

        PrintDoc.PrinterSettings.Copies = 1 ' dt.Item("Copias")

        Try
            'Imagen = Image.FromFile(Base + "\Logo.bmp")
        Catch ex As Exception
            Log("Error Carga Imagen Logo: " + ex.Message)
        End Try


        PrintDoc.Print()
    End Sub


    Private Sub Split_Texto(texto As String, largo As Integer, ByRef tt() As String)
        Dim x As Integer

        Try
            While True
                x = texto.IndexOf(" ")
                If x < 0 Then x = texto.Length - 1

                If IsNothing(tt) Then
                    ReDim tt(0)
                    tt(0) = ""
                End If
                If tt(tt.Length - 1).Length + x > largo Then
                    ReDim Preserve tt(tt.Length)
                    tt(tt.Length - 1) = ""
                End If

                tt(tt.Length - 1) += texto.Substring(0, x + 1)
                texto = texto.Substring(x + 1, texto.Length - (x + 1))

                If texto = "" Then Exit While
            End While
        Catch ex As Exception
            Log("Error Cortar Observaciones:  " + texto + vbCrLf + " Error: " + ex.Message)
            ReDim tt(0)
            tt(0) = texto
        End Try

    End Sub

    Private Sub PrintDoc_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDoc.PrintPage
        Dim fCaption As New System.Drawing.Font("Helvetica Bolt", 8)
        Dim fDato As New System.Drawing.Font("Helvetica", 11, FontStyle.Regular)
        Dim fDatoB As New System.Drawing.Font("Helvetica", 16, FontStyle.Regular)
        Dim fDatoP As New System.Drawing.Font("Helvetica", 9, FontStyle.Regular)
        Dim Pt As New System.Drawing.Point

        e.Graphics.PageUnit = GraphicsUnit.Millimeter

        Try
            Pt.X = 2
            Pt.Y = 2
            e.Graphics.DrawImage(Imagen, Pt.X, Pt.Y, 15, 15)
        Catch ex As Exception
            Log("Error Imagen Logo: " + ex.Message)
        End Try

        Try
            Pt.X = 72
            Pt.Y = 2
            e.Graphics.DrawImage(Imagen, Pt.X, Pt.Y, 26, 26)
        Catch ex As Exception
            Log("Error Imagen QR: " + ex.Message)
        End Try

        Pt.X = 20
        Pt.Y = 5
        e.Graphics.DrawString("Albarán: ", fCaption, Brushes.Black, Pt)
        Pt.X = 35
        Pt.Y = 3
        e.Graphics.DrawString(dt.Item("Albaran"), fDatoB, Brushes.Black, Pt)

        Pt.X = 20
        Pt.Y = 13
        e.Graphics.DrawString("Fecha: ", fCaption, Brushes.Black, Pt)
        Pt.X = 35
        Pt.Y = 11
        e.Graphics.DrawString(dt.Item("Fecha"), fDatoB, Brushes.Black, Pt)

        Pt.X = 3
        Pt.Y = 29
        e.Graphics.DrawString("Remitente: ", fCaption, Brushes.Black, Pt)
        Pt.X = 20
        Pt.Y = 28
        e.Graphics.DrawString(dt.Item("Remitente"), fDato, Brushes.Black, Pt)

        Pt.X = 3
        Pt.Y = 35
        e.Graphics.DrawString("Dir. Envío: ", fCaption, Brushes.Black, Pt)
        Pt.X = 20
        Pt.Y = 34
        e.Graphics.DrawString(dt.Item("DirEnvio"), fDatoP, Brushes.Black, Pt)
        Pt.X = 20
        Pt.Y = 39
        e.Graphics.DrawString(dt.Item("PobEnvio"), fDatoP, Brushes.Black, Pt)
        Pt.X = 20
        Pt.Y = 44
        e.Graphics.DrawString("(" + dt.Item("CPEnvio") + ")" + "    " + dt.Item("PrvEnvio") + "  -" + dt.Item("PaisEnvio") + "-", fDatoP, Brushes.Black, Pt)

        Pt.X = 3
        Pt.Y = 50
        e.Graphics.DrawString("Transpor.: ", fCaption, Brushes.Black, Pt)

        Dim tt1() As String
        Split_Texto(dt.Item("Transportista"), 25, tt1)

        For m = 0 To tt1.Length - 1
            Pt.X = 20
            Pt.Y = 49 + (m * 4)
            e.Graphics.DrawString(tt1(m), fDato, Brushes.Black, Pt)
        Next

        Pt.X = 3
        Pt.Y = Pt.Y + 5
        e.Graphics.DrawString("Cantidad: ", fCaption, Brushes.Black, Pt)

        Pt.X = 20
        e.Graphics.DrawString(Format(dt.Item("Cantidad")), fDato, Brushes.Black, Pt)

        Pt.X = 3
        Pt.Y = 58
        e.Graphics.DrawString("Observa.: ", fCaption, Brushes.Black, Pt)

        Dim tt2() As String
        Split_Texto(dt.Item("Observaciones"), 35, tt2)

        For m = 0 To tt2.Length - 1
            Pt.Y = 63 + (m * 5)
            e.Graphics.DrawString(tt2(m), fDato, Brushes.Black, Pt)
        Next

        Dim myPen As Pen
        myPen = New Pen(Drawing.Color.Black, 1)
        e.Graphics.DrawRectangle(myPen, New Rectangle(74, 50, 26, 22))

        Pt.X = 75
        Pt.Y = 54
        e.Graphics.DrawString("BULTOS ", fDatoB, Brushes.Black, Pt)

        Pt.X = 78
        Pt.Y = 62
        e.Graphics.DrawString(dt.Item("Bulto").ToString + " / " + dt.Item("Bultos").ToString, fDatoB, Brushes.Black, Pt)

        e.HasMorePages = False
    End Sub
End Class