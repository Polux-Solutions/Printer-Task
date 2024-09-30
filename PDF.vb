Imports Ghostscript.NET.Processor

Public Class PDF
    Public dt As DataRow

    Public Sub Ghostscript_PDF()
        Try
            Dim Fichero As String = Datos.Folder + "\TMP\" + dt.Item("PDFID") + ".pdf"
            Dim processor As GhostscriptProcessor = New GhostscriptProcessor()
            Dim switches As List(Of String) = New List(Of String)

            switches.Add("-empty")
            switches.Add("-dPrinted")
            switches.Add("-dBATCH")
            switches.Add("-dNOPAUSE")
            switches.Add("-dNOSAFER")
            switches.Add("-dPDFFitPage")
            switches.Add("-dNumCopies=1")
            switches.Add("-sDEVICE=mswinpr2")
            switches.Add(Convert.ToString("-sOutputFile=%printer%") + dt.Item("Printer"))
            switches.Add("-f")
            switches.Add(Fichero)
            processor.StartProcessing(switches.ToArray(), Nothing)
        Catch ex As Exception
            Log(True, "Error impresión PDF: " + ex.Message)
            Log(False, dt.Item("PDFID"))
            Log(False, "Impresora: " + dt.Item("Printer"))
        End Try

    End Sub


End Class
