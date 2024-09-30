Option Explicit On
Public Class code39
    Private bitsCode As ArrayList

    Public Sub New()
        bitsCode = New ArrayList
        bitsCode.Add(New String(3) {"0001101", "0100111", "1110010", "000000"})
        bitsCode.Add(New String(3) {"0011001", "0110011", "1100110", "001011"})
        bitsCode.Add(New String(3) {"0010011", "0011011", "1101100", "001101"})
        bitsCode.Add(New String(3) {"0111101", "0100001", "1000010", "001110"})
        bitsCode.Add(New String(3) {"0100011", "0011101", "1011100", "010011"})
        bitsCode.Add(New String(3) {"0110001", "0111001", "1001110", "011001"})
        bitsCode.Add(New String(3) {"0101111", "0000101", "1010000", "011100"})
        bitsCode.Add(New String(3) {"0111011", "0010001", "1000100", "010101"})
        bitsCode.Add(New String(3) {"0110111", "0001001", "1001000", "010110"})
        bitsCode.Add(New String(3) {"0001011", "0010111", "1110100", "011010"})
    End Sub

    Public Function Generate(ByVal Code As String) As Image
        Dim a As Integer = 0
        Dim b As Integer = 0
        Dim imgCode As Image
        Dim g As Graphics
        Dim i As Integer
        Dim bCode As Byte()
        Dim bitCode As Byte()
        Dim tmpFont As Font

        If Code.Length <> 12 Or Not IsNumeric(Code.Replace(".", "_").Replace(",", "_")) Then Throw New Exception("Le code doit être composé de 12 chiffres")

        ReDim bCode(12)
        For i = 0 To 11
            bCode(i) = CInt(Code.Substring(i, 1))
            If (i Mod 2) = 1 Then
                b += bCode(i)
            Else
                a += bCode(i)
            End If
        Next

        i = (a + (b * 3)) Mod 10
        If i = 0 Then
            bCode(12) = 0
        Else
            bCode(12) = 10 - i
        End If
        bitCode = getBits(bCode)

        tmpFont = New Font("times new roman", 14, FontStyle.Regular, GraphicsUnit.Pixel)
        imgCode = New Bitmap(110, 50)
        g = Graphics.FromImage(imgCode)
        g.Clear(Color.White)

        g.DrawString(Code.Substring(0, 1), tmpFont, Brushes.Black, 2, 30)
        a = g.MeasureString(Code.Substring(0, 1), tmpFont).Width

        For i = 0 To bitCode.Length - 1
            If i = 2 Then
                g.DrawString(Code.Substring(1, 6), tmpFont, Brushes.Black, a, 30)
            ElseIf i = 48 Then
                g.DrawString(Code.Substring(7, 5) & bCode(12).ToString, tmpFont, Brushes.Black, a, 30)
            End If

            If i = 0 Or i = 2 Or i = 46 Or i = 48 Or i = 92 Or i = 94 Then
                If bitCode(i) = 1 Then 'noir
                    g.DrawLine(Pens.Black, a, 0, a, 40)
                    a += 1
                End If
            Else
                If bitCode(i) = 1 Then 'noir
                    g.DrawLine(Pens.Black, a, 0, a, 30)
                    a += 1
                Else 'blanc
                    a += 1
                End If
            End If
        Next
        g.Flush()
        Return imgCode
    End Function

    Private Function getBits(ByVal bCode As Byte()) As Byte()
        Dim i As Integer
        Dim res As Byte()
        Dim bits As String = "101"
        Dim cle As String = bitsCode(bCode(0))(3)
        For i = 1 To 6
            bits &= bitsCode(bCode(i))(CInt(cle.Substring(i - 1, 1)))
        Next
        bits &= "01010"
        For i = 7 To 12
            bits &= bitsCode(bCode(i))(2)
        Next
        bits += "101"
        ReDim res(bits.Length - 1)
        For i = 0 To bits.Length - 1
            res(i) = Asc(bits.Chars(i)) - 48
        Next
        Return res
    End Function

End Class