Option Explicit On
Option Strict On

Namespace Barcodes
    Public Class Barcode39
        Private Const WIDEBAR_WIDTH As Short = 2
        Private Const NARROWBAR_WIDTH As Short = 1
        Private Const NUM_CHARACTERS As Integer = 43

        Private mEncoding As Hashtable = New Hashtable
        Dim mCodeValue(NUM_CHARACTERS - 1) As Char

        'Additional properties 
        Public ShowString As Boolean
        Public IncludeCheckSumDigit As Boolean
        Public TextFont As New Font("Courier New", 7)
        Public TextColor As Color = Color.Black

        Public Sub New()
            '        Character, symbol
            mEncoding.Add("*", "bWbwBwBwb")
            mEncoding.Add("-", "bWbwbwBwB")
            mEncoding.Add("$", "bWbWbWbwb")
            mEncoding.Add("%", "bwbWbWbWb")
            mEncoding.Add(" ", "bWBwbwBwb")
            mEncoding.Add(".", "BWbwbwBwb")
            mEncoding.Add("/", "bWbWbwbWb")
            mEncoding.Add("+", "bWbwbWbWb")
            mEncoding.Add("0", "bwbWBwBwb")
            mEncoding.Add("1", "BwbWbwbwB")
            mEncoding.Add("2", "bwBWbwbwB")
            mEncoding.Add("3", "BwBWbwbwb")
            mEncoding.Add("4", "bwbWBwbwB")
            mEncoding.Add("5", "BwbWBwbwb")
            mEncoding.Add("6", "bwBWBwbwb")
            mEncoding.Add("7", "bwbWbwBwB")
            mEncoding.Add("8", "BwbWbwBwb")
            mEncoding.Add("9", "bwBWbwBwb")
            mEncoding.Add("A", "BwbwbWbwB")
            mEncoding.Add("B", "bwBwbWbwB")
            mEncoding.Add("C", "BwBwbWbwb")
            mEncoding.Add("D", "bwbwBWbwB")
            mEncoding.Add("E", "BwbwBWbwb")
            mEncoding.Add("F", "bwBwBWbwb")
            mEncoding.Add("G", "bwbwbWBwB")
            mEncoding.Add("H", "BwbwbWBwb")
            mEncoding.Add("I", "bwBwbWBwb")
            mEncoding.Add("J", "bwbwBWBwb")
            mEncoding.Add("K", "BwbwbwbWB")
            mEncoding.Add("L", "bwBwbwbWB")
            mEncoding.Add("M", "BwBwbwbWb")
            mEncoding.Add("N", "bwbwBwbWB")
            mEncoding.Add("O", "BwbwBwbWb")
            mEncoding.Add("P", "bwBwBwbWb")
            mEncoding.Add("Q", "bwbwbwBWB")
            mEncoding.Add("R", "BwbwbwBWb")
            mEncoding.Add("S", "bwBwbwBWb")
            mEncoding.Add("T", "bwbwBwBWb")
            mEncoding.Add("U", "BWbwbwbwB")
            mEncoding.Add("V", "bWBwbwbwB")
            mEncoding.Add("W", "BWBwbwbwb")
            mEncoding.Add("X", "bWbwBwbwB")
            mEncoding.Add("Y", "BWbwBwbwb")
            mEncoding.Add("Z", "bWBwBwbwb")

            mCodeValue(0) = "0"c
            mCodeValue(1) = "1"c
            mCodeValue(2) = "2"c
            mCodeValue(3) = "3"c
            mCodeValue(4) = "4"c
            mCodeValue(5) = "5"c
            mCodeValue(6) = "6"c
            mCodeValue(7) = "7"c
            mCodeValue(8) = "8"c
            mCodeValue(9) = "9"c
            mCodeValue(10) = "A"c
            mCodeValue(11) = "B"c
            mCodeValue(12) = "C"c
            mCodeValue(13) = "D"c
            mCodeValue(14) = "E"c
            mCodeValue(15) = "F"c
            mCodeValue(16) = "G"c
            mCodeValue(17) = "H"c
            mCodeValue(18) = "I"c
            mCodeValue(19) = "J"c
            mCodeValue(20) = "K"c
            mCodeValue(21) = "L"c
            mCodeValue(22) = "M"c
            mCodeValue(23) = "N"c
            mCodeValue(24) = "O"c
            mCodeValue(25) = "P"c
            mCodeValue(26) = "Q"c
            mCodeValue(27) = "R"c
            mCodeValue(28) = "S"c
            mCodeValue(29) = "T"c
            mCodeValue(30) = "U"c
            mCodeValue(31) = "V"c
            mCodeValue(32) = "W"c
            mCodeValue(33) = "X"c
            mCodeValue(34) = "Y"c
            mCodeValue(35) = "Z"c
            mCodeValue(36) = "-"c
            mCodeValue(37) = "."c
            mCodeValue(38) = " "c
            mCodeValue(39) = "$"c
            mCodeValue(40) = "/"c
            mCodeValue(41) = "+"c
            mCodeValue(42) = "%"c
        End Sub

        Public Function GenerateBarcodeImage(ByVal ImageWidth As Integer,
                                             ByVal ImageHeight As Integer,
                                             ByVal OriginalString As String) As Image

            '-- create a image where to paint the bars
            Dim pb As PictureBox
            pb = New PictureBox
            With pb
                .Width = ImageWidth
                .Height = ImageHeight
                pb.Image = New Bitmap(.Width, .Height)
            End With
            '---------------------

            'clear the image and set it to white background
            Dim g As Graphics = Graphics.FromImage(pb.Image)
            g.Clear(Color.White)


            'get the extended string
            Dim ExtString As String
            ExtString = ExtendedString(OriginalString)


            '-- This part format the sring that will be encoded
            '-- The string needs to be surrounded by asterisks 
            '-- to make it a valid Code39 barcode
            Dim EncodedString As String
            Dim ChkSum As Integer
            If IncludeCheckSumDigit = False Then
                EncodedString = String.Format("{0}{1}{0}", "*", ExtString)
            Else
                ChkSum = CheckSum(ExtString)

                EncodedString = String.Format("{0}{1}{2}{0}",
                                              "*", ExtString, mCodeValue(ChkSum))
            End If
            '----------------------

            '-- write the original string at the bottom if ShowString = True
            Dim textBrush As New SolidBrush(TextColor)
            If ShowString Then
                If Not IsNothing(TextFont) Then
                    'calculates the height of the string
                    Dim H As Single = g.MeasureString(OriginalString, TextFont).Height
                    g.DrawString(OriginalString, TextFont, textBrush, 0, ImageHeight - H)
                    ImageHeight = ImageHeight - CShort(H)
                End If
            End If
            '----------------------------------------

            'THIS IS WHERE THE BARCODE DRAWING HAPPENS
            DrawBarcode(g, EncodedString, ImageHeight)

            'IMAGE OBJECT IS RETURNED
            Return pb.Image


        End Function

        Private Sub DrawBarcode(ByVal g As Graphics,
                                ByVal EncodedString As String,
                                ByVal Height As Integer)

            'Start drawing at 0, 0
            Dim XPosition As Short = 0
            Dim YPosition As Short = 0

            'Dim invalidCharacter As Boolean = False
            Dim CurrentSymbol As String = String.Empty
            Dim EncodedSymbol As String
            '-- draw the bars
            For j As Short = 0 To CShort(EncodedString.Length - 1)
                CurrentSymbol = EncodedString.Chars(j)
                EncodedSymbol = mEncoding(CurrentSymbol).ToString

                For i As Short = 0 To CShort(EncodedSymbol.Length - 1)
                    'Dim CurrentCode As String = EncodedSymbol.Substring(i, 1)
                    Dim CurrentCode As Char = EncodedSymbol.Chars(i)

                    g.FillRectangle(getBCSymbolColor(CurrentCode), XPosition, YPosition, getBCSymbolWidth(CurrentCode), Height)

                    XPosition = XPosition + getBCSymbolWidth(CurrentCode)
                Next

                'After each written full symbol we need a whitespace (narrow width)
                g.FillRectangle(getBCSymbolColor("w"c), XPosition, YPosition, getBCSymbolWidth("w"c), Height)
                XPosition = XPosition + getBCSymbolWidth("w"c)

            Next
            '--------------------------


        End Sub


        Private Function getBCSymbolColor(ByVal symbol As Char) As System.Drawing.Brush
            If symbol = "W"c Or symbol = "w"c Then
                Return Brushes.White
            Else
                Return Brushes.Black
            End If
        End Function

        Private Function getBCSymbolWidth(ByVal symbol As Char) As Short
            If symbol = "B"c Or symbol = "W"c Then
                Return WIDEBAR_WIDTH
            Else
                Return NARROWBAR_WIDTH
            End If
        End Function


        Private Function CheckSum(ByVal sCode As String) As Integer
            Dim CurrentSymbol As Char
            Dim Chk As Integer
            For j As Integer = 0 To sCode.Length - 1
                CurrentSymbol = sCode.Chars(j)
                Chk += GetSymbolValue(CurrentSymbol)
            Next
            Return Chk Mod (NUM_CHARACTERS)
        End Function

        Private Function GetSymbolValue(ByVal s As Char) As Integer
            Dim k As Integer

            For k = 0 To NUM_CHARACTERS - 1
                If mCodeValue(k) = s Then
                    Return k
                End If
            Next
            Return Nothing
        End Function


        Private Function ExtendedString(ByVal s As String) As String
            Dim Ch As Char
            Dim KeyChar As Integer
            Dim retVal As String = ""

            For Each Ch In s
                KeyChar = Asc(Ch)
                Select Case KeyChar
                    Case 0
                        retVal &= "%U"
                    Case 1 To 26
                        retVal &= "$" & Chr(64 + KeyChar)
                    Case 27 To 31
                        retVal &= "%" & Chr(65 - 27 + KeyChar)
                    Case 33 To 44
                        retVal &= "/" & Chr(65 - 33 + KeyChar)
                    Case 47
                        retVal &= "/O"
                    Case 58
                        retVal &= "/Z"
                    Case 59 To 63
                        retVal &= "%" & Chr(70 - 59 + KeyChar)
                    Case 64
                        retVal &= "%V"
                    Case 91 To 95
                        retVal &= "%" & Chr(75 - 91 + KeyChar)
                    Case 96
                        retVal &= "%W"
                    Case 97 To 122
                        retVal &= "+" & Chr(65 - 97 + KeyChar)
                    Case 123 To 127
                        retVal &= "%" & Chr(80 - 123 + KeyChar)
                    Case Else
                        retVal &= Ch
                End Select

            Next
            Return retVal

        End Function
    End Class

End Namespace