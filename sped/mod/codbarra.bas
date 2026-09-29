Attribute VB_Name = "codbarra"

Sub DrawBarcode(ByVal bc_string As String, obj As Control, _
    nome As String, Preco As String, _
    Optional bPrintNum As Boolean = True, _
    Optional spaceBetweenBars As Integer = 40)

Dim xpos!, y1!, y2!, DW%, th!, tw, new_string$

'define barcode patterns
Dim bc(90) As String
bc(1) = "1 1221"            'pre-amble
bc(2) = "1 1221"            'post-amble
bc(48) = "11 221"           'digits
bc(49) = "21 112"
bc(50) = "12 112"
bc(51) = "22 111"
bc(52) = "11 212"
bc(53) = "21 211"
bc(54) = "12 211"
bc(55) = "11 122"
bc(56) = "21 121"
bc(57) = "12 121"
                            'capital letters
bc(65) = "211 12"           'A
bc(66) = "121 12"           'B
bc(67) = "221 11"           'C
bc(68) = "112 12"           'D
bc(69) = "212 11"           'E
bc(70) = "122 11"           'F
bc(71) = "111 22"           'G
bc(72) = "211 21"           'H
bc(73) = "121 21"           'I
bc(74) = "112 21"           'J
bc(75) = "2111 2"           'K
bc(76) = "1211 2"           'L
bc(77) = "2211 1"           'M
bc(78) = "1121 2"           'N
bc(79) = "2121 1"           'O
bc(80) = "1221 1"           'P
bc(81) = "1112 2"           'Q
bc(82) = "2112 1"           'R
bc(83) = "1212 1"           'S
bc(84) = "1122 1"           'T
bc(85) = "2 1112"           'U
bc(86) = "1 2112"           'V
bc(87) = "2 2111"           'W
bc(88) = "1 1212"           'X
bc(89) = "2 1211"           'Y
bc(90) = "1 2211"           'Z
                            'Misc
bc(32) = "1 2121"           'space
bc(35) = ""                 '# cannot do!
bc(36) = "1 1 1 11"         '$
bc(37) = "11 1 1 1"         '%
bc(43) = "1 11 1 1"         '+
bc(45) = "1 1122"           '-
bc(47) = "1 1 11 1"         '/
bc(46) = "2 1121"           '.
bc(64) = ""                 '@ cannot do!
bc(65) = "1 1221"           '*



bc_string = UCase(bc_string)


'dimensions
obj.ScaleMode = 3                               'pixels
obj.Cls
obj.Picture = Nothing
DW = CInt(obj.ScaleHeight / spaceBetweenBars)                  'space between bars
If DW < 1 Then DW = 1
'Debug.Print dw
th = obj.TextHeight(bc_string)                  'text height
tw = obj.TextWidth(bc_string)                   'text width
new_string = Chr$(1) & bc_string & Chr$(2)      'add pre-amble, post-amble

y1 = obj.ScaleTop
'y1 = 12
y2 = obj.ScaleTop + obj.ScaleHeight '- 1.5 * th 'mudei de 1.5 para 1
obj.width = 1.1 * Len(new_string) * (15 * DW) * obj.width / obj.ScaleWidth
'obj.FontSize = 8
'P_ImprX obj, Left(nome, 40), 0, False

'draw each character in barcode string

xpos = obj.ScaleLeft
For n = 1 To Len(new_string)
    c = Asc(Mid$(new_string, n, 1))
    If c > 90 Then c = 0
    bc_pattern$ = bc(c)
    
    'draw each bar
    For i = 1 To Len(bc_pattern$)
        Select Case Mid$(bc_pattern$, i, 1)
            Case " "
                'space
                obj.Line (xpos, y1)-(xpos + 1 * DW, y2), &HFFFFFF, BF
                xpos = xpos + DW
                
            Case "1"
                'space
                obj.Line (xpos, y1)-(xpos + 1 * DW, y2), &HFFFFFF, BF
                xpos = xpos + DW
                'line
                obj.Line (xpos, y1)-(xpos + 1 * DW, y2), &H0&, BF
                xpos = xpos + DW
            
            Case "2"
                'space
                obj.Line (xpos, y1)-(xpos + 1 * DW, y2), &HFFFFFF, BF
                xpos = xpos + DW
                'wide line
                obj.Line (xpos, y1)-(xpos + 2 * DW, y2), &H0&, BF
                xpos = xpos + 2 * DW
        End Select
    Next
Next

'1 more space
obj.Line (xpos, y1)-(xpos + 1 * DW, y2), &HFFFFFF, BF
xpos = xpos + DW

'final size and text
obj.width = (xpos + DW) * obj.width / obj.ScaleWidth
'obj.CurrentX = (obj.ScaleWidth - tw) / 2 'centraliza
obj.CurrentX = 0
'obj.CurrentY = y2 + 0.1 * th 'mudei de 0.25 para 0.1
'obj.Print bc_string; 'alterado aki
obj.CurrentY = obj.ScaleHeight - 11
obj.FontName = "ARIAL"
obj.fontsize = 8
If bPrintNum = True Then
    P_ImprX obj, bc_string, 0
End If
P_ImprX obj, Preco & " ", obj.ScaleWidth, False, True
obj.FontName = "arial"
obj.fontsize = 6
P_PrintXY obj, nome, 0, 0

'copy to clipboard
obj.Picture = obj.image
Clipboard.Clear
Clipboard.SetData obj.image, 2



End Sub
Function F_Ean8(Num As String) As String
    Dim i As Integer
    Dim c As Integer
    Dim X As Integer
    If Len(Num) < 7 Then
        Num = Format(Num, "000000#")
    End If
    For i = 1 To 7
        c = Int(Mid(Num, i, 1))
        If i Mod 2 = 0 Then
            X = X + c * 9
        Else
            X = X + c * 7
        End If
    Next
    
    F_Ean8 = Left(Num, 7) & Right(CStr(X), 1)
End Function

Public Function base_10(numero As String) As String
    Dim i, Produto, multiplicador, digito As Integer, valor As Integer
    Produto = 0
    multiplicador = 3
    For i = Len(numero) To 1 Step -1
        valor = val(Mid(numero, i, 1)) * multiplicador
        Produto = Produto + valor
        multiplicador = IIf(multiplicador = 3, 1, 3)
    Next
    digito = 10 - Int(Produto Mod 10)
    digito = IIf(digito = 10 Or digito = 11, 0, digito)
    base_10 = numero & Trim(str(digito))
End Function




