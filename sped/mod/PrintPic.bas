Attribute VB_Name = "PrintPic"
Sub PrintP(obj As Object, conteudo As String, Optional x As Single = -1, Optional y As Single = -1, Optional continua As Boolean = True, Optional tamFont As Integer = -1, _
    Optional Fonte As String = "", Optional BIU As String = "")
    Dim negrito As Boolean
    Dim italico As Boolean
    Dim sublinhado As Boolean
    Dim tamF As Integer
    Dim fon As String
    'BIU = bold, italic, sublinhado
    fon = obj.FontName
    tamF = obj.fontsize
    negrito = obj.FontBold
    obj.AutoRedraw = True
    sublinhado = obj.FontUnderline
    italico = obj.FontItalic
    obj.ScaleMode = 6 'vbmillimeters
    
    If x > -1 Then
        obj.CurrentX = x
    End If
    
    If y > -1 Then
        obj.CurrentY = y
    End If
    If InStr(1, BIU, "B") Then
        
        obj.FontBold = True
    Else
        obj.FontBold = False
    End If
    If InStr(1, BIU, "I") Then
        
        obj.FontItalic = True
    Else
        obj.FontItalic = False
    End If
    If InStr(1, BIU, "U") Then
        obj.FontUnderline = True
    Else
        obj.FontUnderline = False
    End If
    If tamFont > 0 Then
        obj.fontsize = tamFont
    End If
    If Fonte <> "" Then
        obj.FontName = Fonte
    End If
    
    
    If continua Then
        obj.Print conteudo;
      '  obj.CurrentX = obj.CurrentX + obj.TextWidth(conteudo)
    Else
        obj.Print conteudo
    End If
    
    obj.FontBold = negrito
    obj.FontUnderline = sublinhado
    obj.FontItalic = italico
    obj.fontsize = tamF
    obj.FontName = fon
    

End Sub

