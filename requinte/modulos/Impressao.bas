Attribute VB_Name = "Impressao"
Public iPaginaAtual As Integer
Public iQtdPag As Integer

Public Enum pAlign
    centro = 2
    direita = 3
    esquerda = 1
    'justificado = 0
End Enum

Sub callobjP()
    
    On Error Resume Next
    frmTestPrint.Show 1
End Sub

Sub P_ImprLinha(obj As Object, Optional sChar As String = "-", Optional IQuant As Integer = 300)
    P_CentraImpressao obj, String(IQuant, sChar)
End Sub
Sub P_Line(obj As Object, x0 As Single, y0 As Single, x1 As Single, y1 As Single, _
    Optional drwStyle As DrawStyleConstants = DrawStyleConstants.vbSolid)
    Dim dr As Integer
    Dim x As Single
    Dim y As Single
    On Error Resume Next
    x = obj.CurrentX
    y = obj.CurrentY
    dr = obj.DrawStyle
    obj.DrawStyle = drwStyle
    obj.Line (x0, y0)-(x1, y1)
    obj.DrawStyle = dr
    obj.CurrentY = y
    obj.CurrentX = x
    End Sub
Sub P_line1(obj As Object, Optional drwStyle As DrawStyleConstants = DrawStyleConstants.vbSolid)
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 1, obj.CurrentY, drwStyle
End Sub
Sub P_PrintXY(objP As Object, sExpr As String, x As Single, y As Single, Optional bContinuaX = True, Optional btAllign As Byte = 1)
    Dim sY As Single, sX As Single '1-> Esquerda 2-Centra 3- Direita
    objP.ScaleMode = vbMillimeters
    objP.CurrentY = y
    If btAllign = 2 Then
        objP.CurrentX = (objP.ScaleWidth - objP.TextWidth(sExpr)) / 2
    ElseIf btAllign = 3 Then
        objP.CurrentX = x - objP.TextWidth(sExpr)
    Else
        objP.CurrentX = x
    End If
    objP.Print sExpr
    
    If bContinuaX = True Then objP.CurrentY = y
End Sub
Sub P_ImprX(objP As Object, strExpressao As String, CurrentX As Single, Optional blnCont As Boolean = True, _
            Optional btAllign As pAlign = 1, Optional grade As Integer = -1, Optional dist As Single = 1.3, _
            Optional xLinha As Single = -1)
    
    Dim sngX1 As Single
    Dim sngY1 As Single
    Dim HeiExp As Single ' altura da expressao
    Dim WidExp As Single ' Tamanho da expressao
    Dim xRetorno As Single ' o x que retornará para o objeto
    Dim yRetorno As Single ' o x que retornará para o objeto
    Dim xIni As Single
    Dim drSt As DrawStyleConstants
    
    Dim yLinha As Single
    Dim mEsq As Single
    mEsq = 1
    objP.ScaleMode = vbMillimeters
    sngY1 = objP.CurrentY 'o y inicial
    sngX1 = objP.CurrentX ' o x inicial
    
    
    
    
    
    HeiExp = objP.TextHeight(strExpressao)
    WidExp = objP.TextWidth(strExpressao)
    
    
    
    
    yLinha = objP.CurrentY
    
    blnDireita = True
    
    If btAllign = 2 Then 'central
        
        
        CurrentX = CurrentX - (objP.TextWidth(strExpressao) / 2)
        If xLinha = -1 Then
            blnDireita = False
        
        End If
        
        
        
    ElseIf btAllign = 3 Then
        xLinha = CurrentX - WidExp - mEsq
        CurrentX = CurrentX - objP.TextWidth(strExpressao) 'direita
        
        
    Else
        xLinha = CurrentX - mEsq
        CurrentX = CurrentX 'esquerda
        
    End If
    xIni = CurrentX ' o x inicial em que a expressão será impressa
    objP.CurrentX = CurrentX
    If blnCont Then
        objP.Print strExpressao;
        objP.CurrentY = sngY1
    Else
        objP.Print strExpressao
    End If
    xRetorno = objP.CurrentX
    yRetorno = objP.CurrentY
    If grade = 0 Then
        drSt = vbSolid
    Else
        drSt = vbDot
    End If
    
    If grade > -1 Then
        If blnDireita = False Then
            'P_Line objP, xIni - 0.1, sngY1, WidExp + xIni + 0.1, sngY1, vbSolid
             P_Line objP, xIni - dist, sngY1, xIni - dist, sngY1 + HeiExp, drSt
            'P_Line objP, WidExp + xIni + 0.1, sngY1, WidExp + xIni + 0.1, sngY1 + HeiExp, vbSolid
            'P_Line objP, xIni - 0.1, sngY1 + HeiExp, WidExp + xIni + 0.1, sngY1 + HeiExp, vbSolid
        Else
            xIni = xIni - WidExp - dist
            dist = dist + 1
            
            'P_Line objP, WidExp + xIni + dist, sngY1, WidExp + xIni + dist, sngY1 + HeiExp, drSt
            P_Line objP, xLinha, sngY1, xLinha, sngY1 + HeiExp, drSt
            
        End If
    End If
    objP.CurrentY = yRetorno
    objP.CurrentX = xRetorno
End Sub

Sub P_ImprX_old(objP As Object, strExpressao As String, CurrentX As Integer, Optional blnCont As Boolean = True, _
            Optional blnDireita As Boolean = False, Optional grade As Integer = -1, Optional dist As Single = 1.3)
    
    Dim sngX1 As Single
    Dim sngY1 As Single
    Dim HeiExp As Single ' altura da expressao
    Dim WidExp As Single ' Tamanho da expressao
    Dim xRetorno As Single ' o x que retornará para o objeto
    Dim yRetorno As Single ' o x que retornará para o objeto
    Dim xIni As Single
    Dim drSt As DrawStyleConstants
    objP.ScaleMode = vbMillimeters
    sngY1 = objP.CurrentY 'o y inicial
    sngX1 = objP.CurrentX ' o x inicial
    xIni = CurrentX ' o x inicial em que a expressão será impressa
    HeiExp = objP.TextHeight(strExpressao)
    WidExp = objP.TextWidth(strExpressao)
    
    objP.CurrentX = CurrentX
    
    If blnDireita Then
        objP.CurrentX = objP.CurrentX - objP.TextWidth(strExpressao)
    End If
    If blnCont Then
        objP.Print strExpressao;
        objP.CurrentY = sngY1
    Else
        objP.Print strExpressao
    End If
    xRetorno = objP.CurrentX
    yRetorno = objP.CurrentY
    If grade = 0 Then
        drSt = vbSolid
    Else
        drSt = vbDot
    End If
    If grade > -1 Then
        If blnDireita = False Then
            'P_Line objP, xIni - 0.1, sngY1, WidExp + xIni + 0.1, sngY1, vbSolid
             P_Line objP, xIni - dist, sngY1, xIni - dist, sngY1 + HeiExp, drSt
            'P_Line objP, WidExp + xIni + 0.1, sngY1, WidExp + xIni + 0.1, sngY1 + HeiExp, vbSolid
            'P_Line objP, xIni - 0.1, sngY1 + HeiExp, WidExp + xIni + 0.1, sngY1 + HeiExp, vbSolid
        Else
            xIni = xIni - WidExp - dist
            dist = dist + 1
            'P_Line objP, xIni - 0.1, sngY1, WidExp + xIni + 0.1, sngY1, vbSolid
            'P_Line objP, xIni - 0.1, sngY1, xIni - 0.1, sngY1 + HeiExp, vbDot
            P_Line objP, WidExp + xIni + dist, sngY1, WidExp + xIni + dist, sngY1 + HeiExp, drSt
            'P_Line objP, xIni - 0.1, sngY1 + HeiExp, WidExp + xIni + 0.1, sngY1 + HeiExp, vbSolid
        End If
    End If
    objP.CurrentY = yRetorno
    objP.CurrentX = xRetorno
End Sub
Sub P_CentraImpressao(objP As Object, strMens As String, Optional bContinua As Boolean = False)

    objP.CurrentX = (objP.ScaleWidth - objP.TextWidth(strMens)) / 2
    If bContinua Then
        objP.Print strMens;
    Else
        objP.Print strMens
    End If
End Sub

Sub P_DireitaImpressao(objP As Object, strMens As String, tamanho As Single)
    objP.CurrentX = tamanho - objP.TextWidth(strMens)
    objP.Print strMens
End Sub
Sub P_CabRelTimbr(sMsg As String, obj As Object, Optional mostracab As Boolean = False)
    Dim iScale As Integer
    Dim wid As Single
    Dim hei As Single
    Dim iY As Single
    iScale = obj.ScaleMode
    Dim ifont As Integer
    On Error Resume Next
    obj.ScaleMode = 1
    obj.fontsize = 16
    obj.FontName = "arial"
    ifont = obj.fontsize
    If mostracab Then
    P_CentraImpressao obj, strNome_Emp
        obj.fontsize = 10
        P_CentraImpressao obj, strEndereco_Emp & " " & strBairro_Emp & " " & strCidade_Emp & "/" & strUF_Emp
    End If
    
    
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
    
    obj.fontsize = 12
    obj.FontUnderline = True
    iY = obj.CurrentY
    P_CentraImpressao obj, sMsg
    obj.fontsize = 8
    obj.CurrentY = iY
    P_ImprX obj, Date & "-" & Time, 0, False
    obj.FontUnderline = False
    obj.fontsize = 10
    obj.Print ""
    
    
    'obj.CurrentY = obj.CurrentY + 1
End Sub
Sub P_Timbr(sMsg As String, obj As Object)
    obj.ScaleMode = 1
    'obj.PaintPicture mdiFrm.ImageListPrint.ListImages(1).Picture, _
        obj.CurrentX - 10, obj.CurrentY - 5, 800, 700
    
    obj.fontsize = 18
    obj.FontName = "Arial"
    P_PrintXY obj, strNome_Emp, 15, obj.CurrentY
    obj.fontsize = 10
    P_ImprX obj, ConstEmpresaEndereco & "   Fone: " & _
        ConstEmpresaTelefone, 70, False
    P_PrintXY obj, constEmpresaRamo, 15, Printer.CurrentY + 2
    obj.fontsize = 10
    P_ImprX obj, sMsg & "  ", 70, False
    
    'obj.Line (68, obj.CurrentY - 12)-(68, obj.CurrentY)
    'obj.Line (0, obj.CurrentY + 2)-(200, obj.CurrentY + 2)
    'obj.Line (168, 0)-(168, 12)
    'obj.CurrentY = obj.CurrentY + 1
End Sub

Sub P_Rodape(Optional strMens As String = "", Optional blnNovaPagina = True, Optional msgDireita As String = "")
    Printer.FontName = "Times New Roman"
    If F_Vazio(strMens) Then
        strMens = ConstEmpresa
    End If
    Printer.FontName = "arial"
    Printer.fontsize = 10
    Printer.CurrentY = Printer.ScaleHeight - 15
    P_ImprX Printer, String(200, "_"), 0, False
    Printer.fontsize = 10
    Printer.FontBold = True
    
    P_ImprX Printer, msgDireita, 200, True, direita
    P_CentraImpressao Printer, strMens
    Printer.fontsize = 8
    Printer.fontsize = 10
    Printer.FontBold = False
    Printer.Print "" & String(200, "_") & ""
    If blnNovaPagina Then
        Printer.NewPage
    End If
End Sub
Sub P_FimRel(obj As Object, Optional strMens As String = "", Optional blnNovaPagina = True, Optional msgDireita As String = "")
    Dim y As Single
    obj.FontName = "Times New Roman"
    If F_Vazio(strMens) Then
        strMens = strNome_Emp
    End If
    obj.FontName = "arial"
    obj.fontsize = 8
    
    P_ImprX obj, String(200, "_"), 0, False
    obj.fontsize = 10
    obj.FontBold = False
    
    P_ImprX obj, msgDireita, 200, True, direita
    y = obj.CurrentY
    obj.fontsize = 5
    P_ImprX obj, constNome_sis, 0, False
    P_ImprX obj, strNome_Emp & " " & strCidade_Emp & " " & strUF_Emp, 0, False
    'P_ImprX obj, ConstSlogan_Cria, 0, False
    obj.CurrentY = y
    obj.fontsize = 10
    P_CentraImpressao obj, strMens
    obj.fontsize = 8
    obj.fontsize = 10
    obj.FontBold = False
    obj.CurrentY = obj.CurrentY + 0.5
    obj.Print "" & String(200, "-") & ""
    On Error GoTo trata_Erro
    If blnNovaPagina Then
        obj.NewPage
    End If
trata_Erro:
End Sub

Sub P_RodapeVis(obj As Object, Optional strMens As String = "", Optional msgDireita As String = "")
    On Error Resume Next
    obj.FontName = "Times New Roman"
    If F_Vazio(strMens) Then
        strMens = ConstEmpresa
    End If
    obj.FontName = "arial"
    obj.fontsize = 10
    obj.CurrentY = obj.ScaleHeight - 15
    P_ImprX obj, String(200, "_"), 0, False
    obj.fontsize = 10
    obj.FontBold = True
    
    P_ImprX obj, msgDireita, 190, True, direita
    P_ImprX obj, obj.Page, 200, True, direita
    P_CentraImpressao obj, strMens
    obj.fontsize = 8
    obj.fontsize = 10
    obj.FontBold = False
    obj.Print "" & String(190, "_") & ""
    
    
    obj.NewPage
    obj.endoc
    
End Sub




Sub callprinter(Optional bPag As Boolean = False, Optional bQuali As Boolean = False)
    
    
    frmEscolherImpressora.Show 1
    

End Sub

Sub ImprimeImagemArquivo(obj As Object, Arquivo As String, Optional Centralizado As Boolean = True, Optional proporcao As Single = 1)
    
    Dim x As Single
    Dim y As Single
    Dim IM As image
    Dim sPic As New StdPicture
    Dim p As PictureBox
    Dim widPic As Single
    
    Dim escala
    Dim centro As Single
    y = obj.CurrentY
    x = obj.CurrentY
    escala = obj.ScaleMode
    
    Set sPic = LoadPicture(Arquivo)
    'Set p = CreateObject("VB.PictureBox")
    'Set p.Picture = sPic
'    obj.ScaleMode = 9
    
    centro = ((obj.ScaleWidth) - (obj.ScaleX(sPic.width) * proporcao)) / 2
    If Centralizado = True Then
        obj.CurrentX = centro
    End If
    
    obj.PaintPicture sPic, obj.CurrentX, obj.CurrentY, obj.ScaleX(sPic.width * proporcao), obj.ScaleY(sPic.height * proporcao)
    
    obj.CurrentY = obj.CurrentY + obj.ScaleY(sPic.height) * proporcao
    
    obj.ScaleMode = escala

End Sub

Function BuscaImpressoraPadrao() As Integer

    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If Printer.DeviceName = Printers(i).DeviceName Then
            BuscaImpressoraPadrao = i
            Exit Function
        End If
    Next

End Function
Function BuscaImpressoraPeloNome(device As String) As Integer

    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If device = Printers(i).DeviceName Then
            BuscaImpressoraPadrao = i
            Exit Function
        End If
    Next

End Function

Function BuscaNomeImpressoraPorIndice(indice As Integer) As String
    BuscaNomeImpressoraPorIndice = Printers(indice).DeviceName
    
End Function
