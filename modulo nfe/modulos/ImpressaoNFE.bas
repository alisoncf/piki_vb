Attribute VB_Name = "ImpressaoNFe"
Public iPaginaAtual As Integer
Public iQtdPag As Integer
Public CaminhoArquivoConfigImpressaoNFCe As String
Public CaminhoArquivoConfigImpressaoCopiaVenda As String
Public Enum orientacao
    vertical = 0
    horizontal = 1
End Enum

Public Enum pAlign
    centro = 2
    direita = 3
    esquerda = 1
    'justificado = 0
End Enum
Type ImpressoraNFCe
    Indice As Integer
    deviceName As String
End Type
Type tImpressao
    PagIni As Integer
    PagFin As Integer
    Qualidade As Integer
End Type

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
    Dim X As Single
    Dim Y As Single
    On Error Resume Next
    X = obj.CurrentX
    Y = obj.CurrentY
    dr = obj.DrawStyle
    obj.DrawStyle = drwStyle
    obj.Line (x0, y0)-(x1, y1)
    obj.DrawStyle = dr
    obj.CurrentY = Y
    obj.CurrentX = X
    End Sub
Sub P_line1(obj As Object, Optional drwStyle As DrawStyleConstants = DrawStyleConstants.vbSolid)
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 1, obj.CurrentY, drwStyle
End Sub
Sub P_Quadrado(obj As Object, X As Single, Y As Single, Altura As Single, Largura As Single, _
    Optional style As DrawStyleConstants = vbSolid, Optional LineWidth As Integer = 1, _
    Optional cantosArredontados As Boolean = False, Optional Text As String = "")
    Dim pa As Single
    Dim pl As Single
    Dim dr As Integer
    Dim x0 As Integer
    Dim y0 As Integer
'    obj.ScaleMode = Printer.ScaleMode
    x0 = obj.CurrentX
    y0 = obj.CurrentY
    
    'pl = largura * 0.1 * altura / largura
    'pa = altura * 0.1 * altura / largura
    
    
    
    
    dr = obj.DrawStyle
    obj.DrawWidth = LineWidth
    obj.DrawStyle = style
    
    obj.Line (X + pl, Y)-(X + Largura - pl, Y)
    
    obj.Line (X + pl, Y + Altura)-(X + Largura - pl, Y + Altura)
    obj.Line (X, Y + pa)-(X, Y + Altura - pa)
    obj.Line (X + Largura, Y + pa)-(X + Largura, Y + Altura - pa)
    
   'obj.Line (x, y + pa)-(x + pl, y)
   'obj.Line (x, y - pa + altura)-(x + pl, y + altura)
    
   'obj.Line (x + largura - pl, y)-(x + largura, y + pa)
   'obj.Line (x + largura - pl, y + altura)-(x + largura, y + altura - pa)
    
    
    
    obj.CurrentY = Y
    obj.CurrentX = X
    If Text <> "" Then
        obj.CurrentX = X + 0.5
        obj.Print Text;
    End If
    
    obj.CurrentY = y0
    obj.CurrentX = x0
    'obj.Circle (800, 800), 500, , 0.55, 1.15, 1
     
    
End Sub
Sub P_PrintXY(objP As Object, sExpr As String, X As Single, Y As Single, Optional bContinuaX = True, Optional btAllign As Byte = 1)
    Dim sY As Single, sX As Single '1-> Esquerda 2-Centra 3- Direita
    objP.ScaleMode = vbMillimeters
    objP.CurrentY = Y
    If btAllign = 2 Then
        objP.CurrentX = X - (objP.TextWidth(sExpr) / 2)
    ElseIf btAllign = 3 Then
        objP.CurrentX = X - objP.TextWidth(sExpr)
    Else
        objP.CurrentX = X
    End If
    objP.Print sExpr
    
    If bContinuaX = True Then objP.CurrentY = Y
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
Sub P_CentraImpressao(objP As Object, strMens As String, Optional bContinua As Boolean = False, Optional x0 As Single = 0, Optional x1 As Single = 0)
    Dim wid As Single
    If x1 = 0 Then
        x1 = objP.ScaleWidth
    End If
    wid = x1 - x0
    If x0 = 0 Then
        objP.CurrentX = (wid - objP.TextWidth(strMens)) / 2
    Else
        objP.CurrentX = x0 + (wid - objP.TextWidth(strMens)) / 2
    End If
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
Sub P_CabRelTimbr(sMsg As String, obj As Object, Optional mostracab As Boolean = False, Optional cod_emp As String = "")
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
    mostracab = True
    
    
    If cod_emp <> "" Then
        sped_empresa.InicializaEmpresa cod_emp
    End If
    
    
    If mostracab Then
    P_CentraImpressao obj, strCod_Emp & "-" & strNome_Emp
        obj.fontsize = 10
        P_CentraImpressao obj, strEndereco_Emp & " - " & strBairro_Emp & " - " & strCidade_Emp & "/" & strUF_Emp & " - Fone: " & FormataTelefone(strFone_Emp)
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
    P_ImprX obj, strEndereco_Emp & "   Fone: " & _
        ConstEmpresaTelefone, 70, False
    P_PrintXY obj, "", 15, Printer.CurrentY + 2
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
    Dim Y As Single
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
    Y = obj.CurrentY
    obj.fontsize = 5
    P_ImprX obj, "", 0, False
    P_ImprX obj, strNome_Emp & " " & strCidade_Emp & " " & strUF_Emp, 0, False
    'P_ImprX obj, ConstSlogan_Cria, 0, False
    obj.CurrentY = Y
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
    
    Dim X As Single
    Dim Y As Single
    Dim IM As image
    Dim sPic As New StdPicture
    Dim p As PictureBox
    Dim widPic As Single
    
    Dim escala
    Dim centro As Single
    Y = obj.CurrentY
    X = obj.CurrentY
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
        If Printer.deviceName = Printers(i).deviceName Then
            BuscaImpressoraPadrao = i
            Exit Function
        End If
    Next

End Function
Function BuscaImpressoraPeloNome(device As String) As Integer

    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If device = Printers(i).deviceName Then
            BuscaImpressoraPeloNome = i
            Exit Function
        End If
    Next

End Function

Function BuscaNomeImpressoraPorIndice(Indice As Integer) As String
    BuscaNomeImpressoraPorIndice = Printers(Indice).deviceName
    
End Function
Function EscolheImpressora() As Integer
    
    EscolheImpressora = BuscaImpressoraPadrao
    frmEscolherImpressora.tag = "x"
    frmEscolherImpressora.Show 1
    If frmEscolherImpressora.tag <> "" Then
        EscolheImpressora = frmEscolherImpressora.cmbxPrinter.ListIndex
    Else
        EscolheImpressora = -1
    End If
    
    Unload frmEscolherImpressora

End Function
Sub GravaImpressoraPadraoNFCe(Indice As Integer)
    Dim conteudo As String
    conteudo = XMLUtil.XMLHead & CrLf
    conteudo = conteudo & "<PrintNFCe>Contem informações de configuração da impressão do Cupom NFCe" & CrLf
    conteudo = conteudo & "<impressora>" & CrLf
    conteudo = conteudo & "<indice>"
    conteudo = conteudo & Indice
    conteudo = conteudo & "</indice>" & CrLf
    conteudo = conteudo & "<DeviceName>"
    conteudo = conteudo & BuscaNomeImpressoraPorIndice(Indice)
    conteudo = conteudo & "</DeviceName>" & CrLf
    conteudo = conteudo & "</impressora>" & CrLf
    conteudo = conteudo & "</PrintNFCe>"
    GravaArquivoExternoUTF8 CaminhoArquivoConfigImpressaoNFCe, conteudo
End Sub
Sub GravaImpressoraPadraoCopiaVenda(Indice As Integer)
    Dim conteudo As String
    conteudo = XMLUtil.XMLHead & CrLf
    conteudo = conteudo & "<PrintNFCe>Contem informações de configuração da impressão da Cópia de venda" & CrLf
    conteudo = conteudo & "<impressora>" & CrLf
    conteudo = conteudo & "<indice>"
    conteudo = conteudo & Indice
    conteudo = conteudo & "</indice>" & CrLf
    conteudo = conteudo & "<DeviceName>"
    conteudo = conteudo & BuscaNomeImpressoraPorIndice(Indice)
    conteudo = conteudo & "</DeviceName>" & CrLf
    conteudo = conteudo & "</impressora>" & CrLf
    conteudo = conteudo & "</PrintNFCe>"
    GravaArquivoExternoUTF8 CaminhoArquivoConfigImpressaoCopiaVenda, conteudo
End Sub



Function BuscaImpressoraPadraoNFCe() As ImpressoraNFCe
    Dim conteudo As String
    Dim ind As String
    conteudo = LeArquivoExterno(CaminhoArquivoConfigImpressaoNFCe)
    
    If conteudo = "" Then 'não existe o arquivo ou ele está vazio
        BuscaImpressoraPadraoNFCe.Indice = BuscaImpressoraPadrao
        BuscaImpressoraPadraoNFCe.deviceName = BuscaNomeImpressoraPorIndice(BuscaImpressoraPadraoNFCe.Indice)
    Else
        ind = RetornaValorTagXML(conteudo, "<indice>")
        If ind = "" Then
            BuscaImpressoraPadrao
        End If
        BuscaImpressoraPadraoNFCe.Indice = CStr(ind)
        BuscaImpressoraPadraoNFCe.deviceName = RetornaValorTagXML(conteudo, "<DeviceName>")
    End If
    


End Function
Function BuscaImpressoraPadraoCopiaVenda() As ImpressoraNFCe
    Dim conteudo As String
    Dim ind As String
    conteudo = LeArquivoExterno(CaminhoArquivoConfigImpressaoCopiaVenda)
    
    If conteudo = "" Then 'não existe o arquivo ou ele está vazio
        BuscaImpressoraPadraoCopiaVenda.Indice = BuscaImpressoraPadrao
        BuscaImpressoraPadraoCopiaVenda.deviceName = BuscaNomeImpressoraPorIndice(BuscaImpressoraPadraoCopiaVenda.Indice)
    Else
        ind = RetornaValorTagXML(conteudo, "<indice>")
        If ind = "" Then
            BuscaImpressoraPadrao
        End If
        BuscaImpressoraPadraoCopiaVenda.Indice = CStr(ind)
        BuscaImpressoraPadraoCopiaVenda.deviceName = RetornaValorTagXML(conteudo, "<DeviceName>")
    End If
    


End Function

Sub MudaImpressoraPadraoNFCe()
    Dim i As Integer
     i = EscolheImpressora
     If i > -1 Then
        GravaImpressoraPadraoNFCe i
        ImpressoraPadraoCupomNFCe = BuscaImpressoraPadraoNFCe
     End If
End Sub
Sub MudaImpressoraCopiaVenda()
    Dim i As Integer
     i = EscolheImpressora
     If i > -1 Then
        GravaImpressoraPadraoCopiaVenda i
        ImpressoraCopiaVenda = BuscaImpressoraPadraoCopiaVenda
     End If
End Sub

Sub MudaImpressoraPadrao()
    Dim i As Integer
     i = EscolheImpressora
     If i > -1 Then
        Set Printer = Printers(i)
     End If
End Sub
Function P_ImprQuebraTurbo(Text As String, obj As Object, x1 As Single, X2 As Single, _
    Optional bAumenta As Boolean = False, Optional align As Integer = 1, Optional continuaY As Boolean = False) As Single
    On Error Resume Next
    Dim c As New Collection
    Dim sLinha As New Collection
    Dim iCount As Integer
    Dim jCount As Integer
    Dim sH As Single
    Dim sAcumula As String
    Dim pc As PictureBox
    Dim cDif As Single
    Dim iCXAux As Single 'simula o currentx
    Dim cTamPalavra As Single
    Dim yIni As Single
    'Dim l As linha
    'obj.ScaleMode = vbMillimeters
    If TypeName(obj) <> "Printer" Then
        Set pc = obj
    Else
'        Set pc = Printer
    End If
    Set c = F_SeparaTexto(Text)
    iCXAux = x1
    yIni = obj.CurrentY
    Set sLinha = Nothing
    For iCount = 1 To c.Count
        If obj.CurrentY >= obj.ScaleHeight - 2 Then
            obj.height = obj.height * 1.2
        End If
        sH = obj.TextWidth(c(iCount))
        sH = obj.ScaleX(sH, obj.ScaleMode, vbMillimeters)
        If (sH + iCXAux) > X2 Then
            
            iCXAux = x1
            sLinha.Add sAcumula
            sAcumula = ""
        End If
        If c(iCount) <> Chr(13) Then
            cTamPalavra = obj.TextWidth(c(iCount))
            cTamPalavra = obj.ScaleX(cTamPalavra, obj.ScaleMode, vbMillimeters)
            
            iCXAux = iCXAux + cTamPalavra
            sAcumula = sAcumula + c(iCount)
        Else
            'P_ImprX pc, "", 0, False
        End If
        
        'pc = Nothing
        
        
    Next
    If sAcumula <> "" Then
        sLinha.Add sAcumula
    End If
    obj.CurrentX = x1
    If align = 1 Then ' se for esquerda imprime direto
        For iCount = 1 To sLinha.Count
            P_ImprX obj, sLinha(iCount), CInt(x1), False
        Next
    ElseIf align = 2 Then
        For iCount = 1 To sLinha.Count
            P_CentraImpressao obj, sLinha(iCount), False, x1, X2
        Next
    ElseIf align = 3 Then
        For iCount = 1 To sLinha.Count
            P_DireitaImpressao obj, sLinha(iCount), X2 + 4
        Next
    Else 'justificado
        
        For iCount = 1 To sLinha.Count
            
                Set c = Nothing
                Set c = F_SeparaTexto(sLinha(iCount))
                obj.CurrentX = x1
                cDif = obj.TextWidth(sLinha(iCount))
                cDif = obj.ScaleX(cDif, obj.ScaleMode, vbMillimeters)
                cDif = (X2 - 3) - cDif
                For jCount = 1 To c.Count
                    If jCount = c.Count Then
                        If iCount <> sLinha.Count Then
                            P_DireitaImpressao obj, c(jCount), X2 + 4
                        Else
                            P_ImprX obj, c(jCount), obj.CurrentX
                            obj.Print ""
                        End If
                    Else
                    
                        P_ImprX obj, c(jCount), obj.CurrentX
                    End If
                    If iCount <> sLinha.Count Then  'ultima linha não justifica
                        obj.CurrentX = obj.CurrentX + (cDif / (c.Count - 1))
                    
                    End If
                Next
                'obj.Print ""
        Next
    End If
    P_ImprQuebraTurbo = obj.CurrentY
    If continuaY = True Then
        obj.CurrentY = yIni
    End If
End Function


Sub P_ImprQuebra(Text As String, obj As Object, x1 As Single, X2 As Single)
    Dim c As New Collection
    Dim iCount As Integer
    Dim sH As Single
    obj.ScaleMode = vbMillimeters
    Set c = F_SeparaTexto(Text)
    obj.CurrentX = x1
    For iCount = 1 To c.Count
        sH = obj.TextWidth(c(iCount))
        If (sH + obj.CurrentX) > X2 Then
            obj.Print
            obj.CurrentX = x1
        End If
        P_ImprX obj, c(iCount), obj.CurrentX
    Next
End Sub

Sub P_Grades(obj As Object, coord As String, Inicial As Single, final As Single, orient As orientacao, tipo As DrawStyleConstants)
    Dim iCount As Integer
    Dim c As Collection
    Set c = New Collection
    
    Set c = F_SeparaTexto(coord, ", ;,-", False)
    
    
    For iCount = 1 To c.Count
        If orient = vertical Then
            P_Line obj, CSng(c(iCount)), Inicial, CSng(c(iCount)), final, tipo
        Else
            P_Line obj, Inicial, c(iCount), final, c(iCount), tipo
        End If
    
    
    Next
        
    
    
    
End Sub


 
