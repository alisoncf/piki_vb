Attribute VB_Name = "Relatorios"
Public iPaginaAtual As Integer
Public iQtdPag As Integer


Sub callobjP()
    
    On Error Resume Next
    frmTestPrint.Show 1
End Sub



Sub P_CabInv(obj As Object, num_inv As String, data As String, folha As String)
    obj.ScaleMode = Printer.ScaleMode
    obj.FontName = "courier new"
    obj.fontsize = 8
    P_ImprLinha obj, "-", 110
    
    P_CentraImpressao obj, "REGISTRO DE INVENTÁRIO Nº " & num_inv
    P_ImprLinha obj, "-", 110
    P_ImprX obj, "FIRMA..............: " & strNome_Emp, 5, False
    P_ImprX obj, "INSCRIÇÃO ESTADUAL : " & strIE_Emp, 5, True
    P_ImprX obj, "CNPJ...............: " & strCNPJ_Emp, 110, False
    P_ImprX obj, "FOLHA..............: " & folha, 5, True
    P_ImprX obj, "ESTOQUE REFERENTE A: " & data, 110, False
    P_ImprLinha obj, "-", 110
    
End Sub
Sub P_CabInvProd(obj As Object, saldo_transp As String)
    
    obj.FontName = "courier new"
    obj.fontsize = 8
    Dim iX As Single
    Dim iY As Single
    Dim Fx As Single
    Dim Fy As Single
    P_ImprLinha obj, "-", 110
    
    
    iY = obj.CurrentY
    iX = obj.CurrentX
    P_ImprX obj, "CLS.FSC.", 3
    P_ImprX obj, "DISCRIMINAÇÃO: ", 55, True
    P_ImprX obj, "UNIDADE", 115
    P_ImprX obj, "QUANTIDADE", 150, True, 3
    P_ImprX obj, "UNITÁRIO", 172, True, 3
    P_ImprX obj, "TOTAL", 190, False, 3
    P_ImprLinha obj, "-", 110
    P_ImprX obj, "SALDO DE TRANSPORTE===========>: ", 20, True
    P_ImprX obj, saldo_transp, 198, False, 3
    P_ImprLinha obj, "-", 110
    Fy = obj.CurrentY
    Fx = obj.CurrentX
    
    P_Line obj, iX, iY - 1.5, iX, Fy - 1.5, vbDot
    P_Line obj, 19, iY - 1.5, 19, Fy - 1.5, vbDot
    P_Line obj, 114, iY - 1.5, 114, Fy - 1.5, vbDot
    P_Line obj, 129, iY - 1.5, 129, Fy - 1.5, vbDot
    P_Line obj, 154, iY - 1.5, 154, Fy - 1.5, vbDot
    P_Line obj, 176, iY - 1.5, 176, Fy - 1.5, vbDot
    P_Line obj, 202, iY - 1.5, 202, Fy - 1.5, vbDot
    
    
End Sub
Sub P_ItemInv(obj As Object, class As String, discrimina As String, _
    unidade As String, qtd As String, unit As String, total As String)
    
    obj.FontName = "courier new"
    
    obj.fontsize = 8
    
    
    P_ImprX obj, class, 5
    P_ImprX obj, discrimina, 20, True
    P_ImprX obj, unidade, 120, True
    P_ImprX obj, qtd, 150, True, 3
    P_ImprX obj, unit, 175, True, 3
    P_ImprX obj, total, 198, False, 3
    
    
    
    
    
End Sub
Sub P_RodaPeInvProd(obj As Object, saldo_transp As String)
    Dim Fy As Single
    Dim Fx As Single
    Dim iY As Single
    obj.FontName = "courier new"
    P_ImprLinha obj, "-", 110
    iY = 45
    
    
    Fy = obj.CurrentY
    P_ImprX obj, "SALDO A TRANSPORTAR ===========>: ", 20, True
    P_ImprX obj, saldo_transp, 198, False, 3
    P_ImprLinha obj, "-", 110
    Fx = obj.CurrentX
    
    
    
    P_Line obj, 19, iY - 1.5, 19, Fy - 1.5, vbDot
    P_Line obj, 114, iY - 1.5, 114, Fy - 1.5, vbDot
    P_Line obj, 129, iY - 1.5, 129, Fy - 1.5, vbDot
    P_Line obj, 154, iY - 1.5, 154, Fy - 1.5, vbDot
    P_Line obj, 176, iY - 1.5, 176, Fy - 1.5, vbDot
    P_Line obj, 202, iY - 1.5, 202, Fy - 1.5, vbDot
    On Error Resume Next
    'If TypeName(obj) = "printer" Then
        obj.NewPage
    'End If
    Err.Clear
End Sub
Sub P_RodaPeInvFinal(obj As Object, saldo_transp As String)
    
    obj.FontName = "courier new"
    P_ImprLinha obj, "-", 110
    
    
    P_ImprLinha obj, "-", 110
    
    P_ImprX obj, "SALDO DO INVENTÁRIO ===========>: ", 20, True
    P_ImprX obj, saldo_transp, 198, False, 3
    P_ImprLinha obj, "-", 110
    
    
    
    
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

Sub P_DireitaImpressao(objP As Object, strMens As String, tamanho As Single)
    objP.CurrentX = tamanho - objP.TextWidth(strMens)
    objP.Print strMens
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
    P_PrintXY obj, "", 15, Printer.CurrentY + 2
    obj.fontsize = 10
    P_ImprX obj, sMsg & "  ", 70, False
    
    'obj.Line (68, obj.CurrentY - 12)-(68, obj.CurrentY)
    'obj.Line (0, obj.CurrentY + 2)-(200, obj.CurrentY + 2)
    'obj.Line (168, 0)-(168, 12)
    'obj.CurrentY = obj.CurrentY + 1
End Sub





Sub imprimeContrato(seq_doc As String)
    'P_ImprimeContratoHTML seq_doc
    
    'P_ImprimeContratoGrafico seq_doc, frmVisualiza.Pic
    'frmVisualiza.Show 1
    
    
    
    'frmVisualiza.Show 1
    'If frmVisualiza.tag <> "" Then
    '    P_ImprimeContratoGrafico seq_doc, Printer
    '    Printer.EndDoc
    'End If
    'Unload frmVisualiza
    
End Sub

Sub imprimeVenda(obj As Object, num_ven As String)
    Dim rPesq As New ADODB.Recordset
    Dim rTro As New ADODB.Recordset
    Dim strSQL As String
    Dim cTotal As Currency
    Dim yAux As Single
    Dim sVendedor As String
    Dim sCond As String
    Dim cDesc As Currency
    Dim rCli As New ADODB.Recordset
    Dim sNomeCli As String
    Dim sMsg As String
    Dim sEnd As String
    Dim bEntrega As Boolean
    Dim sBai As String
    Dim sFone As String
    Dim sSituacao_Ven As String
    Dim rVendas As New ADODB.Recordset
    Dim sNum_ven As String
    Dim totalPro As Currency
    Dim sObs As String
    Dim sTel As String
    Dim rLan As New ADODB.Recordset
    Dim x As Single
    Dim y As Single
    Dim y2 As Single
    On Error GoTo trata_Erro
    
    
    
    P_Status "Localizando Venda"
    
    strSQL = "select tv.*,tc.*,tf.nome_for from  " & TAB_VENDAS & " as tv "
    
    strSQL = strSQL & " left join " & TAB_CAIXA & " as tc on tv.num_ven=tc.num_ven "
    strSQL = strSQL & " left join " & TAB_FORMASPAG & " as tf on tv.cod_For=tf.cod_for "
    
    
    strSQL = strSQL & " where tv.num_ven='" & num_ven & "' "
    
    abreObjetoAdo rVendas, strSQL
    
    P_Status "Localizando Cliente"
    
    strSQL = "select nome_mar, devolvido_ven,endereco_pro,tc.fantasia_cli,tc.nome_cli,telefone_cli,cnpj_cli,cep_cli,bairro_cli,endereco_cli,logradouro_cli,numero_cli,complemento_cli,cidade_cli,nome_cid,ti.*,"
    strSQL = strSQL & " tv.*,ti.unit_ite*ti.qtd_ite as total,tp.codfab_pro,tf.guerra_fun,tp.nome_pro"
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv INNER join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven inner join " & TAB_CLIENTES & " as tc left join  " & TAB_CIDADES & " as cid on cid.cod_cid=tc.cod_cid  "
    strSQL = strSQL & " on tv.cod_cli=tc.cod_cli inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro  left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " inner join tab_funcionarios as tf on tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " where 1 > 0 "
    strSQL = strSQL & " and tv.num_ven='" & num_ven & "' "
    strSQL = strSQL & " order by codfab_pro "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
    
    
    
    
    P_Status "Preparando impressão"
    
        
'        QRCodeUtil.ImprimeQRCodeArquivo num_ven, "app.Path\qrcodevenda.tmp"
        
        If .RecordCount = 0 Then Exit Sub
        InicializaEmpresa !cod_emp
        obj.fontsize = 14
        'P_PrintXY obj, "Emp: " & F_BuscaUmCampo("cod_emp", TAB_ITVENDAS, "num_ven='" & num_ven & "'"), 180, obj.CurrentY
        
        
        
        sSituacao_Ven = !situacao_ven
        
        
        InicializaEmpresa !cod_emp
        
        
        If !cod_For <> "01" Then
            P_CabRelTimbr "ORÇAMENTO: " & num_ven, obj
        Else
            P_CabRelTimbr "ORÇAMENTO: " & num_ven, obj
        End If
        On Error Resume Next
        frmMenu.bcxVenda.caption = num_ven
        
        obj.PaintPicture frmMenu.bcxVenda.Picture, 145, 11
            
        On Error GoTo trata_Erro
        
        
        If Int(sSituacao_Ven) = 2 Then
            obj.fontsize = 14
            P_CentraImpressao obj, "VENDA A CONCLUIR", False
            P_line1 obj, vbDot
        ElseIf Int(sSituacao_Ven) = 3 Then
            obj.fontsize = 14
            P_CentraImpressao obj, "VENDA A EFETIVAR", False
            P_line1 obj, vbDot
        ElseIf Int(sSituacao_Ven) = 4 Then
            obj.fontsize = 14
            P_CentraImpressao obj, "VENDA EFETIVADA", False
            P_line1 obj, vbDot
        ElseIf Int(sSituacao_Ven) = 5 Then
            obj.fontsize = 14
            P_CentraImpressao obj, "VENDA CANCELADA", False
            P_line1 obj, vbDot
        End If
        
        
        
        
        
        
        obj.fontsize = 9
        P_ImprX obj, "Cliente: " & !cod_cli & "-" & !nomecli_ven & " - " & TiraNull(!fantasia_cli) & "", 0, True
        P_ImprX obj, "CNPJ: " & !cnpj_Cli, 205, False, direita
        
        P_ImprX obj, "Endereço: " & TiraNull(!Logradouro_cli) & " N°" & TiraNull(rPesq!numero_cli) & " " & TiraNull(rPesq!complemento_cli), 0
        P_ImprX obj, "Bairro: " & TiraNull(!bairro_cli), 120, False
        P_ImprX obj, "Cidade: " & TiraNull(!nome_cid), 0
        P_ImprX obj, "Cep: " & TiraNull(!cep_cli), 120, True
        
        P_ImprX obj, "Fone: " & TiraNull(!telefone_cli), 80, True
        P_ImprX obj, "Data: " & !data_ven, 205, False, 3
        If TiraNull(!obs_ven) <> "" Then
            obj.fontsize = 11
            obj.FontBold = True
            P_ImprX obj, "Observação: " & !obs_ven, 0, False
            obj.FontBold = False
        End If
        obj.fontsize = 10
        
        sVendedor = F_BuscaUmCampo("Guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & !abrevenda_usu & "'")
        sNomeCli = !nomecli_ven
        
        If IsNull(!entrega_ven) Then
        ElseIf !entrega_ven = "1" Then
            bEntrega = True
            sEnd = !endereco_ven
            sBai = !bairro_ven
            sFone = Format(!telefonecli_ven)
            sObs = Format(!obs_ven)
            sTel = Format(!telefonecli_ven)
        Else
            
        End If
        
        cTotal = !valor_ven
        sNum_ven = !num_ven
        'sCond = F_BuscaUmCampo("cod_tip", " " & TAB_VENFEC & " ", "num_ven='" & num_ven & "'")
        sCond = !cod_For
        
        If sCond <> "" Then
            sCond = F_BuscaUmCampo("descricao_tip", "tab_tipoven", "cod_tip='" & sCond & "'")
            'P_ImprX obj, "Cond. Pagamento:" & sCond, 0, False
        End If
        P_ImprLinha obj
        P_ImprX obj, "Marca:", 0
       ' P_ImprX obj, "Cód.Fab:", 25
        P_ImprX obj, "End", 45
        'P_ImprX obj, "Cód.It:", 60
        P_ImprX obj, "Descrição:", 75
        
        P_ImprX obj, "Qtd", 165, True, direita
        P_ImprX obj, "Unitário", 180, True, direita
        'P_ImprX obj, "Desconto(R$)", 188, True, direita
        obj.FontBold = True
        obj.FontUnderline = True
        P_ImprX obj, "Valor Total", 200, direita, 3
        obj.FontBold = False
        obj.FontUnderline = False
        obj.Print
        P_ImprLinha obj
        cTotal = 0
        Do While Not .EOF
            
            
            obj.fontsize = 9
            
            y = obj.CurrentY
            If y2 > y Then
                'MsgBox "maior"
                obj.CurrentY = y2
            End If
            
            
            y2 = P_ImprQuebraTurbo(TiraNull(!nome_mar), obj, 1, 20, False, 1, True)
            obj.fontsize = 8
        '    P_ImprX obj, !codfab_pro, 25, , , -1
            
            P_ImprX obj, TiraNull(!endereco_pro), 45, , , -1
            obj.fontsize = 8
         '   P_ImprX obj, !cod_pro, 60, , , -1
             
          '  P_ImprX obj, StrConv(!nome_pro, vbProperCase), 75
            
            y = obj.CurrentY
            y2 = P_ImprQuebraTurbo(!desc_ite, obj, 75, 155, True, 1, True)
           
            obj.CurrentY = y
            
            obj.fontsize = 10
            P_ImprX obj, !qtd_ite, 165, True, direita
            P_ImprX obj, FormatVal(!final_ite / !qtd_ite), 180, True, direita
            
            'P_ImprX obj, FormatVal(!desconto_ite), 183, True, direita
            obj.FontUnderline = True
            obj.FontBold = True
            obj.fontsize = 10
            P_ImprX obj, FormatVal(!final_ite), 200, False, direita
            
            
            obj.FontUnderline = False
            obj.FontBold = False
            obj.fontsize = 8
            'If !pq_itd > 0 Then
            '    P_ImprX obj, "Q", 200, False, direita
            'Else
            '    P_ImprX obj, "", 200, False, direita
            'End If
            If !devolvido_ite > 0 Then
              
                obj.FontItalic = True
                P_ImprX obj, "Devolvido...", 140
                P_ImprX obj, FormatVal(!devolvido_ite), 160, True, True
                P_ImprX obj, "X", 170, True, True
                P_ImprX obj, FormatVal(!final_ite * (!devolvido_ite / !qtd_ite)), 175, False, True
                obj.FontItalic = False
            End If
            
            totalPro = totalPro + (!qtd_ite * !unit_ite)
            cTotal = cTotal + (!final_ite)
            cDesc = cDesc + !Desconto_ite
            
            .MoveNext
        Loop
        P_ImprLinha obj, "_"
        
        'P_ImprX obj, "Reconheço que pagarei a dívida acima.", 20, False
        obj.Print
        
        
        
        
        
        
        strSQL = "select tt.data_tro,tt.seq_tro,tp.nome_pro,ti.* from " & TAB_PRODUTOS & " as tp, tab_trocas as tt, tab_ittrocas as ti"
        strSQL = strSQL & " where ti.seq_tro=tt.seq_tro "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tt.num_ven='" & num_ven & "' "
        
        
        abreObjetoAdo rTro, strSQL
        
        With rTro
            
            Do While Not .EOF
                If .RecordCount > 0 Then
                P_ImprX obj, "TROCA:" & !seq_tro & " Em " & !data_tro, 10, False
                P_ImprX obj, "Produto:" & !nome_pro, 10, False
                P_ImprX obj, "Código:" & !cod_pro, 10, True
                P_ImprX obj, "Antes:" & !qtdant_itt, 50, True
                P_ImprX obj, "Atual:" & !qtdatual_itt, 100, False
                
            End If
                
                .MoveNext
            Loop
            .Close
        End With
        
      '  P_ImprLinha obj, "_"
        
        
        P_CentraImpressao obj, String(60, "_")
               
        
        P_CentraImpressao obj, sNomeCli
        
        
        With rPesq
            
            
            obj.Print ""
            obj.fontsize = 10
            obj.FontBold = True
            P_ImprLinha obj
            P_ImprX obj, "Total dos Produtos: " & FormatVal(totalPro), 5
            P_ImprX obj, " Desconto: " & FormatVal(cDesc), 95
            obj.fontsize = 12
            
            P_ImprX obj, "Valor Final: " & FormatVal(cTotal), 145
            strSQL = F_BuscaUmCampo("cod_for", TAB_VENFEC, "num_ven='" & sNum_ven & "'")
            strSQL = F_BuscaUmCampo("Descricao_TIp", _
                        "tab_tipoven", "cod_tip='" & strSQL & "'")
            If strSQL = "" Or strSQL = "0" Then
                strSQL = "A Vista"
                
            End If
            obj.Print
            obj.fontsize = 10
            P_ImprX obj, "Forma de Pagamento: " & rVendas!nome_For, 5
                       
            
            
            'obj.FontBold = True
            'P_ImprX obj, "Valor Final: " & FormatVal(totalPro - cDesc), 200, False, 3
            'obj.FontBold = False
            'P_ImprLinha obj
            P_ImprX obj, "Comprador: " & rVendas!comprador_ven, 95, False, 0
            P_ImprLinha obj
            
            P_Status "Localizando Documento"
            
            strSQL = "select ti.* from " & TAB_DOC & " as td, " & TAB_ITDOC & " as ti, " & TAB_VENDAS & " as tv "
            strSQL = strSQL & " where tv.num_ven=td.num_ven and td.seq_doc=ti.seq_doc"
            strSQL = strSQL & " and td.tipo_doc='R' and td.num_ven='" & num_ven & "' and esp_doc='V'"
            strSQL = strSQL & " order by venc_itd "
            abreObjetoAdo rLan, strSQL
            With rLan
                strSQL = ""
                If .RecordCount > 0 Then
                    sMsg = "Parcelas: "
                End If
                Do While Not .EOF
                    sMsg = sMsg & .AbsolutePosition & "ª. " & !venc_itd & "- R$ " & FormatVal(!valor_itd) & Chr(9)
                    .MoveNext
                Loop
                .Close
            End With
            
            strSQL = "select * from  " & TAB_VENFEC & "  where num_ven='" & num_ven & "'"
            abreObjetoAdo rLan, strSQL
            strSQL = ""
            With rLan
                Do While Not .EOF
                    strSQL = strSQL & .AbsolutePosition & ". " & !esp_cai & ": " & FormatVal(!valor_ven) & " "
                    .MoveNext
                Loop
                .Close
            End With
            If strSQL <> "" Then
                P_ImprX obj, "Pagamentos: " & strSQL, 5, False
            End If
            
            obj.FontBold = False
            
            strSQL = "select ti.* from " & TAB_DOC & " as td," & TAB_CHEQUES & " as ti, " & TAB_VENDAS & "  as tv "
            strSQL = strSQL & " where tv.num_ven=td.num_ven and td.seq_doc=ti.seq_doc"
            strSQL = strSQL & " and td.tipo_doc='R' and td.num_ven='" & num_ven & "' "
            strSQL = strSQL & " order by venc_che "
            abreObjetoAdo rLan, strSQL
            With rLan
                
                If .RecordCount > 0 Then
                    If strSQL <> "" Then
                        sMsg = sMsg & Chr(9)
                    End If
                    strSQL = strSQL & "Cheques: "
                    
                End If
                Do While Not .EOF
                    sMsg = sMsg & !Venc_che & "- R$ " & FormatVal(!Valor_che) & " nº " & !num_che & Chr(9)
                    .MoveNext
                Loop
                .Close
            End With
            If sMsg <> "" Then
                P_ImprX obj, sMsg, 5, False
                P_ImprLinha obj
            End If
             yAux = obj.CurrentY
            P_ImprX obj, "Vendedor: " & F_BuscaUmCampo("Guerra_Fun", TAB_FUNCIONARIOS, _
                "cod_fun='" & rVendas!abrevenda_usu & "'"), 5, False
                        
            
            
            
           
            
            If sSituacao_Ven = 1 Then
                P_ImprX obj, "(Orçamento em Aberto impresso por " & strNome_Usu & ")", 5, False
            ElseIf sSituacao_Ven = 2 Then
                P_ImprX obj, "(Orçamento Fechado impresso por " & strNome_Usu & ")", 5, False
            ElseIf sSituacao_Ven = 3 Then
                P_ImprX obj, "(Orçamento Aguardando Efetivação por " & strNome_Usu & ")", 5, False
            ElseIf sSituacao_Ven = 4 Then
                P_ImprX obj, "(Orçamento Finalizado impresso por " & strNome_Usu & ")", 5, False
            ElseIf sSituacao_Ven = 5 Then
                P_ImprX obj, "(Orçamento Cancelado impresso por " & strNome_Usu & ")", 5, False
            End If
            
        End With
    
    
        'P_ImprX obj, "Vendedor: " & sVendedor, 5, False
        obj.CurrentY = yAux
        obj.fontsize = 25
        If rVendas!situacao_ven <= 2 Then
            P_ImprX obj, "R$ " & FormatVal(cTotal), 180, False, 3
        
        Else
            P_ImprX obj, "R$ " & FormatVal(cTotal), 180, False, 3
        End If
        
        If rVendas!devolvido_ven > 0 Then
            obj.FontItalic = True
            P_ImprX obj, "Devolvido........: " & FormatVal(rVendas!devolvido_ven), 180, False, 3
            obj.FontItalic = False
        End If
        
        obj.fontsize = 10
        
        If bEntrega Then
            P_line1 obj
            obj.fontsize = 14
            P_line1 obj, vbInsideSolid
            P_ImprX obj, "ENTREGA", 0, False
            obj.fontsize = 10
            'obj.Print
            
            P_ImprX obj, "Endereço: " & sEnd, 0
            P_ImprX obj, "Bairro: " & sBai, 90
            P_ImprX obj, "Fone: " & sFone, 160, False
            
            
            P_line1 obj
            
            
            
        End If
        'If sObs <> "" Then
        '    obj.FontBold = False
        '    obj.fontsize = 11
        '    P_ImprX obj, "Obs: " & sObs, 0, False
        'End If
            obj.FontBold = 8
        P_line1 obj
        obj.FontBold = True
        obj.fontsize = 14
        obj.ForeColor = vbRed
        If Left(sDataBase, 7) = "estoque" Then
            P_CentraImpressao obj, "Devolução 1(um) Dia. Devolução somente com autorização do Vendedor."
        End If
        obj.ForeColor = vbBlack
        P_FimRel obj, "NÃO É VÁLIDO COMO DOCUMENTO FISCAL "
        
        P_Status ""
    rPesq.Close
    rVendas.Close
    End With
    Exit Sub
trata_Erro:
    P_Status ""
    MsgBox Err.Description
End Sub


Sub imprimeVendaHTML(num_ven As String)
    Dim rPesq As New ADODB.Recordset
    Dim rTro As New ADODB.Recordset
    Dim strSQL As String
    Dim cTotal As Currency
    Dim yAux As Single
    Dim sVendedor As String
    Dim sCond As String
    Dim cDesc As Currency
    Dim rCli As New ADODB.Recordset
    Dim sNomeCli As String
    Dim sMsg As String
    Dim sEnd As String
    Dim bEntrega As Boolean
    Dim sBai As String
    Dim sFone As String
    Dim sSituacao_Ven As String
    Dim rVendas As New ADODB.Recordset
    Dim sNum_ven As String
    Dim totalPro As Currency
    Dim sObs As String
    Dim sTel As String
    Dim rLan As New ADODB.Recordset
    Dim x As Single
    Dim y As Single
    Dim y2 As Single
'    On Error GoTo trata_Erro
    
    
    P_Status "Localizando Venda"
    
    strSQL = "select tv.*,tc.*,tf.nome_for from  " & TAB_VENDAS & " as tv "
    
    strSQL = strSQL & " left join " & TAB_CAIXA & " as tc on tv.num_ven=tc.num_ven "
    strSQL = strSQL & " left join " & TAB_FORMASPAG & " as tf on tv.cod_For=tf.cod_for "
    
    
    strSQL = strSQL & " where tv.num_ven='" & num_ven & "' "
    
    abreObjetoAdo rVendas, strSQL
    
    P_Status "Localizando Cliente"
    
    strSQL = "select nome_mar, devolvido_ven,endereco_pro,tc.fantasia_cli,tc.nome_cli,telefone_cli,cnpj_cli,cep_cli,bairro_cli,endereco_cli,logradouro_cli,numero_cli,complemento_cli,cidade_cli,nome_cid,ti.*,"
    strSQL = strSQL & " tv.*,ti.unit_ite*ti.qtd_ite as total,tp.codfab_pro,tf.guerra_fun,tp.nome_pro"
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv INNER join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven inner join " & TAB_CLIENTES & " as tc left join  " & TAB_CIDADES & " as cid on cid.cod_cid=tc.cod_cid  "
    strSQL = strSQL & " on tv.cod_cli=tc.cod_cli inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro  left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " inner join tab_funcionarios as tf on tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " where 1 > 0 "
    strSQL = strSQL & " and tv.num_ven='" & num_ven & "' "
    strSQL = strSQL & " order by codfab_pro "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
    
    
    
    
    P_Status "Preparando Documento"
    
        If .RecordCount = 0 Then Exit Sub
        
        Dim s As String
        Dim path As String
        
        path = html.F_AbreHtml("Venda_" & num_ven)
        html.P_HtmlAbreHead 1
        html.AbreTitle 1, "Venda Nº " & num_ven
        html.FechaHead
        html.AbreBody 1
        html.AbreTable 1, "style=" & entreAspas("width:100%;border: 0px solid")
        html.AbreLinha 1
        html.AbreColuna 1, "style=" & entreAspas("text-align:center")
        html.add_figura 1, App.path & "/pic/logo_topo.jpg"
        html.FechaColuna 1
        html.FechaTable
        
        
        
        html.AbreDiv 1
            html.AbreTable 1, "style=" & entreAspas("width:100%;border: 0px solid; font-family:Arial ; font-size:28px")
                html.AbreLinha 1
                    
                    html.AddColuna 1, "VENDA Nº " & num_ven & " - " & Now, "style=" & entreAspas("text-align:center")
                    html.AddColuna 1, "SITUAÇÃO " & !situacao_ven, "style=" & entreAspas("text-align:center")
                    
                    
                    
                    Print #1, "<br>"
                html.FechaLinha 1
            html.FechaTable
        html.FechaDiv 1
        
        html.AbreDiv 1
            html.LinhaHorizontal
        html.FechaDiv 1
        
        sSituacao_Ven = !situacao_ven
        
        
        html.AbreDiv 1
        html.AbreTable 1, "style=" & entreAspas("width:100%;border: 0px solid; font-family:Arial ; font-size:16px; font-weitght:bold")
            html.AbreLinha 1
                html.AddColuna 1, "Cliente: " & !cod_cli & "-" & rPesq!nomecli_ven, "style=" & entreAspas("text-align:left")
                html.AddColuna 1, "Data: " & Format(rPesq!data_ven, "dd/mm/yyyy"), "style=" & entreAspas("text-align:left")
            html.FechaLinha
        html.FechaTable
        html.FechaDiv 1
        
        If !cod_cli <> 0 Then
        
            html.AbreDiv 1
            html.AbreTable 1, "style=" & entreAspas("width:100%;border: 0px solid; font-family:Arial ; font-size:16px; font-weitght:bold")
                html.AbreLinha 1
                    html.AddColuna 1, "Endereço: " & TiraNull(!Logradouro_cli) & " N°" & TiraNull(rPesq!numero_cli) & " " & TiraNull(rPesq!complemento_cli), "style=" & entreAspas("text-align:left")
                    html.AddColuna 1, "Bairro: " & !bairro_cli, "style=" & entreAspas("text-align:left")
                html.FechaLinha
                html.AbreLinha 1
                    html.AddColuna 1, "Cidade: " & TiraNull(!nome_cid), "style=" & entreAspas("text-align:left")
                    html.AddColuna 1, "UF: " & TiraNull(!uf_cid), "style=" & entreAspas("text-align:left")
                    html.AddColuna 1, "CEP: " & TiraNull(!cep_cli), "style=" & entreAspas("text-align:left")
                    html.AddColuna 1, "Telefone: " & TiraNull(!telefone_cli), "style=" & entreAspas("text-align:left")
                html.FechaLinha
            html.FechaTable
            html.FechaDiv 1
        End If
        
        
        html.AbreDiv 1
        html.LinhaHorizontal
        html.FechaDiv 1
    
        
        If TiraNull(!obs_ven) <> "" Then
            html.AbreDiv 1
            html.AbreTable 1, "style=" & entreAspas("width:100%;border: 0px solid; font-family:Arial ; font-size:16px; font-weitght:bold")
            html.AbreLinha 1
                html.AddColuna 1, "Observação", "style=" & entreAspas("text-align:left")
                html.AddColuna 1, rPesq!obs_ven, "style=" & entreAspas("text-align:left")
            html.FechaLinha
        html.FechaTable
        html.FechaDiv 1
        End If
        
        sVendedor = F_BuscaUmCampo("Guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & !abrevenda_usu & "'")
        sNomeCli = !nomecli_ven
        
        If IsNull(!entrega_ven) Then
        ElseIf !entrega_ven = "1" Then
            bEntrega = True
            sEnd = !endereco_ven
            sBai = !bairro_ven
            sFone = Format(!telefonecli_ven)
            sObs = Format(!obs_ven)
            sTel = Format(!telefonecli_ven)
        Else
            
        End If
        
        cTotal = !valor_ven
        sNum_ven = !num_ven
        sCond = !cod_For
        
        If sCond <> "" Then
            sCond = F_BuscaUmCampo("descricao_tip", "tab_tipoven", "cod_tip='" & sCond & "'")
            'P_ImprX obj, "Cond. Pagamento:" & sCond, 0, False
        End If
        
        If bUsaConfeccao Then
            html.AbreDiv 1
            html.AbreTable 1, "style=" & entreAspas("width:100%;border: 1px solid; font-family:Arial ; font-size:14px;border-collapse:collapse")
            html.AbreLinha 1
                html.AddColuna 1, "Qtde: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Produto: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Detalhamento: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Cor: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Silk: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Bordado: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Vlr Unitário: ", "style=" & entreAspas("text-align:left;border: 1px solid")
                html.AddColuna 1, "Total: ", "style=" & entreAspas("text-align:left;border: 1px solid")
            html.FechaLinha
        
            With rPesq
                Do While Not .EOF
                    html.AbreLinha 1
                        html.AddColuna 1, TiraNull(!qtd_ite), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, TiraNull(!desc_ite), "style=" & entreAspas("text-align:left;border: 1px solid")
                        html.AddColuna 1, TiraNull(!detalhamento), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, TiraNull(!cor), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, TiraNull(!silk), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, TiraNull(!bordado), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, FormatVal(!unit_ite - !Desconto_ite), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, FormatVal(!final_ite), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                    html.FechaLinha
                    If Not IsNull(!final_ite) Then
                        cTotal = cTotal + (!final_ite)
                    End If
                    .MoveNext
                Loop
                .MoveFirst
            End With
        
        html.FechaTable
        html.FechaDiv 1
        
            
        
            
        End If
        
        html.FechaDiv 1
    
        html.FechaBody 1
    
        html.P_FechaHtml 1, True, path
        
        
      
        P_Status ""
        rPesq.Close
        rVendas.Close
    End With
    Exit Sub
trata_Erro:
    P_Status ""
    MsgBox Err.Description
End Sub



Sub P_AddPicVis(obj As Object)
    On Error Resume Next
    If obj.Name = "Pic" Then
         frmVisualiza.imglst.ListImages.Add , , frmVisualiza.Pic.image
        iQtdPag = iQtdPag + 1
        frmVisualiza.Pic.Cls
       
        frmVisualiza.Pic.Picture = Nothing
        obj.CurrentY = yIni
    End If

End Sub

Sub ImprVenda(num_ven As String)
    
    imprimeVenda frmVisualiza.Pic, num_ven
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprimeVenda Printer, num_ven
        Printer.EndDoc

    End If
            Unload frmVisualiza
End Sub

Sub ImprReciboBai(seq_cab As String, Optional abreArquivo As Boolean = True)
    Dim iCount As Integer
    ApontarPrinterToPDF "recibo" & seq_cab & ".pdf", "", abreArquivo
    P_ImprimeReciboBai seq_cab, Printer
    Printer.EndDoc
End Sub
Sub P_ImprimeReciboBai(seq_cab As String, obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rFin As New ADODB.Recordset
    Dim r As representante.representante
    Dim c As cliente
    Dim u As funcionario.funcionario
    On Error GoTo trata_Erro

    
    strSQL = "select valor_bai,desconto_bai,tb.cod_usu as usu,tb.data_bai,tb.cod_cob as cob,td.cod_ter,td.num_doc,td.tipo_ter,td.parcelas_doc,td.esp_doc,ti.* from  " & TAB_DOC & "  as td,  " & TAB_ITDOC & "  as ti, "
    strSQL = strSQL & "baixaitd as tb "
    strSQL = strSQL & " where tb.seq_itd=ti.seq_itd"
    strSQL = strSQL & " and ti.seq_doc=td.seq_doc "
    
    strSQL = strSQL & " and tb.seq_cab='" & seq_cab & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Baixa não encontrada!", vbExclamation
            Exit Sub
        End If
        obj.FontBold = True
        obj.fontsize = 12
        P_CabRelTimbr "COMPROVANTE DE PAGAMENTO", obj
        P_line1 obj
        obj.fontsize = 10
        
        If !tipo_ter = "C" Then
            obj.FontBold = False
            P_ImprX obj, "Cliente: ", 5
            c = Clientes.F_BuscaClienteCodigo(!cod_ter)
            obj.FontBold = True
            P_ImprX obj, c.nome_Cli, 20, False
            
            obj.FontBold = False
            P_ImprX obj, "Endereco: ", 5
            obj.FontBold = True
            P_ImprX obj, c.endereco_cli, 20, False
            obj.FontBold = False
            P_ImprX obj, "Bairro: ", 5
            obj.FontBold = False
            P_ImprX obj, c.bairro_cli & "/" & c.cidade_cli, 20, False
            
'            P_ImprX obj, "Orgão: ", 5
'            obj.FontBold = True
'            P_ImprX obj, c.org.nome_org, 20, False
'
            
            
        ElseIf !tipo_ter = "R" Then
            P_ImprX obj, "Representante: ", 5
            r = representante.F_BuscaRepCodigo(!cod_ter)
            P_ImprX obj, r.nome_rep, 35, False
        ElseIf !tipo_ter = "U" Then
            P_ImprX obj, "Funcionário: ", 5
            u = funcionario.F_SelectFuncionarioCod(!cod_ter)
            P_ImprX obj, u.nome_fun, 35, False
        End If
        
        obj.FontBold = False
        P_ImprX obj, "Data/Hora: " & Now, 5, False
        P_ImprX obj, "Funcionário: " & F_BuscaUmCampo("Guerra_Fun", TAB_FUNCIONARIOS, "cod_fun='" & !usu & "'"), 5, False
        obj.FontBold = True
        P_CentraImpressao obj, "PARCELAS PAGAS"
        obj.FontBold = False
        P_line1 obj
        
        P_ImprX obj, "N°DOC.", 5
        P_ImprX obj, "ID.DOC", 25
        P_ImprX obj, "ID.PARC.", 40
        P_ImprX obj, "Parcela", 60
        P_ImprX obj, "Inicial", 80
        P_ImprX obj, "Cobrança", 100
        P_ImprX obj, "Vencimento", 140
        P_ImprX obj, "Total", 160, True
        P_ImprX obj, "Desconto", 180, False
        
        
        P_line1 obj
        
        Do While Not .EOF
            P_ImprX obj, !num_doc, 5
            P_ImprX obj, !seq_doc, 25
            P_ImprX obj, !seq_itd, 40
            P_ImprX obj, !parcela_itd & "/" & !parcelas_doc, 60
            P_ImprX obj, FormatVal(!valor_itd), 80
            P_ImprX obj, !cob, 100
            P_ImprX obj, !venc_itd, 140
            P_ImprX obj, FormatVal(!valor_bai), 160, True
            P_ImprX obj, FormatVal(!desconto_bai), 180, False
            
            
            
            .MoveNext
        Loop
        .Close
    End With
    
    obj.Print ""
    obj.FontBold = True
    obj.fontsize = 12
    P_CentraImpressao obj, "FORMAS DE PAGAMENTO DAS PARCELAS"
    obj.FontBold = False
    P_line1 obj
    obj.fontsize = 10
    strSQL = "select distinct tb.*,tb2.data_bai from baixafin as tb, baixaitd as tb2 where tb.seq_cab='" & seq_cab & "' "
    strSQL = strSQL & " AND tb2.seq_cab=tb.seq_cab"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        P_ImprX obj, "ID.Fin.", 5
        P_ImprX obj, "Esp.", 20
        P_ImprX obj, "Data.", 40
        
        P_ImprX obj, "Valor", 80, True
        P_ImprX obj, "Desconto", 120, False
        P_line1 obj
        
        Do While Not .EOF
            P_ImprX obj, !seq_fin, 5
            P_ImprX obj, !esp_bai, 20
            P_ImprX obj, !data_bai, 40
            P_ImprX obj, FormatVal(!valor_bai), 80, True
            P_ImprX obj, FormatVal(!desconto_bai), 120, False
            .MoveNext
        Loop
        .Close
    End With
    
    P_ImprX obj, "", 0, False
    P_ImprX obj, "", 0, False
    
    P_line1 obj, vbDot
trata_Erro:
    
End Sub

Sub P_ImprimeListagem(dtobj As data, obj As Object)
    
    Dim sPage As Integer
    Dim iQtdPage As Integer
    Dim iCount As Integer
    Dim iTotal As Integer
    iTotal = 65
    
    With dtobj.Recordset
        If .RecordCount > 0 Then .MoveFirst
        P_CabRelTimbr "Listagem de Preços de " & Date, obj
        obj.fontsize = 10
        P_ImprX obj, "", 5
        P_ImprX obj, "", 20
        P_ImprX obj, "Produto", 5
        P_ImprX obj, "UN.", 140
        P_ImprX obj, "Seção", 148
        P_ImprX obj, "Preço", 190, True, direita
        P_ImprX obj, "Qtd", 200, False, direita
        
        P_ImprLinha obj
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        iQtdPage = (.RecordCount / iTotal)
        If iQtdPage Mod 2 <> 0 Then
            iQtdPage = iQtdPage + 1
        End If
        Do While Not .EOF
            
            If iCount = 55 Then
                iCount = 0
                sPage = sPage + 1
                
                P_RodapeVis obj, strNome_Emp, "Pag " & sPage & " de " & iQtdPage
                P_AddPicVis obj
                If TypeName(obj) = "Printer" Then
                    P_CabRelTimbr "Listagem de Preços de " & Date, obj
                End If
                
                obj.fontsize = 10
                P_ImprX obj, "", 5
                P_ImprX obj, "", 20
                P_ImprX obj, "Produto", 5
                P_ImprX obj, "UN.", 140
                P_ImprX obj, "Seção", 148
                P_ImprX obj, "Preço", 190, True, direita
                P_ImprX obj, "Qtd", 200, False, direita
                P_ImprLinha obj
            End If
            obj.fontsize = 7
            P_ImprX obj, .AbsolutePosition, 0
            obj.fontsize = 10
            P_ImprX obj, !cod_pro, 5
            
            obj.fontsize = 9
            P_ImprX obj, !nome_pro, 22
            obj.fontsize = 10
            P_ImprX obj, !unidade_pro, 140
            obj.fontsize = 9
            P_ImprX obj, !nome_sec, 148
            obj.fontsize = 10
            P_ImprX obj, FormatVal(!pv_pro), 190, True, direita
            P_ImprX obj, Int(!er_pro), 200, False, direita
            P_line1 obj, vbDot
            iCount = iCount + 1
            P_Status "Gerando... " & Round((.AbsolutePosition + 1) / .RecordCount * 100) & "%"
            .MoveNext
        Loop
        P_FimRel obj, strNome_Emp, True, "Pag Final"
    End With
    P_AddPicVis obj
    
End Sub

Sub ImprimePcp(sSeq_Pcp As String)
        P_ImprPcp sSeq_Pcp, frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.tag <> "" Then
            P_ImprPcp sSeq_Pcp, Printer
            Printer.EndDoc
        End If
        Unload frmVisualiza
End Sub
Sub ImprimePCPHtml(seq_pcp As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sArq As String
    Dim sFor As String
    Dim cTotal As Currency
    strSQL = "select ti.cod_for,tf.nome_For,tp.cod_pro,tm.nome_mar,tt.nome_cat,tp.nome_pro,qfa_ipc,upc_ipc,sum(ti.qfa_ipc) as soma "
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp, tab_itenspcp as ti left join " & TAB_FORNECEDORES & " as tf "
    strSQL = strSQL & " on ti.cod_for=tf.cod_for ,"
    strSQL = strSQL & TAB_MARCAS & " as tm, tab_categorias as tt "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " and tt.cod_Cat=tp.cod_cat"
    strSQL = strSQL & " and ti.seq_pcp='" & seq_pcp & "' "
    strSQL = strSQL & " and ti.sts_ipc='S' "
    
    strSQL = strSQL & " group by tp.cod_pro"
    strSQL = strSQL & " order by tf.nome_for, nome_pro "
    
    
    abreObjetoAdo rPesq, strSQL
    
    sArq = html.F_AbreHtml("pcp" & F_SoNumero(Now))
    
    html.P_HtmlAbreHead
    html.P_HtmlAbreTitle 1, "Processo de Compra"
    html.P_HtmlFechaHead
    html.P_HtmlAbreH 1, 4, "Processo de Compra: " & seq_pcp, True
    html.P_HtmlAbreBody
    html.P_LinhaEmBranco
    With rPesq
        Do While Not .EOF
            If sFor <> !cod_For Then
                
                If .AbsolutePosition > adPosBOF Then
                    P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, ""
                    P_HtmlAbreColuna 1, "Total"
                    P_HtmlAbreColuna 1, FormatVal(cTotal)
                    P_HtmlFechaLinha
                    P_HtmlFechaTable
                    cTotal = 0
                End If
                
                P_LinhaEmBranco 1, 6
                P_HtmlAbreTable 1, 0
                P_HtmlAbreColuna 1, "Fornecedor"
                P_HtmlAbreColuna 1, !nome_For
                P_HtmlFechaLinha
                P_HtmlFechaTable
            
            
                P_HtmlAbreTable 1
                P_HtmlAbreLinha 1
                P_HtmlAbreColuna 1, "Código"
                P_HtmlAbreColuna 1, "Produto"
                P_HtmlAbreColuna 1, "Qtd"
                P_HtmlAbreColuna 1, "Unitário"
                P_HtmlAbreColuna 1, "Total"
                P_HtmlFechaLinha
                
                
                
                sFor = !cod_For
                
            End If
            
            
            
            P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, !cod_pro
            P_HtmlAbreColuna 1, !nome_pro
            P_HtmlAbreColuna 1, !qfa_ipc
            P_HtmlAbreColuna 1, FormatVal(!upc_ipc)
            P_HtmlAbreColuna 1, FormatVal(!upc_ipc * !qfa_ipc)
            P_HtmlFechaLinha
            cTotal = cTotal + (!upc_ipc * !qfa_ipc)
            
            .MoveNext
        Loop
        
                    P_HtmlAbreLinha
                    
                    P_HtmlAbreColuna 1, "Total", "center", True, "-1", 2
                    P_HtmlAbreColuna 1, FormatVal(cTotal), "center", True, "-1", "3"
                    P_HtmlFechaLinha
                    P_HtmlFechaTable
                    cTotal = 0
        
        
    End With
    
    
    
    html.P_HtmlFechaBody
    html.P_FechaHtml
    Shell "explorer " & sArq
    
    
    
    
    
    
End Sub
Sub P_ImprPcp(sSeq_Pcp As String, obj As Object)
    Dim iQtd_pag As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sFor As String
    
    
    'fazer a impressão por fornecedor
    
    
    strSQL = "select ti.cod_for,tf.nome_For,tp.cod_pro,tm.nome_mar,tt.nome_cat,tp.nome_pro,sum(ti.qfa_ipc) as soma "
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp, tab_itenspcp as ti left join " & TAB_FORNECEDORES & " as tf "
    strSQL = strSQL & " on ti.cod_for=tf.cod_for ,"
    strSQL = strSQL & TAB_MARCAS & " as tm, tab_categorias as tt "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " and tt.cod_Cat=tp.cod_cat"
    strSQL = strSQL & " and ti.seq_pcp='" & sSeq_Pcp & "' "
    strSQL = strSQL & " and ti.sts_ipc='S' "
    
    strSQL = strSQL & " group by tp.cod_pro"
    strSQL = strSQL & " order by tf.cod_for, nome_pro "
    
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount = 0 Then
        P_Status ""
        MsgBox "Não constam itens nesta compra!", vbExclamation
        Exit Sub
    Else
        If CCur(rPesq.RecordCount) < 30 Then
            iQtd_pag = 1
        ElseIf CCur(rPesq.RecordCount) Mod 30 = 0 Then
            iQtd_pag = CInt(rPesq.RecordCount) / 30
        Else
            iQtd_pag = rPesq.RecordCount / (rPesq.RecordCount) Mod 30
            iQtd_pag = iQtd_pag + 1
        End If
    End If
    
    P_Status "Imprimindo Cotação..."
    
    sFor = rPesq!cod_For
    P_CabRelTimbr strNome_Emp, obj
    
    P_PrintXY obj, "COTAÇÃO: " & sSeq_Pcp, 170, 0 + 20, False
    P_PrintXY obj, "DATA: " & Date, 170, 3.5 + 20, False
    P_PrintXY obj, "HORA: " & Time, 170, 7 + 20, False
    P_ImprX obj, "EMPRESA: " & rPesq!nome_For, 10, False
    obj.fontsize = 12
    'obj.CurrentY = 25
'    P_ImprX obj, "   Solicitamos envio de cotação dos abaixo relacionados: ", 0, False
    P_ImprX obj, "", 0, False
    obj.fontsize = 10
    obj.Print
    obj.FontBold = True
    P_ImprX obj, "Produto", 0, , , 1
    P_ImprX obj, "Ref.:", 90
    P_ImprX obj, "Qtde.", 140, True, True, 1
    P_ImprX obj, "Unit. ", 160, True, True, 1
    P_ImprX obj, "Imp. ", 175, True, True, 1
    P_ImprX obj, "TOTAL", 185, True, , 1
    obj.FontBold = False
    obj.CurrentY = obj.CurrentY + 4
    obj.Line (0, obj.CurrentY)-(200, obj.CurrentY)
    With rPesq
        .MoveFirst
        Do While Not .EOF
            If obj.ScaleHeight <= obj.CurrentY + 50 Then
                'If !sts_ipc = "S" Then
                    If TypeName(obj) = "Printer" Then
                        P_Rodape strEmp_Nome & " - " & strEmp_Cidade & " /" & strEmp_UF, True, "Pag: " & obj.Page & " de " & iQtd_pag
                    End If
                    obj.FontBold = True
                    P_ImprX obj, "Mercadoria ", 0, , , 1
                    P_ImprX obj, "Ref.:", 90, , , 1
                    P_ImprX obj, "Qtde.", 140, True, True, 1
                    P_ImprX obj, "Unit. ", 160, True, True, 1
                    P_ImprX obj, "Imp. ", 175, True, True, 1
                    P_ImprX obj, "TOTAL", 185, True, , 1
                    obj.FontBold = False
                    obj.CurrentY = obj.CurrentY + 4
                    obj.Line (0, obj.CurrentY)-(200, obj.CurrentY)
                'End If
            End If
            
            
            
            If sFor <> !cod_For Then
                If TypeName(obj) = "Printer" Then
                        Printer.NewPage
                End If
                P_ImprX obj, "", 0, False
                P_PrintXY obj, "COTAÇÃO: " & sSeq_Pcp, 170, 0 + 20, False
                P_PrintXY obj, "DATA: " & Date, 170, 3.5 + 20, False
                P_PrintXY obj, "HORA: " & Time, 170, 7 + 20, False
                P_ImprX obj, "EMPRESA: " & rPesq!nome_For, 10, False
            End If
            'If !sts_ipc = "S" Then
                obj.CurrentY = obj.CurrentY + 2
                P_ImprX obj, !nome_pro, 0, , , 1
                P_ImprX obj, "", 90, , , 1
                P_ImprX obj, FormatVal(!soma) & " ", 140, True, True, 1
                P_ImprX obj, "", 160, True, True, 1
                    P_ImprX obj, "", 175, True, True, 1
                P_ImprX obj, "", 185, True, , 1
                obj.CurrentY = obj.CurrentY + 4
                obj.Line (0, obj.CurrentY)-(200, obj.CurrentY)
            'End If
            .MoveNext
        Loop
    End With
    obj.Print
    P_ImprX obj, "Condições de Pagamento: __________________________________________", 0, True
    P_ImprX obj, "Total:______________", 165, False
        On Error Resume Next
        P_FimRel obj, strNome_Emp & " - " & strCidade_Emp & " /" & strUF_Emp, True, "Pag: " & obj.Page & " de " & iQtd_pag
    
    P_Status ""

End Sub

Function P_ImprimeCarne(seq_doc As String, obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipoTer As String
    Dim yIni As Single
    Dim xIni As Single
    Dim iImp As Integer
    Dim rEnd As New ADODB.Recordset
    Dim sNomeTer As String
    Dim sVia As String
    Dim iVia As Integer
    
'    On Error GoTo trata_erro
On Error Resume Next
    
    sVia = F_BuscaUmCampo("count(*)", "tab_docimpress", "seq_doc='" & seq_doc & "'")
    If sVia = "" Then
       sVia = "0"
    End If
    iVia = Int(sVia)
    sVia = Int(sVia) + 1 & "ª Via"
    If iVia > -1 Then
        
        obj.fontsize = 76
        obj.FontName = "arial"
        obj.ForeColor = &H8000000F
        'obj.Print sVia;
        'obj.CurrentY = 100
        P_CentraImpressao obj, sVia, True
        obj.ForeColor = vbBlack
    End If
    sTipoTer = F_BuscaUmCampo("tipo_ter", TAB_DOC, "seq_doc='" & seq_doc & "'")
    If sTipoTer = "" Then
        Exit Function
    End If
    strSQL = F_SelectDocsReceber(sTipoTer, , , , , , , , , seq_doc)
    
    rPesq.Open strSQL, conn
    
    With rPesq
        
        
        
        
        If .RecordCount = 0 Then
            'MsgBox "Documento " & seq_doc & " não encontrado!", vbExclamation
            .Close
            Exit Function
        End If
        
        If sTipoTer = "C" Then
            strSQL = "select tci.nome_cid,tipo_set,cep_cli,logradouro_clu,numero_cli,complemento_cli,endereco_cli,uf_cid,nome_set "
            strSQL = strSQL & " from tab_clientes as tc," & TAB_CIDADES & " as tci, tab_setores as ts "
            strSQL = strSQL & " where tci.cod_cid=ts.cod_cid "
            strSQL = strSQL & " and tc.cod_set=ts.cod_set "
            strSQL = strSQL & " and cod_cli='" & !cod_ter & "' "
            rEnd.Open strSQL, conn_client
            
        Else
            strSQL = "select tci.nome_cid,tipo_set,cep_rep as cep_cli,endereco_rep as endereco_cli,uf_cid,nome_set "
            strSQL = strSQL & " from tab_representantes as tc," & TAB_CIDADES & " as tci, tab_setores as ts "
            strSQL = strSQL & " where tci.cod_cid=ts.cod_cid "
            strSQL = strSQL & " and tc.cod_set=ts.cod_set "
            strSQL = strSQL & " and cod_rep='" & !cod_ter & "' "
            rEnd.Open strSQL, conn_client
        End If
        
        Do While Not .EOF
            sNomeTer = !nome
            If iImp = 4 Then
                P_RodapeVis obj
                P_AddPicVis obj
                iImp = 0
            End If
            iImp = iImp + 1
            frmVisualiza.brcdx.caption = !seq_itd
            
            obj.fontsize = 12
            'obj.Print
            P_ImprX obj, "", 0, False
            obj.fontsize = 8
            P_ImprX obj, sNomeTer, 5, False
            P_ImprX obj, "CPF:" & !cpf, 5, False
            
            
            yIni = obj.CurrentY
            obj.DrawWidth = 2
            
            obj.fontsize = 12
            P_Grades obj, obj.CurrentY + 0.4, 4, 195, 1, DrawStyleConstants.vbSolid
            obj.DrawWidth = 1
            obj.FontBold = False
            P_ImprX obj, "Documento:", 5, True
            obj.fontsize = 8
            
            P_ImprX obj, "Local de Pagamento:", 75, True
            
            obj.fontsize = 10
            P_ImprX obj, strFantasia_Emp, 110, True
            obj.FontBold = True
            obj.fontsize = 12
            P_ImprX obj, seq_doc, 35, False
            obj.FontBold = False
            
            
            
            
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_ImprX obj, "Parcela:", 5, True
            obj.fontsize = 10
            P_ImprX obj, "Parcela:", 75, True, esquerda, -1
            P_ImprX obj, "Nº Seq.:", 90, True, esquerda, 0
            P_ImprX obj, "Doc./Venda:", 110, True, esquerda, 0
            P_ImprX obj, "Dt. Proc.:", 140, True, esquerda, 0
            
            obj.FontBold = True
            obj.fontsize = 12
            P_ImprX obj, !parcela_itd & "/" & !parcelas_doc, 35, False
            
            obj.FontBold = False
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_ImprX obj, "Valor:", 5, True
            obj.FontBold = True
            P_ImprX obj, FormatVal(!reman_itd), 35, True
            
            P_ImprX obj, !parcela_itd & "/" & !parcelas_doc, 75, True, esquerda, -1
            P_ImprX obj, !seq_doc, 90, True, esquerda, 0
            P_ImprX obj, !num_doc, 110, True, esquerda, 0
            P_ImprX obj, !data_doc, 140, False, esquerda, 0
            obj.FontBold = False
            
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_Grades obj, obj.CurrentY, 4, 60, 1, vbSolid
            P_ImprX obj, "Vencimento:", 5, True
            obj.FontBold = True
            P_ImprX obj, !venc_itd, 35, True
            obj.FontBold = False
            obj.fontsize = 10
            P_ImprX obj, "Valor Doc:", 75, True, esquerda, -1
            P_ImprX obj, "Valor Parc.", 100, True, esquerda, 0
            P_ImprX obj, "Seq.Parc", 120, True, esquerda, 0
            P_ImprX obj, "Vencimento", 150, True, esquerda, 0
            obj.fontsize = 12
            P_ImprX obj, "", 0, False
            
            
            obj.FontBold = False
            
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_ImprX obj, "Desconto:", 5, True
            obj.FontBold = True
            P_ImprX obj, FormatVal(!valor_doc), 75, True, esquerda, -1
            P_ImprX obj, FormatVal(!reman_itd), 100, True, esquerda, 0
            P_ImprX obj, !seq_itd, 120, True, esquerda, 0
            P_ImprX obj, !venc_itd, 150, False, esquerda, 0
            
            obj.FontBold = False
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_ImprX obj, "Juros:", 5, True
            obj.fontsize = 8
            obj.FontBold = False
            P_ImprX obj, "Cliente:", 75, True
            obj.fontsize = 10
            P_ImprX obj, !nome, 85, True
            obj.fontsize = 12
            P_ImprX obj, "", 100, False
            P_Grades obj, obj.CurrentY, 4, 60, 1, vbSolid
            P_ImprX obj, "Outros:", 5, True
            obj.fontsize = 8
            obj.FontBold = False
            P_ImprX obj, "CPF:", 75, True
            obj.fontsize = 10
            P_ImprX obj, !cpf, 90, True
            obj.fontsize = 8
            P_ImprX obj, "Cidade:", 120, True
            obj.fontsize = 10
            P_ImprX obj, rEnd!nome_cid & "/" & rEnd!uf_cid, 131, True
            obj.fontsize = 12
            P_ImprX obj, "", 100, False
            
            P_Grades obj, obj.CurrentY, 4, 60, 1, vbSolid
            P_ImprX obj, "Total:", 5, True
            obj.fontsize = 8
            P_ImprX obj, "End: ", 75, True
            
            P_ImprX obj, StrConv(rEnd!Logradouro_cli & " N°" & rEnd!numero_cli & " " & rEnd!complemento_cli, vbProperCase) & " " & StrConv(rEnd!tipo_set, vbProperCase) & " " & StrConv(rEnd!nome_set, vbProperCase) & " CEP: " & rEnd!cep_cli, 82, True
            obj.fontsize = 12
            P_ImprX obj, "", 100, False
            
            
            obj.FontBold = False
            'P_Line obj, 30, yIni, 30, obj.CurrentY
            P_Grades obj, 4, yIni, obj.CurrentY, vertical, vbSolid
            P_Grades obj, 30, yIni, obj.CurrentY, vertical, vbSolid
            P_Grades obj, 60, yIni, obj.CurrentY, vertical, vbSolid
            P_Grades obj, 72, yIni, obj.CurrentY, vertical, vbSolid
            P_Grades obj, 195, yIni, obj.CurrentY, vertical, vbSolid
            'obj.ScaleMode = vbMillimeters
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            
            
            
            
            obj.PaintPicture frmVisualiza.brcdx.Picture, 5, obj.CurrentY
            obj.PaintPicture frmVisualiza.brcdx.Picture, 75, obj.CurrentY
            obj.fontsize = 8
            P_ImprX obj, "Impresso em:" & Now, 195, False, direita
            P_ImprX obj, "", 0, False
            P_Grades obj, obj.CurrentY, 0, 200, 1, vbDot
            If sTipoTer = "R" Then
                
                P_ImprX obj, "Eu, " & sNomeTer & ", reconheço a dívida acima, e estou de acordo com seus vencimentos e valores.", 5, False
                
                P_ImprX obj, strCidade_Emp & ", " & Date & " ______________________________________", 5, False
            End If
            P_ImprX obj, sVia, 0, False
            .MoveNext
        Loop
        .Close
        rEnd.Close
    End With
    
    P_AddPicVis obj
    
    If TypeName(obj) = "Printer" Then
        Printer.EndDoc
        strSQL = "A Impressão Saiu Correta?" & Chr(13)
        strSQL = strSQL & "caso a impressão não tenha sido proveitosa escolha NÃO"
        If MsgBox(strSQL, vbYesNo + vbQuestion) = vbYes Then
            P_GravaDocImperss seq_doc
        End If
    End If
    
    Exit Function
trata_Erro:
    P_MostraErro
End Function
Sub imprimeCarne(seq_doc As String)
    P_ImprimeCarne seq_doc, frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprimeCarne seq_doc, Printer
        Printer.EndDoc
        
    End If
    Unload frmVisualiza
    
    
End Sub


