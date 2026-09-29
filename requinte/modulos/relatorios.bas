Attribute VB_Name = "Relatorios"
Sub imprContagem(seq_cnt As String)
    P_ImprimeContagem seq_cnt, frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprimeContagem seq_cnt, Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Sub P_ImprimeContagem(seq_cnt As String, obj As Object, Optional sOrd = "N")
    'N = nome; S = sequencia; C = Código
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rForn As New ADODB.Recordset
    On Error GoTo trata_Erro
    Dim iAux As Integer
    Dim iCount As Integer
    strSQL = "select tp.nome_pro,tp.unidade_pro,ti.*,tc.data_cnt from tab_produtos as tp, tab_itcont as ti, "
    strSQL = strSQL & " tab_contagens as tc where tc.seq_cnt=ti.seq_cnt "
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro and tc.seq_Cnt='" & seq_cnt & "' "
    If sOrd = "N" Then
        strSQL = strSQL & " order by nome_pro "
    ElseIf sOrd = "S" Then
        strSQL = strSQL & " order by data_itc "
    Else
        strSQL = strSQL & " order by cod_itc "
    End If
    abreObjetoAdo rPesq, strSQL
    
    If rPesq.RecordCount = 0 Then Exit Sub
    P_CabRelTimbr "Contagem: " & seq_cnt & " de: " & rPesq!data_cnt, obj
    
    
    P_ImprLinha obj, "-"
    P_ImprX obj, "Cód:", 0
    P_ImprX obj, "Produto:", 20
    P_ImprX obj, "UN.", 140
    
    P_ImprX obj, "Cont.", 170, True, True
    P_ImprX obj, "Data/hora", 175, False
    
    P_ImprX obj, "¦", 18
    P_ImprX obj, "¦", 138
    P_ImprX obj, "¦", 148
    P_ImprX obj, "¦", 172, True, True
    'P_ImprX obj, "+", 175
    P_ImprLinha obj, "-"
    With rPesq
        Do While Not .EOF
            
            If obj.CurrentY > obj.ScaleHeight - 30 Then
                If TypeName(obj) = "Printer" Then
                    P_RodapeVis obj, constNome_sis & " " & constVersao_Sis, obj.Page
                Else
                    P_RodapeVis obj, constNome_sis & " " & constVersao_Sis
                End If
                P_ImprX obj, "Cód:", 0
                P_ImprX obj, "Produto", 20
                P_ImprX obj, "UN.", 140
                
                P_ImprX obj, "Cont.", 170, True, True
                P_ImprX obj, "Data/hora", 175, False
                
                P_ImprLinha obj, "-"
            End If
            
            P_ImprX obj, "|", 18
            P_ImprX obj, "|", 138
            P_ImprX obj, "|", 148
            P_ImprX obj, "|", 172, True, True
            
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 20
            P_ImprX obj, !unidade_pro, 140
            
            P_ImprX obj, !qtd_itc, 170, True, True
            P_ImprX obj, Format(!data_itc, "dd/mm/yyyy hh:mm:ss"), 175, False
            
            .MoveNext
        Loop
        If TypeName(obj) = "Printer" Then
            P_RodapeVis obj, constNome_sis & " " & constVersao_Sis, obj.Page
        Else
            P_RodapeVis obj, constNome_sis & " " & constVersao_Sis
        End If
    
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Sub P_ImprimePerda(obj As Object, seq_per As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iCount As Integer
    Dim cTotal As Currency
    On erro GoTo trata_Erro
    strSQL = "select tp.*, tf.nome_for, tf.cnpj_for from tab_perdas as tp, tab_fornecedores as tf "
    strSQL = strSQL & " where tf.cod_for=tp.cod_for and tp.seq_per='" & seq_per & "'"
    abreObjetoAdo rPesq, strSQL
    P_CabRelTimbr "Produtos em Perda de " & rPesq!nome_for, obj
    
    strSQL = "select ti.*,tp.nome_pro,tp.cr_pro, (ti.qtd_itp*ti.unit_itp) as total from tab_produtos as tp, "
    strSQL = strSQL & "tab_itperdas ti, tab_perdas as ta where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & "and ta.seq_per=ti.seq_per "
    strSQL = strSQL & " and ta.seq_per='" & seq_per & "' order by nome_pro "
    abreObjetoAdo rPesq, strSQL
    iCount = 1
    
    P_ImprLinha obj, "-", 300
    P_ImprX obj, "...", 0
    P_ImprX obj, "Código", 10
    P_ImprX obj, "Produto", 30
    P_ImprX obj, "Quant.", 130
    P_ImprX obj, "Custo", 150
    P_ImprX obj, "Total", 190, False, 3
    P_ImprLinha obj, "-", 300
    With rPesq
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, constNome_sis & "-" & constVersao_Sis
                P_ImprX obj, "...", 0
                P_ImprX obj, "Código", 10
                P_ImprX obj, "Produto", 30
                P_ImprX obj, "Quant.", 130
                P_ImprX obj, "Custo", 150
                P_ImprX obj, "Total", 190, False, 3
                P_ImprLinha obj, "-", 300
            End If
            P_ImprX obj, CStr(iCount), 0
            P_ImprX obj, Format(!cod_pro, "00#\.####\-#"), 10
            P_ImprX obj, !nome_pro, 30
            P_ImprX obj, FormatVal(!qtd_itp), 130
            P_ImprX obj, FormatVal(!cr_pro), 150
            P_ImprX obj, FormatVal(!total), 190, False, 3
            cTotal = cTotal + FormatVal(!total)
            .MoveNext
            iCount = iCount + 1
        Loop
    End With
    P_ImprLinha obj, "-", 300
    P_ImprX obj, "Total: " & FormatVal(cTotal), 190, False, 3
    P_RodapeVis obj, constNome_sis & "-" & constVersao_Sis
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub


Sub imprimePerda(seq_per As String)
    P_ImprimePerda frmVisualiza.Pic, seq_per
    On Error Resume Next
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        For iCount = 1 To Int(frmVisualiza.txtbxCopias.Text)
        P_ImprimePerda Printer, seq_per
        Printer.EndDoc
        Next
    End If
    Unload frmVisualiza
End Sub

Sub imprimeConsign(seq_con As String)
    Dim iCount As Integer
    
    
        
        
    
    P_ImprimeConsign seq_con, frmVisualiza.Pic
    frmVisualiza.txtbxCopias.Text = "2"
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
        For iCount = 1 To Int(frmVisualiza.txtbxCopias.Text)
            P_ImprimeConsign seq_con, Printer
            Printer.EndDoc
        Next
    End If
    Unload frmVisualiza
    
    
    
End Sub
Sub P_ImprimeConsign(seq_con As String, obj As Object)
    Dim rPesq As New ADODB.Recordset
    Dim rItens As New ADODB.Recordset
    Dim strSQL As String
    Dim cTotal As Currency
    Dim sNome As String
    Dim yIni As Single
    Dim yFinal As Single
    Dim sUsu As String
    
    Dim bTapete As Boolean
    
    strSQL = "select tr.*,tc.* from tab_representantes as tr, tab_consignacoes as tc "
    strSQL = strSQL & " where tc.cod_rep=tr.cod_rep "
    strSQL = strSQL & " and tc.seq_con='" & seq_con & "' "
    strSQL = strSQL & " order by id desc  "
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        
        
        
        
        
        P_CabRelTimbr "Comprovante de Consignação Nº " & seq_con & " de " & !data_con, obj
        P_line1 obj
        P_ImprX obj, "Representante: " & !cod_Rep & " - " & !nome_rep, 5
        P_ImprX obj, "CPF: " & !cpf_rep, 105, False
        P_ImprX obj, "End: " & !endereco_rep, 5, False
        P_ImprX obj, "Bairro: " & !bairro_rep, 5
        P_ImprX obj, "Cidade: " & !cidade_rep & "/" & !uf_rep, 85, True
        P_ImprX obj, "CEP: " & !cep_rep, 170, False
        
        P_line1 obj
        If !retif_con = "S" Then
            
            obj.fontsize = 24
            P_ImprX obj, "RETIFICADORA", 70, False
            obj.fontsize = 10
            P_line1 obj
        
        End If
        
        
        P_line1 obj
        sNome = !nome_rep
    End With
    
    strSQL = "select tp.nome_pro,tp.ref_pro,tc.*,unit_con*reman_con as total from tab_produtos as tp, tab_consignacoes as tc "
    strSQL = strSQL & " where tc.cod_pro=tp.cod_pro "
    'strSQL = strSQL & " and tc.reman_con>0"
    strSQL = strSQL & " and tc.seq_con='" & seq_con & "' "
    strSQL = strSQL & " order by id desc "
    obj.Print ""
    
    rItens.Open strSQL, Conn
    P_line1 obj
    yIni = obj.CurrentY
    P_ImprX obj, "N.", 1, True, 1
    P_ImprX obj, "Ref.", 7, True, 1
    P_ImprX obj, "Produto.", 25, True, 1
    P_ImprX obj, "Qtd.", 160, True, esquerda
    P_ImprX obj, "Unit.", 185, True, direita
    P_ImprX obj, "Total", 205, False, direita
    P_line1 obj
    
    With rItens
        Do While Not .EOF
            If obj.CurrentY > obj.ScaleHeight - 30 Then
                P_RodapeVis obj, "Continua..."
            End If
            sUsu = !cod_usu
            P_ImprX obj, .AbsolutePosition, 1, True, esquerda
            P_ImprX obj, !ref_pro, 7, True, esquerda
            P_ImprX obj, !nome_pro, 25, True, esquerda
            P_ImprX obj, !qtd_con, 160, True, esquerda
            P_ImprX obj, FormatVal(!unit_con), 185, True, direita
            P_ImprX obj, FormatVal(!total), 205, False, direita
            cTotal = cTotal + !total
            P_line1 obj, vbDot
            
            
            If UCase(Left(!nome_pro, 6) = "TAPETE") Then
                bTapete = True
            End If
            
            
            
            .MoveNext
        Loop
        obj.FontBold = True
        yFinal = obj.CurrentY
        obj.Print
        P_ImprX obj, "Atendente: " & F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & sUsu & "'"), 5, True
        P_ImprX obj, "Impresso Por: " & strNome_Usu, 65, True
        P_ImprX obj, "Total: " & FormatVal(cTotal), 190, False, direita
        obj.FontBold = False
        
        
'        P_Grades obj, "24,156,170,184", yIni, yFinal, vertical, vbSolid
        
        P_line1 obj
        obj.Print
        If obj.CurrentY > obj.ScaleHeight - 30 Then
                P_RodapeVis obj, "Continua..."
            End If
            
        strSQL = "Eu, " & sNome & ", "
        strSQL = strSQL & "declaro ter recebido os produtos consignados "
        strSQL = strSQL & "aqui relacionados, nas quantidades e valores acima especificados."
        strSQL = strSQL & " Assinatura.:__________________________________________ em: _____/_____/______"
        
        
        P_ImprQuebra strSQL, obj, 2, 198
        obj.Print
        
        obj.FontBold = True
        obj.fontsize = 12
        If bTapete Then
            strSQL = "Atenção! Pegar o manual do modo de lavagem do(s) tapete(s) para o representante."
        
        
            P_ImprQuebra strSQL, obj, 2, 198
        
        
        End If
        
       obj.FontItalic = False
    obj.Print ""
    obj.fontsize = 14

    P_line1 obj
    
    
    P_CentraImpressao obj, "Não é Válido como Documento Fiscal nem como Garantia de Mercadoria"
    
    P_line1 obj
    obj.fontsize = 10
        
        
      '  P_RodapeVis obj, "NÃO É VÁLIDO COMO DOCUMENTO FISCAL NEM GARANTIA DE MERCADORIA", ""
        .Close
    End With
    rPesq.Close
End Sub



Sub P_ImprimeVenda(obj As Object, num_ven As String, _
    Optional bPrinter As Boolean = True)
    Dim rPesq As New ADODB.Recordset
    Dim rPro As New ADODB.Recordset
    Dim rPre As New ADODB.Recordset
    Dim rExt As New ADODB.Recordset
    Dim cTotal As Currency
    Dim alturaPag As Single
    Dim iPag As Integer
    Dim yAux As Single
    Dim rForn As New ADODB.Recordset
    Dim rLan As New ADODB.Recordset
    Dim strSQL As String
    Dim sMsg As String
    Dim rTroca As New ADODB.Recordset
    Dim cTxJuro As Currency '<<= o juro buscado da forma de pagamento pra aumentar o preço
                            'dos produtos
    'On Error Resume Next
    obj.ScaleMode = 1
    Altura_Pagina = 9000
    
    strSQL = "select tv.*,tf.valorfinal_ven as final, logradouro_cli,numero_cli,complemento_cli,tf.desconto_ven as desconto, "
    strSQL = strSQL & " tc.nome_cli,tc.endereco_cli,tc.uf_cli "
    strSQL = strSQL & " from tab_vendas as tv, tab_clientes as tc,tab_venfec as tf  "
    strSQL = strSQL & " where tc.cod_cli=tv.cod_cli and tv.num_ven='" & num_ven & "' "
    strSQL = strSQL & " and tv.num_ven=tf.num_ven"
    abreObjetoAdo rPesq, strSQL
    strSQL = ""
    If rPesq.RecordCount = 0 Then 'a venda tá aberta
        strSQL = "select tv.*, 0 as  final, 0 as  desconto, "
        strSQL = strSQL & " tc.nome_cli,tc.endereco_cli,tc.uf_cli "
        strSQL = strSQL & " from tab_vendas as tv, tab_clientes as tc "
        strSQL = strSQL & " where tc.cod_cli=tv.cod_cli and tv.num_ven='" & num_ven & "' "
        
        abreObjetoAdo rPesq, strSQL
    End If
    strSQL = (F_BuscaUmCampo("juros_tip", "tab_tipoven", "cod_tip='" & _
        rPesq!cod_tip & "'"))
    
    If strSQL = "" Then
        cTxJuro = 1
    Else
        cTxJuro = CCur(strSQL) / 100
    End If
    'produtos
    strSQL = "select tp.nome_pro,tv.* from tab_produtos as tp, tab_itvendas as tv "
    strSQL = strSQL & " where tv.cod_pro=tp.cod_pro and tv.Num_Ven = '" & num_ven & "' "
    Set rPro = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    'externos
'    strSQL = "select tp.nome_pro,tv.* from tab_produtos as tp, tab_venped as tv "
'    strSQL = strSQL & " where tv.cod_pro=tp.cod_pro and tv.Num_Ven = '" & Num_Ven & "' "
'    Set rExt = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    'Prevenda
'    strSQL = "select tp.nome_pro,tv.* from tab_produtos as tp, tab_prevendas as tv "
'    strSQL = strSQL & " where tv.cod_pro=tp.cod_pro and tv.Num_Ven = '" & Num_Ven & "' "
'    Set rPre = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    P_CabRelTimbr "Orç:" & num_ven & " Cliente:" & rPesq!nome_cli, obj
    'P_ImprX obj, "Venda:" & Num_Ven & " Cliente:" & rPesq!nome_cli, 5, False
    'P_ImprLinha obj
    P_ImprX obj, "Cód.:", 5
    P_ImprX obj, "Descrição:", 20
    P_ImprX obj, "Qtd.:", 150, True, True
    P_ImprX obj, "Unit.:", 160, True, True
    P_ImprX obj, "Total.:", 180, False, 3
    P_ImprLinha obj
    With rPro
        Do While Not .EOF
            
            
            If bPrinter = True Then
                If Printer.CurrentY >= Printer.ScaleHeight - 40 Then
                
                    P_FimRel obj, constNome_sis
                    
                    P_ImprX obj, "Cód.:", 5
                    P_ImprX obj, "Descrição:", 20
                    P_ImprX obj, "Qtd.:", 140, True, True
                    P_ImprX obj, "Unit.:", 150, True, True
                    P_ImprX obj, "Desc.:", 170, True, True
                    P_ImprX obj, "Total.:", 190, False, 3
                    P_ImprLinha obj
                End If
            Else
                obj.ScaleMode = 1
                If obj.CurrentY - (iPag * obj.CurrentY) >= Altura_Pagina + 40 Then
                    P_FimRel obj, constNome_sis
                    P_ImprX obj, "Cód.:", 5
                    P_ImprX obj, "Descrição:", 20
                    P_ImprX obj, "Qtd.:", 135, True, True
                    P_ImprX obj, "Unit.:", 150, True, True
                    P_ImprX obj, "Desc.:", 170, True, True
                    P_ImprX obj, "Total.:", 190, False, 3
                    P_ImprLinha obj
                    P_ImprLinha obj, "_"
                    iPag = iPag + 1
                End If
                yAux = yAux + obj.CurrentY
                
            End If
            
            P_ImprX obj, !cod_pro, 5
            P_ImprX obj, !nome_pro, 20
            P_ImprX obj, FormatVal(!qtd_ite), 135, True, True
            P_ImprX obj, FormatVal(!unit_ite), 150, True, True
            P_ImprX obj, FormatVal(!desconto_ite), 150, True, True
            P_ImprX obj, FormatVal(!final_ite), 190, False, 3
            
            cTotal = cTotal + (!final_ite)
            
            .MoveNext
        Loop
            
        
    End With
    
        With rPesq
            
            
            obj.Print ""

            obj.FontBold = True
            P_ImprLinha obj
            P_ImprX obj, "Total dos Produtos: " & FormatVal(cTotal), 80
            strSQL = F_BuscaUmCampo("cod_tip", "tab_venfec", "num_ven='" & !num_ven & "'")
            strSQL = F_BuscaUmCampo("Descricao_TIp", _
                        "tab_tipoven", "cod_tip='" & strSQL & "'")
            If strSQL = "" Or strSQL = "0" Then
                strSQL = "A Vista"
                
            End If
            P_ImprX obj, "Forma de Pgto: " & strSQL, 5
                       
            
            P_ImprX obj, " Desconto: " & FormatVal(rPesq!desconto), 130
            obj.FontBold = True
            P_ImprX obj, "Valor Final: " & FormatVal(rPesq!final), 200, False, 3
            obj.FontBold = False
            P_ImprLinha obj
            
            
            strSQL = "select ti.* from tab_doc as td, tab_itdoc as ti, tab_vendas as tv "
            strSQL = strSQL & " where tv.num_ven=td.num_ven and td.seq_doc=ti.seq_doc"
            strSQL = strSQL & " and td.tipo_doc='R' and td.num_ven='" & num_ven & "' "
            strSQL = strSQL & " and td.esp_doc='V' "
            strSQL = strSQL & " order by venc_itd "
            Set rLan = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
            With rLan
                strSQL = ""
                If .RecordCount > 0 Then
                    sMsg = "Parcelas: "
                End If
                Do While Not .EOF
                    sMsg = sMsg & !venc_itd & "- R$ " & FormatVal(!valor_itd) & Chr(9)
                    .MoveNext
                Loop
                .Close
            End With
            
            strSQL = "select ti.* from tab_doc as td, tab_cheques as ti, tab_vendas as tv "
            strSQL = strSQL & " where tv.num_ven=td.num_ven and td.seq_doc=ti.seq_doc"
            strSQL = strSQL & " and td.tipo_doc='R' and td.num_ven='" & num_ven & "' "
            strSQL = strSQL & " order by venc_che "
            Set rLan = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
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
            strSQL = "select * from tab_trocas where num_ven='" & num_ven & "'  order by data_tro "
            Set rTroca = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
            With rTroca
                Do While Not .EOF
                    P_ImprX obj, "Troca: " & !seq_tro, 0
                    P_ImprX obj, "Data: " & !data_tro, 30
                    P_ImprX obj, "A Abater: " & FormatVal(!abat_tro), 80
                    P_ImprX obj, "A Acrescer: " & FormatVal(!abat_tro), 110
                    P_ImprX obj, "Final: " & FormatVal(!acres_tro - !abat_tro), 150, False
                    .MoveNext
                Loop
                If .RecordCount > 0 Then
                    P_ImprLinha obj
                End If
                .Close
            End With
            If sMsg <> "" Then
                P_ImprX obj, sMsg, 0, False
                P_ImprLinha obj
            End If
            P_ImprX obj, "Vendedor: " & F_BuscaUmCampo("Guerra_Fun", "tab_funcionarios", _
                "Cod_fun='" & rPesq!cod_fun & "'"), 5, True
            P_ImprX obj, "Orçamento Válido por 5 Dias", 200, False, 3
            
            
            P_ImprLinha obj, "_"
            
            P_CentraImpressao obj, "NÃO É DOCUMENTO FISCAL"
            P_CentraImpressao obj, "NÃO VALE COMO GARANTIA DE MERCADORIA"
            
            P_FimRel obj, constNome_sis & "   -   " & ConstEmpresa
            
        End With
    
    rPesq.Close
    rPro.Close
    
    
    If bPrinter Then
        Printer.EndDoc
    End If
    
End Sub
Sub ImprOrc(num_orc As String)
    imprimeOrc frmVisualiza.Pic, num_orc
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprimeOrc Printer, num_orc
        Printer.EndDoc

    End If
            Unload frmVisualiza
End Sub
Sub ImprMovInt(num_mov As String)
    imprimeMovInt frmVisualiza.Pic, num_mov
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprimeMovInt Printer, num_mov
        Printer.EndDoc

    End If
            Unload frmVisualiza
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
Sub ImprReposicao(seq_rep As String)
    Dim iCount As Integer
    imprimeReposicao frmVisualiza.Pic, seq_rep
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
        For iCount = 1 To Int(frmVisualiza.txtbxCopias.Text)
            imprimeReposicao Printer, seq_rep
            Printer.EndDoc
        Next
    End If
            Unload frmVisualiza
End Sub



Sub ImprConsig(seq_con As String)
    
    P_ImprimeConsign seq_con, frmVisualiza.Pic
    frmVisualiza.txtbxCopias.Text = "2"
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        For iCount = 1 To Int(frmVisualiza.txtbxCopias.Text)
        P_ImprimeConsign seq_con, Printer
        Printer.EndDoc
        Next

    End If
    Unload frmVisualiza
End Sub

Sub imprimeOrc(obj As Object, num_orc As String)
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim cTotal As Currency
    strSQL = "select ti.*,tor.*,ti.unit_ite*ti.qtd_ite as total,tp.nome_pro"
    strSQL = strSQL & " from tab_orc as tor, tab_itorc as ti, "
    strSQL = strSQL & " tab_produtos as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tor.num_orc=ti.num_orc "
    strSQL = strSQL & " and tor.num_orc='" & num_orc & "' "
    strSQL = strSQL & " order by ti.seq_ite "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then Exit Sub
        P_CabRelTimbr "Orçamento: " & num_orc, obj
        P_ImprX obj, "Cliente:" & !cod_cli & " - " & !nomecli_orc, 0, True
        P_ImprX obj, "Data:" & !data_orc, 180, False, 3
        P_ImprX obj, "Validade:" & !validade_orc, 0
        P_ImprX obj, "Cond. Pagamento:" & FormatVal(!obs_orc), 50, False
        P_ImprLinha obj
        P_ImprX obj, "...", 0, , , 1
        P_ImprX obj, "Código:", 10, , , 1
        P_ImprX obj, "Descrição:", 30, , , 1
        P_ImprX obj, "Qtd:", 140, True, True, 1
        P_ImprX obj, "Unit:", 160, True, True, 1
        P_ImprX obj, "Total:", 180, False, 3, 1
        P_ImprLinha obj
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_FimRel obj, "Continua..."
                P_ImprX obj, "...", 0, , , 1
                P_ImprX obj, "Código:", 10, , , 1
                P_ImprX obj, "Descrição:", 30, , , 1
                P_ImprX obj, "Qtd:", 140, True, True, 1
                P_ImprX obj, "Unit:", 160, True, True, 1
                P_ImprX obj, "Total:", 180, False, 3, 1
                P_ImprLinha obj
            End If
            P_ImprX obj, !seq_ite, 0, , , 1
            P_ImprX obj, FormatCodigo(!cod_pro), 10, , , 1
            P_ImprX obj, !nome_pro, 30, , , 1
            P_ImprX obj, FormatVal(!qtd_ite), 140, True, True, 1
            P_ImprX obj, FormatVal(!unit_ite), 160, True, True, 1
            P_ImprX obj, FormatVal(!total), 180, False, 3, 1
            cTotal = cTotal + !total
            
            .MoveNext
        Loop
        P_ImprLinha obj, "_"
        obj.fontsize = 20
        P_ImprX obj, "Total...........: " & FormatVal(cTotal), 180, False, 3
        obj.fontsize = 10
        P_FimRel obj, strNome_Emp
    End With
End Sub
Sub imprimeMovInt(obj As Object, num_mov As String, Optional bPrecoVenda As Boolean = True)
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cTotCust As Currency
    obj.FontName = "arial"
    strSQL = "select ti.*,td.*,tm.*,tp.pv_pro*ti.qtd_ite as total,tp.pv_pro,tp.nome_pro,cr_pro "
    strSQL = strSQL & " from tab_movint as tm, tab_itmov as ti,tab_destmovint as td, "
    strSQL = strSQL & " tab_produtos as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and td.cod_des=tm.cod_des "
    strSQL = strSQL & " and tm.num_mov=ti.num_mov "
    strSQL = strSQL & " and tm.num_mov='" & num_mov & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then Exit Sub
        
        P_CabRelTimbr "Movimento Interno(" & IIf(!tipo_mov = "C", "CONSUMO", "TRANSFERÊNCIA") & "): " & num_mov, obj
                
        P_ImprX obj, "Tipo: " & IIf(!tipo_mov = "C", "CONSUMO", "TRANSFERÊNCIA"), 0, True
        P_ImprX obj, "Destino: " & !nome_des, 40
        
        P_ImprX obj, "Protocolo: " & !protocolo_mov, 120
        P_ImprX obj, "Data Saída: " & !data_mov, 200, False, 3
        
        
        
        P_ImprX obj, "Destino: " & !nome_des, 0, True
        P_ImprX obj, "Funcionário: " & !cod_usu, 80, True
        
        
        P_ImprX obj, "Origem: " & strCNPJ_Emp & " " & strFantasia_Emp, 120, False
        
        P_ImprX obj, "Obs: " & " " & !obs_mov, 10, False
        
        
        P_ImprLinha obj
        P_ImprX obj, "...", 0, , , 1
        P_ImprX obj, "Código:", 10, , , 1
        P_ImprX obj, "Descrição:", 30, , , 1
        P_ImprX obj, "Qtd:", 140, True, True, 1
        P_ImprX obj, "Unit.:", 160, True, True, 1
        
        P_ImprX obj, "Total:", 180, True, True, 1
        
        
            P_ImprX obj, "", 200, False, 3, 1
        
        P_ImprLinha obj
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_FimRel obj, "Continua..."
                P_ImprX obj, "...", 0, , , 1
                P_ImprX obj, "Código:", 10, , , 1
                P_ImprX obj, "Descrição:", 30, , , 1
                P_ImprX obj, "Qtd:", 140, True, True, 1
                P_ImprX obj, "Unit.:", 160, True, True, 1
                P_ImprX obj, "Total:", 180, True, True, 1
                
                P_ImprX obj, "", 200, False, 3, 1
                
                P_ImprLinha obj
            End If
            P_ImprX obj, !seq_ite, 0, , , 1
            P_ImprX obj, !cod_pro, 10, , , 1
            P_ImprX obj, !nome_pro, 30, False, , 1
            P_ImprX obj, FormatVal(!qtd_ite), 140, True, True, 1
            P_ImprX obj, FormatVal(!pv_pro), 160, True, True, 1
            P_ImprX obj, FormatVal(!total), 180, True, True, 1
            
            P_ImprX obj, "", 200, False, 3, 1
            
            cTotal = cTotal + !total
            cTotCust = cTotCust + (!cr_pro * !qtd_ite)
            .MoveNext
        Loop
        P_ImprLinha obj, "_"
        obj.fontsize = 20
        'P_ImprX obj, "Custo...........: " & FormatVal(cTotCust), 5
        P_ImprX obj, "Total...........: " & FormatVal(cTotal), 200, False, 3
        obj.fontsize = 10
        P_FimRel obj, strNome_Emp
    End With
End Sub

Sub ImprimeComprovanteNaoFiscal(num_ven As String, obj As Object)
    Dim pOld
    Dim s As String
    Dim strSQL As String
    pOld = BuscaImpressoraPadrao
    Set Printer = Printers(ImpressoraPadraoCupomNFCe.Indice)
    
    strSQL = "select * from "
    
    obj.FontName = "Arial"
    cQRCode = App.path & "\qrcode.tmp"
    
    Set Printer = pOld
    
    
    imprimeLogo obj
    obj.Print
    
    
    obj.fontsize = 12
    s = strFantasia_Emp
    obj.FontBold = True
    P_CentraImpressao obj, s, False
    
    obj.FontBold = False
    s = strCNPJ_Emp
    P_CentraImpressao obj, "CNPJ:" & F_FormataCNPJ(strCNPJ_Emp) & " IE: " & strIE_Emp, False
    
    
    
    s = strLogradouro_Emp
    s = s & " Nº " & strNumero_Emp
    s = s & " " & strComplemento_Emp
    
    P_CentraImpressao obj, s, False
    
    s = strBairro_Emp
    s = s & " CEP:" & strCEP_Emp
    P_CentraImpressao obj, s, False
    
    s = strCidade_Emp
    s = s & "/" & strUF_Emp
    P_CentraImpressao obj, s, False
    P_CentraImpressao obj, "Fone:" & F_AjustaTelefone(strFone_Emp), False
    
    P_line1 obj, vbDot
    obj.Print
    obj.FontBold = True
    obj.fontsize = "7"
    P_CentraImpressao obj, "SEM VALOR FISCAL"
    P_CentraImpressao obj, "DOCUMENTO DE SIMPLES CONFERÊNCIA"
    obj.Print
    obj.FontBold = False
    P_line1 obj, vbDot
    P_line1 obj, vbDot
    obj.FontBold = True
    obj.Print
    P_CentraImpressao obj, "CLIENTE"
    obj.FontBold = False
    Dim cpf As String
    Dim nome As String
        
        
    
    
    
    
    
    P_line1 obj, vbDashDotDot
    obj.Print
    P_CentraImpressao obj, "Produtos Adquiridos"
    'P_line1 obj, vbDot
    obj.Print
    
    iProd = 1
    obj.FontBold = True
    obj.fontsize = 8
    
    
    P_ImprX obj, "Código", 0
    P_ImprX obj, "Descrição", 14
    'P_ImprX obj, "Und", 10
    P_ImprX obj, "Qtd", 46
    P_ImprX obj, "Unit", 55
    P_ImprX obj, "Total", 63, False
    P_line1 obj, vbDot
    obj.FontBold = False
    obj.FontName = "Arial Narrow"
    det = "1"
    obj.FontName = "Arial"
    P_line1 obj, vbDot
    obj.Print
    obj.FontBold = True
    
    obj.Print
    P_CentraImpressao obj, "Valor Aproximado dos Tributos (Lei. 12.741)"
    obj.FontBold = False
    P_ImprX obj, "Estadual: " & FormatVal(ValorMedioImpostoEstadual / 100 * CCur(vNF)), 10
    P_ImprX obj, "Federal: " & FormatVal(ValorMedioImpostoFederal / 100 * CCur(vNF)), 40, False
    
    
    obj.Print
    obj.FontBold = False
    P_line1 obj, vbDot
    'ImprimeQRCodeArquivo qrCode, cQRCode
    
    
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
    Dim rDoc As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    
    P_Status "Localizando Venda"
    
    strSQL = "select tv.*,tc.*,tf.nome_for from " & TAB_VENDAS & " as tv inner join tab_formaspag as tf  on tv.cod_for = tf.cod_For "
    strSQL = strSQL & " left join " & TAB_CAIXA & " as tc on tc.num_ven=tv.num_ven "
    strSQL = strSQL & " where tv.num_ven='" & num_ven & "' "
    
    abreObjetoAdo rVendas, strSQL
    
    
    
    P_Status "Localizando Cliente"
    
    strSQL = "select tc.nome_cli,cpf_cli,endereco_cli,logradouro_cli,numero_cli,complemento_cli,cidade_cli,ti.*,tv.*,ti.unit_ite*ti.qtd_ite as total,tp.ref_pro,tf.guerra_fun,tp.nome_pro"
    strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti, tab_clientes as tc, "
    strSQL = strSQL & " tab_produtos as tp,tab_funcionarios as tf where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and tv.num_ven='" & num_ven & "' "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " and tf.cod_fun=abrevenda_usu "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
    
    P_Status "Preparando impressão"
    
        If .RecordCount = 0 Then Exit Sub
        sSituacao_Ven = !situacao_ven
        If !cod_For <> "01" Then
            P_CabRelTimbr "ORÇAMENTO: " & num_ven, obj
        Else
            P_CabRelTimbr "ORÇAMENTO: " & num_ven, obj
        End If
        If Int(sSituacao_Ven) < 4 Then
            obj.fontsize = 24
            P_CentraImpressao obj, "VENDA NÃO CONCLUÍDA", False
            P_line1 obj, vbDot
        ElseIf Int(sSituacao_Ven) = 5 Then
            obj.fontsize = 14
            P_CentraImpressao obj, "VENDA CANCELADA", False
            P_line1 obj, vbDot
        End If
        obj.fontsize = 10
        P_ImprX obj, "Cliente:" & !cod_cli & " - " & !nomecli_ven, 0, True
        P_ImprX obj, "CPF:" & !cpf_cli, 120, True
        P_ImprX obj, "Data:" & !data_ven, 200, False, 3
        P_ImprX obj, "Endereço:" & !Logradouro_cli & " N°" & rPesq!numero_cli & " " & rPesq!complemento_cli, 0
        P_ImprX obj, "Bairro:" & "", 120, False
        P_ImprX obj, "Cidade:" & !cidade_cli, 0
        P_ImprX obj, "Cep:" & "", 60, True
        P_ImprX obj, "Fone:" & TiraNull(!telefonecli_ven), 80, False
        sVendedor = !guerra_fun
        sNomeCli = !nomecli_ven
        
        
        If !entrega_ven = "S" Then
            bEntrega = True
            sEnd = TiraNull(!endereco_ven)
            sBai = TiraNull(!bairro_ven)
            sFone = TiraNull(!fone_ven)
        End If
        
        cTotal = !valor_ven
        sNum_ven = !num_ven
        sCond = F_BuscaUmCampo("cod_tip", "tab_venfec", "num_ven='" & num_ven & "'")
        sCond = !cod_For
        
        If sCond <> "" Then
            sCond = F_BuscaUmCampo("descricao_tip", "tab_tipoven", "cod_tip='" & sCond & "'")
            P_ImprX obj, "Cond. Pagamento:" & sCond, 0, False
        End If
        P_ImprLinha obj
        P_ImprX obj, "...", 0
        P_ImprX obj, "Ref:", 10
        P_ImprX obj, "Descrição:", 30
        P_ImprX obj, "Qtd:", 110, True, True
        P_ImprX obj, "Unit:", 130, True, True
        P_ImprX obj, "Prom:", 150, True, True
        P_ImprX obj, "Desc:", 170, True, True
        P_ImprX obj, "Final:", 195, False, 3
        P_ImprLinha obj
        cTotal = 0
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_FimRel obj, "Continua..."
                P_ImprX obj, "...", 0
                P_ImprX obj, "Ref:", 10
                P_ImprX obj, "Descrição:", 30
                P_ImprX obj, "Qtd:", 110, True, True
                P_ImprX obj, "Prom:", 130, True, True
                P_ImprX obj, "Prom:", 150, True, True
                P_ImprX obj, "Desc:", 170, True, True
                P_ImprX obj, "Total:", 195, False, 3
                P_ImprLinha obj
            End If
            P_ImprX obj, !seq_ite, 0, , , 1
            P_ImprX obj, (!ref_pro), 10, , , 1
            obj.fontsize = 8
            P_ImprX obj, StrConv(!nome_pro, vbProperCase), 30
            obj.fontsize = 10
            P_ImprX obj, !qtd_ite, 115, True, True
            P_ImprX obj, FormatVal(!unit_ite), 130, True, True
            If !pq_itd > 0 Then
                P_ImprX obj, FormatVal(!pq_itd), 150, True, True
            End If
            P_ImprX obj, FormatVal(!desconto_ite), 170, True, True
            P_ImprX obj, FormatVal(!final_ite), 195, True, 3
            If !pq_itd > 0 Then
                P_ImprX obj, "Q", 200, False, direita
            Else
                P_ImprX obj, "", 200, False, direita
            End If
            totalPro = totalPro + (!qtd_ite * !unit_ite)
            cTotal = cTotal + (!final_ite)
            cDesc = cDesc + !desconto_ite
            
            .MoveNext
        Loop
        P_ImprLinha obj, "_"
        
        'P_ImprX obj, "Reconheço que pagarei a dívida acima.", 20, False
        obj.Print
        
        
        
        
        
        
        strSQL = "select tt.data_tro,tt.seq_tro,tp.nome_pro,ti.* from tab_produtos as tp, tab_trocas as tt, tab_ittrocas as ti"
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
                P_ImprX obj, "Trouxe:" & !qtdant_itt, 50, True
                P_ImprX obj, "Levou:" & !qtdatual_itt, 100, False
                
            End If
                
                .MoveNext
            Loop
            .Close
        End With
        
        P_ImprLinha obj, "_"
        
        
        P_CentraImpressao obj, String(60, "_")
               
        
        P_CentraImpressao obj, sNomeCli
        
        
        With rPesq
            
            
            obj.Print ""
            obj.fontsize = 12
            obj.FontBold = True
            P_ImprLinha obj
            P_ImprX obj, "Total dos Produtos (sem desconto): " & FormatVal(totalPro), 10
            P_ImprX obj, "Vlr Final: " & FormatVal(cTotal), 120
            strSQL = F_BuscaUmCampo("cod_for", "tab_venfec", "num_ven='" & sNum_ven & "'")
            strSQL = F_BuscaUmCampo("Descricao_TIp", _
                        "tab_tipoven", "cod_tip='" & strSQL & "'")
            If strSQL = "" Or strSQL = "0" Then
                strSQL = "A Vista"
                
            End If
            obj.Print " <---"
            obj.fontsize = 10
            P_ImprX obj, "Forma de Pgto: " & rVendas!nome_for, 5
                       
            
            P_ImprX obj, " Desconto: " & FormatVal(cDesc), 130
            obj.FontBold = True
            P_ImprX obj, "Valor Final: " & FormatVal(totalPro - cDesc), 200, False, 3
            obj.FontBold = False
            P_ImprLinha obj
            
            P_Status "Localizando Documento"
            
            strSQL = "select ti.* from tab_doc as td, tab_itdoc as ti, tab_vendas as tv "
            strSQL = strSQL & " where tv.num_ven=td.num_ven and td.seq_doc=ti.seq_doc"
            strSQL = strSQL & " and td.tipo_doc='R' and td.num_ven='" & num_ven & "' and esp_doc='V'"
            strSQL = strSQL & " order by venc_itd "
            
            abreObjetoAdo rDoc, strSQL
            With rDoc
                strSQL = ""
                If .RecordCount > 0 Then
                    sMsg = "Parcelas: "
                End If
                Do While Not .EOF
                    sMsg = sMsg & !venc_itd & "- R$ " & FormatVal(!valor_itd) & Chr(9)
                    .MoveNext
                Loop
                .Close
            End With
            
            strSQL = "select * from tab_venfec where num_ven='" & num_ven & "'"
            abreObjetoAdo rDoc, strSQL
            strSQL = ""
            With rDoc
                Do While Not .EOF
                    strSQL = strSQL & !esp_cai & ": " & FormatVal(!valor_ven) & " "
                    .MoveNext
                Loop
                .Close
            End With
            If strSQL <> "" Then
                P_ImprX obj, "Pagamentos: " & strSQL, 5, False
            End If
            
            
            
            strSQL = "select ti.* from tab_doc as td, tab_cheques as ti, tab_vendas as tv "
            strSQL = strSQL & " where tv.num_ven=td.num_ven and td.seq_doc=ti.seq_doc"
            strSQL = strSQL & " and td.tipo_doc='R' and td.num_ven='" & num_ven & "' "
            strSQL = strSQL & " order by venc_che "
            abreObjetoAdo rDoc, strSQL
            With rDoc
                
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
                P_ImprQuebraTurbo sMsg, obj, 5, 190, False
                P_ImprLinha obj
            End If
            P_ImprX obj, "Vendedor: " & F_BuscaUmCampo("Guerra_Fun", "tab_funcionarios", _
                "cod_fun='" & rVendas!abrevenda_usu & "'"), 5, False
                        
            
            
            
            yAux = obj.CurrentY
            
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
    
    
        P_ImprX obj, "Vendedor: " & sVendedor, 5, False
        obj.CurrentY = yAux - 5
        obj.fontsize = 14
        If rVendas!valorentrega_ven > 0 Then
            P_ImprX obj, "Taxa de Entrega: " & FormatVal(rVendas!valorentrega_ven), 180, False, 3
        End If
        obj.fontsize = 20
        If rVendas!situacao_ven <= 2 Then
            P_ImprX obj, "Total...........: " & FormatVal(cTotal + rVendas!valorentrega_ven), 180, False, 3
        
        Else
            
            P_ImprX obj, "Total...........: " & FormatVal(rVendas!valor_ven + rVendas!valorentrega_ven), 180, False, 3
        End If
        obj.fontsize = 10
        
        If bEntrega Then
            P_line1 obj
            obj.fontsize = 14
            P_CentraImpressao obj, "ENTREGA"
            obj.fontsize = 8
            obj.Print
            P_ImprX obj, "Endereço: " & sEnd, 0
            P_ImprX obj, "Bairro: " & sBai, 90
            P_ImprX obj, "Fone: " & sFone, 160, False
            P_line1 obj
            
        End If
        
        
        P_FimRel obj, "NÃO É VÁLIDO COMO DOCUMENTO FISCAL"
        P_Status ""
    rPesq.Close
    rVendas.Close
    End With
    Exit Sub
trata_Erro:
    P_Status ""
    P_MostraErro
End Sub
Sub imprimeReposicao(obj As Object, seq_rep As String)
    Dim rPesq As New ADODB.Recordset
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
    Dim sAtendente As String

    'On Error Resume Next
    
    strSQL = "select tr.*,tf.guerra_fun as solicitante from tab_reposicao as tr, "
    
    strSQL = strSQL & "tab_funcionarios as tf "
    strSQL = strSQL & " where tf.cod_fun=tr.solicitante_rep "
    strSQL = strSQL & " and tr.seq_rep='" & seq_rep & "' "
    
    
    
    Set rVendas = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    
    sAtendente = F_BuscaUmCampo("guerra_fun", "tab_Funcionarios", "cod_fun='" & rVendas!atendente_rep & "'")
    
    
    strSQL = Reposicao.F_BuscaItensRep(seq_rep)
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then Exit Sub
        sSituacao_Ven = rVendas!status_rep
        
        P_CabRelTimbr "SOLICITAÇÃO DE REPOSIÇÃO: " & seq_rep, obj
        
        P_ImprX obj, "Solicitante:" & rVendas!solicitante, 5, False
        
        P_ImprX obj, "Data:" & rVendas!abertura_rep, 5, False, esquerda
        
        P_ImprX obj, "Atendente:" & sAtendente, 5, False, esquerda
        
        
        
        
        
        
        
        
        
        P_ImprLinha obj
        P_ImprX obj, "...", 0
        P_ImprX obj, "Ref:", 10
        P_ImprX obj, "Descrição:", 30
        P_ImprX obj, "Qtd:", 150, True, direita
        P_ImprX obj, "Aferição", 200, False, direita
        
        P_ImprLinha obj
        cTotal = 0
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_FimRel obj, "Continua..."
                P_ImprX obj, "...", 0
                P_ImprX obj, "Ref:", 10
                P_ImprX obj, "Descrição:", 30
                P_ImprX obj, "Qtd:", 150, True, direita
                P_ImprX obj, "Aferição", 200, False, direita
                
                P_ImprLinha obj
            End If
            P_ImprX obj, !seq_ite, 0
            P_ImprX obj, !ref_pro, 10
            P_ImprX obj, !nome_pro, 30
            P_ImprX obj, !qtd_ite, 150, True, direita
            P_ImprX obj, "_________________", 200, False, direita
                
            
            .MoveNext
        Loop
        obj.Print
        P_ImprLinha obj, "_"
        
        
        
            
            
            If sSituacao_Ven = 1 Then
                P_ImprX obj, "(Solicitação em Aberto impresso por " & strNome_Usu & ")", 5, False
            ElseIf sSituacao_Ven = 2 Then
                P_ImprX obj, "(Solicitação Aguardando Atendimento por " & strNome_Usu & ")", 5, False
            ElseIf sSituacao_Ven = 3 Then
                P_ImprX obj, "(Solicitação Aguardando Efetivação por " & strNome_Usu & ")", 5, False
            ElseIf sSituacao_Ven = 4 Then
                P_ImprX obj, "(Solicitação Finalizada impresso por " & strNome_Usu & ")", 5, False
            ElseIf sSituacao_Ven = 5 Then
                P_ImprX obj, "(Solicitação Cancelada impresso por " & strNome_Usu & ")", 5, False
            End If
            
        End With
        obj.Print
    
        
        
        
        
        
        
        
        
        
        P_FimRel obj, "DOCUMENTO APENAS PARA CONTROLE INTERNO"
    rPesq.Close
    rVendas.Close
    
End Sub

Sub imprimeListagemPrecos(dtobj As data)
    P_ImprimeListagem dtobj, frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprimeListagem dtobj, Printer
        Printer.EndDoc
        
    End If
    Unload frmVisualiza
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
            P_ImprX obj, !Nome_Sec, 148
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

Sub P_ImprNotaFiscal(obj As Object, seq_not As String, _
    Optional tipo_ent As Integer = 0)
    Dim strSQL As String
    Dim sNome_For As String
    Dim rPesq As New ADODB.Recordset
    Dim iCount As Integer
   ' On Error Resume Next
    strSQL = "select tp.cod_pro,ti.*,"
    strSQL = strSQL & "sum(ti.qtd_ite*ti.unit_ite) as soma, te.*,tp.nome_pro "
    strSQL = strSQL & " from tab_produtos as tp, "
    If tipo_ent = 0 Then
        strSQL = strSQL & "  tab_cmp as te, tab_itcmp as ti "
    Else
        strSQL = strSQL & " tab_compras as te, tab_itcompras as ti "
    End If
    strSQL = strSQL & " where ti.seq_not=te.seq_not "
    
    
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and te.seq_not='" & seq_not & "' "
    strSQL = strSQL & " group by ti.cod_pro "
    strSQL = strSQL & " order by ti.seq_ite "
    abreObjetoAdo rPesq, strSQL
    sNome_For = F_BuscaUmCampo("Nome_For", "tab_fornecedores", "cod_for='" & _
        cod_For & "'")
    With rPesq
        If .RecordCount < 1 Then Exit Sub
        If tipo_ent = 0 Then
            P_CabRelTimbr "Extrato de Nota Fiscal " & Num_not, obj
        Else
            P_CabRelTimbr "Extrato de Compra " & Num_not, obj
        End If
        obj.fontsize = 12
        
        P_ImprX obj, sNome_For, 0, False
        P_ImprLinha obj
        obj.fontsize = 9
        P_ImprX obj, "Número: " & Num_not & " S: " & serie_not, 0, True
        P_ImprX obj, "Valor: " & FormatVal(!valor_not), 45
        P_ImprX obj, "BC: " & FormatVal(!bc_not), 70
        P_ImprX obj, "ICMS: " & FormatVal(!icms_not), 90
        P_ImprX obj, "IPI: " & FormatVal(!ipi_not), 115, True
        P_ImprX obj, "Desp: " & FormatVal(!despesas_not), 135, True
        P_ImprX obj, "Frete: " & FormatVal(!frete_not), 155, True
        P_ImprX obj, "Entrada: " & Format(!entrada_not), 175, False
        P_ImprLinha obj, "-"
        P_ImprX obj, "Cód", 0
        P_ImprX obj, "Produto", 18
    '    P_ImprX obj, "Lote", 100, True, True
        P_ImprX obj, "Unitário", 130, True, True
        P_ImprX obj, "Qtd", 150, True, True
        'P_ImprX obj, "IPI", 165, True, True
        P_ImprX obj, "Gastos", 162, True, True
        P_ImprX obj, "Custo", 190, False, 3
        P_ImprLinha obj
        Do While Not .EOF
            iCount = iCount + 1
            If iCount = 41 Then
                iCount = 0
                P_RodapeVis obj, strNome_Emp
                P_ImprX obj, "Cód", 0
                P_ImprX obj, "Produto", 18
            '    P_ImprX obj, "Lote", 100, True, True
                P_ImprX obj, "Unitário", 130, True, True
                P_ImprX obj, "Qtd", 150, True, True
                'P_ImprX obj, "IPI", 165, True, True
                P_ImprX obj, "Gastos", 162, True, True
                P_ImprX obj, "Custo", 190, False, 3
                P_ImprLinha obj
            End If
            obj.fontsize = 9
            P_ImprX obj, !cod_pro, 0
            P_ImprX obj, !nome_pro, 18
           ' P_ImprX obj, !lote_ite, 100
            P_ImprX obj, FormatVal(!unit_ite, "###,##0.0000"), 130, True, True
            P_ImprX obj, !qtd_ite, 150, True, True
            'P_ImprX obj, FormatVal(!ipi_ite, "###,##0.0000"), 165, True, True
            P_ImprX obj, FormatVal(!gastos_ite, "###,##0.0000"), 162, True, True
            P_ImprX obj, FormatVal(!custo_ite), 190, False, 3
                
            .MoveNext
        Loop
    P_RodapeVis obj, constNome_sis & "-" & constVersao_Sis
    End With
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
Sub P_ImprPcp(sSeq_Pcp As String, obj As Object)
    Dim iQtd_pag As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
    
    
    
    strSQL = "select tp.cod_pro,tm.nome_mar,tt.nome_cat,tp.nome_pro,sum(ti.qfa_ipc) as soma "
    strSQL = strSQL & " from tab_produtos as tp, tab_itenspcp as ti, "
    strSQL = strSQL & " tab_marcas as tm, tab_categorias as tt "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " and tt.cod_Cat=tp.cod_cat"
    strSQL = strSQL & " and ti.seq_pcp='" & sSeq_Pcp & "' "
    strSQL = strSQL & " and ti.sts_ipc='S' "
    
    strSQL = strSQL & " group by tp.cod_pro"
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
    
    P_CabRelTimbr strNome_Emp, obj
    
    P_PrintXY obj, "COTAÇÃO: " & sSeq_Pcp, 170, 0 + 20, False
    P_PrintXY obj, "DATA: " & Date, 170, 3.5 + 20, False
    P_PrintXY obj, "HORA: " & Time, 170, 7 + 20, False
    obj.fontsize = 12
    obj.CurrentY = 25
    P_ImprX obj, "   Solicitamos envio de cotação dos abaixo relacionados: ", 0, False
    
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
Sub imprEspelhoNota(nota_sai As String, serie_sai As String)
    P_ImprEspelhoNota nota_sai, serie_sai, frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprEspelhoNota nota_sai, serie_sai, Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub P_ImprEspelhoNota(nota_sai As String, serie_sai As String, obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select ts.*,tc.* from tab_saidas as ts, tab_clientes as tc "
    strSQL = strSQL & " where ts.cod_cli=tc.cod_cli "
    strSQL = strSQL & " and ts.nota_sai='" & nota_sai & "' "
    strSQL = strSQL & " and serie_sai='" & serie_sai & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        P_CabRelTimbr "Espelho de Nota Fisca de Saída Nº " & nota_sai & " Série " & serie_sai, obj
        P_ImprX obj, "Cliente: " & !nome_cli, 0
        P_ImprX obj, "CPF/CNPJ: " & Format(!cpf_cli), 60, False
        obj.FontBold = True
        P_ImprX obj, "Valor da Nota: " & FormatVal(!valor_sai), 200, False, 3
        obj.FontBold = False
        P_ImprX obj, "Endereço: " & Format(rPesq!Logradouro_cli & " N°" & !numero_cli & " " & !complemento_cli) & " Bairro: " & Format(!bairro_cli) & " CEP: " & Format(!cep_cli), 0, False
        P_ImprX obj, "Cidade: " & !cidade_cli & " UF: " & !uf_cli, 0, True
        P_ImprX obj, "Emissão: " & !emissao_sai, 200, False, 3
        P_ImprLinha obj
    End With
    obj.fontsize = 9
    strSQL = "select tp.*,ti.* from tab_itsai as ti,"
    strSQL = strSQL & " tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.nota_sai='" & nota_sai & "' "
    strSQL = strSQL & " and ti.serie_sai='" & serie_sai & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        P_ImprX obj, "Cód. ", 0
        P_ImprX obj, "Descrição", 20
        P_ImprX obj, "Un.", 120
        P_ImprX obj, "Ref.", 130
        P_ImprX obj, "Qtd.", 170, True, True
        P_ImprX obj, "Unit.", 185, True, True
        P_ImprX obj, "Total", 200, False, 3
        P_ImprLinha obj
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, "Nota " & nota_sai & " (continua...)"
                P_ImprX obj, "Cód. ", 0
                P_ImprX obj, "Descrição", 20
                P_ImprX obj, "Un.", 120
                P_ImprX obj, "Ref.", 130
                P_ImprX obj, "Qtd.", 170, True, True
                P_ImprX obj, "Unit.", 185, True, True
                P_ImprX obj, "Total", 200, False, 3
            End If
            P_ImprX obj, FormatCodigo(!cod_pro), 0
            P_ImprX obj, !nome_pro, 20
            P_ImprX obj, !unidade_pro, 120
            P_ImprX obj, !codfab_pro, 130
            P_ImprX obj, !qtd_its, 170, True, True
            P_ImprX obj, FormatVal(!unit_its), 185, True, True
            P_ImprX obj, FormatVal(!unit_its * !qtd_its), 200, False, 3
            
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_PreencheTamanhoPapel(obj As Object)
    obj.Clear
    obj.AddItem ""
        
End Sub
Sub P_TamanhoPapel(Tam As Integer)
  '  Printer.ScaleMode = 0
    If Tam = 1 Then
        Printer.PaperSize = vbPRPSLetter
    ElseIf Tam = 9 Then
        Printer.PaperSize = vbPRPSA4
    ElseIf Tam = 0 Then
        Printer.PaperSize = 256
        Printer.width = Printer.ScaleX(140, vbMillimeters, vbTwips)
        Printer.height = Printer.ScaleX(210, vbMillimeters, vbTwips)
        
        MsgBox Printer.width & " x " & Printer.height
    End If


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
Sub P_CabImo(obj As Object)
    obj.ScaleMode = Printer.ScaleMode
    obj.FontName = "courier new"
    obj.fontsize = 8
 '   P_ImprLinha obj, "-", 110
    P_CentraImpressao obj, "A T I V O   I M O B I L I Z A D O"
'    P_ImprLinha obj, "-", 110
    
    
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
Sub imprAcertoData(seq_rep As String, data_ace As String)
    
    P_Msg "Imprimindo acerto..."
    P_imprimeAcertoData seq_rep, data_ace, frmVisualiza.Pic
    P_Msg ""
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_imprimeAcertoData seq_rep, data_ace, Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub P_imprimeAcertoData(seq_rep As String, data_ace As String, obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sNome As String
    On Error GoTo trata_Erro
    strSQL = "select tr.nome_rep,ta.seq_ace from tab_acertos as ta,tab_representantes as tr "
    strSQL = strSQL & " where ta.Cod_rep=tr.Cod_rep "
    strSQL = strSQL & " and data_ace=" & F_GravaData(data_ace)
    strSQL = strSQL & " and tr.cod_rep='" & seq_rep & "' "
    strSQL = strSQL & " order by ta.seq_ace "
    
    
    abreObjetoAdo rPesq, strSQL
    
    If obj.CurrentY >= obj.ScaleHeight - 60 Then
        P_RodapeVis obj, strNome_Emp, iPaginaAtual & "/" & iQtdPag
        P_AddPicVis obj
    End If
    
    With rPesq
        
        If .RecordCount = 0 Then
            MsgBox "Nao encontrada!", vbExclamation
            .Close
            Exit Sub
        End If
        P_CabRelTimbr "GUIA DE ACERTO - " & data_ace, obj
        obj.fontsize = 14
        P_ImprX obj, "Representante: " & !nome_rep, 0
        
        
        P_ImprX obj, "", 40, False
        P_line1 obj
        sNome = !nome_rep
        'obj.Print
        Do While Not .EOF
        If obj.CurrentY >= obj.ScaleHeight - 60 Then
            P_RodapeVis obj, strNome_Emp, iPaginaAtual & "/" & iQtdPag
            P_AddPicVis obj
        End If
        
        
            P_imprimeAcerto !seq_ace, obj
            
            .MoveNext
        Loop
        .Close
    End With
    'On Error Resume Next
    
    
    obj.Print
    P_CentraImpressao obj, String(40, "_")
    P_CentraImpressao obj, sNome
    
    P_RodapeVis obj, strNome_Emp, iPaginaAtual & "/" & iQtdPag + 1
    
    If obj.CurrentY >= obj.ScaleHeight - 60 Then
            P_RodapeVis obj, strNome_Emp, iPaginaAtual & "/" & iQtdPag
            P_AddPicVis obj
    End If
    Exit Sub
    
trata_Erro:
    P_MostraErro
End Sub
Sub imprimeacerto(seq_ace As String)
    P_imprimeAcerto seq_ace, frmVisualiza.Pic
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
            P_imprimeAcerto seq_ace, Printer
            Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Sub P_imprimeAcerto(seq_ace As String, obj As Object)

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iEsq As Single
    Dim sCob As String
    Dim sEntr As Currency
    Dim cTotal As Currency
    Dim sTipoTer As String
    
    sTipoTer = F_BuscaUmCampo("tipo_ter", "tab_acertos", "seq_ace='" & seq_ace & "'")
    If sTipoTer = "C" Then
    
    Else
    
    End If
    If sTipoTer = "C" Then
        strSQL = "select ta.status_ace,ta.entrada_ace,tc.nome_cli,tc.endereco_cli,tt.nome_tip,tc.bairro_cli,tc.cidade_cli,"
        strSQL = strSQL & "ta.comissao_ace,ta.diferenca_ace,ta.cobranca_ace,ta.data_ace,hora_ace,cod_for,valor_ace,cod_usu,cobranca_ace,"
        strSQL = strSQL & "ti.*,tp.nome_pro,tp.ref_pro from tab_produtos as tp, tab_itacertos as ti,"
        strSQL = strSQL & "tab_acertos as ta, tab_clientes as tc, tab_tipocob as tt "
        strSQL = strSQL & " where tc.cod_cli=ta.cod_cli "
        strSQL = strSQL & " and ta.seq_ace=ti.seq_ace "
        strSQL = strSQL & " and tt.cod_tip=cobranca_ace "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and ta.seq_ace='" & seq_ace & "' "
        strSQL = strSQL & " order by seq_ite "
    Else
        strSQL = "select ta.status_ace,tc.nome_rep as nome_cli,ta.entrada_ace,tt.nome_tip,endereco_rep as endereco_cli,bairro_rep as bairro_cli,cidade_rep as cidade_cli,"
        strSQL = strSQL & "ta.comissao_ace,ta.diferenca_ace,ta.cobranca_ace,ta.data_ace,hora_ace,cod_for,valor_ace,cod_usu,cobranca_ace,"
        strSQL = strSQL & "ti.*,tp.nome_pro,tp.ref_pro from tab_produtos as tp, tab_itacertos as ti,"
        strSQL = strSQL & "tab_acertos as ta, tab_representantes as tc, tab_tipocob as tt "
        strSQL = strSQL & " where tc.cod_rep=ta.cod_cli "
        strSQL = strSQL & " and tt.cod_tip=cobranca_ace"
        strSQL = strSQL & " and ta.seq_ace=ti.seq_ace "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and ta.seq_ace='" & seq_ace & "' "
        strSQL = strSQL & " order by seq_ite "

    End If
    abreObjetoAdo rPesq, strSQL
    iEsq = 1
    
    
    
    
    With rPesq
        If .RecordCount > 0 Then
            sEntr = !entrada_ace
            
            If !status_ace = "1" Then
                P_ImprX obj, "Este acerto encontra-se em aberto!", 0, False
            ElseIf !status_ace = "3" Then
                P_ImprX obj, "Este acerto foi cancelado!", 0, False
            End If
            
            obj.FontName = "times new roman"
            obj.fontsize = 12
            obj.Print
        
            sCob = !nome_tip
        
            P_ImprX obj, "Acerto:" & !seq_ace, 0 + iEsq, True
            obj.FontBold = True
            P_ImprX obj, StrConv(!nome_cli, vbProperCase), 30 + iEsq, True
            obj.fontsize = 10
        
            P_ImprX obj, "Comissão " & FormatVal(!comissao_ace), 140 + iEsq, True, 3
            P_ImprX obj, "Dif: " & FormatVal(!diferenca_ace), 160 + iEsq, True, 3
        
            obj.fontsize = 16
        
            P_ImprX obj, "R$ " & FormatVal(!valor_ace), 195 + iEsq, False, 3
            obj.FontBold = False
            obj.FontName = "arial"
            P_line1 obj
            obj.fontsize = 10
            P_ImprX obj, "Produto:", 0 + iEsq
            P_ImprX obj, "Ref:", 100 + iEsq
            P_ImprX obj, "Qtd:", 120 + iEsq, True, True
            P_ImprX obj, "Unit:", 140 + iEsq, True, True
            P_ImprX obj, "Total:", 180 + iEsq, False, 3
            P_line1 obj, vbDot
        
            Do While Not .EOF
            
                If obj.CurrentY >= obj.ScaleHeight - 30 Then
                    P_RodapeVis obj
                    P_AddPicVis obj
                End If
            
            P_ImprX obj, StrConv(!nome_pro, vbProperCase), 0 + iEsq
            P_ImprX obj, !ref_pro, 100 + iEsq
            P_ImprX obj, !qtd_ite, 120 + iEsq, True, True
            P_ImprX obj, FormatVal(!unit_ite), 140 + iEsq, True, True
            P_ImprX obj, FormatVal(!unit_ite * !qtd_ite), 180 + iEsq, False, 3
            cTotal = cTotal + (!unit_ite * !qtd_ite)
            .MoveNext
        Loop
        P_line1 obj
        obj.FontBold = True
        P_ImprX obj, "Total..........:", 0 + iEsq
        P_ImprX obj, FormatVal(cTotal), 180 + iEsq, False, 3
        obj.FontBold = False
        P_line1 obj
        
    End If
        
        .Close
    End With
    
    
    
    
    strSQL = "select td.num_doc,ti.* from tab_itdoc as ti, tab_doc as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.esp_doc='A' "
    strSQL = strSQL & " and td.num_ven='" & seq_ace & "' "
    strSQL = strSQL & " and tipo_doc='R' "
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        obj.fontsize = 9
        obj.FontBold = True
        P_ImprX obj, "Duplicatas Geradas", 0 + iEsq, True
        P_ImprX obj, "Cobrança: " & sCob, 40 + iEsq, False, False
    
    End If
    obj.FontBold = False
    With rPesq
        strSQL = ""
        Do While Not .EOF
             If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj
                P_AddPicVis obj
             End If
             strSQL = strSQL & .AbsolutePosition & "ª Parc." & !venc_itd & " R$ " & FormatVal(!valor_itd) & "   "
            
            .MoveNext
            If Not .EOF Then strSQL = strSQL & ""
        Loop
        .Close
    End With
    P_ImprQuebra strSQL, obj, 0 + iEsq, obj.ScaleWidth - 10
    
    strSQL = "select td.num_doc,ti.num_che,valor_che,venc_che,cpf_che from tab_cheques as ti, tab_doc as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.esp_doc='A' "
    strSQL = strSQL & " and td.num_ven='" & seq_ace & "' "
    strSQL = strSQL & " and tipo_doc='R' "
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        obj.fontsize = 9
        obj.FontBold = True
        obj.Print
        P_ImprX obj, "Cheques Gerados", 0 + iEsq, False
        
    
    End If
    obj.FontBold = False
    With rPesq
        strSQL = ""
        Do While Not .EOF
             If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj
                P_AddPicVis obj
             End If
             strSQL = strSQL & "Ch: " & !num_che & " Venc: " & !Venc_che & "  R$ " & FormatVal(!Valor_che) & "   "
            
            .MoveNext
            If Not .EOF Then strSQL = strSQL & ""
        Loop
        .Close
    End With
    
    P_ImprQuebra strSQL, obj, 0 + iEsq, obj.ScaleWidth - 10
    
    P_ImprX obj, "", 0, False
    
    
    strSQL = "select tf.* from tab_caixa as tf,tab_acertofin as ta  "
    strSQL = strSQL & " where tf.seq_cai=ta.seq_fin "
    strSQL = strSQL & " and ta.seq_ace='" & seq_ace & "' "
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        obj.fontsize = 9
        obj.FontBold = True
        P_ImprX obj, "Pagamento a Vista " & rPesq!esp_cai, 0 + iEsq, False
        
    
    End If
    
    
    
    obj.FontBold = False
    With rPesq
        strSQL = ""
        Do While Not .EOF
             If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj
                P_AddPicVis obj
             End If
             strSQL = strSQL & "Valor " & FormatVal(!valor_Cai) & "   "
            
            .MoveNext
            If Not .EOF Then strSQL = strSQL & ""
        Loop
        .Close
    End With
    P_ImprQuebra strSQL, obj, 0 + iEsq, obj.ScaleWidth - 10
    
    
    
    strSQL = "select tv.*,tf.nome_fin from tab_venfin as tv, "
    strSQL = strSQL & "tab_financeiras as tf "
    strSQL = strSQL & " where tv.cod_fin=tf.cod_fin and tv.tipo_fin='A' "
    strSQL = strSQL & " and num_ven='" & seq_ace & "' "
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        obj.fontsize = 9
        obj.FontBold = True
        P_ImprX obj, "Pagamento Em Cartão", 0 + iEsq, False
        
    
    End If
    obj.FontBold = False
    With rPesq
        strSQL = ""
        Do While Not .EOF
             If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj
                P_AddPicVis obj
             End If
             strSQL = strSQL & "Cartão " & !nome_fin & " Valor " & FormatVal(!valor_fin) & "   "
            
            .MoveNext
            If Not .EOF Then strSQL = strSQL & ""
        Loop
        .Close
    End With
    P_ImprQuebra strSQL, obj, 0 + iEsq, obj.ScaleWidth - 10
    
    
    If sEntr > 0 Then
        P_ImprX obj, "Entrada: " & FormatVal(sEntr), 10, False
    End If
    
    
    
    
    
    obj.Print
    P_line1 obj
    
    'P_AddPicVis obj
End Sub
Sub P_Grades1(obj As Object, coord As String, Inicial As Single, final As Single, orient As orientacao, tipo As DrawStyleConstants)
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
Sub ImprReciboBai(seq_cab As String)
    Dim iCount As Integer
    P_ImprimeReciboBai seq_cab, frmVisualiza.Pic
    frmVisualiza.txtbxCopias.Text = "2"
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        For iCount = 1 To Int(frmVisualiza.txtbxCopias.Text)
            P_ImprimeReciboBai seq_cab, Printer
            Printer.EndDoc
        Next
    End If
    Unload frmVisualiza
End Sub
Sub P_ImprimeReciboBai(seq_cab As String, obj As Object, Optional meia_pagina As Boolean = False)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rFin As New ADODB.Recordset
    Dim r As representante.representante
    Dim c As cliente
    Dim u As funcionario.funcionario
    
    Dim totalParc As Currency
    
    strSQL = "select tb.*,tb.cod_usu as usu,tb.data_bai,tb.cod_cob as cob,"
    strSQL = strSQL & " td.cod_ter,td.num_doc,td.tipo_ter,td.parcelas_doc,td.esp_doc,ti.* from tab_doc as td, tab_itdoc as ti, "
    strSQL = strSQL & " tab_baixaitd2 as tb "
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
            P_ImprX obj, c.nome_cli, 20, False
            
            obj.FontBold = False
            P_ImprX obj, "Endereco: ", 5
            obj.FontBold = True
            P_ImprX obj, c.endereco_cli, 20, False
            obj.FontBold = False
            P_ImprX obj, "Bairro: ", 5
            obj.FontBold = False
            P_ImprX obj, c.set.nome & "/" & c.set.Cidade, 20, False
            
            P_ImprX obj, "Orgão: ", 5
            obj.FontBold = True
            P_ImprX obj, c.org.nome_org, 20, False
            
            
            
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
        P_ImprX obj, "Funcionário: " & F_BuscaUmCampo("Guerra_Fun", "tab_funcionarios", "cod_fun='" & !usu & "'"), 5, False
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
        P_ImprX obj, "Venc.Inicial", 140
        P_ImprX obj, "Total", 160, True
        P_ImprX obj, "Desc", 175, False
        
        
        P_line1 obj
        
        Do While Not .EOF
            P_ImprX obj, !num_doc, 5
            P_ImprX obj, !seq_doc, 25
            P_ImprX obj, !seq_itd, 40
            P_ImprX obj, !parcela_itd & "/" & !parcelas_doc, 60
            P_ImprX obj, FormatVal(!valor_itd), 80
            P_ImprX obj, !cob, 100
            P_ImprX obj, !vencini_itd, 140
            P_ImprX obj, FormatVal(!valor_bai), 160, True
            P_ImprX obj, FormatVal(!desconto_bai), 175, False
            
            
            
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
    strSQL = "select distinct tb.*,tb2.data_bai from tab_baixafin as tb, tab_baixaitd2 as tb2 where tb.seq_cab='" & seq_cab & "' "
    strSQL = strSQL & " AND tb2.seq_cab=tb.seq_cab"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        P_ImprX obj, "ID.Fin.", 5
        P_ImprX obj, "Esp.", 20
        P_ImprX obj, "Data.", 40
        
        P_ImprX obj, "Valor", 80, True
        P_ImprX obj, "Desconto", 120, False
        P_line1 obj
        totalParc = 0
        Do While Not .EOF
            P_ImprX obj, !seq_fin, 5
            P_ImprX obj, !esp_bai, 20
            P_ImprX obj, !data_bai, 40
            P_ImprX obj, FormatVal(!valor_bai), 80, True
            P_ImprX obj, FormatVal(!desconto_bai), 120, False
            totalParc = totalParc + (!valor_bai)
            
            .MoveNext
        Loop
        .Close
    End With
    obj.FontBold = True
    P_ImprX obj, "Total", 40, True
    P_ImprX obj, FormatVal(totalParc), 80, False
    
    P_ImprX obj, "", 0, False
    obj.FontBold = False
    P_line1 obj, vbDot
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
    Dim p As IPictureDisp
    On Error Resume Next
    Set p = LoadPicture(strCaminhoQrCodeDoc)
    
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
    sTipoTer = F_BuscaUmCampo("tipo_ter", "tab_doc", "seq_doc='" & seq_doc & "'")
    If sTipoTer = "" Then
        Exit Function
    End If
    strSQL = F_SelectDocsReceber(sTipoTer, , , , , , , , , seq_doc)
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        
        
        
        
        If .RecordCount = 0 Then
            'MsgBox "Documento " & seq_doc & " não encontrado!", vbExclamation
            .Close
            Exit Function
        End If
        
        If sTipoTer = "C" Then
            strSQL = "select tci.nome_cid,tipo_set,cep_cli,logradouro_cli,numero_cli,complemento_cli,endereco_cli,uf_cid,nome_set "
            strSQL = strSQL & " from tab_clientes as tc,tab_cidades as tci, tab_setores as ts "
            strSQL = strSQL & " where tci.cod_cid=ts.cod_cid "
            strSQL = strSQL & " and tc.cod_set=ts.cod_set "
            strSQL = strSQL & " and cod_cli='" & !cod_ter & "' "
            abreObjetoAdo rEnd, strSQL
            
        Else
            strSQL = "select tci.nome_cid,tipo_set,cep_rep as cep_cli,endereco_rep as endereco_cli,uf_cid,nome_set "
            strSQL = strSQL & " from tab_representantes as tc,tab_cidades as tci, tab_setores as ts "
            strSQL = strSQL & " where tci.cod_cid=ts.cod_cid "
            strSQL = strSQL & " and tc.cod_set=ts.cod_set "
            strSQL = strSQL & " and cod_rep='" & !cod_ter & "' "
            abreObjetoAdo rEnd, strSQL
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
            
            
            On Error Resume Next
            obj.PaintPicture p, 175, yIni, 20, 20
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
Sub P_ImprimeContratoGrafico(seq_doc As String, obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rClau As New ADODB.Recordset
    Dim rPro As New ADODB.Recordset
    Dim tamImprNormal As Integer
    Dim tamImprCond As Integer
    Dim x As Single
    Dim i As Single
    Dim sAltura As Single
    Dim sLargura As Single
    Dim novaAltura As Single
    Dim sBolsa As String
    Dim sClausulas As Integer
    Dim rCli As New ADODB.Recordset
    Dim c As Clientes.cliente
    Dim r As representante.representante
    Dim rDoc As New ADODB.Recordset
    Dim sCli As String
    Dim sCPF As String
    Dim sTipoTer As String
    Dim yAux As Single
    Dim yAux1 As Single
    Dim mEsq As Single
    '    On Error GoTo trata_ERro
    mEsq = 10
  
  '  On Error Resume Next
    'A ESCALA SEMPRE TEM QUE SER PASSADA EM VBTWIPS, PORTANTO É NECESSÁRIO TRANSFORMAR
    obj.FontName = "arial"
 '   Printer.PrintQuality = -2
    
    
    
    strSQL = "select td.*,ti.* from tab_doc as td, tab_itdoc as ti where ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " and td.seq_doc='" & seq_doc & "' order by venc_itd  "
    FechaObjeto rDoc
    abreObjetoAdo rDoc, strSQL
    
    If rDoc!tipo_ter = "C" Then
        c = Clientes.F_BuscaClienteCodigo(rDoc!cod_ter)
    Else
        
    End If
    
    
    strSQL = "select * from tab_itvendas where num_ven='" & rDoc!num_ven & "'"
    rPro.Open strSQL, Conn_Client
    
    
    
    
    strSQL = "select * from tab_contratos where seq_con=1"
    abreObjetoAdo rPesq, strSQL
    strSQL = "select * from tab_clausulas where seq_con=1"
    rClau.Open strSQL, Conn_Client
    
    
    'If obj.Name = "Printer" Then
    '    obj.PaperSize = 121
    'End If
    
    obj.ScaleMode = Printer.ScaleMode
    sAltura = 90 + (rClau.RecordCount * 30)  '10 cm
    sLargura = 90 '17 cm
    'sBolsa = rPesq!bolsa_loc
    'If TypeName(obj) = "Printer" Then
    '    obj.ScaleMode = vbMillimeters
    '    obj.PaperSize = 119
    'End If
    
    sAltura = Printer.ScaleY(sAltura, vbMillimeters, vbTwips)
    sLargura = Printer.ScaleX(sLargura, vbMillimeters, vbTwips)
    obj.FontName = "arial"
    obj.fontsize = 10
    'obj.Height = sAltura
    'obj.Width = sLargura
    'obj.ScaleHeight = 80.76855 + (rClau.RecordCount * 30)
    'obj.ScaleWidth = 163.2304
    
    obj.FontBold = True
    P_CabRelTimbr "", obj, True
    P_ImprQuebraTurbo rPesq!cab_con, obj, 1, obj.ScaleWidth - 10, False, 2
    obj.FontBold = False
    obj.Print
    
    
    
    
    
    
    yAux = obj.CurrentY
    P_line1 obj
    obj.Print
    P_ImprX obj, "", 0, False
    
    
    yAux = obj.CurrentY
    P_Line obj, mEsq, yAux, 200 - mEsq, yAux
    
    obj.FontBold = True
    P_ImprX obj, "COMPRADOR(A)", mEsq, False
    obj.FontBold = False
    obj.Print
    P_ImprX obj, c.nome_cli, 1 + mEsq, False
    P_ImprX obj, "CPF: " & c.cnpj_Cli, 1 + mEsq, True
    P_ImprX obj, "CI: " & c.ie_Cli, 60 + mEsq, False
    P_ImprX obj, "Endereço: " & c.endereco_cli, 1 + mEsq, False
    P_ImprX obj, "Bairro: " & c.set.nome, 1 + mEsq, True
    P_ImprX obj, "Cidade: " & c.set.Cidade & " /" & c.set.UF, 200 - mEsq, False, direita
    
    P_Line obj, mEsq, obj.CurrentY, 200 - mEsq, obj.CurrentY
    
    P_Grades obj, mEsq & "," & 200 - mEsq, yAux, obj.CurrentY, vertical, vbSolid
    
    obj.Print
    
    yAux = obj.CurrentY
    
    P_Line obj, mEsq, yAux, 200 - mEsq, yAux
    
    P_Line obj, mEsq, obj.CurrentY, 200 - mEsq, obj.CurrentY
    obj.FontBold = True
    P_ImprX obj, "VENDEDORA: ", 11, False
    
    obj.Print
    
    obj.FontBold = False
    P_ImprX obj, strNome_Emp, 11, False
    P_ImprX obj, "CNPJ: " & strCNPJ_Emp, 11, True
    P_ImprX obj, "CI: " & strIE_Emp, 71, False
    P_ImprX obj, "Endereço: " & strEndereco_Emp & " Bairro: " & strBairro_Emp & " - " & strCidade_Emp & "/" & strUF_Emp, 11, False
    P_Line obj, mEsq, obj.CurrentY, 200 - mEsq, obj.CurrentY
    
    P_Line obj, mEsq, obj.CurrentY, 200 - mEsq, obj.CurrentY
    
    P_Grades obj, mEsq & "," & 200 - mEsq, yAux, obj.CurrentY, vertical, vbSolid
    
    'P_ImprX obj, "Bairro: " & c.set.nome, 1, True
    'P_ImprX obj, "Cidade: " & c.set.cidade & " /" & c.set.uf, 200, False, direita
    'P_line1 obj
    obj.Print
    
    yAux1 = obj.CurrentY
    
    P_Line obj, mEsq, yAux1, 200 - mEsq, yAux1
    obj.FontBold = True
    P_ImprX obj, "VENCIMENTOS DAS PARCELAS - DOCUMENTO/VENDA: " & rDoc!num_doc & " DE " & rDoc!data_doc, 11, True
    obj.FontBold = False
    P_ImprX obj, "", 91, False
    obj.Print
    yAux = obj.CurrentY
    
    P_Line obj, 10, obj.CurrentY, 71, obj.CurrentY
    P_ImprX obj, "Parcela", 11, True
    P_ImprX obj, "Vencimento", 30, True
    P_ImprX obj, "Valor", 70, False, direita
    
    P_Line obj, 10, obj.CurrentY, 71, obj.CurrentY
    
    With rDoc
        Do While Not .EOF
            P_ImprX obj, .AbsolutePosition + 1, 11
            P_ImprX obj, !venc_itd, 30
            P_ImprX obj, FormatVal(!valor_itd), 71, False, direita
            .MoveNext
        Loop
    End With
    P_Line obj, 60, yAux, 60, obj.CurrentY
    P_Grades obj, "29,71", yAux, obj.CurrentY, vertical, vbSolid
    
    
    P_Line obj, mEsq, obj.CurrentY, 200 - mEsq, obj.CurrentY
    
    P_Grades obj, mEsq & "," & 200 - mEsq, yAux1, obj.CurrentY, vertical, vbSolid
    
'    obj.CurrentY = yAux
'    P_ImprX obj, "...", 61, True
'    P_ImprX obj, "Produtos", 71, True
'    P_ImprX obj, "Qtd", 160, True, direita
'    P_ImprX obj, "Unit", 170, True, direita
'    P_ImprX obj, "Desc", 185, True, direita
'    P_ImprX obj, "Total", 200, False, direita
'
'    P_Line obj, 0, obj.CurrentY, 60, obj.CurrentY
'    With rPro
'        Do While Not .EOF
'            P_ImprX obj, .AbsolutePosition + 1, 61, True
'            P_ImprX obj, !desc_ite, 71, True
'            P_ImprX obj, !qtd_ite, 160, True, direita
'            P_ImprX obj, FormatVal(!unit_ite), 170, True, direita
'            P_ImprX obj, FormatVal(!desconto_ite), 185, True, direita
'            P_ImprX obj, FormatVal(!final_ite), 200, False, direita
'
'            .MoveNext
'        Loop
'    End With
'    P_Line obj, 60, yAux, 60, obj.CurrentY
'    P_line1 obj
    
    
    obj.Print
    
    P_ImprQuebraTurbo rPesq!preambulo_con, obj, 14, obj.ScaleWidth - 10, False, 14
    obj.Print
    
    With rClau
        Do While Not .EOF
            
            P_ImprQuebra !cab_cla, obj, 12, obj.ScaleWidth - 12
            
            strSQL = F_BuscaUmCampo("texto_cla", "tab_clausulas", "seq_cla=" & !seq_cla & " and seq_con=" & !seq_con)
             P_ImprQuebraTexto strSQL, obj, 15, obj.ScaleWidth - 25
                    
            obj.Print
            
            .MoveNext
        Loop
    End With
    obj.Print
    P_CentraImpressao obj, String(45, "_")
    P_CentraImpressao obj, "COMPRADOR(A) - " & c.nome_cli
    obj.Print
    obj.Print
    P_CentraImpressao obj, String(45, "_")
    P_CentraImpressao obj, "VENDEDORA - " & strNome_Emp
    
    obj.Print
    obj.Print
    P_CentraImpressao obj, String(45, "_")
    P_CentraImpressao obj, "AVALISTA  "
    
    
    rPesq.Close
    rClau.Close
    rDoc.Close
'    rCli.Close
    rPro.Close
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_ImprimeContratoHTML(seq_doc As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rClau As New ADODB.Recordset
    Dim rPro As New ADODB.Recordset
    Dim tamImprNormal As Integer
    Dim tamImprCond As Integer
    Dim x As Single
    Dim i As Single
    Dim sAltura As Single
    Dim sLargura As Single
    Dim novaAltura As Single
    Dim sBolsa As String
    Dim sClausulas As Integer
    Dim rCli As New ADODB.Recordset
    Dim c As Clientes.cliente
    Dim r As representante.representante
    Dim rDoc As New ADODB.Recordset
    Dim sCli As String
    Dim sCPF As String
    Dim sTipoTer As String
    Dim yAux As Single
    Dim yAux1 As Single
    Dim mEsq As Single
    Dim sNomeArq As String
    Dim tamTable As String
    Dim tamColuna As String
    Dim valorDoc As Currency
    Dim bCheque As Boolean
    mEsq = 10
  
    tamTable = 800
    tamColuna = 800
  '  On Error Resume Next
    'A ESCALA SEMPRE TEM QUE SER PASSADA EM VBTWIPS, PORTANTO É NECESSÁRIO TRANSFORMAR
    'obj.FontName = "arial"
 '   Printer.PrintQuality = -2
    sNomeArq = F_AbreHtml("contrato" & F_SoNumero(Now))
    
    P_HtmlAbreHead 1
    P_HtmlAbreTitle 1, "CONTRATO DE COMPRA E VENDA "
    P_HtmlFechaTitle 1
    P_HtmlFechaHead 1
    P_HtmlAbreBody 1
    
    
    
    
    strSQL = "select td.valor_doc,data_doc,td.num_doc,num_ven,tipo_ter,td.parcelas_doc,td.cod_ter,ti.* from tab_doc as td, tab_itdoc as ti where ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " and td.seq_doc='" & seq_doc & "' order by venc_itd  "
    abreObjetoAdo rDoc, strSQL
    
    
    If rDoc.RecordCount = 0 Then
        'é cheque
        strSQL = "select td.valor_doc,td.num_doc,td.num_ven,tipo_ter,td.parcelas_doc,td.cod_ter,ti.* from tab_doc as td, tab_cheques as ti where ti.seq_doc=td.seq_doc "
        strSQL = strSQL & " and td.seq_doc='" & seq_doc & "' order by venc_che  "
        bCheque = True
        abreObjetoAdo rDoc, strSQL
        
    End If
    
    If rDoc.RecordCount = 0 Then
        MsgBox "Não encontrado!", vbExclamation
        Exit Sub
    End If
    If rDoc!tipo_ter = "C" Then
        c = Clientes.F_BuscaClienteCodigo(rDoc!cod_ter)
    Else
        
    End If
    valorDoc = rDoc!valor_doc
    
    
    strSQL = "select * from tab_itvendas where num_ven='" & rDoc!num_ven & "'"
    FechaObjeto rPro
    abreObjetoAdo rPro, strSQL
    
    
    
    
    strSQL = "select * from tab_contratos where seq_con=1"
    abreObjetoAdo rPesq, strSQL
    strSQL = "select * from tab_clausulas where seq_con=1"
    abreObjetoAdo rClau, strSQL
    
        
    P_HtmlAbreCenter 1
    
    
    P_HtmlAbreTable 1, 0
    P_HtmlAbreLinha
    P_HtmlAbreColuna 1, "", "center", False
    P_HtmlAbreH 1, 4, strNome_Emp, True
    P_HtmlAbreH 1, 4, rPesq!cab_con, True
    P_HtmlFechaColuna 1
    P_HtmlFechaLinha 1
    
    
    P_HtmlAbreTable 1, 1, tamTable  'tabela principal
        P_HtmlAbreTable 1, 1, tamTable '1
            P_HtmlAbreLinha 1
                P_HtmlAbreColuna 1, "COMPRADOR(A)", , False, tamTable
            P_HtmlFechaLinha 1
            
            P_HtmlAbreTable 1, 0, tamTable '2
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, "CPF:", "left", , "20%"
                    P_HtmlAbreColuna 1, "CI:", "left", , "30%"
                    P_HtmlAbreColuna 1, "Nome:" & "&nbsp", "left", , "30%"
                    P_HtmlAbreColuna 1, "Data Nasc.:", "left", , "20%"
                P_HtmlFechaLinha 1
                
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, c.cnpj_Cli, "left"
                    P_HtmlAbreColuna 1, c.ie_Cli, "left"
                    P_HtmlAbreColuna 1, c.nome_cli, "left"
                    P_HtmlAbreColuna 1, c.datanasc_cli, "left"
                P_HtmlFechaLinha 1
                
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, "Cidade:", "left"
                    P_HtmlAbreColuna 1, "CEP:", "left"
                    P_HtmlAbreColuna 1, "Endereço:", "left"
                    P_HtmlAbreColuna 1, "Bairro", "left"
                P_HtmlFechaLinha 1
            
        
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, c.set.Cidade, "left"
                    P_HtmlAbreColuna 1, c.cep_cli, "left"
                    P_HtmlAbreColuna 1, c.endereco_cli & "&nbsp", "left"
                    P_HtmlAbreColuna 1, c.set.nome, "left"
                P_HtmlFechaLinha 1
    
            P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, "-"
            P_HtmlFechaLinha
    
            P_HtmlFechaTable 1 '2
    
        P_HtmlFechaTable 1 '1
    
        P_HtmlAbreTable 1, 1, tamTable '3
            P_HtmlAbreTable 1, 1, tamTable '4
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, "VENDEDORA", , False, tamTable
                P_HtmlFechaLinha 1
                
                P_HtmlAbreTable 1, 0, tamTable '5
                
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, "CNPJ:", "left", , "20%"
                    P_HtmlAbreColuna 1, "IE:", "left", , "20%"
                    P_HtmlAbreColuna 1, "Nome:" & "&nbsp", "left", , "50%"
                    P_HtmlAbreColuna 1, "", "left", , "10%"
                P_HtmlFechaLinha 1
                
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, strCNPJ_Emp, "left"
                    P_HtmlAbreColuna 1, strIE_Emp, "left"
                    P_HtmlAbreColuna 1, strNome_Emp, "left"
                    P_HtmlAbreColuna 1, "", "left"
                P_HtmlFechaLinha 1
                
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, "Cidade:", "left"
                    P_HtmlAbreColuna 1, "CEP:", "left"
                    P_HtmlAbreColuna 1, "Endereço:", "left"
                    P_HtmlAbreColuna 1, "Bairro", "left"
                P_HtmlFechaLinha 1
            
        
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, strCidade_Emp, "left"
                    P_HtmlAbreColuna 1, strCEP_Emp, "left"
                    P_HtmlAbreColuna 1, strEndereco_Emp & "&nbsp", "left"
                    P_HtmlAbreColuna 1, strBairro_Emp, "left"
                P_HtmlFechaLinha 1
                
        P_HtmlAbreLinha 1
        P_HtmlAbreColuna 1, "-"
        P_HtmlFechaLinha
        P_HtmlFechaTable 1 'principal
    
    
    
    
    'P_ImprX obj, strNome_Emp, 11, False
    'P_ImprX obj, "CNPJ: " & strCNPJ_Emp, 11, True
    'P_ImprX obj, "CI: " & strIE_Emp, 71, False
    'P_ImprX obj, "Endereço: " & strEndereco_Emp & " Bairro: " & strBairro_Emp & " - " & strCidade_Emp & "/" & strUF_Emp, 11, False
    'P_Line obj, mEsq, obj.CurrentY, 200 - mEsq, obj.CurrentY
    
   ' P_Line obj, mEsq, obj.CurrentY, 200 - mEsq, obj.CurrentY
    
   ' P_Grades obj, mEsq & "," & 200 - mEsq, yAux, obj.CurrentY, vertical, vbSolid
    
    'P_ImprX obj, "Bairro: " & c.set.nome, 1, True
    'P_ImprX obj, "Cidade: " & c.set.cidade & " /" & c.set.uf, 200, False, direita
    'P_line1 obj
    'obj.Print
    
    
    
    
    P_HtmlAbreTable 1
    
    P_HtmlAbreLinha 1
    
    P_HtmlAbreColuna 1, "<b>DOCUMENTO: " & rDoc!num_doc & " ID:" & rDoc!seq_doc & " Data: " & rDoc!data_doc & "</b>", "center", True, -1, 4
    P_HtmlFechaLinha 1
    P_HtmlAbreLinha 1
    
    P_HtmlAbreColuna 1, "<b>VENCIMENTOS DAS PARCELAS</b>", "center", True, -1, 4
    P_HtmlFechaLinha 1
    
    
    P_HtmlAbreLinha 1
    
    P_HtmlAbreColuna 1, "Parcela"
    P_HtmlAbreColuna 1, "Id.Parc."
    P_HtmlAbreColuna 1, "Vencimento"
    P_HtmlAbreColuna 1, "Valor"
        
    P_HtmlFechaLinha
    
    With rDoc
        Do While Not .EOF
            P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, .AbsolutePosition
             
            If bCheque Then
                P_HtmlAbreColuna 1, !seq_che
                P_HtmlAbreColuna 1, !Venc_che
                P_HtmlAbreColuna 1, FormatVal(!Valor_che)
            Else
                P_HtmlAbreColuna 1, !seq_itd
                P_HtmlAbreColuna 1, !venc_itd
            P_HtmlAbreColuna 1, FormatVal(!valor_itd)
            End If
            
            P_HtmlFechaLinha
            .MoveNext
        Loop
        P_HtmlAbreColuna 1, "..."
        P_HtmlAbreColuna 1, "..."
        P_HtmlAbreColuna 1, "..."
        P_HtmlAbreColuna 1, FormatVal(valorDoc)
        P_HtmlFechaTable
    End With
    
    P_HtmlAbreTable 1, 0
    
    P_HtmlAbreLinha 1
    
    P_HtmlAbreColuna 1, "", "justify", False, "-1", 2
    
    P_HtmlAbreP rPesq!preambulo_con, 2, 1, "justify"
    
    P_HtmlFechaColuna
    P_HtmlFechaLinha
    
    P_HtmlFechaTable
    
    P_HtmlAbreTable 1, 0
    
    
    
    With rClau
        Do While Not .EOF
            
            P_HtmlAbreLinha
            'P_HtmlAbreColuna 1, "-"
            P_HtmlFechaColuna
            
            P_HtmlAbreLinha
            
            P_HtmlAbreColuna 1, "<b>" & !cab_cla & "</b>", "center", True, "-1", 2
            
            
            strSQL = F_BuscaUmCampo("texto_cla", "tab_clausulas", "seq_cla=" & !seq_cla & " and seq_con=" & !seq_con)
            P_HtmlFechaLinha
            P_HtmlAbreLinha
             
             
            P_HtmlAbreColuna 1, strSQL
            
            P_HtmlFechaLinha
            
            .MoveNext
        Loop
    End With
    
    P_HtmlAbreLinha
    P_HtmlAbreColuna 1, "-"
    P_HtmlFechaLinha
    
    P_HtmlFechaTable
    
    
    
    
    'P_HtmlBR
    
    P_HtmlAbreTable 1, 0, 600
    
    P_HtmlAbreLinha 1
     
    P_HtmlAbreColuna 1, "<br><b>COMPRADOR(A) - " & c.nome_cli & "</b>", "center", True, "-1", 0, -1
    P_HtmlFechaLinha
    P_HtmlFechaTable
    
    
    
    P_HtmlAbreTable 1, 0, 600
    P_HtmlAbreLinha 1
    P_HtmlAbreColuna 1, "<br><b>VENDEDORA - " & strNome_Emp & "</b>", "center", True, "-1", 0, -1
    P_HtmlFechaLinha
    P_HtmlFechaTable
    
    
    P_HtmlAbreTable 1, 0, 600
    P_HtmlAbreLinha 1
    P_HtmlAbreColuna 1, "<br><b>AVALISTA<b>", "center", True, "-1", 0, -1
    P_HtmlFechaLinha
    P_HtmlFechaTable
    
    
    rPesq.Close
    rClau.Close
    rDoc.Close

    rPro.Close
    P_HtmlFechaCenter 1
    P_HtmlFechaBody 1
    P_FechaHtml 1
    Close #1
    shell "explorer " & sNomeArq
    
    
    
End Sub


Sub imprimeContrato(seq_doc As String)
    P_ImprimeContratoHTML seq_doc
    
        
End Sub
Sub imprimeContratoHTML(seq_doc As String)
    P_ImprimeContratoHTML seq_doc
    
End Sub

Sub P_RelContasAReceberPorRepresentante(Optional dataI As String = "", Optional dataF As String = "", _
    Optional cod_Rep As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    P_Status "Aguarde..."
    strSQL = "select tr.nome_rep as Representante,sum(ti.reman_itd) as Total"
    strSQL = strSQL & " from tab_Representantes as tr, tab_doc as td, tab_itdoc as ti "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc and ti.reman_itd>0 "
    strSQL = strSQL & " and td.cod_ter=tr.cod_rep "
    
    strSQL = strSQL & " and td.tipo_ter='R' "
    If dataI <> "" Then
        strSQL = strSQL & " and venc_itd>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and venc_itd<=" & F_GravaData(dataF)
    End If
    If cod_Rep <> "" Then
        strSQL = strSQL & " and tr.Cod_rep ='" & cod_Rep & "' "
    End If
    strSQL = strSQL & " group by tr.cod_rep "
    strSQL = strSQL & " order by nome_rep"
    abreObjetoAdo rPesq, strSQL
    P_rPesqToHtml "RELATÓRIO DE DÉBITOS DE REPRESENTANTES", rPesq, "De: " & dataI & " Até: " & dataF
    rPesq.Close
    P_Status ""
End Sub
Sub P_RelAcertosPorRepresentante(Optional dataI As String = "", Optional dataF As String = "", Optional cod_Rep As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    P_Status "Aguarde..."
    strSQL = "select tr.nome_rep as Representante,sum(ta.valor_ace) as Total"
    strSQL = strSQL & " from tab_Representantes as tr, tab_acertos as ta "
    strSQL = strSQL & " where ta.cod_rep=tr.cod_rep "
    
    
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ace>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ace<=" & F_GravaData(dataF)
    End If
    If cod_Rep <> "" Then
        strSQL = strSQL & " and tr.Cod_rep ='" & cod_Rep & "' "
    End If
    strSQL = strSQL & " group by tr.cod_rep "
    strSQL = strSQL & " order by nome_rep"
    abreObjetoAdo rPesq, strSQL
    P_rPesqToHtml "RELATÓRIO DE ACERTOS POR REPRESENTANTE", rPesq, _
        "De: " & dataI & " Até: " & dataF, "", "1"
    rPesq.Close
    P_Status ""
End Sub
Sub P_RelResumoAcertos(Optional dataI As String = "", Optional dataF As String = "", Optional cod_Rep As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipoTer As String
    
    
    P_Status "Aguarde..."
    
    
    
    strSQL = "select ta.seq_ace as Num,ta.data_ace as Data, ta.hora_ace as Hora,"
    strSQL = strSQL & " ta.cod_cli as Codigo, ta.tipo_ter as Tipo,tt.nome_tip as Cob, "
    strSQL = strSQL & " ta.valor_ace as valor, ta.comissao_ace as Comiss, ta.diferenca_ace as dif, "
    strSQL = strSQL & " ta.entrada_ace as Entrada from tab_tipocob as tt, tab_acertos as ta "
    strSQL = strSQL & " where ta.cobranca_ace=tt.cod_tip "
    
    
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ace>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ace<=" & F_GravaData(dataF)
    End If
    If cod_Rep <> "" Then
        strSQL = strSQL & " and ta.Cod_rep ='" & cod_Rep & "' "
    End If
    'strSQL = strSQL & " group by tr.cod_rep "
    strSQL = strSQL & " order by data_ace,seq_ace"
    abreObjetoAdo rPesq, strSQL
    P_rPesqToHtml "RELATÓRIO DE RESUMO DE ACERTOS POR REPRESENTANTE", rPesq, _
        "De: " & dataI & " Até: " & dataF, "Representante: " & cod_Rep & " " & _
            F_BuscaUmCampo("nome_rep", "tab_representantes", "cod_rep='" & cod_Rep & "'"), "6,7,8"
    rPesq.Close
    P_Status ""
End Sub


Sub P_RelResumoComissPer(dataI As String, dataF As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    P_Status "Aguarde..."
    strSQL = "select distinct tr.nome_rep as Representante,sum(valor_com) as Total,if(status_com=1,sum(valor_com),0) as Aberto,"
    strSQL = strSQL & " if(status_com=2,sum(valor_com),0) as Pago "
    strSQL = strSQL & " from tab_Representantes as tr, tab_comissparc as tp "
    strSQL = strSQL & " where tp.cod_rep=tr.cod_rep "
    If dataI <> "" Then
        strSQL = strSQL & " and venc_com>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and venc_com<=" & F_GravaData(dataF)
    End If
    
    strSQL = strSQL & " and status_com=1 "
    
    
    
    
    strSQL = strSQL & " group by tr.cod_rep "
    strSQL = strSQL & " order by nome_rep"
    abreObjetoAdo rPesq, strSQL
    P_rPesqToHtml "RESUMO DE COMISSÕES PERMANENTES EM ABERTO", rPesq, "Em " & Date
    rPesq.Close
    P_Status ""
End Sub
Sub P_ImprUltimoComprPagComiss(cod_Rep)
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select max(seq_cai) as seq_cai from tab_comissparc where cod_rep='" & cod_Rep & "' "
    
    abreObjetoAdo rPesq, strSQL
    
      
    
    P_ImprimePagamentoComissRepr rPesq!seq_cai
End Sub
Sub P_ImprimePagamentoComissRepr(seq_cai As String, Optional omitir_parcelas As Boolean = True)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sNomeArq As String
    Dim sNomeRep As String
    On Error GoTo trata_Erro
    Dim sTab As String
    Dim cTotal As Currency
    For iCount = 1 To 2
    If iCount = 1 Then
        sTab = "tab_itdoc"
    Else
        sTab = "tab_cheques"
    End If
    If sTab = "tab_itdoc" Then
        strSQL = "select concat(parcela_itd,'/',parcelas_doc) as parcela,tc.*,tx.*,tr.nome_rep,"
        strSQL = strSQL & "if(status_com='2',valor_com,0) as comiss_pg,"
        strSQL = strSQL & "ti.reman_itd,"
        strSQL = strSQL & "tcli.nome_cli, ta.data_ace,tc.seq_ace,ti.valor_itd,td.num_doc,td.parcelas_doc,td.valor_doc "
        strSQL = strSQL & " from tab_doc as td, tab_comissparc as tc, tab_acertos as ta, tab_clientes as tcli,tab_representantes as tr,"
        strSQL = strSQL & sTab & " as ti,tab_caixa as tx  "
        strSQL = strSQL & " where ti.seq_doc=td.seq_doc "
        strSQL = strSQL & " and tc.seq_itd=ti.seq_itd "
        strSQL = strSQL & " and ta.cod_cli=tcli.cod_cli"
        strSQL = strSQL & " and tr.cod_rep=ta.cod_rep "
        strSQL = strSQL & " and tc.seq_cai=tx.seq_cai"
        strSQL = strSQL & " and tc.seq_cai='" & seq_cai & "'"
    Else
        strSQL = "select concat(parc_che,'/',parcelas_doc) as parcela,tc.*,tx.*,tr.nome_rep,"
        strSQL = strSQL & "if(status_com='2',valor_com,0) as comiss_pg,tcli.nome_cli,"
        strSQL = strSQL & "if(dep_che='0000-00-00',valor_che,0) as reman_itd,"
        strSQL = strSQL & " ta.data_ace,tc.seq_ace,ti.valor_che as valor_itd,td.num_doc,td.parcelas_doc,td.valor_doc "
        strSQL = strSQL & " from tab_doc as td, tab_comissparc as tc, tab_acertos as ta,tab_clientes as tcli,tab_representantes as tr, "
        strSQL = strSQL & sTab & " as ti,tab_caixa as tx  "
        strSQL = strSQL & " where ti.seq_doc=td.seq_doc "
        strSQL = strSQL & " and tc.seq_itd=ti.seq_che "
        strSQL = strSQL & " and ta.cod_cli=tcli.cod_cli"
        strSQL = strSQL & " and tc.seq_cai=tx.seq_cai"
        strSQL = strSQL & " and tc.seq_cai='" & seq_cai & "'"
        strSQL = strSQL & " and tr.cod_rep=ta.cod_rep "
    End If
    
    
    
    
    strSQL = strSQL & " and tipo_com=" & IIf(sTab = "tab_itdoc", "'1'", "'2'")
    
    
    
    strSQL = strSQL & " and ta.seq_ace=tc.seq_ace "
    abreObjetoAdo rPesq, strSQL
    sNomeArq = "RelComissPar.Html"
    With rPesq
        If .RecordCount > 0 Then
            sNomeArq = F_AbreHtml(sNomeArq)
            
            P_HtmlAbreHead
            P_HtmlAbreTitle 1, "Comprovante de Pagamento"
            P_HtmlFechaHead
            P_HtmlAbreBody 1
            sNomeRep = rPesq!nome_rep
            P_HtmlAbreTable 1, 1, 1000
            P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "", , False
            P_HtmlAbreH 1, 2, "Comprovante de Pagamento - Repr: " & !nome_rep, True
            P_HtmlAbreH 1, 4, !historico_cai & " Valor: " & FormatVal(!valor_Cai) & " Data " & !data_cai, True
            
            If omitir_parcelas = False Then
                P_HtmlAbreTable 1
                    P_HtmlAbreLinha 1
                        P_HtmlAbreColuna 1, "Acerto", "Center", True
                        P_HtmlAbreColuna 1, "Vencimento", "Center", True
                        P_HtmlAbreColuna 1, "Cliente", "Left", True
                        P_HtmlAbreColuna 1, "Valor", "Center", True
                    P_HtmlFechaLinha 1
                Do While Not .EOF
                    P_HtmlAbreLinha 1
                        P_HtmlAbreColuna 1, !seq_ace, "Center", True
                        P_HtmlAbreColuna 1, !venc_com, "Center", True
                        P_HtmlAbreColuna 1, !nome_cli, "Left", True
                        P_HtmlAbreColuna 1, FormatVal(!valor_com), "Center", True
                        cTotal = cTotal + !valor_com
                    P_HtmlFechaLinha 1
                            
                    
                
                    .MoveNext
                Loop
            
                P_HtmlAbreLinha 1
                    P_HtmlAbreColuna 1, "Total", "Center", True
                    P_HtmlAbreColuna 1, "...", "Center", True
                    P_HtmlAbreColuna 1, "...", "Left", True
                    P_HtmlAbreColuna 1, FormatVal(cTotal), "Center", True
                P_HtmlFechaLinha 1
            
            
                P_HtmlFechaTable
            End If
            P_HtmlAbreP "Declaro ter recebido os valores acima especificados.", 3, , "center", , True
            P_HtmlAbreP "&nbsp", 3
            P_HtmlAbreP sNomeRep, 3, , "center", , True
            P_HtmlFechaColuna 1
            P_HtmlFechaTable  '
            
            P_HtmlFechaBody
            P_FechaHtml 1, True, sNomeArq
            .Close
        End If
    End With
    
    Next
Exit Sub
trata_Erro:
P_MostraErro
On Error Resume Next
    Close #1
End Sub


    

Sub ImprimeNotaEntrada(seq_not As String, tipo As Integer)
    
    P_ImprNotaFiscal frmVisualiza.Pic, seq_not, 0
         frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprNotaFiscal Printer, seq_not, 0
        Printer.EndDoc
    End If
    Unload frmVisualiza
            
End Sub

Sub imprimeTrocaAvulsa(id As String, obj As Object)
    On Error Resume Next
    Dim strSQL As String
    Dim sSeq As String
    Dim iCount As Integer
    Dim cQtd_Ant As Currency
    Dim cQtd_Atual As Currency
    Dim sPro As String
    Dim sCod_Pro As String
    Dim cFinal As Currency
    Dim sSeq_ite As String
    Dim sUnit_Ite As String
    
    On Error GoTo trata_Erro
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select tt.*,ti.*,tp.nome_pro,tp.ref_pro from tab_produtos as tp, tab_trocaavulsa as tt, "
    strSQL = strSQL & " tab_trocaavulsait as ti where ti.id_tro=tt.id "
    strSQL = strSQL & " and ti.codigo=tp.cod_pro "
    strSQL = strSQL & " and tt.id=" & id
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
    
        P_CabRelTimbr "TROCA AVULSA " & id, obj
        P_ImprX obj, "Data: " & !data, 0
        P_ImprX obj, "Cliente: " & !cliente, 30, True
        P_ImprX obj, "Vendedor:" & !vendedor, 130, False
        
        
        P_ImprLinha obj
        P_ImprX obj, "Valor Entrada: " & FormatVal(!valor_in), 0
        P_ImprX obj, "Valor Saída: " & FormatVal(!valor_out), 60
        
        P_ImprX obj, "Dif: " & FormatVal(!valor_out - !valor_in), 100
        
        P_ImprX obj, "Desconto: " & FormatVal(!desconto), 140, False
        P_ImprX obj, "Obs: " & !obs, 5, False
        
        
        P_ImprLinha obj, "="
        With rPesq
            
            obj.FontBold = True
            obj.fontsize = 9
            P_ImprX obj, "...", 0
            P_ImprX obj, "Ref", 5
            P_ImprX obj, "Descrição", 35
            P_ImprX obj, "Qtd", 135, True, True
            P_ImprX obj, "Unit", 155, True, True
            P_ImprX obj, "Total", 175, True, True
            P_ImprX obj, "Tipo", 195, False, True
            
            obj.FontBold = False
            P_ImprLinha obj
            Do While Not .EOF
                
                obj.fontsize = 9
                P_ImprX obj, .AbsolutePosition, 0
                P_ImprX obj, !ref_pro, 5
                obj.fontsize = 8
                P_ImprX obj, !nome_pro, 35
                obj.fontsize = 10
                P_ImprX obj, !qtd, 135, True, True
                
                P_ImprX obj, FormatVal(!unit), 155, True, True
                P_ImprX obj, FormatVal(!total), 175, True, True
                P_ImprX obj, !tipo, 195, False, True
                
                .MoveNext
            Loop
            P_ImprLinha obj, "_"
        End With
    End With
    Exit Sub
trata_Erro:
P_MostraErro
End Sub
Sub imprTrocaAvulsa(id As String)
    On Error GoTo trata_Erro
    imprimeTrocaAvulsa id, frmVisualiza.Pic
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
        imprimeTrocaAvulsa id, Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    
    Exit Sub
    
    
trata_Erro:
    P_MostraErro

End Sub
    
