Attribute VB_Name = "Vendas_antigo"
Function F_selectVendas(Optional cod_cli As String = "", Optional dataI As String = "", Optional dataF As String = "", _
    Optional situacao_ven As String = "situacao_ven>1", Optional abreVendaUsu As String = "", _
    Optional sOrderby As String = "num_ven desc", Optional groupBy As String = "") As String
    Dim strSQL As String
    Dim sCampoData As String
    strSQL = "select tfo.descricao_tip,tc.nome_cli, tv.*, tv.valor_ven+tv.desconto_ven as total, tf.guerra_fun "
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " " & TAB_VENDAS & " as tv, tab_funcionarios as tf,tab_tipoven as tfo  "
    
    strSQL = strSQL & " where tf.cod_fun=tv.cod_fun "
    
    strSQL = strSQL & " and tfo.cod_tip=tv.cod_tip "
    If situacao_ven <> "" Then
        strSQL = strSQL & " and " & situacao_ven
    End If
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    If abreVendaUsu <> "" Then
        strSQL = strSQL & " and cod_fun='" & abreVendaUsu & "' "
    End If
    
    If cod_cli <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & cod_cli & "' "
    End If
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(dataF & " 23:59:59")
    End If
    If groupBy <> "" Then
        strSQL = strSQL & " group by " & groupBy
    End If
    F_selectVendas = strSQL
End Function

Function F_selectVendasAgrupCli(Optional cod_cli As String = "", Optional dataI As String = "", Optional dataF As String = "", _
    Optional situacao_ven As String = "situacao_ven>4", Optional abreVendaUsu As String = "", _
    Optional formaP As String = "", Optional sOrderby As String = "nome_cli") As String
    Dim strSQL As String
    Dim sCampoData As String
    strSQL = "select tc.nome_cli, tv.*, sum(tv.valor_ven)as total,"
    strSQL = strSQL & " sum(tv.desconto_ven) as desconto "
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " " & TAB_VENDAS & " as tv, tab_formaspag as tfo  "
    
    strSQL = strSQL & " where tfo.cod_for=tv.cod_for "
    strSQL = strSQL & " and tv.cod_cli=tc.cod_cli "
    If formaP <> "" Then
    
        strSQL = strSQL & " and tfo.cod_for='" & formaP & "' "
    End If
    If situacao_ven <> "" Then
        strSQL = strSQL & " and " & situacao_ven
    End If
    
    If abreVendaUsu <> "" Then
        strSQL = strSQL & " and abrevenda_usu='" & abreVendaUsu & "' "
    End If
    
    If cod_cli <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & cod_cli & "' "
    End If
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(dataF & " 23:59:59")
    End If
    
    strSQL = strSQL & " group by tv.cod_cli"
    
    F_selectVendasAgrupCli = strSQL
End Function
Sub imprimeLogo(obj As Object)
    Dim Path As String
    On Error Resume Next
    Path = App.Path & "\pic\logoprint.jpg"
    obj.PaintPicture LoadPicture(Path), 0, 0
    obj.Print
    obj.Print
    
    
End Sub
Sub ImprimeComprovanteNaoFiscal(num_ven As String, Optional bViz As Boolean = False)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim crlf_ As String
    crlf_ = CrLf
  '  On Error GoTo trata_Erro
    Dim impressora As Object
    
    
    If bViz = False Then
        Set impressora = Printer
    Else
    
        Set impressora = frmVisualiza.Pic
    End If
    
    
    
    
    Dim Centra As Integer
    Dim cTotal As Currency
    Dim valor_ven As Currency
    Dim nome_Cli As String
    Centra = 60
    
    strSQL = "select tv.num_ven,tv.valor_ven,tv.data_ven,tv.cod_fun,tv.cod_usu,tv.valor_ven,tv.desconto_ven,"
    strSQL = strSQL & "tv.nomecli_ven,tv.entrega,tv.endereco,ti.* from " & TAB_VENDAS & " as tv inner join " & TAB_ITVENDAS & " as ti "
    strSQL = strSQL & " on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " where tv.num_ven='" & num_ven & "' "
    strSQL = strSQL & " and tv.situacao_ven='4' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Venda não encontrada em situação efetivada.", vbExclamation
            
            .Close
            Exit Sub
        End If
        
        
        imprimeLogo impressora
        If TypeName(impressora) = "Printer" Then
            impressora.Print
        Else
            impressora.Print "->"
        End If
        
        P_ImprX impressora, F_Centra(strNome_Emp, 90), 0, False
        P_ImprX impressora, F_Centra(strEndereco_Emp, 90), 0, False
        P_ImprX impressora, F_Centra("FONE:" & strFone_Emp, 90), 0, False
        P_ImprX impressora, F_Centra("*", Centra, "*"), 0, False
        P_line1 impressora
        impressora.FontBold = True
        impressora.Print
        impressora.Print
        P_PrintXY impressora, F_Centra("ROMANEIO DE ENTREGA", 50, "*"), 0, impressora.CurrentY, 2
        impressora.Print
        P_PrintXY impressora, F_Centra("NÃO É DOCUMENTO FISCAL", 50, "*"), 0, impressora.CurrentY, 2
        P_PrintXY impressora, F_Centra("NÃO É VÁLIDO COMO GARANTIA DE MERCADORIA", 40, "*"), 0, impressora.CurrentY, 2
        impressora.FontBold = False
        impressora.Print
        impressora.Print
        P_ImprX impressora, "VENDA:" & !num_ven & " DATA: " & !data_ven & " Op:" & !cod_fun, 0, False
        impressora.FontBold = True
                P_ImprQuebra "CLIENTE:" & !nomecli_ven, impressora, 0, impressora.ScaleWidth
                nome_Cli = !nomecli_ven
        impressora.Print
        P_ImprX impressora, "ENDEREÇO:" & !Endereco, 0, False
        impressora.FontBold = False
        impressora.Print
        P_line1 impressora
        impressora.Print
        P_line1 impressora
        valor_ven = !valor_ven
        
        P_ImprX impressora, "Item|Código|Qtd x Unitário = Total|Descrição", 0, False
        P_line1 impressora, vbDash
        Do While Not .EOF
             
            P_ImprQuebra .AbsolutePosition & ".|" & !cod_pro & "| " & FormatVal(!qtd_ite, "###,##0.000") & " X " & Format(!unit_ite) & " = " & Format(!qtd_ite * !unit_ite) & "|" & !descricao_ite, impressora, 0, impressora.ScaleWidth
           impressora.Print
            cTotal = cTotal + (!unit_ite * !qtd_ite)
            
            .MoveNext
        Loop
        impressora.Print
        P_line1 impressora
        impressora.Print
        impressora.FontBold = True
        P_ImprX impressora, "Itens: " & .RecordCount, 0, False
        P_ImprX impressora, "Soma: " & FormatVal(cTotal), 0, False
        P_ImprX impressora, "Valor Final: " & FormatVal(valor_ven), 0, False
        P_line1 impressora, vbInsideSolid
        impressora.FontBold = False
        impressora.Print
        impressora.fontsize = 10
        
        
        P_ImprX impressora, "DECLARO TER RECEBIDO OS PRODUTOS ACIMA", 0, False
        
        
        
        
        
        
        
        P_PrintXY impressora, F_Centra("______________________________", 40, "_"), 0, impressora.CurrentY, 2
        P_PrintXY impressora, F_Centra(nome_Cli, 40, "*"), 0, impressora.CurrentY, 2
        impressora.Print
        
        
        P_PrintXY impressora, F_Centra("NÃO É DOCUMENTO FISCAL", 50, "*"), 0, impressora.CurrentY, 2
        P_PrintXY impressora, F_Centra("NÃO É VÁLIDO COMO GARANTIA DE MERCADORIA", 40, "*"), 0, impressora.CurrentY, 2
        P_ImprX impressora, F_Centra("Obrigado! Volte sempre!", 70, "*"), 0, False
        impressora.Print
        impressora.Print Now
        
    End With
    
    
    
    
'    If TypeName(impressora) = "PictureBox" Then
'        frmVisualiza.Show 1
'        Unload frmVisualiza
'    Else
        impressora.EndDoc
'    End If
    
Exit Sub
trata_Erro:
MostraErro


End Sub
Function F_SelectItensCompradosVendas(cod_cli As String, Optional dataI As String = "", _
    Optional dataF As String = "") As String
    Dim strSQL As String
    strSQL = "select tfo.descricao_tip,tc.nome_cli, ti.*,tv.num_ven,tv.data_ven,"
    strSQL = strSQL & "tv.valor_ven,desconto_ven, tf.guerra_fun "
    strSQL = strSQL & " from tab_clientes as tc,"
    strSQL = strSQL & " " & TAB_VENDAS & " as tv, tab_funcionarios as tf,tab_tipoven as tfo,  "
    strSQL = strSQL & " " & TAB_ITVENDAS & " as ti "
    strSQL = strSQL & " where tf.cod_fun=tv.cod_fun"
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tfo.cod_tip=tv.cod_tip "
    
    strSQL = strSQL & " and situacao_ven=4 "
    
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    If abreVendaUsu <> "" Then
        strSQL = strSQL & " and tv.cod_fun='" & abreVendaUsu & "' "
    End If
    
    If cod_cli <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & cod_cli & "' "
    End If
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(dataF & " 23:59:59")
    End If
    F_SelectItensCompradosVendas = strSQL
End Function


Function InformaCPFClienteVenda() As String
    frmCPFVenda.Show 1
    InformaCPFClienteVenda = frmCPFVenda.txtcpf.Text
    Unload frmCPFVenda
    
End Function
Function F_VerificaVendaAberta(num_ven As String) As String
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & num_ven & "'")
    If strSQL <> strIP_Usu Then
    
        F_VerificaVendaAberta = strSQL
    End If
End Function

