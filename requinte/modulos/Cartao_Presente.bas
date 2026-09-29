Attribute VB_Name = "Cartao_Presente"
Public Type CartaoPresente
    id As String
    numero As String
    nome As String
    presenteado As String
    valor As String
    vencimento As String
    data As String
    status As String
    venda As String
End Type


Function BuscaCartaoPresente(numero As String) As CartaoPresente
    frmEdiCartao_Presente.txtNumero.Text = numero
    frmEdiCartao_Presente.cmbStatus.ListIndex = 0
    
    frmEdiCartao_Presente.tag = "x"
    frmEdiCartao_Presente.Show 1
    
    If frmEdiCartao_Presente.tag <> "" Then
        With frmEdiCartao_Presente.ado
            BuscaCartaoPresente.id = !id
            BuscaCartaoPresente.data = !data
            BuscaCartaoPresente.nome = !nome
            BuscaCartaoPresente.numero = !numero
            BuscaCartaoPresente.presenteado = TiraNull(!presenteado)
            BuscaCartaoPresente.status = !status
            BuscaCartaoPresente.valor = !valor
            BuscaCartaoPresente.venda = TiraNull(!num_ven)
            BuscaCartaoPresente.vencimento = TiraNull(!vencimento)
            
            
        End With
    End If
    Unload frmEdiCartao_Presente
End Function
Function GravaCartao_Presente(nome As String, cpf As String, contato As String, _
    presenteado As String, obs As String, data As String, vencimento As String, _
    numero As String, valor As String, cod_For As String) As String

    
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into cartao_presente(nome,cpf,contato,presenteado,obs,data,vencimento,numero,"
    strSQL = strSQL & "valor,cod_for,utilizado,cod_usu,status)values("
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & "'" & cpf & "',"
    strSQL = strSQL & "'" & contato & "',"
    strSQL = strSQL & "'" & presenteado & "',"
    strSQL = strSQL & "'" & obs & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & F_GravaData(vencimento) & ","
    strSQL = strSQL & "'" & numero & "',"
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & "'" & cod_For & "',"
    strSQL = strSQL & "0,"
    strSQL = strSQL & "'" & strCod_Usu & "','1')"
    Conn.Execute strSQL
    
    GravaCartao_Presente = F_BuscaUmCampo("last_insert_id(id)", "cartao_presente", "")
    
trata_Erro:
    
    
End Function
    
Function AlteraCartao_Presente(id As String, nome As String, cpf As String, contato As String, _
    presenteado As String, obs As String, vencimento As String) As Boolean


    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update carta_presente set "
    strSQL = strSQL & "nome='" & nome & "',"
    strSQL = strSQL & "cpf='" & cpf & "',"
    strSQL = strSQL & "contato='" & contato & "',"
    strSQL = strSQL & "presenteado='" & presenteado & "',"
    strSQL = strSQL & "obs='" & obs & "',"
    
    strSQL = strSQL & "vencimento=" & F_GravaData(vencimento) & ","

    
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where id='" & id & "'"
    
    Conn.Execute strSQL
    AlteraCartao_Presente = True
trata_Erro:
    
    
End Function
Function ExcluiCartao_Presente(id As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "delete from carta_presente "
    strSQL = strSQL & " where id='" & id & "'"
    
    Conn.Execute strSQL
    ExcluiCartao_Presente = True
trata_Erro:
    
End Function
Function AlteraVencCartaoPresente(id As String, venc As String)
        Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update cartao_presente set vencimento=" & F_GravaData(venc)
    strSQL = strSQL & " where id='" & id & "'"
    
    Conn.Execute strSQL
    AlteraVencCartaoPresente = True
trata_Erro:

End Function
Function BaixaCartaoPresente(id As String, utilizado As String, num_ven As String)
    
        Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update cartao_presente set utilizado=" & F_GravaDec(utilizado) & ","
    strSQL = strSQL & "num_ven='" & num_ven & "', "
    strSQL = strSQL & "status='2'"
    strSQL = strSQL & " where id='" & id & "'"
    
     Conn.Execute strSQL
    BaixaCartaoPresente = True
trata_Erro:

End Function


Sub ImprimeCartao_Presente(id As String)
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim Arquivo As String
    Dim style As String
    
    strSQL = "select * from cartao_presente     "
    strSQL = strSQL & " where id='" & id & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount <= 0 Then
            MsgBox "Produto não encontrado!", vbExclamation
            .Close
            Exit Sub
        End If
        Arquivo = html.F_AbreHtml("CartaoPresente_" & id, "Cartão Presente")
        html.P_HtmlAbreBody
        
        html.Paragrafo 1, "Cartão-Presente Comprovante de Transação", True, "left", "14"
        
        html.P_HtmlAbreTable 1, 1
        html.P_HtmlAbreLinha 1
        html.P_HtmlAbreColuna 1, "Transação"
        html.P_HtmlAbreColuna 1, "Nº Cartão"
        html.P_HtmlAbreColuna 1, "Valor"
        html.P_HtmlAbreColuna 1, "Nome"
        html.P_HtmlAbreColuna 1, "Contato"
        html.P_HtmlAbreColuna 1, "Data/Hora"
        html.P_HtmlAbreColuna 1, "Vencimento"
        html.P_HtmlFechaLinha 1
        html.P_HtmlAbreLinha 1
        html.P_HtmlAbreColuna 1, !id
        html.P_HtmlAbreColuna 1, !numero
        html.P_HtmlAbreColuna 1, FormatVal(!valor)
        html.P_HtmlAbreColuna 1, !nome
        html.P_HtmlAbreColuna 1, !contato
        html.P_HtmlAbreColuna 1, !data
        html.P_HtmlAbreColuna 1, !vencimento
        html.P_HtmlFechaLinha 1
        html.P_HtmlFechaTable 1
        html.P_HtmlBR 1
        
        style = "style=" & aspas & "position: absolute; width:30px;" & aspas
        html.Paragrafo 1, strNome_Emp & " - Impresso por " & strNome_Usu & " em " & Now, True
        
        html.IMG 1, App.path & "\pic\logoprint.jpg", style
        
        html.P_HtmlFechaBody
        
        
        html.P_FechaHtml 1, True, Arquivo
                
        .Close
    End With
    
    
   

    

End Sub
Function CartaoPresenteEhValido(numero As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    strSQL = "select * from cartao_presente where numero='" & numero & "' order by id desc limit 1 "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            If !status = "1" Then
                MsgBox "O cartão " & numero & " está ativo no nome de " & !nome & " no valor de " & FormatVal(!valor) & " Validade: " & !vencimento, vbExclamation
                CartaoPresenteEhValido = False
            Else
                CartaoPresenteEhValido = True
            End If
        
        Else
            If MsgBox("Este cartão nunca foi usado. Deseja ativá-lo?", vbYesNo + vbQuestion) = vbYes Then
                CartaoPresenteEhValido = True
            Else
                CartaoPresenteEhValido = False
            End If
        End If
    End With
      
    
    
End Function

