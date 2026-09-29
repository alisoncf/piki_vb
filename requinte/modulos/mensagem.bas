Attribute VB_Name = "Mensagem"
Sub PostaMensagem(tipo_ter As String, cod_ter As String, mensagem As String)
    Dim strSQL As String
    strSQL = "insert into tab_recados(seq_rec,de_rec,para_rec,tipo_ter,mensagem_rec,data_rec,status_rec)values("
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & cod_ter & "',"
    strSQL = strSQL & "'" & tipo_ter & "',"
    strSQL = strSQL & "'" & mensagem & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'1')"
    Conn.Execute strSQL
End Sub
Sub alteraStatusMsg(seq_rec As String, status As String)
    Dim strSQL As String
    strSQL = "update tab_recados set status_rec='" & status & "'"
    strSQL = strSQL & " where seq_rec='" & seq_rec & "'"
    Conn.Execute strSQL
End Sub
Sub ExcluiMsg(seq_rec As String)
    Dim strSQL As String
    strSQL = "delete from tab_recados "
    strSQL = strSQL & " where seq_rec='" & seq_rec & "'"
    Conn.Execute strSQL
End Sub
Sub BuscaMensagem(tipo_ter As String, cod_ter As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If tipo_ter = "U" Then
        strSQL = "select tf.nome_fun,tf2.guerra_fun as de,tr.* "
        strSQL = strSQL & " from tab_recados as tr,tab_funcionarios as tf, tab_funcionarios as tf2"
        strSQL = strSQL & " where tr.de_rec=tf2.cod_fun "
        strSQL = strSQL & " and tr.para_rec=tf.cod_fun "
        strSQL = strSQL & " and status_rec='1' "
        strSQL = strSQL & " and tf.cod_fun='" & cod_ter & "'"
        abreObjetoAdo rPesq, strSQL
    
        With rPesq
            If .RecordCount > 0 Then
                strSQL = "Você tem " & .RecordCount & " mensagens não lidas na sua caixa de mensagem." & Chr(13)
                strSQL = strSQL & "Deseja abrir agora?"
                If MsgBox(strSQL, vbYesNo + vbQuestion) = vbYes Then
                    frmNovoRecado.tag = "U"
                    
                    frmNovoRecado.Show 0, mdiFrm
                    
                    frmNovoRecado.SetFocus
                End If
                .Close
            End If
        End With
    
    Else
        strSQL = "select tf.nome_rep,tf2.guerra_fun as de,tr.* "
        strSQL = strSQL & " from tab_recados as tr,tab_representantes as tf, tab_funcionarios as tf2"
        strSQL = strSQL & " where tr.de_rec=tf2.cod_fun "
        strSQL = strSQL & " and tr.para_rec=tf.cod_rep "
        strSQL = strSQL & " and status_rec='1' "
        strSQL = strSQL & " and tf.cod_rep='" & cod_ter & "'"
        
        
        abreObjetoAdo rPesq, strSQL
    
        With rPesq
            If .RecordCount > 0 Then
            '    strSQL = "Este representante tem " & .RecordCount & " mensagens não lidas na sua caixa de mensagem." & Chr(13)
            '    strSQL = strSQL & "Deseja abrir agora?"
            '    If MsgBox(strSQL, vbYesNo + vbQuestion) = vbYes Then
            '        On Error Resume Next
            '        frmNovoRecado.tag = "x"
            '        frmNovoRecado.tag = "R"
            '
            '        frmNovoRecado.txtbxTerceiro.tag = cod_ter
            '        'On Error Resume Next
            '        frmNovoRecado.Show 0, mdiFrm
            '        frmNovoRecado.SetFocus
            '    End If
                .Close
            End If
        End With
    
    End If
    
End Sub
Function F_SelectMensagem(tipo_ter As String, cod_ter As String, Optional todas As Boolean = False) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If tipo_ter = "U" Then
        strSQL = "select tf.nome_fun,tf2.guerra_fun as de,tr.* "
        strSQL = strSQL & " from tab_recados as tr,tab_funcionarios as tf, tab_funcionarios as tf2"
        strSQL = strSQL & " where tr.de_rec=tf2.cod_fun "
        strSQL = strSQL & " and tr.para_rec=tf.cod_fun "
        strSQL = strSQL & " and tf.cod_fun='" & cod_ter & "' "
        If todas = False Then
            strSQL = strSQL & " and status_rec='1' "
        End If
    Else
        strSQL = "select tf.nome_rep,tf2.guerra_fun as de,tr.* "
        strSQL = strSQL & " from tab_recados as tr,tab_representantes as tf, tab_funcionarios as tf2"
        strSQL = strSQL & " where tr.de_rec=tf2.cod_fun "
        strSQL = strSQL & " and tr.para_rec=tf.cod_rep "
        strSQL = strSQL & " and tf.cod_rep='" & cod_ter & "'"
        If todas = False Then
            strSQL = strSQL & " and status_rec='1' "
        End If
    
    End If
    F_SelectMensagem = strSQL
End Function

Sub GravaMensagem(para As String, mensagem As String)
    Dim strSQL As String
    strSQL = "insert into mensagens(de,para,mensagem,status,data)values("
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & para & "',"
    strSQL = strSQL & "'" & mensagem & "',"
    strSQL = strSQL & "1,"
    strSQL = strSQL & F_GravaDataTempo(Now) & ") "
    Conn.Execute strSQL

End Sub


Sub MudaStatusMensagem(id As String, status As String)
    Dim strSQL As String
    strSQL = "update mensagens set status=" & status
    
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL

End Sub
Function RetornoMsgNaoLida(para As String) As String
    RetornoMsgNaoLida = F_BuscaUmCampo("count(*)", "mensagens", "para='" & para & "' and status=1")

End Function
Function RetornoMsg(para As String) As String
    RetornoMsg = F_BuscaUmCampo("count(*)", "mensagens", "para='" & para & "'")

End Function


