Attribute VB_Name = "Contrato"
Public caminho_contratos As String

Function buscaDocs_Contrato(seq_doc As String) As ADODB.Recordset
    Dim strSQL As String
    
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from documentos_contrato where seq_doc='" & seq_doc & "' and (excluido  is null OR excluido<>'S') "
    abreObjetoAdo rPesq, strSQL
    
    
    Set buscaDocs_Contrato = rPesq
    
    
End Function
Sub insere_documento_contrato(seq_doc As String, arquivo As String, hash As String)
    Dim strSQL As String
    If hash <> "" Then
        If F_BuscaUmCampo("hash", "documentos_contrato", "seq_doc='" & seq_doc & "' and hash='" & hash & "'") <> "" Then
            MsgBox "já existe um documento com este conteúdo", vbExclamation
            Exit Sub
        End If
    End If
    
    
    strSQL = "Insert into documentos_contrato(seq_doc,caminho,data,cod_usu,hash)values("
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & "'" & AjustaCaminho(arquivo) & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & hash & "')"
    Conn.Execute strSQL
End Sub
Sub excluir_arquivo(seq_doc As String, arquivo As String)
    Dim strSQL As String
    strSQL = "update documentos_contrato set excluido ='S' where seq_doc='" & seq_doc & "'"
    strSQL = strSQL & " and caminho like '%" & arquivo & "' "
    Conn.Execute strSQL
End Sub
