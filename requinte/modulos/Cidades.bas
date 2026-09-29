Attribute VB_Name = "Cidades"
Public Type Cidade
    nome_cid As String
    cod_cid As String
    uf_cid As String
    COD_MUN As String
End Type


Public Function F_GravaCidade(nome_cid As String, uf_cid As String) As String
    Dim sCod As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    sCod = F_BuscaUmCampo("cod_cid", "tab_cidades", "nome_cid='" & nome_cid & "'")
    If sCod = "" Then
        sCod = MaxKeySQL("tab_cidades", "cod_cid", "", "0000#")
    
    
        strSQL = "insert into tab_cidades(cod_cid,nome_cid,uf_cid)values("
        strSQL = strSQL & "'" & sCod & "',"
        strSQL = strSQL & "'" & nome_cid & "',"
        strSQL = strSQL & "'" & uf_cid & "')"
        Banco.Execute strSQL
    End If
    F_GravaCidade = sCod
    Exit Function
trata_Erro:
    
End Function

Public Function F_AlteraCidade(cod_cid As String, nome_cid As String, uf_cid As String) As Boolean
    Dim sCod As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = " update tab_cidades set"
    
    strSQL = strSQL & " nome_cid = '" & nome_cid & "',"
    strSQL = strSQL & "uf_cid='" & uf_cid & "' "
    strSQL = strSQL & " where cod_cid='" & cod_cid & "' "
    Banco.Execute strSQL
    F_AlteraCidade = True
    Exit Function
trata_Erro:
    
End Function
Function F_BuscaCidades(expr As String) As Cidade
    On Error GoTo trata_Erro
    frmEdiCid.tag = "B"
    frmEdiCid.txtbxNome.Text = expr
    frmEdiCid.Show 1
    
    If frmEdiCid.tag <> "" Then
        F_BuscaCidades.cod_cid = frmEdiCid.ado!cod_cid
        F_BuscaCidades.nome_cid = frmEdiCid.ado!nome_cid
        F_BuscaCidades.uf_cid = frmEdiCid.ado!uf_cid
        F_BuscaCidades.COD_MUN = frmEdiCid.ado!COD_MUN
    End If

trata_Erro:

End Function
Sub P_PreencheCidades(obj As Object, Optional uf As String = "")
    Dim sWhere As String
    If uf <> "" Then
        sWhere = "uf_cid='" & uf & "' "
    End If
    P_Preenche obj, "tab_cidades", "nome_cid", sWhere
    

End Sub
Function BuscaCidadesIBGE(Optional nome As String = "") As Cidades.Cidade
        frmCidadesIBGE.tag = "x"
        frmCidadesIBGE.txtbxNome.Text = nome
        frmCidadesIBGE.Show 1
        If frmCidadesIBGE.tag <> "" Then
            BuscaCidadesIBGE.cod_cid = frmCidadesIBGE.ado!cod_cid
            BuscaCidadesIBGE.nome_cid = frmCidadesIBGE.ado!nome_cid
            BuscaCidadesIBGE.uf_cid = frmCidadesIBGE.ado!uf_cid
            
            
            Unload frmCidadesIBGE
        Else
            Unload frmCidadesIBGE
        End If
       
       
End Function
Function F_BuscaCidadeCodigo(codigo As String) As Cidade
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from " & TAB_CIDADES & " where cod_cid='" & codigo & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            F_BuscaCidadeCodigo.cod_cid = !cod_cid
            F_BuscaCidadeCodigo.nome_cid = !nome_cid
            F_BuscaCidadeCodigo.uf_cid = !uf_cid
            .MoveNext
        Loop
        .Close
    End With
    
End Function

