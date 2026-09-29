Attribute VB_Name = "SpedParticipante"
Type participante
    codigo As String
    nome As String
    COD_PAIS As String
    CNPJ As String
    cpf As String
    IE As String
    COD_MUN As String
    SUFRAMA As String
    endereco As String
    numero As String
    complemento As String
    bairro As String
    cod_cli As String
    cep As String
    cod_for As String
    UF As String

End Type
Function ExisteDocSpedParticipante(num_doc As String, cod_ter As String, tipo_doc As String, Optional seq_Doc As String = "") As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = " select td.seq_doc from  " & TAB_DOC & " as td "
    
    
    strSQL = strSQL & " left join sped_participante as sp on sp.cod_cli=td.cod_ter or sp.cod_for=td.cod_ter "
    strSQL = strSQL & " where tipo_doc='" & tipo_doc & "' "
    strSQL = strSQL & " and td.cod_ter='" & cod_ter & "' "
    strSQL = strSQL & " and num_doc='" & num_doc & "' "
    strSQL = strSQL & " and seq_doc <>  '" & seq_Doc & "'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            
           ExisteDocSpedParticipante = !seq_Doc
            .Close
        End If
        
    End With
    
End Function

Function GravaParticipante(nome As String, COD_PAIS As String, CNPJ As String, cpf As String, IE As String, _
COD_MUN As String, SUFRAMA As String, endereco As String, numero As String, _
complemento As String, bairro As String, cod_cli As String, cod_for As String, cep As String, UF As String) As Boolean
    On Error GoTo trata_Erro
    
    Dim strSQL As String
    cpf = F_SoNumero(cpf)
    CNPJ = F_SoNumero(CNPJ)
    cep = F_SoNumero(cep)
    strSQL = "insert into sped_participante(NOME, COD_PAIS, CNPJ, CPF, IE, COD_MUN, SUFRAMA, ENDERECO, "
    strSQL = strSQL & "NUMERO, COMPLEMENTO, BAIRRO, cod_cli, cep, cod_for,uf)values("
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & "'" & COD_PAIS & "',"
    strSQL = strSQL & "'" & CNPJ & "',"
    strSQL = strSQL & "'" & cpf & "',"
    strSQL = strSQL & "'" & IE & "',"
    strSQL = strSQL & "'" & COD_MUN & "',"
    strSQL = strSQL & "'" & SUFRAMA & "',"
    strSQL = strSQL & "'" & endereco & "',"
    strSQL = strSQL & "'" & numero & "',"
    strSQL = strSQL & "'" & complemento & "',"
    strSQL = strSQL & "'" & bairro & "',"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & cep & "',"
    strSQL = strSQL & "'" & cod_for & "',"
    strSQL = strSQL & "'" & UF & "')"
    conn.Execute strSQL
    
    GravaParticipante = True
    Exit Function
trata_Erro:
    P_MostraErro
End Function

Function AlteraParticipante(codigo As String, nome As String, COD_PAIS As String, CNPJ As String, cpf As String, IE As String, _
COD_MUN As String, SUFRAMA As String, endereco As String, numero As String, _
complemento As String, bairro As String, cod_cli As String, cod_for As String, cep As String, UF As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "update  sped_participante set "
    strSQL = strSQL & "nome='" & nome & "',"
    strSQL = strSQL & "COD_PAIS='" & COD_PAIS & "',"
    strSQL = strSQL & "cnpj='" & CNPJ & "',"
    strSQL = strSQL & "cpf='" & cpf & "',"
    strSQL = strSQL & "ie='" & IE & "',"
    strSQL = strSQL & "cod_mun='" & COD_MUN & "',"
    strSQL = strSQL & "suframa='" & SUFRAMA & "',"
    strSQL = strSQL & "endereco='" & endereco & "',"
    strSQL = strSQL & "numero='" & numero & "',"
    strSQL = strSQL & "complemento='" & complemento & "',"
    strSQL = strSQL & "bairro='" & bairro & "',"
    strSQL = strSQL & "cod_cli='" & cod_cli & "',"
    strSQL = strSQL & "uf='" & UF & "',"
    strSQL = strSQL & "cep='" & cep & "',"
    strSQL = strSQL & "cod_for='" & cod_for & "'"
    strSQL = strSQL & " where codigo=" & codigo
    conn.Execute strSQL
    
    AlteraParticipante = True
    Exit Function
trata_Erro:

End Function
Function BuscaParticipanteCodigo(cpfcnpj As String) As String
    cpfcnpj = F_SoNumero(cpfcnpj)
    If cpfcnpj = "" Then Exit Function
    BuscaParticipanteCodigo = F_BuscaUmCampo("codigo", "sped_participante", "cpf='" & cpfcnpj & "' or cnpj='" & cpfcnpj & "'")
    
End Function
Function F_BuscaParticipante(expr As String) As participante
        frmEdiParticipante.tag = "x"
    frmEdiParticipante.WindowState = 2
    frmEdiParticipante.txtPesquisar.Text = expr
    frmEdiParticipante.Show 1
    
        
    If frmEdiParticipante.tag <> "" Then
        With frmEdiParticipante.ado
            F_BuscaParticipante.cep = Format(!cep)
            F_BuscaParticipante.bairro = !bairro
            F_BuscaParticipante.CNPJ = !CNPJ
            F_BuscaParticipante.cod_cli = !cod_cli
            F_BuscaParticipante.cod_for = !cod_for
            F_BuscaParticipante.COD_MUN = !COD_MUN
            F_BuscaParticipante.COD_PAIS = !COD_PAIS
            F_BuscaParticipante.codigo = !codigo
            F_BuscaParticipante.complemento = !complemento
            F_BuscaParticipante.cpf = !cpf
            F_BuscaParticipante.endereco = !endereco
            F_BuscaParticipante.IE = !IE
            F_BuscaParticipante.nome = !nome
            F_BuscaParticipante.numero = !numero
            F_BuscaParticipante.SUFRAMA = !SUFRAMA
            F_BuscaParticipante.UF = TiraNull(!UF)
            
        End With
    End If
    Unload frmEdiParticipante
    
    
End Function

