Attribute VB_Name = "SpedParticipante"
Type participante
    codigo As String
    nome As String
    COD_PAIS As String
    CNPJ As String
    CPF As String
    IE As String
    COD_MUN As String
    SUFRAMA As String
    Endereco As String
    numero As String
    complemento As String
    bairro As String
    cod_cli As String
    cep As String
    cod_for As String


End Type
Function GravaParticipante(nome As String, COD_PAIS As String, CNPJ As String, CPF As String, IE As String, _
COD_MUN As String, SUFRAMA As String, Endereco As String, numero As String, _
complemento As String, bairro As String, cod_cli As String, cod_for As String, cep As String) As Boolean
    On Error GoTo trata_erro
    
    Dim strSQL As String
    CPF = F_SoNumero(CPF)
    CNPJ = F_SoNumero(CNPJ)
    cep = F_SoNumero(cep)
    strSQL = "insert into sped_participante(NOME, COD_PAIS, CNPJ, CPF, IE, COD_MUN, SUFRAMA, ENDERECO, "
    strSQL = strSQL & "NUMERO, COMPLEMENTO, BAIRRO, cod_cli, cep, cod_for)values("
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & "'" & COD_PAIS & "',"
    strSQL = strSQL & "'" & CNPJ & "',"
    strSQL = strSQL & "'" & CPF & "',"
    strSQL = strSQL & "'" & IE & "',"
    strSQL = strSQL & "'" & COD_MUN & "',"
    strSQL = strSQL & "'" & SUFRAMA & "',"
    strSQL = strSQL & "'" & Endereco & "',"
    strSQL = strSQL & "'" & numero & "',"
    strSQL = strSQL & "'" & complemento & "',"
    strSQL = strSQL & "'" & bairro & "',"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & cep & "',"
    strSQL = strSQL & "'" & cod_for & "')"
    conn.Execute strSQL
    
    GravaParticipante = True
    Exit Function
trata_erro:

End Function

Function AlteraParticipante(codigo As String, nome As String, COD_PAIS As String, CNPJ As String, CPF As String, IE As String, _
COD_MUN As String, SUFRAMA As String, Endereco As String, numero As String, _
complemento As String, bairro As String, cod_cli As String, cod_for As String, cep As String) As Boolean
    On Error GoTo trata_erro
    Dim strSQL As String
    strSQL = "update into sped_participante set "
    strSQL = strSQL & "nome='" & nome & "',"
    strSQL = strSQL & "COD_PAIS='" & COD_PAIS & "',"
    strSQL = strSQL & "cnpj='" & CNPJ & "',"
    strSQL = strSQL & "cpf='" & CPF & "',"
    strSQL = strSQL & "ie='" & IE & "',"
    strSQL = strSQL & "cod_mun='" & COD_MUN & "',"
    strSQL = strSQL & "suframa='" & SUFRAMA & "',"
    strSQL = strSQL & "endereco='" & Endereco & "',"
    strSQL = strSQL & "numero='" & numero & "',"
    strSQL = strSQL & "complemento='" & complemento & "',"
    strSQL = strSQL & "bairro='" & bairro & "',"
    strSQL = strSQL & "cod_cli='" & cod_cli & "',"
    strSQL = strSQL & "cep='" & cep & "',"
    strSQL = strSQL & "cod_for='" & cod_for & "'"
    strSQL = strSQL & " where codigo=" & codigo
    conn.Execute strSQL
    
    AlteraParticipante = True
    Exit Function
trata_erro:

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
            F_BuscaParticipante.CPF = !CPF
            F_BuscaParticipante.Endereco = !Endereco
            F_BuscaParticipante.IE = !IE
            F_BuscaParticipante.nome = !nome
            F_BuscaParticipante.numero = !numero
            F_BuscaParticipante.SUFRAMA = !SUFRAMA
            
            
        End With
    End If
    Unload frmEdiParticipante
    
    
End Function

