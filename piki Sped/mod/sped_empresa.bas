Attribute VB_Name = "sped_empresa"

Public Type sped_empresa
    codigo As String
    CNPJ As String
    apelido As String
    cpf As String
    UF As String
    IE As String
    nome As String
    COD_MUN As String
    IM As String
    SUFRAMA As String
    ind_perfil As String
    ind_atv As String
    fantasia As String
    cep As String
    endereco As String
    numero As String
    complemento As String
    bairro As String
    fone As String
    fax As String
    email As String
    caminho_certificado As String
    caminho_qrcodedoc As String
    senha_certificado As String
    idCSC As String
    token As String
    
    
End Type
Public empresaSelecionada As sped_empresa


Public Function BuscaSpedEmpresa(cod_emp As String) As sped_empresa.sped_empresa
    Dim fr As New frmEdiSpedEmpresa
    fr.tag = "x"
    fr.txtPesquisar.Text = cod_emp
    fr.Show 1
    If fr.tag <> "" Then
        With fr.ado
            BuscaSpedEmpresa.codigo = !codigo
            BuscaSpedEmpresa.CNPJ = !CNPJ
            BuscaSpedEmpresa.apelido = !apelido
            BuscaSpedEmpresa.cpf = !cpf
            BuscaSpedEmpresa.UF = !UF
            BuscaSpedEmpresa.IE = !IE
            BuscaSpedEmpresa.nome = !nome
            BuscaSpedEmpresa.COD_MUN = !COD_MUN
            BuscaSpedEmpresa.IM = !IM
            BuscaSpedEmpresa.SUFRAMA = !SUFRAMA
            BuscaSpedEmpresa.ind_perfil = !ind_perfil
            BuscaSpedEmpresa.ind_atv = !ind_atv
            BuscaSpedEmpresa.fantasia = !fantasia
            BuscaSpedEmpresa.cep = !cep
            BuscaSpedEmpresa.endereco = !endereco
            BuscaSpedEmpresa.numero = !numero
            BuscaSpedEmpresa.complemento = !complemento
            BuscaSpedEmpresa.bairro = !bairro
            BuscaSpedEmpresa.fone = !fone
            BuscaSpedEmpresa.fax = !fax
            BuscaSpedEmpresa.email = !email
            BuscaSpedEmpresa.caminho_certificado = !caminho_certificado
            BuscaSpedEmpresa.caminho_qrcodedoc = TiraNull(!caminho_qrcodedoc)
            BuscaSpedEmpresa.senha_certificado = TiraNull(!senha_certificado)
            BuscaSpedEmpresa.idCSC = !idCSC
            BuscaSpedEmpresa.token = !token
        End With
    End If
    Unload fr
End Function



Public Sub gravasped_empresa(codigo As String, CNPJ As String, apelido As String, _
        cpf As String, UF As String, IE As String, nome As String, COD_MUN As String, IM As String, _
        SUFRAMA As String, ind_perfil As String, ind_atv As String, fantasia As String, _
        cep As String, endereco As String, numero As String, complemento As String, bairro As String, _
        fone As String, fax As String, email As String, caminho_certificado As String, senha_certificado As String, _
        idCSC As String, token As String, caminho_logo As String, caminho_qrcodedoc As String, CRT As String, cor As String)
        On Error GoTo erro
        Dim strSQL As String
        strSQL = " insert into sped_empresa(codigo,cnpj,apelido,cpf,UF,IE,nome,cod_mun,im,suframa,ind_perfil,"
        strSQL = strSQL & " ind_atv,fantasia,cep,endereco,numero,complemento,"
        strSQL = strSQL & " bairro,fone,fax,email,caminho_certificado,senha_certificado,idcsc,token,caminho_logo,caminho_qrcodedoc, "
        strSQL = strSQL & " crt,cor)values("
        strSQL = strSQL & "'" & codigo & "',"
        strSQL = strSQL & "'" & CNPJ & "',"
        strSQL = strSQL & "'" & apelido & "',"
        strSQL = strSQL & "'" & cpf & "',"
        strSQL = strSQL & "'" & UF & "',"
        strSQL = strSQL & "'" & IE & "',"
        strSQL = strSQL & "'" & nome & "',"
        strSQL = strSQL & "'" & COD_MUN & "',"
        strSQL = strSQL & "'" & IM & "',"
        strSQL = strSQL & "'" & SUFRAMA & "',"
        strSQL = strSQL & "'" & ind_perfil & "',"
        strSQL = strSQL & "'" & ind_atv & "',"
        strSQL = strSQL & "'" & fantasia & "',"
        strSQL = strSQL & "'" & cep & "',"
        strSQL = strSQL & "'" & endereco & "',"
        strSQL = strSQL & "'" & numero & "',"
        strSQL = strSQL & "'" & complemento & "',"
        strSQL = strSQL & "'" & bairro & "',"
        strSQL = strSQL & "'" & fone & "',"
        strSQL = strSQL & "'" & fax & "',"
        strSQL = strSQL & "'" & email & "',"

        strSQL = strSQL & "'" & AjustaCaminho(caminho_certificado) & "',"
        strSQL = strSQL & "'" & senha_certificado & "',"
        strSQL = strSQL & "'" & idCSC & "',"
        strSQL = strSQL & "'" & token & "',"
        strSQL = strSQL & "'" & AjustaCaminho(caminho_logo) & "',"
        strSQL = strSQL & "'" & AjustaCaminho(caminho_qrcodedoc) & "',"
        strSQL = strSQL & "'" & CRT & "',"
        strSQL = strSQL & "'" & cor & "'"
        strSQL = strSQL & ")"
        Conn.Execute strSQL
Exit Sub
erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub Alterasped_empresa(codigo As String, CNPJ As String, apelido As String, cpf As String, UF As String, IE As String, _
    nome As String, COD_MUN As String, IM As String, SUFRAMA As String, ind_perfil As String, ind_atv As String, _
    fantasia As String, cep As String, endereco As String, numero As String, complemento As String, bairro As String, _
    fone As String, fax As String, email As String, caminho_certificado As String, senha_certificado As String, _
        idCSC As String, token As String, caminho_logo As String, caminho_qrcodedoc As String, CRT As String, cor As String)
Dim strSQL As String
strSQL = " update sped_empresa set "
'strSQL = strSQL & "codigo='" & codigo & "',"
strSQL = strSQL & "cnpj='" & CNPJ & "',"
strSQL = strSQL & "apelido='" & apelido & "',"
strSQL = strSQL & "cpf='" & cpf & "',"
strSQL = strSQL & "UF='" & UF & "',"
strSQL = strSQL & "IE='" & IE & "',"
strSQL = strSQL & "nome='" & nome & "',"
strSQL = strSQL & "cod_mun='" & COD_MUN & "',"
strSQL = strSQL & "im='" & IM & "',"
strSQL = strSQL & "suframa='" & SUFRAMA & "',"
strSQL = strSQL & "ind_perfil='" & ind_perfil & "',"
strSQL = strSQL & "ind_atv='" & ind_atv & "',"
strSQL = strSQL & "fantasia='" & fantasia & "',"
strSQL = strSQL & "cep='" & cep & "',"
strSQL = strSQL & "endereco='" & endereco & "',"
strSQL = strSQL & "numero='" & numero & "',"
strSQL = strSQL & "complemento='" & complemento & "',"
strSQL = strSQL & "bairro='" & bairro & "',"
strSQL = strSQL & "fone='" & fone & "',"
strSQL = strSQL & "fax='" & fax & "',"

strSQL = strSQL & "email='" & email & "',"
strSQL = strSQL & "caminho_certificado='" & AjustaCaminho(caminho_certificado) & "',"

strSQL = strSQL & "caminho_logo='" & AjustaCaminho(caminho_logo) & "',"
strSQL = strSQL & "senha_certificado='" & senha_certificado & "',"
strSQL = strSQL & "idcsc='" & idCSC & "',"
strSQL = strSQL & "caminho_logo='" & AjustaCaminho(caminho_logo) & "',"
strSQL = strSQL & "caminho_qrcodedoc='" & AjustaCaminho(caminho_qrcodedoc) & "',"
strSQL = strSQL & "crt='" & CRT & "',"
strSQL = strSQL & "cor='" & cor & "',"
strSQL = strSQL & "token='" & token & "'"
        
strSQL = strSQL & " where codigo='" & codigo & "' "
Conn.Execute strSQL

End Sub
Function SelecionarSped_Empresa(Optional busca As String = "") As sped_empresa.sped_empresa
    On Error GoTo trata_Erro
    
    frmEdiSpedEmpresa.tag = "x"
    frmEdiSpedEmpresa.txtPesquisar.Text = busca
    frmEdiSpedEmpresa.Show 1
    
    If frmEdiSpedEmpresa.tag <> "" Then
        With frmEdiSpedEmpresa.ado
        
            SelecionarSped_Empresa.apelido = TiraNull(!apelido)
            SelecionarSped_Empresa.bairro = TiraNull(!bairro)
            SelecionarSped_Empresa.cep = TiraNull(!cep)
            SelecionarSped_Empresa.CNPJ = TiraNull(!CNPJ)
            SelecionarSped_Empresa.COD_MUN = TiraNull(!COD_MUN)
            SelecionarSped_Empresa.codigo = TiraNull(!codigo)
            SelecionarSped_Empresa.complemento = TiraNull(!complemento)
            SelecionarSped_Empresa.cpf = TiraNull(!cpf)
            SelecionarSped_Empresa.email = TiraNull(!email)
            SelecionarSped_Empresa.endereco = TiraNull(!endereco)
            SelecionarSped_Empresa.fantasia = TiraNull(!fantasia)
            SelecionarSped_Empresa.fax = TiraNull(!fax)
            SelecionarSped_Empresa.fone = TiraNull(!fone)
            SelecionarSped_Empresa.IE = TiraNull(!IE)
            SelecionarSped_Empresa.IM = TiraNull(!IM)
            SelecionarSped_Empresa.ind_atv = TiraNull(!ind_atv)
            SelecionarSped_Empresa.ind_perfil = TiraNull(!ind_perfil)
            SelecionarSped_Empresa.nome = TiraNull(!nome)
            SelecionarSped_Empresa.numero = TiraNull(!numero)
            SelecionarSped_Empresa.SUFRAMA = TiraNull(!SUFRAMA)
            SelecionarSped_Empresa.UF = TiraNull(!UF)
        End With
        Unload frmEdiSpedEmpresa
    End If
    Exit Function
trata_Erro:
    
End Function



Sub InicializaEmpresa(cod_emp As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from sped_empresa where codigo='" & cod_emp & "' "
    'strImprFiscal = "D" 'daruma
    abreObjetoAdo rPesq, strSQL
    
    If rPesq.RecordCount = 0 Then
        
        MsgBox "Empresa não encontrada!" & cod_emp, vbExclamation
        strCod_Emp = ""
    Else
        With rPesq
            strNumero_Emp = TiraNull(!numero)
            strComplemento_Emp = TiraNull(!complemento)
            strLogradouro_Emp = TiraNull(!endereco)
          
            'strNumero_Emp = "S/N"
            
            strNome_Emp = !nome
            strCod_Emp = !codigo
            strCRT_emp = TiraNull(!CRT)
            strCaminho_Logo = TiraNull(!caminho_logo)
            strEndereco_Emp = TiraNull(!endereco)
            strBairro_Emp = TiraNull(!bairro)
           ' strCidade_Emp = tiranull(!cidade)
           ' strIP_Servidor = !ipserver_par
            strUF_Emp = TiraNull(!UF)
            strCEP_Emp = TiraNull(!cep)
            strFone_Emp = TiraNull(!fone)
            strFax_Emp = TiraNull(!fone)
            strCNPJ_Emp = F_SoNumero(!CNPJ)
            strIE_Emp = TiraNull(!IE)
            strFantasia_Emp = TiraNull(!fantasia)
            strCodigoEstadoIBGE = "52" 'goiás
            strCodigoPaisIBGE = "1058" 'brasil
            strCodigoMunicipioIBGE = TiraNull(!COD_MUN)
            strCor_Emp = TiraNull(!cor)
            strCaminhoQrCodeDoc = TiraNull(!caminho_qrcodedoc)
            OpcaoEmpresaCRT = TiraNull(!CRT)
            
        '    strArquivoServidor = !caminhoexec_par
            .Close
        End With
    End If

    IniciaAcbr
End Sub
