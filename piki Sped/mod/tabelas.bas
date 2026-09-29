Attribute VB_Name = "tabelas"
Public TAB_VENDAS As String
Public TAB_ITVENDAS As String
Public TAB_CLIENTES As String
Public TAB_FORNECEDORES As String
Public TAB_PRODUTOS As String
Public TAB_DEFUSU As String
Public TAB_VENFIN As String
Public TAB_VENFEC As String
Public TAB_DOC   As String
Public TAB_BAIXAITD    As String
Public TAB_BAIXAFIN    As String
Public TAB_ITDOC    As String
Public TAB_CAIXA As String
Public TAB_CAIXAS As String
Public TAB_TURNOS As String
Public TAB_TURNCAI As String
Public TAB_FINANCEIRO As String
Public TAB_FORMASPAG As String
Public TAB_CHEQUES As String
Public TAB_LOCALCHE As String
Public TAB_CIDADES As String
Public TAB_PLANOCONTAS As String
Public TAB_CATEGORIAS As String
Public TAB_SUBCAT As String
Public TAB_SECOES As String
Public TAB_SUBSEC As String
Public TAB_MARCAS As String
Public TAB_BANCO As String
Public TAB_IMPOSTOS As String
Public TAB_DEVOLUCOES As String
Public TAB_ITDEV As String
Public TAB_FUNCIONARIOS As String
Public TAB_USUARIOS As String
Public TAB_PARAMETROS As String



Sub iniciaTabelas()
    TAB_PARAMETROS = "tab_parametros"
    TAB_VENDAS = "vendas"
    TAB_ITVENDAS = "itvendas"
    TAB_CLIENTES = "clientes"
    TAB_FORNECEDORES = "fornecedores"
    TAB_PRODUTOS = "produtos"
    
    TAB_VENFEC = "venfec"
    TAB_VENFIN = "venfin"
    TAB_ITDOC = "itdoc"
    TAB_BAIXAITD = "baixaitd"
    TAB_BAIXAFIN = "baixafin"
    TAB_DOC = "doc"
    TAB_CAIXA = "caixa"
    TAB_CAIXAS = "tab_caixas"
    TAB_TURNOS = "turnos"
    TAB_TURNCAI = "turncai"
    TAB_FORMASPAG = "tab_formaspag"
    TAB_CHEQUES = "cheques"
    TAB_CIDADES = "TAB_cidades"
    TAB_PLANOCONTAS = "planocontas"
    TAB_LOCALCHE = "localche"
    TAB_FINANCEIRO = "financeiro"
    TAB_CATEGORIAS = "tab_categorias"
    TAB_SUBCAT = "tab_subcat"
    TAB_MARCAS = "marcas"
    TAB_SECOES = "tab_secoes"
    TAB_SUBSEC = "tab_subsec"
    TAB_CIDADES = "tab_cidades"
    TAB_CAIXA = "caixa"
    TAB_IMPOSTOS = "tab_impostos"
    TAB_DEVOLUCOES = "devolucoes"
    TAB_ITDEV = "itdev"
    TAB_FUNCIONARIOS = "tab_funcionarios"
    TAB_USUARIOS = "tab_usuarios"
    TAB_DEFUSU = "defusu"
End Sub

Sub iniciaVariaveis()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    IniciaAcbr
    IniciaConfiguracoesNFe
    iniciaParametrosSPED
    bUsaVendaSupermercado = False
    
    If BuscaParametroLista("ControlaEstoque") = "N" Then
        bControlaEstoque = False
    
    Else
        bControlaEstoque = True
    End If
    
    
    bUsaEstoqueSeparado = True
    
    strContaDevolucoes = "006"
    
    IndiceInicioLeituraBalanca = 8
    IndiceFinalLeituraBalanca = 12
    BalancaPreco = True
     strCorFundoUsu = buscaCorFundoUso(strCod_Usu)
    
    strSQL = "select * from tab_parametros "
    'strImprFiscal = "D" 'daruma
    abreObjetoAdo rPesq, strSQL
    
    If rPesq.RecordCount = 0 Then
        
        
    Else
        With rPesq
            strNumero_Emp = Format(!numero)
            strComplemento_Emp = Format(!complemento)
            'strNumero_Emp = "S/N"
            
            If !usacaixaseparado = "1" Or !usacaixaseparado = "S" Then
                bUsaCaixaSeparado = True
            End If
            
            If !usavendasupermercado = "S" Or !usavendasupermercado = "1" Then
                bUsaVendaSupermercado = True
            End If
            If !usaimpedirvendadeprodutosemestoque = "S" Then
                bUsaImpedirVendaDeProdutoSemEstoque = True
            End If
            
            
            
            
            strFormaPagamentoPadrao = !formapagamentopadrao
            strCodigoEstadoIBGE = "52" 'goiás
            strCodigoPaisIBGE = "1058" 'brasil
            strCodigoMunicipioIBGE = "5201108"
            strCRT_emp = SIMPLES_NACIONAL
            strTipo_IcmsNFe = !tipoicms_nfe
            strCST_Padrao = F_SoNumero(!tipoicms_nfe)
            strNome_Emp = !nome
            strCod_Emp = "01"
            strLogradouro_Emp = !Logradouro
            strEndereco_Emp = !Logradouro & " " & !numero & " " & !complemento
            strBairro_Emp = !bairro
            strCidade_Emp = !Cidade
            'strIP_Servidor = !ipserver_par
            strUF_Emp = !UF
            strCEP_Emp = F_SoNumero(!cep)
            strFone_Emp = F_SoNumero(!fone)
            'strFax_Emp = !fax
            strCNPJ_Emp = F_SoNumero(!CNPJ)
            strIE_Emp = F_SoNumero(!IE)
            'strIP_TerminalTEF = !ipterminaltef_par
            'strId_Estab = !idestabtef_par
            strFantasia_Emp = !fantasia
            strNomeResp_Emp = Format(!responsavel)
            'strNomeMoeda_Par = !NomeMoeda_Par
            'strsimbMoeda_Par = !simbMoeda_Par
            'strPluraMoeda_Par = !PLuralMoeda_PAr
            crcTxFederativa_Par = !tx_federal
            crcTxAdm = !tx_adm
            crcTxMargemSeg_Par = !MargemSeg
            'strPlaVendasEfetuadas = !plavendas_par
            strPlaRecebDupl = "10001"
            strPlaPagDupl = "20001"
            strPlaNaoClassC = "10001"
            strPlaNaoClassD = "20001"
            
            strArquivoServidor = !caminhoexec_par
            'If !usatef_par = "S" Then
                bUsaTEF = False
                bUsaECF = False
            'Else
            '    bUsaTEF = False
            'End If
            'strArquivoServidor = !caminhoexec_par
            
            .Close
        End With
    End If
    
    P_BuscaDefUsu
    
    
    OpcaoEmpresaCRT = SIMPLES_NACIONAL
    
    InicializaEmpresa strCod_Emp
    
    buscaCaminhoImagem
    Exit Sub
trata_Erro:
    P_MostraErro
    MsgBox Err.Description
    P_Fim
End Sub
