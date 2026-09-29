Attribute VB_Name = "Global"
Global intCUPOM As Integer
Global strNome_Emp As String
Global strFantasia_Emp As String
Global strEndereco_Emp As String
Global strBairro_Emp As String
Global strCidade_Emp As String
Global strComplemento_Emp As String
Global strNumero_Emp As String
Global strLogradouro_Emp  As String
Global strUF_Emp As String
Global strCEP_Emp As String
Global strFone_Emp As String
Global strFax_Emp As String
Global strCNPJ_Emp As String
Global strIE_Emp As String
Global strNomeResp_Emp As String
Global strCRCResp_Emp As String
Global strNomeMoeda_Par As String
Global strsimbMoeda_Par As String
Global strPluraMoeda_Par As String
Global crcTxFederativa_Par As Currency
Global crcTxMargemSeg_Par As Currency
Global crcTxAdm As Currency
Global crcTxDiversas As Currency
Global crcComissVen As Currency 'comissão sobre as vendas
Global crcDesconto_Avista As Currency
Global strCod_Usu As String
Global strNome_Usu As String 'o nome
Global strHora As String 'teste
Global strNomeBanco As String
Global strLocalBanco As String
Global strUserBanco As String
Global strEmpresaLogada As String
Global strCnpjEmpresaLogada As String
Global strFantasiaEmpresaLogada As String


Global strCorFundoUsu As String

Global strUsaComissaoAcerto As String
Global SomenteLeitura As Boolean


Global strArquivoServidor As String
Public strImprFiscal As String
Global strCod_Car As String
Global strNome_car As String
Global strAbrevenda_usu As String
Global strAbrenota_usu As String
Global strEfetnota_usu As String
Global strAbresaida_usu As String
Global strLiberarCli_usu As String
Global strLancaixa_usu As String
Global strLandoc_usu As String
Global strLanche_usu As String
Global strBaixadoc_usu As String
Global strBaixache_usu As String
Global strEdicai_usu As String
Global strExclcai_usu As String
Global strExcldoc_usu As String
Global strCadprod_usu As String
Global strAltprod_usu As String
Global strAltpreco_usu As String
Global strNivel_Usu As String

Global strVisaoCaixa_usu As String
Global strVisaoConsignado_usu As String
Global strVisaoVendas_usu As String
Global strVisaoCrediario_usu As String
Global strVisaoGeral_usu As String
Global strVisaoPacote_usu As String
Global strVisaoFinanceiro_usu As String
Global strVisaoEstoque_usu As String
Global strProrrogarPar_Usu As String
Global strAlterarTerceiroDoc_Usu As String
Global strBaixarcheque_Usu As String
Global strExcluircheque_Usu As String
Global strEfetivarVenda_Usu As String
Global strCreditario_Usu As String
Global strVisualizarvenda_Usu As String
Global strLancarFin_Usu As String
Global strEditarFin_Usu As String
Global strExcluirFin_Usu As String
Global strTransferenciaInterna_usu As String
Global strEfetuarDevCons_Usu As String
Global strEfetuarTransfCons_Usu As String
Global strEfetuarAcerto_usu As String
Global strExcluirAcerto_usu As String
Global strAlterarComRepr_Usu As String
Global strAlterarVlrParc_Usu As String
Global strBaixarComRepr_usu As String
Global strCadastrarRepr_Usu As String
Global strCadFun_Usu As String
Global strCadfor_Usu As String
Global strCadusu_usu As String
Global strCadcli_usu As String
Global strCadcco_usu As String
Global strExclVenda_usu As String
Global strDesbloquearVenda_usu As String
Global strCadformaspag_usu As String
Global strAltformaspag_usu As String
Global strViscusto_usu As String
Global strVisrelest_usu As String
Global strCorrigeEstCons_Usu As String
Global strVisrelfin_usu As String
Global strFechacai_usu As String
Global strFechames_usu As String
Global strContrpre_usu As String
Global crcDesconto_usu As Currency
Global crcDescontoBaixa_usu As Currency
Global strIP_Usu As String 'o ip da mákina
Global strHost_Usu As String 'nome do computador
Global strCod_Cai As String 'o código do caixa buscado pelo ip
Global strSerie_Nfe As String 'serie do caixa
Global strCod_Tur As String
Global crcCpm As Currency 'o valor da venda Tj blue
Global strIP_Servidor As String
Global strCod_Emp As String 'codigo da empresa
Global picVis(9)  As PictureBox
Public Const constSigla_sis = "Piki "
Public Const constNome_sis = ".:Piki:."
Public Const constVersao_Sis = "V. 2018"
Public Const ConstNome_Cria = "Xpert Softwares"
Public Const ConstSlogan_Cria = "Soluções para sua empresa"
Public Const ConstFone_Cria = "(..62) 3098-3005"
Public Const constMsg_Usu1 = "Obrigado! Volte Sempre!"
Public Const constMsg_Usu2 = "Pense Positivo!"

Public Const ConstEmpresaEndereco = ""
Public Const ConstEmpresaTelefone = "(..62) 481-0000"
Public Const constEmpresaRamo = "ENXOVAIS"
Public Const constContaCaixa = "001" 'CÓDIGO DA CONTA CAIXA PADRÃO
Public strNomedaPromocao As String

Function IsLoadTag(Nomefrm As String, tag As String) As Boolean
    Dim i As Integer
    Dim sText As String
    For i = 0 To Forms.Count - 1
        If UCase(Forms(i).Name) = UCase(Nomefrm) Then
            If UCase(Forms(i).tag) = UCase(tag) Then
                IsLoadTag = True
                Exit Function
            End If
        End If
    Next
    IsLoadTag = False
End Function
