Attribute VB_Name = "daruma"

 
 Global Int_Retorno As Integer
 Global Int_Ack As Integer
 Global Int_St1 As Integer
 Global Int_St2 As Integer
 Global Str_ErroExtendido As String
 Global Int_Tipo_Aliquota As Integer
 Global Str_Informacao As String
 Global Int_Informacao As Integer
 
 
     'Metodos Cupom
 
 Public Declare Function Daruma_FI_AbreCupom Lib "Daruma32.dll" (ByVal Cpf_ou_Cnpj As String) As Integer
 Public Declare Function Daruma_FI_VendeItem Lib "Daruma32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
 Public Declare Function Daruma_FI_VendeItemDepartamento Lib "Daruma32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
 Public Declare Function Daruma_FI_VendeItemTresDecimais Lib "Daruma32.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal Desconto_ou_Acrescimo As String, ByVal Percentual_Desconto_ou_Acrescimo As String) As Integer
 Public Declare Function Daruma_FI_FechaCupomResumido Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
 Public Declare Function Daruma_FI_IniciaFechamentoCupom Lib "Daruma32.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
 Public Declare Function Daruma_FI_EfetuaFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
 Public Declare Function Daruma_FI_EfetuaFormaPagamentoDescricaoForma Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal TextoLivre As String) As Integer
 Public Declare Function Daruma_FI_IdentificaConsumidor Lib "Daruma32.dll" (ByVal NomeConsumidor As String, ByVal Endereco As String, ByVal Cpf_ou_Cnpj As String) As Integer
 Public Declare Function Daruma_FI_TerminaFechamentoCupom Lib "Daruma32.dll" (ByVal Mensagem As String) As Integer
 Public Declare Function Daruma_FI_FechaCupom Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
 Public Declare Function Daruma_FI_CancelaItemAnterior Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_CancelaItemGenerico Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
 Public Declare Function Daruma_FI_CancelaCupom Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_AumentaDescricaoItem Lib "Daruma32.dll" (ByVal Descricao As String) As Integer
 Public Declare Function Daruma_FI_UsaUnidadeMedida Lib "Daruma32.dll" (ByVal UnidadeMedida As String) As Integer
 Public Declare Function Daruma_FI_EmitirCupomAdicional Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_EstornoFormasPagamento Lib "Daruma32.dll" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer
  
 
 'Metodos para Recebimentos e Relatorios
 Public Declare Function Daruma_FI_AbreComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
 Public Declare Function Daruma_FI_UsaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal Texto As String) As Integer
 Public Declare Function Daruma_FI_FechaComprovanteNaoFiscalVinculado Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_RelatorioGerencial Lib "Daruma32.dll" (ByVal Texto As String) As Integer
 Public Declare Function Daruma_FI_AbreRelatorioGerencial Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_EnviarTextoCNF Lib "Daruma32.dll" (ByVal Texto As String) As Integer
 Public Declare Function Daruma_FI_FechaRelatorioGerencial Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_RecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal IndiceTotalizador As String, ByVal ValorRecebimento As String, ByVal FormaPagamento As String) As Integer
 Public Declare Function Daruma_FI_AbreRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal IndiceTotalizador As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal ValorRecebimento As String, ByVal TextoLivreto As String) As Integer
 Public Declare Function Daruma_FI_EfetuaFormaPagamentoNaoFiscal Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal ObsLivre As String) As Integer
 Public Declare Function Daruma_FI_Sangria Lib "Daruma32.dll" (ByVal Valor As String) As Integer
 Public Declare Function Daruma_FI_Suprimento Lib "Daruma32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
 Public Declare Function Daruma_FI_FundoCaixa Lib "Daruma32.dll" (ByVal ValorPagamento As String, ByVal FormaPagamento As String) As Integer
 Public Declare Function Daruma_FI_LeituraX Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_ReducaoZ Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
 Public Declare Function Daruma_FI_ReducaoZAjustaDataHora Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
 Public Declare Function Daruma_FI_LeituraMemoriaFiscalData Lib "Daruma32.dll" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
 Public Declare Function Daruma_FI_LeituraMemoriaFiscalReducao Lib "Daruma32.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String) As Integer
 Public Declare Function Daruma_FI_LeituraMemoriaFiscalSerialData Lib "Daruma32.dll" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
 Public Declare Function Daruma_FI_LeituraMemoriaFiscalSerialReducao Lib "Daruma32.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String) As Integer
 
 'Metodos Modo Gaveta, Autentica e Outras
 Public Declare Function Daruma_FI_VerificaDocAutenticacao Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_Autenticacao Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_AutenticacaoStr Lib "Daruma32.dll" (ByVal AtenticacaoStr As String) As Integer
 Public Declare Function Daruma_FI_VerificaEstadoGaveta Lib "Daruma32.dll" (ByRef EstadoGaveta As Integer) As Integer
 Public Declare Function Daruma_FI_VerificaEstadoGavetaStr Lib "Daruma32.dll" (ByVal EstadoGaveta As String) As Integer
 Public Declare Function Daruma_FI_AcionaGaveta Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_AbrePortaSerial Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_FechaPortaSerial Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_AberturaDoDia Lib "Daruma32.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
 Public Declare Function Daruma_FI_FechamentoDoDia Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_ImprimeConfiguracoesImpressora Lib "Daruma32.dll" () As Integer
  
 
 'Metodos Prog e Config
 
 Public Declare Function Daruma_FI_ProgramaAliquota Lib "Daruma32.dll" (ByVal Aliquota As String, ByVal ICMS_ou_ISS As Integer) As Integer
 Public Declare Function Daruma_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "Daruma32.dll" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
 Public Declare Function Daruma_FI_ProgramaFormasPagamento Lib "Daruma32.dll" (ByVal DescricaoFormasPgto As String) As Integer
 Public Declare Function Daruma_FI_ProgramaOperador Lib "Daruma32.dll" (ByVal NomeOperador As String) As Integer
 Public Declare Function Daruma_FI_ProgramaArredondamento Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_ProgramaTruncamento Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_LinhasEntreCupons Lib "Daruma32.dll" (ByVal Linhas As Integer) As Integer
 Public Declare Function Daruma_FI_EspacoEntreLinhas Lib "Daruma32.dll" (ByVal Dots As Integer) As Integer
 Public Declare Function Daruma_FI_ProgramaHorarioVerao Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_EqualizaFormasPgto Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_ProgramaVinculados Lib "Daruma32.dll" (ByVal Descricao As String) As Integer
 Public Declare Function Daruma_FI_ProgFormasPagtoSemVincular Lib "Daruma32.dll" (ByVal Descricao As String) As Integer
 
  
 
 'Metodos de Configuracao do ECF
 
 Public Declare Function Daruma_FI_CfgFechaAutomaticoCupom Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgRedZAutomatico Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgLeituraXAuto Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgImpEstGavVendas Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgCalcArredondamento Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgHorarioVerao Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgSensorAut Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgCupomAdicional Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgPermMensPromCNF Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgEspacamentoCupons Lib "Daruma32.dll" (ByVal DistanciaCupons As String) As Integer
 Public Declare Function Daruma_FI_CfgHoraMinReducaoZ Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_CfgLimiarNearEnd Lib "Daruma32.dll" (ByVal NumeroLinhas As String) As Integer
 Public Declare Function Daruma_FI_CfgLegProdutos Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 
 'Metodos para Configuracao do Registry
 Public Declare Function Daruma_Registry_AlteraRegistry Lib "Daruma32.dll" (ByVal NomeChave As String, ByVal ValorChave As String) As Integer
 Public Declare Function Daruma_Registry_Porta Lib "Daruma32.dll" (ByVal NomePorta As String) As Integer
 Public Declare Function Daruma_Registry_Path Lib "Daruma32.dll" (ByVal Path As String) As Integer
 Public Declare Function Daruma_Registry_Status Lib "Daruma32.dll" (ByVal Status As String) As Integer
 Public Declare Function Daruma_Registry_StatusFuncao Lib "Daruma32.dll" (ByVal StatusFuncao As String) As Integer
 Public Declare Function Daruma_Registry_Retorno Lib "Daruma32.dll" (ByVal Retorno As String) As Integer
 Public Declare Function Daruma_Registry_ControlePorta Lib "Daruma32.dll" (ByVal ControlePorta As String) As Integer
 Public Declare Function Daruma_Registry_ModoGaveta Lib "Daruma32.dll" (ByVal ModoGaveta As String) As Integer
 Public Declare Function Daruma_Registry_Log Lib "Daruma32.dll" (ByVal Log As String) As Integer
 Public Declare Function Daruma_Registry_NomeLog Lib "Daruma32.dll" (ByVal NomeLog As String) As Integer
 Public Declare Function Daruma_Registry_Separador Lib "Daruma32.dll" (ByVal Separador As String) As Integer
 Public Declare Function Daruma_Registry_SeparaMsgPromo Lib "Daruma32.dll" (ByVal SeparaMsgPromo As String) As Integer
 Public Declare Function Daruma_Registry_ZAutomatica Lib "Daruma32.dll" (ByVal ZAutomatica As String) As Integer
 Public Declare Function Daruma_Registry_XAutomatica Lib "Daruma32.dll" (ByVal XAutomatica As String) As Integer
 Public Declare Function Daruma_Registry_VendeItemUmaLinha Lib "Daruma32.dll" (ByVal VendeItem1Lin13Dig As String) As Integer
 Public Declare Function Daruma_Registry_Default Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_Registry_RetornaValor Lib "Daruma32.dll" (ByVal Produto As String, ByVal Chave As String, ByVal Valor As String) As Integer
 Public Declare Function Daruma_Registry_TerminalServer Lib "Daruma32.dll" (ByVal TerminalServer As String) As Integer
 Public Declare Function Daruma_Registry_ErroExtendidoOk Lib "Daruma32.dll" (ByVal ValorErro As String) As Integer
 Public Declare Function Daruma_Registry_AbrirDiaFiscal Lib "Daruma32.dll" (ByVal AbrirDiaFiscal As String) As Integer
 Public Declare Function Daruma_Registry_VendaAutomatica Lib "Daruma32.dll" (ByVal VendaAutomatica As String) As Integer
 Public Declare Function Daruma_Registry_IgnorarPoucoPapel Lib "Daruma32.dll" (ByVal IgnorarPoucoPapel As String) As Integer
 Public Declare Function Daruma_Registry_ImprimeRegistry Lib "Daruma32.dll" (ByVal Produto As String) As Integer
 
  
 
 'Metodos de Status
 
 Public Declare Function Daruma_FI_StatusCupomFiscal Lib "Daruma32.dll" (ByVal StsCF As String) As Integer
 Public Declare Function Daruma_FI_StatusRelatorioGerencial Lib "Daruma32.dll" (ByVal StsGerencial As String) As Integer
 Public Declare Function Daruma_FI_StatusComprovanteNaoFiscalVinculado Lib "Daruma32.dll" (ByVal StsCNFV As String) As Integer
 Public Declare Function Daruma_FI_VerificaImpressoraLigada Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_VerificaTotalizadoresParciais Lib "Daruma32.dll" (ByVal Totalizadores_Parciais As String) As Integer
 Public Declare Function Daruma_FI_VerificaModoOperacao Lib "Daruma32.dll" (ByVal Modo As String) As Integer
 Public Declare Function Daruma_FI_VerificaTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal Totalizadores As String) As Integer
 Public Declare Function Daruma_FI_VerificaTotalizadoresNaoFiscaisEx Lib "Daruma32.dll" (ByVal Totalizadores As String) As Integer
 Public Declare Function Daruma_FI_VerificaTruncamento Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_VerificaAliquotasIss Lib "Daruma32.dll" (ByVal AliquotasIss As String) As Integer
 Public Declare Function Daruma_FI_VerificaRecebimentoNaoFiscal Lib "Daruma32.dll" (ByVal Recebimentos As String) As Integer
 Public Declare Function Daruma_FI_VerificaTipoImpressora Lib "Daruma32.dll" (ByRef TipoImpressora As Integer) As Integer
 Public Declare Function Daruma_FI_VerificaIndiceAliquotasIss Lib "Daruma32.dll" (ByVal AliquotaIss As String) As Integer
 Public Declare Function Daruma_FI_VerificaModeloECF Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_VerificaDescricaoFormasPagamento Lib "Daruma32.dll" (ByVal Descricao As String) As Integer
 Public Declare Function Daruma_FI_VerificaXPendente Lib "Daruma32.dll" (ByVal XPendente As String) As Integer
 Public Declare Function Daruma_FI_VerificaZPendente Lib "Daruma32.dll" (ByVal ZPendente As String) As Integer
 Public Declare Function Daruma_FI_VerificaDiaAberto Lib "Daruma32.dll" (ByVal DiaAberto As String) As Integer
 Public Declare Function Daruma_FI_VerificaHorarioVerao Lib "Daruma32.dll" (ByVal HorarioVerao As String) As Integer
 Public Declare Function Daruma_FI_VerificaFormasPagamento Lib "Daruma32.dll" (ByVal FormasPagto As String) As Integer
 Public Declare Function Daruma_FI_VerificaFormasPagamentoEx Lib "Daruma32.dll" (ByVal FormasPagtoEx As String) As Integer
 Public Declare Function Daruma_FI_VerificaEpromConectada Lib "Daruma32.dll" (ByVal Flag As String) As Integer
 Public Declare Function Daruma_FI_VerificaEstadoImpressora Lib "Daruma32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
 
 
 'Metodos de Informacao do ECF e Contadores
 
 Public Declare Function Daruma_FI_ClicheProprietario Lib "Daruma32.dll" (ByVal Cliche As String) As Integer
 Public Declare Function Daruma_FI_ClicheProprietarioEx Lib "Daruma32.dll" (ByVal ClicheProprietarioEx As String) As Integer
 Public Declare Function Daruma_FI_NumeroCaixa Lib "Daruma32.dll" (ByVal NumeroCaixa As String) As Integer
 Public Declare Function Daruma_FI_NumeroLoja Lib "Daruma32.dll" (ByVal NumeroLoja As String) As Integer
 Public Declare Function Daruma_FI_NumeroSerie Lib "Daruma32.dll" (ByVal NumeroSerie As String) As Integer
 Public Declare Function Daruma_FI_VersaoFirmware Lib "Daruma32.dll" (ByVal VersaoFirmware As String) As Integer
 Public Declare Function Daruma_FI_CGC_IE Lib "Daruma32.dll" (ByVal CGC As String, ByVal IE As String) As Integer
 Public Declare Function Daruma_FI_LerAliquotasComIndice Lib "Daruma32.dll" (ByVal AliquotasComIndice As String) As Integer
 Public Declare Function Daruma_FI_NumeroCupom Lib "Daruma32.dll" (ByVal NumeroCupom As String) As Integer
 Public Declare Function Daruma_FI_COO Lib "Daruma32.dll" (ByVal Coo_Inicial As String, ByVal COO_Final As String) As Integer
 Public Declare Function Daruma_FI_MinutosImprimindo Lib "Daruma32.dll" (ByVal Minutos As String) As Integer
 Public Declare Function Daruma_FI_MinutosLigada Lib "Daruma32.dll" (ByVal Minutos As String) As Integer
 Public Declare Function Daruma_FI_NumeroSubstituicoesProprietario Lib "Daruma32.dll" (ByVal NumeroSubstituicoes As String) As Integer
 Public Declare Function Daruma_FI_NumeroIntervencoes Lib "Daruma32.dll" (ByVal NumeroIntervencoes As String) As Integer
 Public Declare Function Daruma_FI_NumeroReducoes Lib "Daruma32.dll" (ByVal NumeroReducoes As String) As Integer
 Public Declare Function Daruma_FI_NumeroCuponsCancelados Lib "Daruma32.dll" (ByVal NumeroCancelamentos As String) As Integer
 Public Declare Function Daruma_FI_NumeroOperacoesNaoFiscais Lib "Daruma32.dll" (ByVal NumeroOperacoes As String) As Integer
 Public Declare Function Daruma_FI_DataHoraImpressora Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
 Public Declare Function Daruma_FI_DataHoraReducao Lib "Daruma32.dll" (ByVal Data As String, ByVal Hora As String) As Integer
 Public Declare Function Daruma_FI_DataMovimento Lib "Daruma32.dll" (ByVal Data As String) As Integer
 Public Declare Function Daruma_FI_ContadoresTotalizadoresNaoFiscais Lib "Daruma32.dll" (ByVal Contadores As String) As Integer
 
 'Metodos Totalizadores Gerais
 
 Public Declare Function Daruma_FI_GrandeTotal Lib "Daruma32.dll" (ByVal GrandeTotal As String) As Integer
 Public Declare Function Daruma_FI_Descontos Lib "Daruma32.dll" (ByVal ValorDescontos As String) As Integer
 Public Declare Function Daruma_FI_Acrescimos Lib "Daruma32.dll" (ByVal ValorAcrescimos As String) As Integer
 Public Declare Function Daruma_FI_Cancelamentos Lib "Daruma32.dll" (ByVal ValorCancelamentos As String) As Integer
 Public Declare Function Daruma_FI_DadosUltimaReducao Lib "Daruma32.dll" (ByVal DadosReducao As String) As Integer
 Public Declare Function Daruma_FI_SubTotal Lib "Daruma32.dll" (ByVal SubTotal As String) As Integer
 Public Declare Function Daruma_FI_RetornoAliquotas Lib "Daruma32.dll" (ByVal cAliquotas As String) As Integer
 Public Declare Function Daruma_FI_ValorPagoUltimoCupom Lib "Daruma32.dll" (ByVal ValorCupom As String) As Integer
 Public Declare Function Daruma_FI_ValorFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
 Public Declare Function Daruma_FI_ValorTotalizadorNaoFiscal Lib "Daruma32.dll" (ByVal Totalizador As String, ByVal ValorTotalizador As String) As Integer
 Public Declare Function Daruma_FI_UltimoItemVendido Lib "Daruma32.dll" (ByVal NumeroItem As String) As Integer
 Public Declare Function Daruma_FI_UltimaFormaPagamento Lib "Daruma32.dll" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
 Public Declare Function Daruma_FI_TipoUltimoDocumento Lib "Daruma32.dll" (ByVal TipoUltimoDoc As String) As Integer
 
 
 'Metodos Relatorios Fiscais e Relatorios
 
 Public Declare Function Daruma_FI_MapaResumo Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_RelatorioTipo60Analitico Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_RelatorioTipo60Mestre Lib "Daruma32.dll" () As Integer
 Public Declare Function Daruma_FI_FlagsFiscais Lib "Daruma32.dll" (ByRef Flag As Integer) As Integer
 Public Declare Function Daruma_FI_PalavraStatus Lib "Daruma32.dll" (ByVal PalavraStatus As String) As Integer
 Public Declare Function Daruma_FI_PalavraStatusBinario Lib "Daruma32.dll" (ByVal PalavraStatusBinario As String) As Integer
 Public Declare Function Daruma_FI_SimboloMoeda Lib "Daruma32.dll" (ByVal SimboloMoeda As String) As Integer
 Public Declare Function Daruma_FI_RetornoImpressora Lib "Daruma32.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
 Public Declare Function Daruma_FI_RetornaErroExtendido Lib "Daruma32.dll" (ByVal ErroExtendido As String) As Integer
 Public Declare Function Daruma_FI_RetornaAcrescimoNF Lib "Daruma32.dll" (ByVal AcrescimoNF As String) As Integer
 Public Declare Function Daruma_FI_RetornaCFCancelados Lib "Daruma32.dll" (ByVal CNCancelados As String) As Integer
 Public Declare Function Daruma_FI_RetornaCNFCancelados Lib "Daruma32.dll" (ByVal CNFCancelados As String) As Integer
 Public Declare Function Daruma_FI_RetornaCLX Lib "Daruma32.dll" (ByVal RetornaCLX As String) As Integer
 Public Declare Function Daruma_FI_RetornaCNFNV Lib "Daruma32.dll" (ByVal RetornaCNFNV As String) As Integer
 Public Declare Function Daruma_FI_RetornaCNFV Lib "Daruma32.dll" (ByVal RetornaCNFV As String) As Integer
 Public Declare Function Daruma_FI_RetornaCRO Lib "Daruma32.dll" (ByVal RetornaCRO As String) As Integer
 Public Declare Function Daruma_FI_RetornaCRZ Lib "Daruma32.dll" (ByVal RetornaCRZ As String) As Integer
 Public Declare Function Daruma_FI_RetornaCRZRestante Lib "Daruma32.dll" (ByVal RetornaCRZRestante As String) As Integer
 Public Declare Function Daruma_FI_RetornaCancelamentoNF Lib "Daruma32.dll" (ByVal CancelamentoNF As String) As Integer
 Public Declare Function Daruma_FI_RetornaDescontoNF Lib "Daruma32.dll" (ByVal DescontoNF As String) As Integer
 Public Declare Function Daruma_FI_RetornaGNF Lib "Daruma32.dll" (ByVal RetornaGNF As String) As Integer
 Public Declare Function Daruma_FI_RetornaTempoImprimindo Lib "Daruma32.dll" (ByVal TempoImprimindo As String) As Integer
 Public Declare Function Daruma_FI_RetornaTempoLigado Lib "Daruma32.dll" (ByVal TempoLigado As String) As Integer
 Public Declare Function Daruma_FI_RetornaTotalPagamentos Lib "Daruma32.dll" (ByVal TotalPagamentos As String) As Integer
 Public Declare Function Daruma_FI_RetornaTroco Lib "Daruma32.dll" (ByVal Troco As String) As Integer
 Public Declare Function Daruma_FI_RetornaValorComprovanteNaoFiscal Lib "Daruma32.dll" (ByVal IndiceRegistrCNF As String, ByVal Ref_Valor As String) As Integer
 Public Declare Function Daruma_FI_RetornaIndiceComprovanteNaoFiscal Lib "Daruma32.dll" (ByVal DescricaoRegistrCNF As String, ByVal Ref_Indice As String) As Integer
 Public Declare Function Daruma_FI_RetornaRegistradoresNaoFiscais Lib "Daruma32.dll" (ByVal RegistrNaoFiscais As String) As Integer
 Public Declare Function Daruma_FI_RetornaRegistradoresFiscais Lib "Daruma32.dll" (ByVal RegistradoresFiscais As String) As Integer
 
 'Metodos TEF
 
 Public Declare Function Daruma_TEF_EsperarArquivo Lib "Daruma32.dll" (ByVal PathArquivo As String, ByVal Tempo As String, ByVal Travar As String) As Integer
 
 Public Declare Function Daruma_TEF_ImprimirResposta Lib "Daruma32.dll" (ByVal PathArquivo As String, ByVal Forma As String, ByVal Travar As String) As Integer
 
 Public Declare Function Daruma_TEF_ImprimirRespostaCartao Lib "Daruma32.dll" (ByVal PathArquivo As String, ByVal Forma As String, ByVal Travar As String, ByVal ValorPagamento As String) As Integer
 
 Public Declare Function Daruma_TEF_FechaRelatorio Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_TEF_SetFocus Lib "Daruma32.dll" (ByVal TituloJanela As String) As Integer
 
 Public Declare Function Daruma_TEF_TravarTeclado Lib "Daruma32.dll" (ByVal Travar As String) As Integer
 
  
 
 '**********************************************************************************************************************'
 
 ' '
 
 '                   FS2000 '
 
 ' '
 
 '**********************************************************************************************************************'
 
  
 
 'Metodos exclusivos
 
 Public Declare Function Daruma_Registry_FS2000_CupomAdicional Lib "Daruma32.dll" (ByVal CupomAdicional As String) As Integer
 
 Public Declare Function Daruma_Registry_FS2000_TempoEsperaCheque Lib "Daruma32.dll" (ByVal TempodeEspera As String) As Integer
 
 Public Declare Function Daruma_FI2000_DescontoSobreItemVendido Lib "Daruma32.dll" (ByVal NumeroItem As String, ByVal TipoDesconto As String, ByVal ValorDesconto As String) As Integer
 
 Public Declare Function Daruma_FI2000_AcrescimosICMSISS Lib "Daruma32.dll" (ByVal AcrescICMS As String, ByVal AcrescISS As String) As Integer
 
 Public Declare Function Daruma_FI2000_CancelamentosICMSISS Lib "Daruma32.dll" (ByVal CancelICMS As String, ByVal CancelISS As String) As Integer
 
 Public Declare Function Daruma_FI2000_DescontosICMSISS Lib "Daruma32.dll" (ByVal DescICMS As String, ByVal DescISS As String) As Integer
 
 Public Declare Function Daruma_FI2000_LeituraInformacaoUltimosCNF Lib "Daruma32.dll" (ByVal UltimosCNF As String) As Integer
 
 Public Declare Function Daruma_FI2000_LeituraInformacaoUltimoDoc Lib "Daruma32.dll" (ByVal TipoUltimoDoc As String, ByVal ValorUltimoDoc As String) As Integer
 
 Public Declare Function Daruma_FI2000_VerificaRelatorioGerencial Lib "Daruma32.dll" (ByVal Gerencial As String) As Integer
 
 Public Declare Function Daruma_FI2000_CriaRelatorioGerencial Lib "Daruma32.dll" (ByVal NomeGerencial As String) As Integer
 
 Public Declare Function Daruma_FI2000_AbreRelatorioGerencial Lib "Daruma32.dll" (ByVal Indice As String) As Integer
 
 Public Declare Function Daruma_FI2000_CancelamentoCNFV Lib "Daruma32.dll" (ByVal COO_CNFV As String) As Integer
 
 Public Declare Function Daruma_FI2000_SegundaViaCNFVinculado Lib "Daruma32.dll" () As Integer
 
  
 
 'Metodos para cheques
 
 Public Declare Function Daruma_FI2000_StatusCheque Lib "Daruma32.dll" (ByVal StatusCheque As String) As Integer
 
 Public Declare Function Daruma_FI2000_ImprimirCheque Lib "Daruma32.dll" (ByVal Banco As String, ByVal Cidade As String, ByVal Data As String, ByVal Favorecido As String, ByVal Valor As String, ByVal PosicaoCheque As String) As Integer
 
 Public Declare Function Daruma_FI2000_ImprimirVersoCheque Lib "Daruma32.dll" (ByVal VersoCheque As String) As Integer
 
 Public Declare Function Daruma_FI2000_LiberarCheque Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_FI2000_LeituraCodigoMICR Lib "Daruma32.dll" (ByVal CodigoMICR As String) As Integer
 
 Public Declare Function Daruma_FI2000_CancelarCheque Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_FI2000_LeituraTabelaCheque Lib "Daruma32.dll" (ByVal TabelaCheque As String) As Integer
 
 Public Declare Function Daruma_FI2000_CarregarCheque Lib "Daruma32.dll" (ByVal NumeroBanco As String) As Integer
 
 Public Declare Function Daruma_FI2000_CorrigirGeometriaCheque Lib "Daruma32.dll" (ByVal NumeroBanco As String, ByVal GeometriaCheque As String) As Integer
 
  
 
 '*********************************************************************************************************************'
 
 ' '
 
 '                   TA1000 '
 
 ' '
 
 '*********************************************************************************************************************'
 
  
 
 'Metodos para Registry
 
 Public Declare Function Daruma_Registry_TA1000_Porta Lib "Daruma32.dll" (ByVal Porta As String) As Integer
 
 Public Declare Function Daruma_Registry_TA1000_PathProdutos Lib "Daruma32.dll" (ByVal PathProdutos As String) As Integer
 
 Public Declare Function Daruma_Registry_TA1000_PathUsuarios Lib "Daruma32.dll" (ByVal PathUsuarios As String) As Integer
 
 Public Declare Function Daruma_Registry_TA1000_NumeroItensEnviados Lib "Daruma32.dll" (ByVal NumeroItensEnviados As String) As Integer
 
 Public Declare Function Daruma_Registry_TA1000_PathRelatorios Lib "Daruma32.dll" (ByVal PathRelatorios As String) As Integer
 
  
 
 'Metodos para Produtos
 
 Public Declare Function Daruma_TA1000_CadastrarProdutos Lib "Daruma32.dll" (ByVal Descricao As String, ByVal Codigo As String, ByVal DecimaisPreco As String, ByVal DecimaisQuantidade As String, ByVal Preco As String, ByVal DescontoAcrescimo As String, ByVal ValorDescontoAcrescimo As String, ByVal UnidadeMedida As String, ByVal Aliquota As String, ByVal ProximoProduto As String, ByVal ProdutoAnterior As String, ByVal Estoque As String) As Integer
 
 Public Declare Function Daruma_TA1000_ConsultarProdutos Lib "Daruma32.dll" (ByVal Descricao As String, ByVal Codigo As String, ByVal DecimaisPreco As String, ByVal DecimaisQuantidade As String, ByVal Preco As String, ByVal DescontoAcrescimo As String, ByVal ValorDescontoAcrescimo As String, ByVal UnidadeMedida As String, ByVal Aliquota As String, ByVal ProximoProduto As String, ByVal ProdutoAnterior As String, ByVal Estoque As String) As Integer
 
 Public Declare Function Daruma_TA1000_AlterarProdutos Lib "Daruma32.dll" (ByVal Codigo_Consultar As String, ByVal Descricao As String, ByVal Codigo As String, ByVal DecimaisPreco As String, ByVal DecimaisQuantidade As String, ByVal Preco As String, ByVal DescontoAcrescimo As String, ByVal ValorDescontoAcrescimo As String, ByVal UnidadeMedida As String, ByVal Aliquota As String, ByVal ProximoProduto As String, ByVal ProdutoAnterior As String, ByVal Estoque As String) As Integer
 
 Public Declare Function Daruma_TA1000_EliminarProdutos Lib "Daruma32.dll" (ByVal Codigo As String) As Integer
 
 Public Declare Function Daruma_TA1000_EnviarBancoProdutos Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_TA1000_ReceberBancoProdutos Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_TA1000_ReceberProdutosVendidos Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_TA1000_ZerarProdutos Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_TA1000_ZerarProdutosVendidos Lib "Daruma32.dll" () As Integer
 
  
 
  
 
 'Metodos para Usuarios
 
 Public Declare Function Daruma_TA1000_EnviarBancoUsuarios Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_TA1000_ReceberBancoUsuarios Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_TA1000_ZerarUsuarios Lib "Daruma32.dll" () As Integer
 
 Public Declare Function Daruma_TA1000_CadastrarUsuarios Lib "Daruma32.dll" (ByVal Nome As String, ByVal CPF As String, ByVal CodigoConvenio As String, ByVal CodigoUsuario As String, ByVal UsuarioAnterior As String, ByVal ProximoUsuario As String) As Integer
 
 Public Declare Function Daruma_TA1000_ConsultarUsuarios Lib "Daruma32.dll" (ByVal Codigo_Consultar As String, ByVal Nome As String, ByVal CPF As String, ByVal CodigoConvenio As String, ByVal CodigoUsuario As String, ByVal UsuarioAnterior As String, ByVal ProximoUsuario As String) As Integer
 
 Public Declare Function Daruma_TA1000_AlterarUsuarios Lib "Daruma32.dll" (ByVal Codigo_Consultar As String, ByVal Nome As String, ByVal CPF As String, ByVal CodigoConvenio As String, ByVal CodigoUsuario As String, ByVal UsuarioAnterior As String, ByVal ProximoUsuario As String) As Integer
 
 Public Declare Function Daruma_TA1000_EliminarUsuarios Lib "Daruma32.dll" (ByVal Codigo As String) As Integer
 
 Public Declare Function Daruma_TA1000_LeStatusTransferencia Lib "Daruma32.dll" () As Integer
 
  
 
  
 
 '*********************************************************************************************************************'
 
 ' '
 
 '                   DUAL '
 
 ' '
 
 '*********************************************************************************************************************'
 
  
 
 'Metodos para Registry
 
 Public Declare Function Daruma_Registry_DUAL_Enter Lib "Daruma32.dll" (ByVal Enter As String) As Integer
 
 Public Declare Function Daruma_Registry_DUAL_Espera Lib "Daruma32.dll" (ByVal Espera As String) As Integer
 
 Public Declare Function Daruma_Registry_DUAL_ModoEscrita Lib "Daruma32.dll" (ByVal ModoEscrita As String) As Integer
 
 Public Declare Function Daruma_Registry_DUAL_Porta Lib "Daruma32.dll" (ByVal Porta As String) As Integer
 
 Public Declare Function Daruma_Registry_DUAL_Tabulacao Lib "Daruma32.dll" (ByVal Tabulacao As String) As Integer
 
  
 
 'Metodos para Status
 
 Public Declare Function Daruma_DUAL_VerificaStatus Lib "Daruma32.dll" () As Integer 'Verificar Status
 
 Public Declare Function Daruma_DUAL_VerificaDocumento Lib "Daruma32.dll" () As Integer 'Verifica Documento
 
  
 
 'Metodos para Autenticacao a Impressao
 
 Public Declare Function Daruma_DUAL_ImprimirTexto Lib "Daruma32.dll" (ByVal TextoLivre As String, ByVal TamanhoTexto As Integer) As Integer 'Imprimir Texto Livre
 
 Public Declare Function Daruma_DUAL_Autenticar Lib "Daruma32.dll" (ByVal NumVias As String, ByVal Texto As String, ByVal TempoAguardar As String) As Integer 'Autenticar
 
 Public Declare Function Daruma_DUAL_AcionaGaveta Lib "Daruma32.dll" () As Integer 'AcionaGaveta
 
  
 
  
 
 'Declaracoes globais
 
 Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
  
 
  
 
 ' Funcoes globais
 
 Public Function Daruma_MostrarRetorno()
 
 Int_Ack = 0
 
 Int_St1 = 0
 
 Int_St2 = 0
 
 Daruma_FI_RetornoImpressora Int_Ack, Int_St1, Int_St2
 
  
 
 Str_ErroExtendido = Space(4)
 
 Daruma_FI_RetornaErroExtendido Str_ErroExtendido
 
  
 
 MsgBox "Retorno do Metodo = " + CStr(Int_Retorno) + Chr(13) + Chr(10) _
 + "Ack = " + CStr(Int_Ack) + Chr(13) + Chr(10) _
 + "St1 = " + CStr(Int_St1) + Chr(13) + Chr(10) _
 + "St2 = " + CStr(Int_St2) + Chr(13) + Chr(10) _
 + "Erro Extendido = " + Str_ErroExtendido, , "Daruma Framework Retorno do Metodo"
 
  
 
 End Function
 
  
 
 'Retornos da DUAL
 
 Public Function Retorno_DUAL()
 
  
 
 If Int_Retorno = 1 Then
 
 MsgBox "1(um) - Impressora OK!", vbInformation, "Daruma Framework"
 
 End If
 
 If Int_Retorno = -50 Then
 
 MsgBox "-50 - Impressora OFF-LINE!", vbCritical, "Daruma Framework"
 
 End If
 
 If Int_Retorno = -51 Then
 
 MsgBox "-51 - Impressora Sem Papel!", vbCritical, "Daruma Framework"
 
 End If
 
 If Int_Retorno = -27 Then
 
 MsgBox "-27 - Erro Generico!", vbCritical, "Daruma Framework"
 
 End If
 
 If Int_Retorno = 0 Then
 
 MsgBox "0 - Impressora Desligada!", vbCritical, "Daruma Framework"
 
 End If
 
 End Function
 
  
 
 'Retornos do TA1000
 
 Public Function Retorno_TA1000()
 
  
 
 If Int_Retorno = 1 Then
 
 MsgBox "1(um) - Metodo Executado com Sucesso!", vbInformation, "Daruma Framework"
 
 Else
 
 MsgBox CStr(Int_Retorno) + " Erro Generico!", vbCritical, "Daruma Framework"
 
 End If
 
 End Function
 


