Attribute VB_Name = "BemaFi32"
Public Declare Function Bematech_FI_NumeroSerie Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BEMAFI32.DLL" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BEMAFI32.DLL" (ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_ResetaImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraX Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraXSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AcionaGuilhotinaMFD Lib "BEMAFI32.DLL" (ByVal iQtd As Integer) As Integer
Public Declare Function Bematech_FI_AbreCupom Lib "BEMAFI32.DLL" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_VendeItem Lib "BEMAFI32.DLL" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaCupomResumido Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_ReducaoZ Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_FechaCupom Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BEMAFI32.DLL" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AumentaDescricaoItem Lib "BEMAFI32.DLL" (ByVal Descricao As String) As Integer
Public Declare Function Bematech_FI_UsaUnidadeMedida Lib "BEMAFI32.DLL" (ByVal UnidadeMedida As String) As Integer
Public Declare Function Bematech_FI_AlteraSimboloMoeda Lib "BEMAFI32.DLL" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_ProgramaAliquota Lib "BEMAFI32.DLL" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaHorarioVerao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_NomeiaDepartamento Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Departamento As String) As Integer
Public Declare Function Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms Lib "BEMAFI32.DLL" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
Public Declare Function Bematech_FI_ProgramaArredondamento Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaTruncamento Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LinhasEntreCupons Lib "BEMAFI32.DLL" (ByVal Linhas As Integer) As Integer
Public Declare Function Bematech_FI_EspacoEntreLinhas Lib "BEMAFI32.DLL" (ByVal Dots As Integer) As Integer
Public Declare Function Bematech_FI_RelatorioGerencial Lib "BEMAFI32.DLL" (ByVal cTexto As String) As Integer
Public Declare Function Bematech_FI_FechaRelatorioGerencial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RecebimentoNaoFiscal Lib "BEMAFI32.DLL" (ByVal IndiceTotalizador As String, ByVal valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_AbreComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal valor As String, ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_UsaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" (ByVal Texto As String) As Integer
Public Declare Function Bematech_FI_FechaComprovanteNaoFiscalVinculado Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_Sangria Lib "BEMAFI32.DLL" (ByVal valor As String) As Integer
Public Declare Function Bematech_FI_Suprimento Lib "BEMAFI32.DLL" (ByVal valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalReducao Lib "BEMAFI32.DLL" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialData Lib "BEMAFI32.DLL" (ByVal cDataInicial As String, ByVal cDataFinal As String) As Integer
Public Declare Function Bematech_FI_LeituraMemoriaFiscalSerialReducao Lib "BEMAFI32.DLL" (ByVal cReducaoInicial As String, ByVal cReducaoFinal As String) As Integer
Public Declare Function Bematech_FI_VersaoFirmware Lib "BEMAFI32.DLL" (ByVal VersaoFirmware As String) As Integer
Public Declare Function Bematech_FI_CGC_IE Lib "BEMAFI32.DLL" (ByVal CGC As String, ByVal IE As String) As Integer
Public Declare Function Bematech_FI_GrandeTotal Lib "BEMAFI32.DLL" (ByVal GrandeTotal As String) As Integer
Public Declare Function Bematech_FI_Cancelamentos Lib "BEMAFI32.DLL" (ByVal ValorCancelamentos As String) As Integer
Public Declare Function Bematech_FI_Descontos Lib "BEMAFI32.DLL" (ByVal ValorDescontos As String) As Integer
Public Declare Function Bematech_FI_NumeroOperacoesNaoFiscais Lib "BEMAFI32.DLL" (ByVal NumeroOperacoes As String) As Integer
Public Declare Function Bematech_FI_NumeroCuponsCancelados Lib "BEMAFI32.DLL" (ByVal NumeroCancelamentos As String) As Integer
Public Declare Function Bematech_FI_NumeroIntervencoes Lib "BEMAFI32.DLL" (ByVal NumeroIntervencoes As String) As Integer
Public Declare Function Bematech_FI_NumeroReducoes Lib "BEMAFI32.DLL" (ByVal NumeroReducoes As String) As Integer
Public Declare Function Bematech_FI_NumeroSubstituicoesProprietario Lib "BEMAFI32.DLL" (ByVal NumeroSubstituicoes As String) As Integer
Public Declare Function Bematech_FI_UltimoItemVendido Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_ClicheProprietario Lib "BEMAFI32.DLL" (ByVal Cliche As String) As Integer
Public Declare Function Bematech_FI_NumeroCaixa Lib "BEMAFI32.DLL" (ByVal NumeroCaixa As String) As Integer
Public Declare Function Bematech_FI_NumeroLoja Lib "BEMAFI32.DLL" (ByVal NumeroLoja As String) As Integer
Public Declare Function Bematech_FI_SimboloMoeda Lib "BEMAFI32.DLL" (ByVal SimboloMoeda As String) As Integer
Public Declare Function Bematech_FI_MinutosLigada Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_MinutosImprimindo Lib "BEMAFI32.DLL" (ByVal Minutos As String) As Integer
Public Declare Function Bematech_FI_VerificaModoOperacao Lib "BEMAFI32.DLL" (ByVal Modo As String) As Integer
Public Declare Function Bematech_FI_VerificaEpromConectada Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_FlagsFiscais Lib "BEMAFI32.DLL" (ByRef Flag As Integer) As Integer
Public Declare Function Bematech_FI_ValorPagoUltimoCupom Lib "BEMAFI32.DLL" (ByVal ValorCupom As String) As Integer
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_ContadoresTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Contadores As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_DataHoraReducao Lib "BEMAFI32.DLL" (ByVal Data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_DataMovimento Lib "BEMAFI32.DLL" (ByVal Data As String) As Integer
Public Declare Function Bematech_FI_VerificaTruncamento Lib "BEMAFI32.DLL" (ByVal Flag As String) As Integer
Public Declare Function Bematech_FI_Acrescimos Lib "BEMAFI32.DLL" (ByVal ValorAcrescimos As String) As Integer
Public Declare Function Bematech_FI_ContadorBilhetePassagem Lib "BEMAFI32.DLL" (ByVal ContadorPassagem As String) As Integer
Public Declare Function Bematech_FI_VerificaAliquotasIss Lib "BEMAFI32.DLL" (ByVal AliquotasIss As String) As Integer
Public Declare Function Bematech_FI_VerificaFormasPagamento Lib "BEMAFI32.DLL" (ByVal Formas As String) As Integer
Public Declare Function Bematech_FI_VerificaRecebimentoNaoFiscal Lib "BEMAFI32.DLL" (ByVal Recebimentos As String) As Integer
Public Declare Function Bematech_FI_VerificaDepartamentos Lib "BEMAFI32.DLL" (ByVal Departamentos As String) As Integer
Public Declare Function Bematech_FI_VerificaTipoImpressora Lib "BEMAFI32.DLL" (ByRef TipoImpressora As Integer) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresParciais Lib "BEMAFI32.DLL" (ByVal cTotalizadores As String) As Integer
Public Declare Function Bematech_FI_RetornoAliquotas Lib "BEMAFI32.DLL" (ByVal cAliquotas As String) As Integer
Public Declare Function Bematech_FI_VerificaEstadoImpressora Lib "BEMAFI32.DLL" (ByRef iAck As Integer, ByRef ist1 As Integer, ByRef ist2 As Integer) As Integer
Public Declare Function Bematech_FI_DadosUltimaReducao Lib "BEMAFI32.DLL" (ByVal DadosReducao As String) As Integer
Public Declare Function Bematech_FI_MonitoramentoPapel Lib "BEMAFI32.DLL" (ByRef Linhas As Integer) As Integer
Public Declare Function Bematech_FI_Autenticacao Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ProgramaCaracterAutenticacao Lib "BEMAFI32.DLL" (ByVal Parametros As String) As Integer
Public Declare Function Bematech_FI_AcionaGaveta Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaEstadoGaveta Lib "BEMAFI32.DLL" (ByRef EstadoGaveta As Integer) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaSingular Lib "BEMAFI32.DLL" (ByVal MoedaSingular As String) As Integer
Public Declare Function Bematech_FI_ProgramaMoedaPlural Lib "BEMAFI32.DLL" (ByVal MoedaPlural As String) As Integer
Public Declare Function Bematech_FI_CancelaImpressaoCheque Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaStatusCheque Lib "BEMAFI32.DLL" (ByRef StatusCheque As Integer) As Integer
Public Declare Function Bematech_FI_ImprimeCheque Lib "BEMAFI32.DLL" (ByVal Banco As String, ByVal valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_ImprimeCopiaCheque Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_IncluiCidadeFavorecido Lib "BEMAFI32.DLL" (ByVal Cidade As String, ByVal Favorecido As String) As Integer
Public Declare Function Bematech_FI_EstornoFormasPagamento Lib "BEMAFI32.DLL" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal valor As String) As Integer
Public Declare Function Bematech_FI_ForcaImpactoAgulhas Lib "BEMAFI32.DLL" (ByVal ForcaImpacto As Integer) As Integer
Public Declare Function Bematech_FI_RetornoImpressora Lib "BEMAFI32.DLL" (ByRef iAck As Integer, ByRef ist1 As Integer, ByRef ist2 As Integer) As Integer
Public Declare Function Bematech_FI_FechaPortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_VerificaImpressoraLigada Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_IniciaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Public Declare Function Bematech_FI_EfetuaFormaPagamentoDescricaoForma Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoOpcional As String) As Integer
Public Declare Function Bematech_FI_TerminaFechamentoCupom Lib "BEMAFI32.DLL" (ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_AbreBilhetePassagem Lib "BEMAFI32.DLL" (ByVal ImprimeValorFinal As String, ByVal ImprimeEnfatizado As String, ByVal LocalEmbarque As String, ByVal Destino As String, ByVal Linha As String, ByVal Prefixo As String, ByVal Agente As String, ByVal Agencia As String, ByVal Data As String, ByVal Hora As String, ByVal Poltrona As String, ByVal Plataforma As String) As Integer
Public Declare Function Bematech_FI_MapaResumo Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Analitico Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_RelatorioTipo60Mestre Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeConfiguracoesImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ImprimeDepartamentos Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AberturaDoDia Lib "BEMAFI32.DLL" (ByVal valor As String, ByVal FormaPagamento As String) As Integer
Public Declare Function Bematech_FI_FechamentoDoDia Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_ValorFormaPagamento Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal ValorForma As String) As Integer
Public Declare Function Bematech_FI_ValorTotalizadorNaoFiscal Lib "BEMAFI32.DLL" (ByVal Totalizador As String, ByVal ValorTotalizador As String) As Integer
Public Declare Function Bematech_FI_VerificaIndiceAliquotasIss Lib "BEMAFI32.DLL" (ByVal IndiceAliquotasIss As String) As Integer

'Funções para Impressora restaurante
Public Declare Function Bematech_FIR_RegistraVenda Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_CancelaVenda Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_ConferenciaMesa Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_AbreConferenciaMesa Lib "BEMAFI32.DLL" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_FechaConferenciaMesa Lib "BEMAFI32.DLL" (ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaMesa Lib "BEMAFI32.DLL" (ByVal MesaOrigem As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_AbreCupomRestaurante Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_ContaDividida Lib "BEMAFI32.DLL" (ByVal NumeroCupons As String, ByVal ValorPago As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomContaDividida Lib "BEMAFI32.DLL" (ByVal NumeroCupons As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal FormasPagamento As String, ByVal ValorFormasPagamento As String, ByVal ValorPagoCliente As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaItem Lib "BEMAFI32.DLL" (ByVal MesaOrigem As String, ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertas Lib "BEMAFI32.DLL" (ByVal TipoRelatorio As Integer) As Integer
Public Declare Function Bematech_FIR_ImprimeCardapio Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_RelatorioMesasAbertasSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_CardapioPelaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FIR_RegistroVendaSerial Lib "BEMAFI32.DLL" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_VerificaMemoriaLivre Lib "BEMAFI32.DLL" (ByVal Bytes As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomRestaurante Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomResumidoRestaurante Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal mensagem As String) As Integer

' Funções para o TEF

Public Declare Function Bematech_FITEF_Status Lib "BEMAFI32.DLL" (ByVal Identificacao As String) As Integer
Public Declare Function Bematech_FITEF_VendaCartao Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_ConfirmaVenda Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal Header As String) As Integer
Public Declare Function Bematech_FITEF_NaoConfirmaVendaImpressao Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_CancelaVendaCartao Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal Nsu As String, ByVal NumeroCupom As String, ByVal Hora As String, ByVal Data As String, ByVal Rede As String) As Integer
Public Declare Function Bematech_FITEF_ImprimeTEF Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal FormaPagamento As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_ImprimeRelatorio Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FITEF_ADM Lib "BEMAFI32.DLL" (ByVal Identificacao As String) As Integer
Public Declare Function Bematech_FITEF_VendaCompleta Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal FormaPagamento As String, ByVal Texto As String) As Integer
Public Declare Function Bematech_FITEF_ConfiguraDiretorioTef Lib "BEMAFI32.DLL" (ByVal PathArqReq As String, ByVal PathArqResp As String) As Integer

Public Type EstadoImp
    Retorno As Integer
    msg As String
End Type
Function F_RetornoVariaveis() As Integer
    Dim i As Integer
    i = 0
    Dim iRet As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_FlagsFiscais(iRet)
    Else
    
    
        i = Bematech_FI_FlagsFiscais(iRet)
    End If
    F_RetornoVariaveis = iRet
End Function
Function F_EstadoImpressora() As EstadoImp
    Dim i As Integer
    Dim iAck As Integer, ist1 As Integer, ist2 As Integer
    i = Bematech_FI_VerificaImpressoraLigada
    If i <> 1 Then
        F_EstadoImpressora.msg = "DESLIGADA"
        F_EstadoImpressora.Retorno = i
        Exit Function
    End If
    i = Bematech_FI_VerificaEstadoImpressora(iAck, ist1, ist2)
    If iAck = 0 Then
        F_EstadoImpressora.msg = "OK"
        F_EstadoImpressora.Retorno = -1
        Exit Function
    End If
    F_EstadoImpressora.Retorno = ist1
    If ist1 = 128 Then
        F_EstadoImpressora.msg = "fim de papel"
        
    ElseIf ist1 = 64 Then
        F_EstadoImpressora.msg = "pouco papel"
    ElseIf ist1 = 32 Then
        F_EstadoImpressora.msg = "erro no relógio"
    ElseIf ist1 = 16 Then
        F_EstadoImpressora.msg = "impressora em erro"
    ElseIf ist1 = 8 Then
        F_EstadoImpressora.msg = "primeiro dado deCMD não foi ESC(1BH)"
    ElseIf ist1 = 4 Then
        F_EstadoImpressora.msg = "comando inexistente"
    ElseIf ist1 = 2 Then
        F_EstadoImpressora.msg = "cupom aberto"
    ElseIf ist1 = 1 Then
        F_EstadoImpressora.msg = "número de parâmetros CMD inválido"
    End If
End Function
