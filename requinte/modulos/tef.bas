Attribute VB_Name = "tef"
Public Declare Function BlockInput Lib "User32" (ByVal fBlock As Long) As Long

'constantes de diretório
Public Const DiretorioTEFreq = "C:\tef_dial\req\"
Public Const DiretorioTEFresp = "C:\tef_dial\resp\"
'Public Const DiretorioTEFreq = "C:\hipertef\req\"
'Public Const DiretorioTEFresp = "C:\hipertef\resp\"
Public strREDE As String
Public strAUTORIZACAO As String

Public Buffer As String ' Buffer Generico

Public ValorTef As String  'Valor do Cartao a ser passado para o TEF
Public NumeroCupomTef As String  'Numero do Cupom Fiscal para a Compra do TEF
'constantes de diretório
'Declaracaoes da Dll de comunicacao Bematech   - BemaFi32.dll
Public Declare Function Bematech_FI_NumeroSerie Lib "BEMAFI32.DLL" (ByVal NumeroSerie As String) As Integer
Public Declare Function Bematech_FI_SubTotal Lib "BEMAFI32.DLL" (ByVal SubTotal As String) As Integer
Public Declare Function Bematech_FI_NumeroCupom Lib "BEMAFI32.DLL" (ByVal NumeroCupom As String) As Integer
Public Declare Function Bematech_FI_ResetaImpressora Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbrePortaSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraX Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_LeituraXSerial Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_AbreCupom Lib "BEMAFI32.DLL" (ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FI_VendeItem Lib "BEMAFI32.DLL" (ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal desconto As String) As Integer
Public Declare Function Bematech_FI_CancelaItemAnterior Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_CancelaItemGenerico Lib "BEMAFI32.DLL" (ByVal NumeroItem As String) As Integer
Public Declare Function Bematech_FI_CancelaCupom Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FI_FechaCupomResumido Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_ReducaoZ Lib "BEMAFI32.DLL" (ByVal data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_FechaCupom Lib "BEMAFI32.DLL" (ByVal FormaPagamento As String, ByVal DiscontoAcrecimo As String, ByVal TipoDescontoAcrecimo As String, ByVal ValorAcrecimoDesconto As String, ByVal ValorPago As String, ByVal mensagem As String) As Integer
Public Declare Function Bematech_FI_VendeItemDepartamento Lib "BEMAFI32.DLL" (ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
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
Public Declare Function Bematech_FI_DataHoraImpressora Lib "BEMAFI32.DLL" (ByVal data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_ContadoresTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Contadores As String) As Integer
Public Declare Function Bematech_FI_VerificaTotalizadoresNaoFiscais Lib "BEMAFI32.DLL" (ByVal Totalizadores As String) As Integer
Public Declare Function Bematech_FI_DataHoraReducao Lib "BEMAFI32.DLL" (ByVal data As String, ByVal Hora As String) As Integer
Public Declare Function Bematech_FI_DataMovimento Lib "BEMAFI32.DLL" (ByVal data As String) As Integer
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
Public Declare Function Bematech_FI_ImprimeCheque Lib "BEMAFI32.DLL" (ByVal Banco As String, ByVal valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal data As String, ByVal mensagem As String) As Integer
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
Public Declare Function Bematech_FI_AbreBilhetePassagem Lib "BEMAFI32.DLL" (ByVal ImprimeValorFinal As String, ByVal ImprimeEnfatizado As String, ByVal LocalEmbarque As String, ByVal Destino As String, ByVal linha As String, ByVal Prefixo As String, ByVal Agente As String, ByVal agencia As String, ByVal data As String, ByVal Hora As String, ByVal Poltrona As String, ByVal Plataforma As String) As Integer
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
Public Declare Function Bematech_FIR_RegistraVenda Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_CancelaVenda Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_ConferenciaMesa Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_AbreConferenciaMesa Lib "BEMAFI32.DLL" (ByVal Mesa As String) As Integer
Public Declare Function Bematech_FIR_FechaConferenciaMesa Lib "BEMAFI32.DLL" (ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaMesa Lib "BEMAFI32.DLL" (ByVal MesaOrigem As String, ByVal MesaDestino As String) As Integer
Public Declare Function Bematech_FIR_AbreCupomRestaurante Lib "BEMAFI32.DLL" (ByVal Mesa As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_ContaDividida Lib "BEMAFI32.DLL" (ByVal NumeroCupons As String, ByVal ValorPago As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_FechaCupomContaDividida Lib "BEMAFI32.DLL" (ByVal NumeroCupons As String, ByVal FlagAcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal FormasPagamento As String, ByVal ValorFormasPagamento As String, ByVal ValorPagoCliente As String, ByVal CGC_CPF As String) As Integer
Public Declare Function Bematech_FIR_TransferenciaItem Lib "BEMAFI32.DLL" (ByVal MesaOrigem As String, ByVal codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal Quantidade As String, ByVal ValorUnitario As String, ByVal FlagAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal MesaDestino As String) As Integer
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
Public Declare Function Bematech_FITEF_CancelaVendaCartao Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal Nsu As String, ByVal NumeroCupom As String, ByVal Hora As String, ByVal data As String, ByVal Rede As String) As Integer
Public Declare Function Bematech_FITEF_ImprimeTEF Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal FormaPagamento As String, ByVal ValorCompra As String) As Integer
Public Declare Function Bematech_FITEF_ImprimeRelatorio Lib "BEMAFI32.DLL" () As Integer
Public Declare Function Bematech_FITEF_ADM Lib "BEMAFI32.DLL" (ByVal Identificacao As String) As Integer
Public Declare Function Bematech_FITEF_VendaCompleta Lib "BEMAFI32.DLL" (ByVal Identificacao As String, ByVal ValorCompra As String, ByVal FormaPagamento As String, ByVal Texto As String) As Integer
Public Declare Function Bematech_FITEF_ConfiguraDiretorioTef Lib "BEMAFI32.DLL" (ByVal PathArqReq As String, ByVal PathArqResp As String) As Integer

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpAplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'-------------------------------------
'Funcao para Criar arquivo para o GP e Copiar para o GP
Public Function CriaArquivo(sBufferArquivo As String, sPath As String)
On Error GoTo trata_Erro
'Cria temp para TEF
Dim arq As String

arq = "c:\intpos.tmp"


FechaArquivoExterno 1
    


Open arq For Binary As #1
    Put #1, , sBufferArquivo
Close #1
'Copia para o Diretorio do GP
'Open sPath & "INTPOS.001" For Binary As #1
'    Put #1, , sBufferArquivo
'Close #1

'On Error Resume Next

FileCopy "C:\INTPOS.TMP", sPath & "INTPOS.001"
Exit Function
trata_Erro:
    MsgBox Err.Description, vbCritical
End Function

Public Function EsperaArquivo(sNome As String, sTime As String, Minutos As Integer) As Integer
Dim TimeIni, TimeFinal As Integer

'Criei o time para Minutos e segundos
If (Minutos = 1) Then
    TimeIni = Minute(Time)
    TimeFinal = TimeIni + sTime
Else
    Sleep (sTime * 1000)
End If
    
DoEvents
     Do While ((Dir(DiretorioTEFresp & sNome)) = "")
        If Minutos = 1 Then
            TimeIni = Minute(Time)
            If (TimeIni >= TimeFinal) Then: EsperaArquivo = -1: Exit Function 'Retorna Time Out
        Else
            EsperaArquivo = -1: Exit Function 'Retorna Time Out
        End If
        
    Loop
DoEvents
    EsperaArquivo = 1
End Function
Function ContadorTEF()
    ContadorTEF = ChaveSQL("tab_tef", "contador_tef", "0######")
End Function
Public Function PreparaVenda(sValor As String, sNumeroCupom As String)
    Dim s As String
    Buffer = "000-000 = CRT" + Chr(13) + Chr(10) + _
             "001-000 = " + ContadorTEF + Chr(13) + Chr(10) + _
             "002-000 = " + F_SoNumero(sNumeroCupom) + Chr(13) + Chr(10) + _
             "003-000 = " + FormatValTef(Trim(sValor)) + Chr(13) + Chr(10) + _
             "004-000 = 0" + Chr(13) + Chr(10) + _
             "999-999 = 0"
             
             CriaArquivo Buffer, DiretorioTEFreq

End Function
Public Function FinalizaTef(valor As String) As Boolean
    Dim str As String
    Dim iCount As Integer
    Dim Retorno As Integer
    Dim sAux As String
    Dim bLigou As Boolean
    Dim Estado As Integer
    Dim sMsg As String
    Dim s009 As String
    Dim iAck As Integer, ist1 As Integer, ist2 As Integer
    Dim sRede As String
    Dim i As Integer
    'Efetua a operacao de Leitura da Resposta e Tratamento das strings do Arquivo
    str = frmTroco.caption
    bLigou = False 'se a impressora foi desligada, e depois ligada
    Open DiretorioTEFresp & "INTPOS.001" For Input As #1
        Buffer = ""
        Do While Not EOF(1)
    
            Line Input #1, sBufArq
            If (Mid(sBufArq, 1, 3) = "030") Then
                DoEvents
                    
                 'frmTroco.lblAviso.caption = Mid(sBufArq, 11, Len(sBufArq) - 10)
                    'frmVenda.pcMsg.Visible = True
                    'P_CentraObejct frmTroco.lblAviso, frmTroco
                DoEvents
                Exit Do
            End If
            If (Mid(sBufArq, 1, 3) = "009") Then
                s009 = s009 + Trim(Right(sBufArq, 2)) + Chr(13) + Chr(10)
            End If
            If (Mid(sBufArq, 1, 3) = "028") Then
                sMsg = sMsg + Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)
            End If
            If (Mid(sBufArq, 1, 3) = "029") Then
                Buffer = Buffer + Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)  'Nao pego os campos e nao pego as aspas
            End If
            If (Mid(sBufArq, 1, 3) = "010") Then
                strREDE = Mid(sBufArq, 11, Len(sBufArq) - 10) 'Nao pego os campos e nao pego as aspas
            End If
            If (Mid(sBufArq, 1, 3) = "013") Then
                strAUTORIZACAO = Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)  'Nao pego os campos e nao pego as aspas
            End If
        Loop
    Close #1

    'Copia de Seguranca do TEF
    Open "C:\TEFTEMP1.TXT" For Binary As #1
    Put #1, , Buffer
    Close #1

'Abrir o Arquivo de BackUp e vamos jogar linha a linha na Impressora
'aqui devo buscar se a transação foi ou não efetuada
    
    
    
    
    
    Retorno = 0
    Do While Retorno <> 1
        Trava_Teclado
        Retorno = ecf_VerificaImpressoraLigada
        If Retorno <> 1 Then
            Habilita_Teclado
            If MsgBox("Impressora não responde!, tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                'aqui
                NaoConfirmaVenda
                FinalizaTef = False
                ecf_CancelaCupom
                Exit Function
            Else
                Trava_Teclado
            End If
        End If
    Loop
    
    ecf_abreComprovanteNaoFiscalVinculado "Cartao", FormatVal(valor), strNum_Cpm
    
    For iCount = 1 To 2
        If iCount = 2 Then
            For i = 1 To 3
                ecf_UsaComprovanteNaoFiscalVinculado "-> " & Chr(13) + Chr(10)
            Next
            i = Bematech_FI_AcionaGuilhotinaMFD(10)
            For i = 1 To 30
                ecf_UsaComprovanteNaoFiscalVinculado "-> "
            Next
        End If
        Open "C:\TEFTEMP1.TXT" For Input As #1
        
        Do While Not EOF(1)
        
    
    
            Line Input #1, sBufArq
            Retorno = 0
            Do While Retorno <> 1 Or ist1 >= 128
                
                
                
                Retorno = ecf_RetornoImpressora(iAck, ist1, ist2)
                
                
                If Retorno <> 1 Or ist1 >= 128 Then
                    '128 é falta de papel
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        Close #1
                        NaoConfirmaVenda
                        FinalizaTef = False
                        ecf_CancelaCupom
                        Exit Function
                    Else
                        Trava_Teclado
                        bLigou = True
                    End If
                
                End If
                
            Loop
            If bLigou Then 'vai entrar no caso de não ter respondido uma vez e ter optado por sim e ter respondido...
                
                    ecf_fechaComprovanteNaoFiscalVinculado
                    Close #1
                    ImprimeGerencial ' leitura X + duas vias do Comprovante TEF
                    ConfirmaVenda
                    FinalizaTef = True
                    Exit Function
            Else
                If sBufArq = "" Then
                    sBufArq = "    "
                End If
                ecf_UsaComprovanteNaoFiscalVinculado (sBufArq)
                
            End If
        
        
        Loop
        Close #1
        'If iCount = 1 Then 'para destacar o papel
        '    Sleep 5000
        'End If
        
    Next
    ecf_fechaComprovanteNaoFiscalVinculado

    'Termina OK
    Habilita_Teclado
    ConfirmaVenda
    FinalizaTef = True
End Function
Public Function FinalizaTefNaoFiscal() As Boolean
    Dim str As String
    Dim iCount As Integer
    Dim Retorno As Integer
    Dim sAux As String
    Dim bLigou As Boolean
    Dim Estado As Integer
    Dim sMsg As String
    Dim s009 As String
    Dim iAck As Integer, ist1 As Integer, ist2 As Integer
    Dim sRede As String
    Dim i As Integer
    'Efetua a operacao de Leitura da Resposta e Tratamento das strings do Arquivo
    str = frmTroco.caption
    bLigou = False 'se a impressora foi desligada, e depois ligada
    Open DiretorioTEFresp & "INTPOS.001" For Input As #1
        Buffer = ""
        Do While Not EOF(1)
    
            Line Input #1, sBufArq
            If (Mid(sBufArq, 1, 3) = "030") Then
                DoEvents
                    
                 'frmTroco.lblAviso.caption = Mid(sBufArq, 11, Len(sBufArq) - 10)
                    'frmVenda.pcMsg.Visible = True
                    'P_CentraObejct frmTroco.lblAviso, frmTroco
                DoEvents
                Exit Do
            End If
            If (Mid(sBufArq, 1, 3) = "009") Then
                s009 = s009 + Trim(Right(sBufArq, 2)) + Chr(13) + Chr(10)
            End If
            If (Mid(sBufArq, 1, 3) = "028") Then
                
                sMsg = sMsg + Mid(sBufArq, 10) + Chr(13) + Chr(10)
            End If
            If (Mid(sBufArq, 1, 3) = "029") Then
                Buffer = Buffer + Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)  'Nao pego os campos e nao pego as aspas
            End If
            If (Mid(sBufArq, 1, 3) = "010") Then
                strREDE = Mid(sBufArq, 11, Len(sBufArq) - 10) 'Nao pego os campos e nao pego as aspas
            End If
            If (Mid(sBufArq, 1, 3) = "013") Then
                strAUTORIZACAO = Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)  'Nao pego os campos e nao pego as aspas
            End If
        Loop
    Close #1

    'Copia de Seguranca do TEF
    Open "C:\TEFTEMP1.TXT" For Binary As #1
    Put #1, , Buffer
    Close #1

    
    
    
    
    
    
    
    
    
    ConfirmaVenda
    ImprimeComprovanteNaoFiscalTEF Buffer, True
    
    
    
    
    
    Habilita_Teclado
    
    
    
    
    
    
    FinalizaTefNaoFiscal = True
End Function


Public Function FinalizaTef_Old(valor As String) As Boolean
    Dim str As String
    Dim iCount As Integer
    Dim Retorno As Integer
    Dim sAux As String
    Dim bLigou As Boolean
    Dim Estado As Integer
    Dim sMsg As String
    Dim s009 As String
    Dim iAck As Integer, ist1 As Integer, ist2 As Integer
    Dim sRede As String
    'Efetua a operacao de Leitura da Resposta e Tratamento das strings do Arquivo
    str = frmTroco.caption
    bLigou = False 'se a impressora foi desligada, e depois ligada
    Open DiretorioTEFresp & "INTPOS.001" For Input As #1
        Buffer = ""
        Do While Not EOF(1)
    
            Line Input #1, sBufArq
            If (Mid(sBufArq, 1, 3) = "030") Then
                DoEvents
                    
                 'frmTroco.lblAviso.caption = Mid(sBufArq, 11, Len(sBufArq) - 10)
                    'frmVenda.pcMsg.Visible = True
                    'P_CentraObejct frmTroco.lblAviso, frmTroco
                DoEvents
                Exit Do
            End If
            If (Mid(sBufArq, 1, 3) = "009") Then
                s009 = s009 + Trim(Right(sBufArq, 2)) + Chr(13) + Chr(10)
            End If
            If (Mid(sBufArq, 1, 3) = "028") Then
                sMsg = sMsg + Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)
            End If
            If (Mid(sBufArq, 1, 3) = "029") Then
                Buffer = Buffer + Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)  'Nao pego os campos e nao pego as aspas
            End If
            If (Mid(sBufArq, 1, 3) = "010") Then
                strREDE = Mid(sBufArq, 11, Len(sBufArq) - 10) 'Nao pego os campos e nao pego as aspas
            End If
            If (Mid(sBufArq, 1, 3) = "013") Then
                strAUTORIZACAO = Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)  'Nao pego os campos e nao pego as aspas
            End If
        Loop
    Close #1

    'Copia de Seguranca do TEF
    Open "C:\TEFTEMP1.TXT" For Binary As #1
    Put #1, , Buffer
    Close #1

'Abrir o Arquivo de BackUp e vamos jogar linha a linha na Impressora
'aqui devo buscar se a transação foi ou não efetuada
    
    
    
    
    
    Retorno = 0
    Do While Retorno <> 1
        Trava_Teclado
        Retorno = ecf_VerificaImpressoraLigada
        If Retorno <> 1 Then
            Habilita_Teclado
            If MsgBox("Impressora não responde!, tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                'aqui
                NaoConfirmaVenda
                FinalizaTef_Old = False
                ecf_CancelaCupom
                Exit Function
            Else
                Trava_Teclado
            End If
        End If
    Loop
    
    
    For iCount = 1 To 3
        ecf_fechaComprovanteNaoFiscalVinculado
        P_Status "Imprimindo " & iCount & "a Via do Cupom TEF"
        Bematech_FI_AbreComprovanteNaoFiscalVinculado "Cartao", FormatVal(valor), strNum_Cpm
        If iCount = 2 Then
            Bematech_FI_RelatorioGerencial " "
            Bematech_FI_RelatorioGerencial " "
            Bematech_FI_RelatorioGerencial " "
            Bematech_FI_RelatorioGerencial " "
        End If
        Open "C:\TEFTEMP1.TXT" For Input As #1
        
        Do While Not EOF(1)
        
    
    
            Line Input #1, sBufArq
            Retorno = 0
            Do While Retorno <> 1 Or ist1 >= 128
                
                
                
                Retorno = Bematech_FI_RetornoImpressora(iAck, ist1, ist2)
                
                
                If Retorno <> 1 Or ist1 >= 128 Then
                    '128 é falta de papel
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        Close #1
                        NaoConfirmaVenda
                        FinalizaTef_Old = False
                        Bematech_FI_CancelaCupom
                        Exit Function
                    Else
                        Trava_Teclado
                        bLigou = True
                    End If
                
                End If
                
            Loop
            If bLigou Then 'vai entrar no caso de não ter respondido uma vez e ter optado por sim e ter respondido...
                
                    Bematech_FI_FechaComprovanteNaoFiscalVinculado
                    Close #1
                    ImprimeGerencial ' leitura X + duas vias do Comprovante TEF
                    ConfirmaVenda
                    FinalizaTef_Old = True
                    Exit Function
            Else
                If sBufArq = "" Then
                    sBufArq = "    "
                End If
                Bematech_FI_UsaComprovanteNaoFiscalVinculado (sBufArq)
                
            End If
        
        
        Loop
        Close #1
        
        
        'If iCount = 1 Then 'para destacar o papel
        '    Sleep 5000
        'End If
        Bematech_FI_FechaComprovanteNaoFiscalVinculado
        'MsgBox "Pressione Ok para início da 2a Via"
    Next
    
    Bematech_FI_FechaComprovanteNaoFiscalVinculado

    'Termina OK
    Habilita_Teclado
    ConfirmaVenda
    FinalizaTef_Old = True
End Function
Sub VerificaIntPos()
    On Error GoTo trata_Erro
    
    If Dir(DiretorioTEFresp & "\intpos.001") <> "" Then
        NaoConfirmaVenda
    End If
    If Dir(DiretorioTEFresp & "\intpos.sts") <> "" Then
        Kill DiretorioTEFresp & "\intpos.sts"
    End If
    If Dir(DiretorioTEFresp & "\imprime.txt") <> "" Then
        Kill DiretorioTEFresp & "\imprime.txt"
    End If
    If Dir(DiretorioTEFresp & "\ativo.001") <> "" Then
        Kill DiretorioTEFresp & "\ativo.001"
    End If
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Public Sub ConfirmaVenda()
Dim Buffer2 As String
Dim Buffer As String
On Error Resume Next
Open DiretorioTEFresp & "INTPOS.001" For Input As #1
Buffer = ""
Buffer2 = ""
   Do While Not EOF(1)
      Line Input #1, sBufArq
      If Mid(sBufArq, 1, 3) = "001" Or _
         Mid(sBufArq, 1, 3) = "010" Or _
         Mid(sBufArq, 1, 3) = "012" Or _
         Mid(sBufArq, 1, 3) = "027" Or _
         Mid(sBufArq, 1, 3) = "999" Then
            Buffer = Buffer + sBufArq & Chr(13) & Chr(10)
      End If
   Loop
  
   Close #1
   
   Buffer2 = "000-000 = CNF" & Chr(13) & Chr(10) + Buffer
   
   CriaArquivo Buffer2, DiretorioTEFreq
   
   Sleep (2000)
   
    If ((Dir(DiretorioTEFresp & "INTPOS.001")) <> "") Then
        Kill DiretorioTEFresp & "INTPOS.001"
    End If
   On Error Resume Next
   Kill "C:\TEFTEMP1.TXT"
   Kill "C:\INTPOS.TMP"
   killIntPosResp
  'MsgBox "Autorização Efetivada! Venda Confirmada", vbInformation

End Sub

Public Sub NaoConfirmaVenda()
Dim Buffer As String
Dim Buffer2 As String
Dim sNsu As String
Dim sRede As String
Dim sValor As String
Dim sMsg As String
Dim sOp As String

Open DiretorioTEFresp & "INTPOS.001" For Input As #1
   
   Buffer = ""
   Buffer2 = ""
   Do While Not EOF(1)
      Line Input #1, sBufArq
    If Mid(sBufArq, 1, 3) = "012" Then
        sNsu = Mid(sBufArq, 11, Len(sBufArq) - 10)
    End If
    If Mid(sBufArq, 1, 3) = "010" Then
        sRede = Mid(sBufArq, 11, Len(sBufArq) - 10)
    End If
    If Mid(sBufArq, 1, 3) = "003" Then
        sValor = Mid(sBufArq, 11, Len(sBufArq) - 10)
    End If
    If Mid(sBufArq, 1, 3) = "030" Then
        sMsg = Mid(sBufArq, 11, Len(sBufArq) - 10)
    End If
    If Mid(sBufArq, 1, 3) = "009" Then
        sOp = Mid(sBufArq, 11, Len(sBufArq) - 10)
    End If
      If Mid(sBufArq, 1, 3) = "001" Or _
         Mid(sBufArq, 1, 3) = "010" Or _
         Mid(sBufArq, 1, 3) = "012" Or _
         Mid(sBufArq, 1, 3) = "027" Or _
         Mid(sBufArq, 1, 3) = "999" Then
            
            
            
            Buffer = Buffer + sBufArq & Chr(13) & Chr(10)
      End If
   Loop
  
   Close #1
   
   Buffer2 = "000-000 = NCN" & Chr(13) & Chr(10) + Buffer
   
   On Error Resume Next
   
    If Not IsLoad("frmVenda") Then
         P_Status "Transação cancelada " & "NSU: " & sNsu & " Rede: " & sRede & " Valor: " & FormatVal(CCur(sValor) / 100)
    Else
         'frmVenda.pcMsg = True
         'frmTroco.lblAviso.caption = "Transação cancelada " & "NSU: " & sNsu & " Rede: " & sRede & " Valor: " & FormatVal(CCur(sValor) / 100)
    End If
    
    'frmVenda.pcMsg.Visible = True
     'P_CentraObejct frmTroco.lblAviso, frmTroco
    CriaArquivo Buffer2, DiretorioTEFreq
    
    Sleep (2000)
    
    If ((Dir(DiretorioTEFresp & "INTPOS.001")) <> "") Then
      Kill DiretorioTEFresp & "INTPOS.001"
    End If
    On Error Resume Next
    Kill "C:\TEFTEMP1.TXT"
    Kill "C:\INTPOS.TMP"
    Bematech_FI_FechaRelatorioGerencial
    
    'sMsg = "Última transação TEF cancelada! " & Chr(13) & " Rede: " & sRede
    If sOp <> "0" Then
        If sMsg <> "TRANSACAO CONCLUIDA COM SUCESSO" Then
            sMsg = sMsg & Chr(13)
            frmVenda.SetFocus
            MsgBox sMsg, vbCritical
    
            P_Status ""
    
            Exit Sub
        Else
            sMsg = "Transação cancelada" & Chr(13)
            frmVenda.SetFocus
            MsgBox sMsg, vbCritical
    
            P_Status ""
            Exit Sub
        End If
    End If
    
    sMsg = "Transação Cancelada!"
    If sRede <> "" Then
         sMsg = sMsg & Chr(13) & " Rede: " & sRede
         
    End If
    
    If sNsu <> "" Then
         sMsg = sMsg & " NSU: " & sNsu
         
    End If
    
    If sValor <> "" Then
         sMsg = sMsg & " Valor: " & FormatVal(CCur(sValor) / 100)
    End If

    frmVenda.SetFocus
    MsgBox sMsg, vbCritical

    P_Status ""

End Sub
Public Function ImprimeGerencial()
    Dim iResp As Integer
    Dim iCount As Integer
    If Bematech_FI_FechaRelatorioGerencial = 0 Then
        Bematech_FI_FechaRelatorioGerencial
    End If
    For iCount = 1 To 2
        Trava_Teclado
        If iCount = 2 Then
            Bematech_FI_RelatorioGerencial " "
            Bematech_FI_RelatorioGerencial " "
            Bematech_FI_RelatorioGerencial " "
            Bematech_FI_RelatorioGerencial " "
        End If
        'Open "C:\TEFTEMP1.TXT" For Input As #1
        Open DiretorioTEFresp & "intpos.001" For Input As #1
        
        Do While Not EOF(1)
        
            Line Input #1, sBufArq
            If Left(sBufArq, 3) = "029" Then
                Bematech_FI_RelatorioGerencial Mid(sBufArq, 12, Len(sBufArq) - 12)
            End If
            Retorno = Bematech_FI_VerificaImpressoraLigada() ' ARRUMEI AKI
            Do While Retorno <> 1
                Retorno = Bematech_FI_VerificaImpressoraLigada()
                If Retorno <> 1 Then
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        strOk = False
                        Close #1
                        Exit Function
                    End If
                Else
                        Trava_Teclado
                        Close #1
                        ImprimeGerencial
                        strOk = True
                        Exit Function
                End If
            Loop
        Loop
        Close #1
    Next
    
    Bematech_FI_FechaRelatorioGerencial
    Habilita_Teclado
End Function
Public Function ImprimeGerencialADM() As Boolean
    Dim iResp As Integer
    Dim iCount As Integer
    Dim ist1 As Integer
    Dim ist2 As Integer
    Dim iAck As Integer
    Dim iRetorno As Integer
    Dim sNsu As String
    Dim sRede As String
    Dim sValor As String
    Trava_Teclado
    If ecf_FechaRelatorioGerencial = 0 Then
        ecf_FechaRelatorioGerencial
    End If
    For iCount = 1 To 2
        If iCount = 2 Then
            'Bematech_FI_RelatorioGerencial " "
            'Bematech_FI_RelatorioGerencial " "
            'Bematech_FI_RelatorioGerencial " "
            'ecf_RelatorioGerencial " "
        End If
                
        Open DiretorioTEFresp & "Intpos.001" For Input As #1
        
        Do While Not EOF(1)
            Trava_Teclado
            Line Input #1, sBufArq
            
            Retorno = 0
            Do While Retorno <> 1 Or ist1 >= 128
                
                Retorno = ecf_VerificaEstadoImpressora(iAck, ist1, ist2)
                
                If Retorno <> 1 Or ist1 >= 128 Then
                    '128 é falta de papel
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        
                        Close #1
                        NaoConfirmaVenda
                        ImprimeGerencialADM = False
                        bLigou = True
                        'Bematech_FI_CancelaCupom
                        ecf_FechaRelatorioGerencial
                        Exit Function
                    Else
                        Trava_Teclado
                        bLigou = True
                    End If
                
                End If
                
            Loop
            If bLigou Then 'vai entrar no caso de não ter respondido uma vez e ter optado por sim e ter respondido...
                    
                    Close #1
                    strOk = True
                    Trava_Teclado
                    ImprimeGerencial ' leitura X + duas vias do Comprovante TEF
                    If strOk = True Then
                        ImprimeGerencialADM = True
                        
            
                    Else
                        ImprimeGerencialADM = False
                    End If
                    Exit Function
            Else
                
                If Left(sBufArq, 3) = "029" Then
                    sBufArq = Mid(sBufArq, 12, Len(sBufArq) - 12)
                    If sBufArq = "" Then
                        sBufArq = "    "
                    End If
                    ecf_RelatorioGerencial (sBufArq)
                    
                
                End If
            End If
            
        Loop
        Close #1
    Next
    ecf_FechaRelatorioGerencial
    killIntPosResp
    ImprimeGerencialADM = True
    Habilita_Teclado
End Function

Function F_VerificaGP() As Boolean
    On Error GoTo trata_Erro
    'P_ContadorTEF
    P_Status "Criando arquivo..."
    Buffer = "000-000 = ATV" + Chr(13) + Chr(10) + _
             "001-000 =" + ContadorTEF + Chr(13) + Chr(10) + _
             "999-999 = 0"
             
    CriaArquivo Buffer, DiretorioTEFreq
    P_Status "Aguardando Retorno..."
    Retorno = EsperaArquivo("INTPOS.STS", 6, 0)
    P_Status ""
    If (Retorno = -1) Then
        F_VerificaGP = False
    Else
        F_VerificaGP = True
    End If
    Exit Function
trata_Erro:
    MsgBox Err.Description, vbCritical
    F_VerificaGP = False
    P_Status ""
End Function
Sub P_AdmTEF()
    Dim Retorno As Integer
    Dim Buffer As String
    Dim bOk As String
    Dim sBufArq As String
    Dim sAux As String
    Dim bReimp As Boolean
    Dim sMsg As String
    Dim sText As String
    Dim iResp As Integer
    Dim sRel() As String
    Dim iCount As Integer
    Dim cPage As Currency
    
'    If ecf_FechaRelatorioGerencial = 0 Then
        
 '   End If
    
    Dim cont As String
    cont = ContadorTEF
    P_Status "Enviando dados para o Gerenciador Padrão..."
    Buffer = "000-000 = ADM" + Chr(13) + Chr(10) + _
             "001-000 = " + cont + Chr(13) + Chr(10) + _
             "999-999 = 0"
             
    CriaArquivo Buffer, DiretorioTEFreq
    Retorno = EsperaArquivo("Intpos.001", 5, 1)
    P_Status ""
    bOk = True
    sText = ""
    If (Retorno = 1) Then
        Open DiretorioTEFresp & "intpos.001" For Input As #1
        sAux = ""

        Do While Not EOF(1)
        
        Line Input #1, sBufArq
        If Left(sBufArq, 3) = "009" Then
            sAux = Right(sBufArq, 1)
        End If
        If Left(sBufArq, 3) = "029" Then
            sText = sText + Mid(sBufArq, 12, Len(sBufArq) - 12) & Chr(13) + Chr(10)
        End If
        If Left(sBufArq, 3) = "030" Then
            sMsg = sMsg & Mid(sBufArq, 11, Len(sBufArq)) & Chr(13) + Chr(10)
        End If
        
    Loop
    Close #1
    
    
    
    If sAux = "0" Then
        bOk = True
    Else
        bOk = False
    End If
    If bOk = True Then
        Trava_Teclado
        If sText <> "" Then ' Caso haja impressão linha 029
            P_Status sMsg
            If ImprimeGerencialADM Then
                ConfirmaVenda
            Else
                ecf_FechaRelatorioGerencial
            End If
            If sMsg <> "" Then
                'mdiFrm.SetFocus
                'MsgBox sMsg, vbInformation
            End If
        Else
            Habilita_Teclado
            mdiFrm.SetFocus
            MsgBox sMsg, vbInformation
            killIntPosResp
        End If
    Else
        Habilita_Teclado
        mdiFrm.SetFocus
        MsgBox sMsg, vbCritical
        
        killIntPosResp
        
    End If
    
    Else
        Habilita_Teclado
        MsgBox "O Gerenciador Padrão não está ativo!", vbExclamation + vbSystemModal
    End If
    Habilita_Teclado
    P_Status ""
End Sub

Function F_CupomTEF(valor As String) As Boolean
    Trava_Teclado
    NumeroCupomTef = Space(14)
    Bematech_FI_NumeroCupom (NumeroCupomTef)
    ValorTef = valor

    PreparaVenda ValorTef, NumeroCupomTef

    DoEvents
    Retorno = EsperaArquivo("INTPOS.001", 2, 1)

    If (Retorno = 1) Then
        Bematech_FI_AbreCupom "12134"
        Bematech_FI_VendeItem "1234456", "CocaCola", "FF", "I", "10", 2, "1,00", "%", "0"
        Bematech_FI_IniciaFechamentoCupom "A", "%", "0"
        Bematech_FI_EfetuaFormaPagamento "Dinheiro", "2,00"
        Bematech_FI_EfetuaFormaPagamento "Cartao", "8,00"
        Bematech_FI_TerminaFechamentoCupom constMsg_Usu1
        FinalizaTef ValorTef
    
    Else
        Habilita_Teclado
        MsgBox "Não foi possível realizar a transação TEF", vbSystemModal
        Kill DiretorioTEFreq & "INTPOS.001"
    End If
    Habilita_Teclado
End Function
Function F_ValidaTEF(valor As String) As Boolean
    
    NumeroCupomTef = Space(14)
    Bematech_FI_NumeroCupom (NumeroCupomTef)
    ValorTef = valor

    PreparaVenda ValorTef, NumeroCupomTef

    DoEvents
    Retorno = EsperaArquivo("INTPOS.001", 2, 1)

    If (Retorno = 1) Then
        F_ValidaTEF = True
    
    Else
        'MsgBox ("Não foi possível realizar a transação TEF")
        F_ValidaTEF = False
        Kill DiretorioTEFreq & "INTPOS.001"
    End If

End Function

Sub P_VerificaFormasP()


Dim Formas As String
    Dim FormasAux As String
    Dim Retorno As Integer
    If (LocalRetorno = "1") Then 'Grava retorno em arquivo
        Formas = Space(1)
    Else
        Formas = Space(3016)
    End If
    
    Retorno = Bematech_FI_VerificaFormasPagamento(Formas)
    FormasAux = vbCr & vbLf & vbLf & Formas
    'Call VerificaRetornoImpressora("Formas de Pagamento: ", FormasAux, "Informações da Impressora")
    MsgBox Formas
End Sub
Function F_RecebNaoFiscalTef(sIndice As String, sValor As String) As Boolean
    'Esta função emite o cupom fiscal da nota
    
    Dim sBufArq As String
    Dim Buffer As String
    Dim bOk As Boolean
    Dim sMsg As String
    Dim Retorno As Integer
    
    
    Do While Retorno <> 1
        
        Retorno = Bematech_FI_VerificaImpressoraLigada()
        If Retorno <> 1 Then
            Habilita_Teclado
            If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                F_RecebNaoFiscalTef = False
                Exit Function
            End If
        End If
    Loop
    
    Retorno = Bematech_FI_RecebimentoNaoFiscal(sIndice, sValor, "cartao")
    If Retorno <> 1 Then
        MsgBox "O Recebimento não foi efetuado!", vbCritical
        F_RecebNaoFiscalTef = False
        Exit Function
    End If
    Retorno = 0 'colou aki
    
    
    NumeroCupomTef = Space(14)
    Bematech_FI_NumeroCupom (NumeroCupomTef)
    ValorTef = sValor

    PreparaVenda Format(ValorTef / 100, "###,##0.00"), NumeroCupomTef
    
    DoEvents
    
    Retorno = EsperaArquivo("INTPOS.001", 5, 1) 'baixar para sete segundos
    
    If Retorno = -1 Then
        
        MsgBox "Gerenciador padrão não está ativo!", vbCritical + vbSystemModal
        Exit Function
    End If

    Trava_Teclado
    If (Retorno = 1 Or -1) Then
        Open DiretorioTEFresp & "INTPOS.001" For Input As #1
            Buffer = ""
            bOk = True
            Do While Not EOF(1)
    
                Line Input #1, sBufArq
                If (Mid(sBufArq, 1, 3) = "009") Then
                    If Right(sBufArq, 1) <> "0" Then
                        'MsgBox "Não foi possível concluir a transação!", vbInformation
                        bOk = False
                    End If
                End If
                If (Mid(sBufArq, 1, 3) = "030") Then
                    sMsg = Mid(sBufArq, 10, Len(sBufArq) - 9)
                    'frmTroco.lblAviso.caption = sMsg
                    'frmVenda.pcMsg.Visible = True
                    frmVenda.Refresh
                    'P_CentraObejct frmTroco.lblAviso, frmTroco
                End If
            Loop
        Close #1
        If bOk = True Then
            Retorno = 0
            Do While Retorno <> 1
                Trava_Teclado
                Retorno = Bematech_FI_VerificaImpressoraLigada()
                If Retorno <> 1 Then
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        NaoConfirmaVenda
                        F_RecebNaoFiscalTef = False
                        'Bematech_FI_CancelaCupom
                        'Bematech_FI_FechaRelatorioGerencial
                        Exit Function
                    Else
                        Trava_Teclado
                    End If
                End If
            Loop
            If cDinheiro > 0 Then
                Bematech_FI_EfetuaFormaPagamento "Dinheiro", Format(cDinheiro, "###,##0.00")
            End If
            If cAprazo > 0 Then
                Bematech_FI_EfetuaFormaPagamento "A Prazo", Format(cAprazo, "###,##0.00")
            End If
            Retorno = 0
            Do While Retorno <> 1
                Trava_Teclado
                Retorno = Bematech_FI_VerificaImpressoraLigada()
                If Retorno <> 1 Then
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        F_RecebNaoFiscalTef = False
                        Exit Function
                    Else
                        Trava_Teclado
                    End If
                    
                End If
            Loop
            crcValorPago = CCur(sValor)
            Bematech_FI_EfetuaFormaPagamento "Cartao", Format(crcValorPago, "###,##0.00")
            
            Bematech_FI_TerminaFechamentoCupom constMsg_Usu1
            
            If FinalizaTef(CStr(crcValorPago)) Then ' abre a emissão dos comprovantes TEF
                
                F_RecebNaoFiscalTef = True
            Else
                F_RecebNaoFiscalTef = False
            End If
            Habilita_Teclado
        Else
            Habilita_Teclado
            strSQL = "Não foi possível concluir a transação!" & Chr(13)
            strSQL = strSQL & "Mensagem: '" & sMsg & "'" & Chr(13)
            
            MsgBox strSQL, vbCritical + vbSystemModal
            'Bematech_FI_CancelaCupom
            'NaoConfirmaVenda ' - so mandar ncn qd. desl. micro\imp.
            F_RecebNaoFiscalTef = False
            Habilita_Teclado
            If ((Dir(DiretorioTEFresp & "INTPOS.001")) <> "") Then
                Kill DiretorioTEFresp & "INTPOS.001"
            End If
            On Error Resume Next
            Kill "C:\TEFTEMP1.TXT"
            Kill "C:\INTPOS.TMP"
            
        End If
    Else
        MsgBox ("Erro no TEF Selecione outra forma de Pagamento")
        F_RecebNaoFiscalTef = False
        Habilita_Teclado
    End If
Habilita_Teclado
End Function

'End Function

Function F_VendaCupomTEF(operacao As String, sNum_ven As String, Optional strNome_Cli As String = "", _
Optional strCNPJ As String = "", Optional crcValorPago As Currency = 0, _
Optional bVendaCartao As Boolean = True, Optional cDinheiro As Currency = 0, _
Optional cAprazo As Currency = 0, Optional crcDesconto As Currency = 0) As Boolean
    'Esta função emite o cupom fiscal da nota
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim strST As String
    Dim sBufArq As String
    Dim Buffer As String
    Dim bOk As Boolean
    Dim sMsg As String
    Dim Retorno As Integer
    Dim ack As Integer
    Dim st1 As Integer
    Dim st2 As Integer
    Dim rImp As retorno_impressora
    On Error GoTo trata_Erro
    killIntPosResp
    If strNum_Cpm = "" Then
        MsgBox "Não há um cupom fiscal Gerado!", vbExclamation
        Exit Function
    End If
    
    Do While Retorno <> 1
        P_Msg "Conectando com ECF", "Transferência Eletrônica de Fundos"
        Retorno = ecf_VerificaImpressoraLigada
        If Retorno <> 1 Then
            Habilita_Teclado
            If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                P_Msg "", ""
                F_VendaCupomTEF = False
                Exit Function
            End If
        End If
    Loop
    
        
    Retorno = 0 'colou aki
    Do While Retorno <> 1
        P_Msg "Conectando com ECF", "Transferência Eletrônica de Fundos"
        Retorno = ecf_VerificaImpressoraLigada()
        If Retorno <> 1 Then
            
            If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                F_VendaCupomTEF = False
                P_Msg "", ""
                Exit Function
            End If
        End If
    Loop
    If operacao = "B" Then
        rImp = ecf_retornoImprComandos()
        
        If rImp.cupomaberto = False Then
            If Not F_EmiteCupomVenda(sNum_ven, 17, strNome_Cli, _
                    strCNPJ, crcValorPago, False, True) Then
                F_VendaCupomTEF = False
                Exit Function
            End If
        End If
    End If
        
   
       
    NumeroCupomTef = "      "
    i = ecf_NumeroCupom(NumeroCupomTef)
    ValorTef = CStr(crcValorPago)
    i = ecf_IniciaFechamentoCupom("D", "$", CStr(crcDesconto))
    PreparaVenda Format(ValorTef, "###,##0.00"), strNum_Cpm
    
    DoEvents
    P_Msg "Aguardando Arquivo", "Transferência Eletrônica de Fundos"
    Retorno = EsperaArquivo("INTPOS.001", 5, 1) 'baixar para sete segundos
    
    If Retorno = -1 Then
        P_Msg
            MsgBox "Gerenciador padrão não está ativo!", vbCritical + vbSystemModal
        Exit Function
    End If
    
    Trava_Teclado
    If (Retorno = 1 Or -1) Then
        Open DiretorioTEFresp & "INTPOS.001" For Input As #1
            Buffer = ""
            bOk = True
            Do While Not EOF(1)
    
                Line Input #1, sBufArq
                If (Mid(sBufArq, 1, 3) = "009") Then
                    If Right(sBufArq, 1) <> "0" Then
                        'MsgBox "Não foi possível concluir a transação!", vbInformation
                        bOk = False
                    End If
                End If
                If (Mid(sBufArq, 1, 3) = "030") Then
                    sMsg = Mid(sBufArq, 10, Len(sBufArq) - 9)
                    P_Msg sMsg, ""
                    
                    
                    'P_CentraObejct frmTroco.lblAviso, frmTroco
                End If
            Loop
        Close #1
        If bOk = True Then
            Retorno = 0
            Do While Retorno <> 1
                Trava_Teclado
                P_Msg "Conectando ECF", "Transferência Eletrônica de Fundos"
                Retorno = ecf_VerificaImpressoraLigada()
                If Retorno <> 1 Then
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        NaoConfirmaVenda
                        F_VendaCupomTEF = False
                        P_Msg
                        'Bematech_FI_CancelaCupom
                        'Bematech_FI_FechaRelatorioGerencial
                        Exit Function
                    Else
                        Trava_Teclado
                    End If
                End If
            Loop
            
            'ecf_verificaImpressora
            crcDesconto = 0
            
                        
            If cDinheiro > 0 Then
                P_Msg "Efetuando Forma de Pagamento", "Transferência Eletrônica de Fundos"
                ecf_EfetuaFormaPagamento "Dinheiro", Format(cDinheiro, "###,##0.00")
            End If
            If cAprazo > 0 Then
                ecf_EfetuaFormaPagamento "A Prazo", Format(cAprazo, "###,##0.00")
            End If
            
            Retorno = 0
            Do While Retorno <> 1
                Trava_Teclado
                Retorno = ecf_VerificaImpressoraLigada()
                If Retorno <> 1 Then
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        F_VendaCupomTEF = False
                        Exit Function
                    Else
                        Trava_Teclado
                    End If
                    
                End If
            Loop
            
            ecf_EfetuaFormaPagamento "Cartao", Format(crcValorPago, "###,##0.00")
            'ecf_verificaImpressora
            ecf_TerminaFechamentoCupom constMsg_Usu1
            
            If FinalizaTef(CStr(crcValorPago)) Then ' abre a emissão dos comprovantes TEF
                
                F_VendaCupomTEF = True
            Else
                F_VendaCupomTEF = False
            End If
            Habilita_Teclado
        Else
            Habilita_Teclado
            strSQL = "Não foi possível concluir a transação!" & Chr(13)
            strSQL = strSQL & "Mensagem: '" & sMsg & "'" & Chr(13)
            
            MsgBox strSQL, vbCritical + vbSystemModal
            'Bematech_FI_CancelaCupom
            'NaoConfirmaVenda ' - so mandar ncn qd. desl. micro\imp.
            F_VendaCupomTEF = False
            
            If ((Dir(DiretorioTEFresp & "INTPOS.001")) <> "") Then
                Kill DiretorioTEFresp & "INTPOS.001"
            End If
            On Error Resume Next
            Kill "C:\TEFTEMP1.TXT"
            Kill "C:\INTPOS.TMP"
            
        End If
    Else
        MsgBox ("Erro no TEF Selecione outra forma de Pagamento")
        F_VendaCupomTEF = False
        Habilita_Teclado
    End If
    P_Msg
Habilita_Teclado
Exit Function
trata_Erro:
    Habilita_Teclado
    P_Msg
    P_Status
    P_MostraErro
End Function

Function FormatValTef(sValor) As String
    On Error GoTo trata_Erro
    FormatValTef = F_SoNumero(FormatVal(sValor))
    Exit Function
trata_Erro:
    FormatValTef = sValor
End Function

Sub killIntPosResp()
    On Error Resume Next
    Kill DiretorioTEFresp & "intpos.001"
End Sub
Sub Trava_Teclado()
'    BlockInput True

End Sub

Sub Habilita_Teclado()
    BlockInput False

End Sub
