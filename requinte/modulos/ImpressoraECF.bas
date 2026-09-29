Attribute VB_Name = "ImpressoraECF"
Option Explicit
'Constantes de Comando da Impressora
Global conAbre_Protocolo_Fiscal As String
Global conFecha_Protocolo_Fiscal As String
  
'Constante para porta de comunicação
Global strPorta_Com_Fscl As String

'Constantes de valores de retorno para FormataTx
Public Const conFTXSucesso_Fscl = 0
Public Const conFTXErro_Comunicacao_Fscl = 1
Public Const conFTXParametro_Invalido_Fscl = -2
Public Const conFTXVersao_Antiga_Fscl = -3

'Constantes de valores de retorno para FechaPorta
Public Const conFCHPErro_Fscl = 0
Public Const conFCHPSucesso_Fscl = 1

'Constantes de valores de retorno para IniPortaStrr
Public Const conINPErro_Fscl = 0
Public Const conINPSucesso_Fscl = 1

'Constantes relativas aos comandos de configuração aceitos pela impressora
Public Const conCMDAlt_Simbolo_Moeda = "01|" 'Alteração da Moeda Corrente
Public Const conCMDAdd_Aliquota_Tributaria = "07|" 'Adição de alíquota tributária
Public Const conCMDPro_Horario_Verao = "18" 'Altera o horário para horário de verão
Public Const conCMDNom_Tot_Nao_ICMS = "40|" 'Nomeia totalizadores parciais não sujeitos ao ICMS
Public Const conCMDPro_Truncamento = "39|"  'Programa truncamento/arredondamento
Public Const conCMDPro_LinhasCupon = "61|" 'Programa número de linhas entre cupons
Public Const conCMDNom_Departamento = "65|" 'Nomeia um departamento
Public Const conCMDHab_Cupon_Adicional = "68|" 'Habilita cupon Adicional
Public Const conCMDReset = "70" 'Reseta a impressora

'Constantes relativas aos comandos de configuração aceitos pela impressora
Public Const conCMDAbre_Cupon = "00|" 'Abre cupon Fiscal
Public Const conCMDVenda_Item = "09|" 'Venda de item
Public Const conCMDAcrescimo_Financeiro = "62 50|" 'Acréscimo Financeiro - IOF
Public Const conCMDVenda_Item_3Casas = "56|" 'Venda de item com 3 casas decimais
Public Const conCMDCanc_Item_Anterior = "13" 'Cancelamento do item anterior
Public Const conCMDCand_Item_Generico = "31|" 'Cancelamento de item genérico
Public Const conCMDVenda_Item_Ent_Depto = "63|" 'Venda de Item com entrada de Departamento, Desconto e Unidade de medida
Public Const conCMDInicia_Fechamento_Cupon = "32|" 'Inicia o fechamento de cupon
Public Const conCMDProVer_Forma_Pgto = "71|" 'Programa ou verifica forma de pagamento
Public Const conCMDEfe_Forma_Pgto = "72|" 'Efetua forma de pagamento
Public Const conCMDPro_Desc_Forma_Pgto = "73|" 'Programa a descrição das formas de pagamento
Public Const conCMDTer_Fechamento_Cupon_CForma = "34|" 'Programa o fechamento do cupon com forma de pagamento
Public Const conCMDCancela_Cupon = "14" 'Cancela o cupon
Public Const conCMDPro_Unid_Medida = "62|51|" 'Programa a unidade de medida para o próximo item
Public Const conCMDAumenta_Descricao = "62|52|" 'Aumenta a descrição para o próximo item
Public Const conCMDEmite_Reducao_Z = "05|" 'Emite a redução Z e fecha o movimento do dia
Public Const conCMDEmite_Leitura_X = "06" 'Emite a leitura "X"
Public Const conCMDEmite_Leitura_Fiscal = "08|" 'Emite leitura da memória fiscal

'Constante de Informações da Impressora Fiscal
Public Const conINFLe_Est_Impre = "19" 'Leitura do Estado da impressora
Public Const conINFAliquota = "26" 'Retorno das Alíquotas Tributadas
Public Const conINFTot_Parciais = "27" 'Retorno dos Totalizadores Parciais
Public Const conINFSub_Total = "29" 'Retorno do SubTotal
Public Const conINFNum_Cupom = "30" 'Retorno do Número do cupom
Public Const conINFVariaveis = "35|" 'Retorno das Variáveis

'Variáveis utilizadas com o comando 35
Public Const conVARNum_Serie = "00" 'Número de série
Public Const conVARVer_FirmWare = "01" 'Versão do FirmWare
Public Const conVARCGC_IE = "02" 'Retorna o CNPJ e IE
Public Const conVARGrandeTotal = "03" 'Retorna o Grande Total
Public Const conVARCancelamentos = "04" 'Retorna os cancelamentos
Public Const conVARDescontos = "05" 'Descontos
Public Const conVARContadorSequencial = "06" 'Contador Sequencial
Public Const conVARNum_Oper_NFiscais = "07" 'Número de operações ñ fiscais
Public Const conVARNum_Cupons_Cancelados = "08" 'Número de cupons cancelados
Public Const conVARNum_Reducoes = "09"  'Número de Reduções
Public Const conVARNum_Int_Tec = "10" 'Número de Intervenções Técnicas
Public Const conVARSubst_Proprientario = "11" 'Número Substituições de Proprietários
Public Const conVARNum_Ult_ItemVendido = "12" 'Número do último item vendido
Public Const conVARCliche_Propr = "13" 'Clichê do proprietário
Public Const conVARNum_Caixa = "14" 'Número do Caixa
Public Const conVARNum_Loja = "15"  'Número da Loja
Public Const conVARMoeda = "16" 'Moeda
Public Const conVARFlagsFiscais = "17" 'Flags fiscais
Public Const conVARMin_Ligada = "18" 'Minutos Ligada
Public Const conVARMin_Imprimindo = "19" 'Minutos Imprimindo
Public Const conVARFlag_Int_Tec = "20" 'Flag Intervenção Técnica
Public Const conVARFlag_Eprom_Conectada = "21" 'Flag de Eprom Conectada
Public Const conVARValor_Pago_Ult_Cupom = "22" 'Valor pago no último cupom
Public Const conVARData_Hora = "23" 'Data e hora atual
Public Const conVARCont_Tot_NaoICMS = "24" 'Contadores não sujeitos ao ICMS
Public Const conVARDescr_Tot_NaoICMS = "25" 'Descrição dos contadores não sujeitos ao ICMS
Public Const conVARData_Reducao = "26" 'Data da última redução
Public Const conVARData_Movimento = "27" 'Data do movimento
Public Const conVARFlag_Truncamento = "28" 'Flag Truncamento
Public Const conVARFlag_Vinc_ISS = "29" 'Flags de Vinculação do ISS
Public Const conVARTot_Acrescimo = "30" 'Flag Totalizador de Acrescimo
Public Const conVARCont_Bilhet_Passagem = "31" 'Contador Bilhet_Passagem
Public Const conVARFormas_Pagamento = "32" 'Formas de pagamento
Public Const conVARRecebimentos = "33" 'Recebimentos (CNF não vinculado)
Public Const conVARDepartamentos = "34" 'Departamentos
Public Const conVARTipo_Impressora = "35" 'Tipo da impressora


'Declaração das funções da DLL
Declare Function IniPortaStr Lib "MP20FI32.DLL" (ByVal Porta As String) As Integer
Declare Function FormataTX Lib "MP20FI32.DLL" (ByVal Buffer As String) As Integer
Declare Function FechaPorta Lib "MP20FI32.DLL" () As Integer
Declare Function Status_Mp20FI Lib "MP20FI32.DLL" (ByRef Retorno As Integer, ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Declare Function Le_Variaveis Lib "MP20FI32.DLL" (ByVal var As String) As Integer



'Funções de Uso do Sistema c/ a Impressora
Function F_AbrePortaECF() As Boolean
    ' Esta função retorna verdadeiro se conseguir abrir comunicação com a
    'impressora fiscal
    
    If F_Vazio(strPorta_Com_Fscl) Then
       'Não existe ECF
       'A variável de porta de comunicação está vazia retorna falso
       F_AbrePortaECF = False
       Exit Function
    End If
    
    F_AbrePortaECF = _
    F_StatusComunicacao(IniPortaStr(strPorta_Com_Fscl), "A") 'Abre
End Function

Function F_FechaPortaECF() As Boolean
    ' Esta função retorna verdadeiro se conseguir abrir comunicação com a
    'impressora fiscal
    F_FechaPortaECF = _
    F_StatusComunicacao(FechaPorta, "F") 'Fecha
End Function

Function F_ComandoECF(strExpr As String) As Boolean
    '  Esta função envia os comandos para a impressora fiscal retornando
    'verdadeiro caso tudo esteja OK ou falso
    
    F_ComandoECF = F_StatusFormataTX(FormataTX(strExpr))
    
End Function
Function F_StatusComunicacao(Retorno As Integer, strTipo As String) _
As Boolean
'    Esta função retorna verdadeiro caso o estado de comunicação com a
'impressora (abertura e fechamento de porta) esteja correto
strTipo = UCase(strTipo)
If Left(strTipo, 1) = "A" Then
   strTipo = "Abrindo Porta de Comunicação"
Else
   strTipo = "Fechando Porta de Comunicação"
End If
   

If Retorno <= 0 Then
   F_StatusComunicacao = False
   MsgBox "Erro de comunicação com a impressora!", vbCritical, strTipo
   Exit Function
End If
F_StatusComunicacao = True

End Function



Function F_StatusFormataTX(Retorno As Integer) As Boolean
'    Esta função retorna verdadeiro se o comando da Função FORMATATX
'foi passado e recebido corretamente.
F_StatusFormataTX = False
If Retorno = conFTXErro_Comunicacao_Fscl Then
   MsgBox "Erro na comunicação com a impressora!", vbCritical
ElseIf Retorno = conFTXParametro_Invalido_Fscl Then
   MsgBox "Erro na passagem de parâmetros!", vbCritical
ElseIf Retorno = conFTXVersao_Antiga_Fscl Then
   MsgBox "Erro de compatibilidade. Impressora não suporta o comando!", _
   vbCritical
Else
   'Tudo all right
   F_StatusFormataTX = True
End If
End Function

Function F_AbrirCupomFscl(Optional strCPF As String = "") As Boolean
'  Esta função Abre um cupom fiscal retornando verdadeiro se tudo estiver
'certo.
'  Atenção esta função deixa a porta de comunicação aberta
    Dim strExpr As String
    Dim Retorno As Integer
    
    'Iniciando comunicação com a impressora
    If Not F_AbrePortaECF Then
       F_AbrirCupomFscl = False
       Exit Function
    End If
    
    strExpr = conAbre_Protocolo_Fiscal & conCMDAbre_Cupon & strCPF
    strExpr = strExpr & conFecha_Protocolo_Fiscal
    
        
    'Iguala com o valor de Status de FormataTX
    F_AbrirCupomFscl = F_ComandoECF(strExpr)
    
    
End Function

Function F_VendaItem(strCod As String, strDescricao As String, _
strSituacaoTributaria As String, crcQuantidade As Currency, _
crcUnitario As Currency, Optional strUnidade As String = "PC", _
Optional crcDesconto As Currency = 0) _
As Boolean
'      Esta função só pode ser executada se o cupom fiscal e a porta de
'comunicação estiverem abertos

    
    'Efetua venda de um item no cupom fiscal
    'Fazer até qdo a mercadoria tiver quantidade
    Do While crcQuantidade > 0
        
        Dim Ret As Integer
        Dim strQuantidade As String
        Dim strUnitario As String
        Dim strDesconto As String
    
        If crcQuantidade > 9999 Then
           strQuantidade = F_SoNumero(Format(9999, "0###.000"))
        Else
           strQuantidade = F_SoNumero(Format(crcQuantidade, "0###.000"))
        End If
        crcQuantidade = crcQuantidade - 9999
        strUnitario = F_SoNumero(Format(crcUnitario, "0#####.00"))
        If frmTroco.pc.BackColor = vbGreen Then
            strDesconto = F_SoNumero(Format(crcDesconto, "0#####.00"))
        Else
            strDesconto = F_SoNumero(Format(0, "0#####.00"))
        End If
        strUnidade = UCase(strUnidade)
    
        'Quantidade não pode ser maior que 9999
        Dim strExpr As String
    
        strCod = F_SoNumero(strCod)
        strCod = Format(strCod, "0000##\.####\-#")
        strDescricao = Trim(strDescricao)
    
        If Len(strDescricao) > 29 Then
            'Envia comando para a impressora Fiscal
            If Not F_ComandoECF(conAbre_Protocolo_Fiscal & _
            conCMDAumenta_Descricao & _
            strDescricao & conFecha_Protocolo_Fiscal) Then
          
                'Diferente de sucesso
                FechaPorta
                F_VendaItem = False
                Exit Function
            End If
        End If
    
        'Unidade de Medida da mercadoria
        strExpr = conAbre_Protocolo_Fiscal
        strExpr = strExpr & conCMDPro_Unid_Medida & strUnidade
        strExpr = strExpr & conFecha_Protocolo_Fiscal
    
    
        'Envia o comando verificando se foi executado
        If Not F_ComandoECF(strExpr) Then
        F_VendaItem = False
            FechaPorta
            Exit Function
        End If
    
        strExpr = conAbre_Protocolo_Fiscal
        strExpr = strExpr & conCMDVenda_Item
        strExpr = strExpr & strCod & "|"
        strExpr = strExpr & Left(strDescricao, 29) & "|"
        strExpr = strExpr & strSituacaoTributaria & "|"
        strExpr = strExpr & strQuantidade & "|"
        strExpr = strExpr & strUnitario & "|"
        strExpr = strExpr & strDesconto
        strExpr = strExpr & conFecha_Protocolo_Fiscal
    
        'Iguala cfe a execução do comando
        F_VendaItem = F_ComandoECF(strExpr)
    Loop
    
End Function

Function F_IniFechaCupomFsl(strTipo As String, Optional crcDescAcresc _
As Currency = 0) As Boolean
'   Esta função inicia o fechamento do cupom fiscal, que juntamente com a
'porta de comunicação deve estar aberto
'Ela também fecha a porta de comunicação com a impressora
    Dim strExpr As String
    Dim strValor As String
    Dim Ret As Integer
    'Formatando valor
    
    
    
    strValor = F_SoNumero(Format(crcDescAcresc, "0###########.00"))
    
    strExpr = conAbre_Protocolo_Fiscal
    strExpr = strExpr & conCMDInicia_Fechamento_Cupon
    strExpr = strExpr & LCase(strTipo) & "|"
    strExpr = strExpr & strValor
    strExpr = strExpr & conFecha_Protocolo_Fiscal
    
    'Iguala conforme o comando executado
    F_IniFechaCupomFsl = F_ComandoECF(strExpr)
    
    If Not F_FechaPortaECF Then
       'Ocorreu um erro ao liberar a porta de comunicação
       F_IniFechaCupomFsl = False
    End If
    
End Function

Function F_FormaPagamento(strTipo As String, crcValor As Currency, _
Optional strDescricao As String = "") As Boolean
'Esta função especifica a forma de pagamento para o Cupom Fiscal
    
    Dim Ret As Integer
    Dim strValor As String
    Dim strExpr As String
    
    If Not F_AbrePortaECF Then
       'Ocorreu um erro ao abrir porta para comunicação
       F_FormaPagamento = False
    End If
    
    strValor = F_SoNumero(Format(crcValor, "0###########.00"))
    strExpr = conAbre_Protocolo_Fiscal
    strExpr = strExpr & conCMDEfe_Forma_Pgto
    strExpr = strExpr & strTipo & "|"
    strExpr = strExpr & strValor
    If Not F_Vazio(strDescricao) Then
       strExpr = strExpr & "|" & strDescricao
    End If
 
    strExpr = strExpr & conFecha_Protocolo_Fiscal
    
    F_FormaPagamento = F_ComandoECF(strExpr)
    
    If Not F_FechaPortaECF Then
       'Ocorreu um erro ao fechar porta de comunicação
       F_FormaPagamento = False
    End If
    
End Function

Function F_TermFechamentoCupom(Optional strMens As String = "") _
As Boolean
'Termina o fechamento do Cupom Fiscal
    If Not F_AbrePortaECF Then
       'Ocorreu um erro ao abrir porta para comunicação
       F_TermFechamentoCupom = False
    End If
    If F_Vazio(strMens) Then
      strMens = "" & Chr(13) & Chr(10) & _
      "Agradece a sua preferência!" & Chr(13) & Chr(10) & _
      "Tenha " & IIf(Left(F_Cumprimenta(), 3) = "Bom", "um ", "uma ") & _
      F_Cumprimenta() & "|" & Chr(13) & Chr(10)
    End If
    
    Dim Ret As Integer
     
    
    'Iguala conforme o sucesso de execução do comando
    F_TermFechamentoCupom = F_ComandoECF(conAbre_Protocolo_Fiscal & _
    conCMDTer_Fechamento_Cupon_CForma & strMens & _
    conFecha_Protocolo_Fiscal)
    
    'Fechando comunicação com a impressora fiscal
    If Not F_FechaPortaECF Then
       F_TermFechamentoCupom = False
    End If
     

End Function

Function F_LeFormasPagamento() As String
  ''On Error  GoTo TrataErro
  '  Esta função lê as formas de pagamento gravadas em arquivo e
  'retorna em String o conteúdo lido.
  
  Dim intRet As Integer
  Dim strExpr As String
 
  
  'Iniciando porta de comunicação c/ a impressora
  If Not F_AbrePortaECF Then
     'Ocorreu erro
     F_LeFormasPagamento = ""
     FechaPorta
     P_Reset_Impr_Fscl
     Exit Function
  End If
  
    
  'Executa o comando informando o sucesso ou insucesso do mesmo
  If Not F_ComandoECF(conAbre_Protocolo_Fiscal & _
  conINFVariaveis & conVARFormas_Pagamento & _
  conFecha_Protocolo_Fiscal) Then
     
     'Caso a função FORMATATX retornar erro, saia da função
     F_LeFormasPagamento = ""
     FechaPorta
     Exit Function
     
  End If
  
  'Pega o retorno da impressora
  strExpr = F_RetornoArquivo
  
  If strExpr = "" Then
     'Não retornou nada
     F_LeFormasPagamento = ""
     FechaPorta
     Exit Function
  End If
     
  'Formatando StrExpr para mostrar somente os caracteres válidos
  strExpr = _
  Left(strExpr, InStr(1, strExpr, "Valor Recebido") - 1)
  
  
  'Retorna String com as descrições retirando os caracteres
  'chr(0) e chr(1)
  F_LeFormasPagamento = _
  F_Trunca(F_Trunca(strExpr, Chr(1)), Chr(0))
  
  
  If Not F_FechaPortaECF Then
     intRet = FechaPorta
  End If
  Exit Function

TrataErro:
   MsgBox Err.Description, vbCritical, "Impressora Fiscal!"
End Function


Sub P_Reset_Impr_Fscl()
    Dim intRet As Integer
    intRet = F_ComandoECF(Chr(27) & "|70|" & Chr(27))
End Sub

Function F_ProgramaFormaPagamento(strDescricao As String) As Boolean
'   Esta função programa as formas de pagamento
'na impressora fiscal
Dim intRet As Integer
Dim strExpr As String

F_ProgramaFormaPagamento = False

If Not F_AbrePortaECF Then
   'Erro de abertura
   Exit Function
End If

strExpr = conAbre_Protocolo_Fiscal
strExpr = strExpr & conCMDProVer_Forma_Pgto & UCase(strDescricao)
strExpr = strExpr & conFecha_Protocolo_Fiscal

'Retorna valor cfe o sucesso da função FORMATATX
F_ProgramaFormaPagamento = F_ComandoECF(strExpr)

If Not F_ProgramaFormaPagamento Then
   FechaPorta
   Exit Function
End If

'Retorna valor cfe o Sucesso da liberação da porta
F_ProgramaFormaPagamento = F_FechaPortaECF

End Function

Sub P_GravaFormaPagamento()
'Este procedimento grava algumas formas de pagamento

Dim strExpr As String
Dim blnRet As Boolean

'Pegando as formas de pagamentos já gravadas
strExpr = F_LeFormasPagamento

If InStr(1, strExpr, "Dinheiro") = 0 Then
   blnRet = F_ProgramaFormaPagamento("Dinheiro")
   If Not blnRet Then
      Exit Sub
   End If
End If
If InStr(1, strExpr, "A Prazo") = 0 Then
   blnRet = F_ProgramaFormaPagamento("A Prazo")
   If Not blnRet Then
      Exit Sub
   End If
End If
End Sub

Sub P_PreencheFormaPagamento(cmbxOBJ As ComboBox)
'   Este procedimento preenche um combo box com as formas de pagamentos
'dísponiveis na impressora fiscal

   Dim intCont As Integer
   Dim strExpr As String
   strExpr = F_LeFormasPagamento
   cmbxOBJ.Clear
   For intCont = 1 To Len(strExpr) Step 16
       cmbxOBJ.AddItem Mid(strExpr, intCont, 16)
   Next
   
End Sub

Sub P_PreencheECFTaxasICMS(cmbxOBJ As ComboBox)
   ''On Error  GoTo TrataErro
   Dim strExpr As String
   Dim intReg As Integer
   Dim strExpr2 As String
   cmbxOBJ.Clear
   If Not F_AbrePortaECF Then
      Exit Sub
   End If
   
        
   strExpr = conAbre_Protocolo_Fiscal
   strExpr = strExpr & conINFAliquota
   strExpr = strExpr & conFecha_Protocolo_Fiscal
   
   intReg = F_ComandoECF(strExpr)  'Iniciando
   strExpr = F_RetornoArquivo 'Pegando valores de retorno
      'Pegando quais alíquotas pertecem ao ISS
   If strExpr = "" Then
      Exit Sub
   End If
   For intReg = 3 To (CCur(Left(strExpr, 2) - 1) * 4) + 3 Step 4
       cmbxOBJ.AddItem Mid(strExpr, intReg, 2) & " %"
   Next
   If Not F_FechaPortaECF Then
        Exit Sub
   End If
   Exit Sub

TrataErro:
   P_MostraErro

End Sub

Function F_NumeroCupom() As String
'Esta função retorna o número do cupom fiscal atual
  ''On Error  GoTo TrataErro
  Dim strExpr As String
  Dim intRet As Integer
  
  F_NumeroCupom = ""
  If Not F_AbrePortaECF Then
     Exit Function
  End If
  
  'Expressão para indicar o número do cupom
  strExpr = conAbre_Protocolo_Fiscal
  strExpr = strExpr & conINFNum_Cupom
  strExpr = strExpr & conFecha_Protocolo_Fiscal
    
  If Not F_ComandoECF(strExpr) Then
     FechaPorta
     Exit Function
  End If
  strExpr = F_RetornoArquivo
  intRet = FechaPorta
  
  F_NumeroCupom = Format(strExpr + 1, "0#######")
  Exit Function

TrataErro:
   P_MostraErro

End Function

Function F_RetornoArquivo() As String
'Esta função retorna o conteúdo do Arquivo MP20FI.RET
'Arquivo de retorno da impressora Fiscal
   ''On Error  GoTo TrataErro
   Dim strExpr As String
   
   F_RetornoArquivo = ""
   If Dir("Mp20Fi.Ret") = "" Then
      'Não achou o arquivo retorna
      Exit Function
   End If
   
   'Abrindo arquivo
   Open "Mp20Fi.Ret" For Input As #1
   Line Input #1, strExpr
   strExpr = F_ConvDOS(strExpr)
   Close #1
   
   'Apagando arquivo
   Kill "Mp20Fi.Ret"
   F_RetornoArquivo = strExpr
   Exit Function

TrataErro:
   P_MostraErro "Impressora Fiscal"
End Function

Function F_ECFSubTotal() As Currency
    'Esta função retorna o valor do Sub-Total do Cupom atual
    'O cupom deve estar aberto c/ a porta COM fechada
    
    'Abrindo comunicação com a impressora fiscal
    If Not F_AbrePortaECF Then
       F_ECFSubTotal = -1 'Erro
       Exit Function
    End If
    
    Dim strExpr As String
    
    strExpr = conAbre_Protocolo_Fiscal
    strExpr = strExpr & conINFSub_Total
    strExpr = strExpr & conFecha_Protocolo_Fiscal
    
    If Not F_ComandoECF(strExpr) Then
       'Ocorreu erro retorna zero
       FechaPorta
       F_ECFSubTotal = 0
       Exit Function
    End If
    
    'Abre arquivo
    strExpr = F_RetornoArquivo
    
    'Retorna o valor como o arquivo grava 100 = 1,00 divide-se por 100
    F_ECFSubTotal = CCur(strExpr) / 100
    
    If Not F_FechaPortaECF Then
    'Ocorreu um erro ao fechar a porta de comunicaçãO
       F_ECFSubTotal = -1
       Exit Function
    End If
End Function

Function F_ReducaoZ() As Boolean
  'Esta função efetua a REDUÇÃO "Z"
  'Retornando verdadeiro se conseguiu
  Dim strExpr As String
  
  F_ReducaoZ = False
  If Format(F_ECFData, "DD/MM/YYYY") <> Format(Date, "DD/MM/YYYY") Then
     'A data do computador está diferente do ECF
     
     If MsgBox("A data do computador está diferente com a do cupom fiscal!" & _
     Chr(13) & "Data e hora atuais do ECF: " & F_ECFData & _
     Chr(13) & "Data e hora atuais da CPU: " & Now & _
     Chr(13) & "Corrigir a Data e Hora do computador? ", vbQuestion + vbYesNo) = vbYes Then
        'O usuário solicitou o ajuste da data da máquina
        
        
        Time = Format(F_ECFData, "HH:MM:SS") 'Ajusta a hora do sistema
        Date = CDate(F_ECFData)              'Ajusta a data do sistema
        
     Else
       'Data diferente não posso efetuar Redução Z
       Exit Function
     End If
  End If
  
  strExpr = Format(F_ECFData, "HH:MM:SS")
  If Hour(strExpr) <> Hour(Time) Then
     
     If MsgBox("   Há diferença entre o relógio da Impressora Fiscal e" & _
     Chr(13) & "o relógio interno da CPU!!" & _
     Chr(13) & "ECF.......: " & Format(F_ECFData, "HH:MM:SS") & _
     Chr(13) & "Computador: " & Time & _
     Chr(13) & "Corrigir o relógio do computador com o ECF?", _
     vbQuestion + vbYesNo) = vbYes Then
      
       'Corrige a hora do sistema  pelo ECF
       Time = Format(F_ECFData, "HH:MM:SS")
       
     Else
       'Hora incorreta, então não faça redução Z
       Exit Function
     End If
  End If
  
  If Minute(strExpr) - Minute(Time) > 5 Or _
     Minute(Time) - Minute(strExpr) > 5 Then
     
     If MsgBox("   Há diferença entre o relógio da Impressora Fiscal e" & _
     Chr(13) & "o relógio interno da CPU!!" & _
     Chr(13) & "ECF.......: " & Format(F_ECFData, "HH:MM:SS") & _
     Chr(13) & "Computador: " & Time & _
     Chr(13) & "Corrigir o relógio do computador com o ECF?", _
     vbQuestion + vbYesNo) = vbYes Then
      
       'Corrige a hora do sistema  pelo ECF
       Time = Format(F_ECFData, "HH:MM:SS")
       
     Else
       'Hora incorreta, então não faça redução Z
       Exit Function
     End If
     
  End If
  
        
        
  If Not F_AbrePortaECF Then
     'Não conseguiu abrir a porta de comunicação com a impressora
     Exit Function
  End If
  
  
  strExpr = conAbre_Protocolo_Fiscal
  strExpr = strExpr & "05|"
  
  'Tenta ajustar o relógio
  strExpr = strExpr & Format(Now, "DDMMYY|HHMMSS")
  
  strExpr = strExpr & conFecha_Protocolo_Fiscal
  
  
     
  
  F_ReducaoZ = F_ComandoECF(strExpr)
  
  
  FechaPorta
  
  
End Function

Function F_HouveReducaoZData() As Boolean
   'Indica se já foi ou não efetuada a redução Z na data
   Dim strExpr As String
   F_HouveReducaoZData = Format(F_DataUltReducaoZ, "DD/MM/YYYY") = _
   Format(F_ECFData, "DD/MM/YYYY")
End Function

Function F_ECFStatus() As String

   'Esta Retorna o estado da impressora caso houver
   
   Dim Estado As Integer
   Dim intRet As Integer, ACK As Integer, ST1 As Integer, ST2 As Integer
   Dim strMens As String
   
   F_ECFStatus = ""
   F_AbrePortaECF
   F_ComandoECF (Chr(27) & "|19|" & Chr(27))
   Estado = Status_Mp20FI(intRet, ACK, ST1, ST2)
   If intRet = conFTXErro_Comunicacao_Fscl Then Beep
   If Not F_StatusFormataTX(intRet) Then
      'Ocorreu erro de comunicação
      F_ECFStatus = "Erro de comunicação!" & Chr(13) & Chr(10)
   End If
   
   If ACK <> 6 Then
      MsgBox "Erro ao se comunicar com a impressora!", vbCritical
      F_ECFStatus = F_ECFStatus & _
      "Recepção de dados incompreendidos!" & Chr(13) & Chr(10)
      Exit Function
   End If
   
   strMens = F_ECFSt1St2(ST1, ST2)
      
   FechaPorta
   F_ECFStatus = UCase(F_ECFStatus & strMens)
   If F_Vazio(F_ECFStatus) Then
      F_ECFStatus = "TUDO OK!"
   End If
End Function

Function F_DataUltReducaoZ() As String
   'Retorna a data e hora da última redução Z executada
   Dim strExpr As String
   F_DataUltReducaoZ = ""
   
   'Abrindo comunicação com a impressora
   If Not F_AbrePortaECF Then
      Exit Function
   End If
   
   'Busca a data da última redução fiscal
   F_ComandoECF (conAbre_Protocolo_Fiscal & conINFVariaveis & _
   conVARData_Reducao & conFecha_Protocolo_Fiscal)
   
   If Not F_FechaPortaECF Then
      Exit Function
   End If
   
   strExpr = F_RetornoArquivo
   
   If F_Vazio(strExpr) Then
      Exit Function
   End If
   
   F_DataUltReducaoZ = Format(Left(strExpr, 2) & "/" & _
   Mid(strExpr, 3, 2) & "/" & Mid(strExpr, 5, 2) & " " & _
   Mid(strExpr, 7, 2) & ":" & Mid(strExpr, 9, 2) & ":" & _
   Right(strExpr, 2), "DD/MM/YYYY HH:MM:SS")
   
End Function

Function F_ECFData() As String
   'Retorna a data e hora da impressora fiscal
   
   Dim strExpr As String
   F_ECFData = ""
   
   'Abrindo porta comunicação com a impressora
   If Not F_AbrePortaECF Then
      Exit Function
   End If
   
   'Busca a data da última redução fiscal
   
   F_ComandoECF (conAbre_Protocolo_Fiscal & _
   conINFVariaveis & conVARData_Hora & conFecha_Protocolo_Fiscal)
      
   
   If Not F_FechaPortaECF Then
      Exit Function
   End If
   
   strExpr = F_RetornoArquivo
   
   If F_Vazio(strExpr) Then
      Exit Function
   End If
   
   F_ECFData = Format(Left(strExpr, 2) & "/" & _
   Mid(strExpr, 3, 2) & "/" & Mid(strExpr, 5, 2) & " " & _
   Mid(strExpr, 7, 2) & ":" & Mid(strExpr, 9, 2) & ":" & _
   Right(strExpr, 2), "DD/MM/YYYY HH:MM:SS")

End Function

Function F_ECFInformacoesFiscais() As String
    '   Esta função retorna os flags fiscais
    Dim strExpr As String
    Dim intRet As Integer
    Dim strLinha As String
    
    strLinha = Chr(13) & Chr(10)
    
    F_ECFInformacoesFiscais = ""
    If Not F_AbrePortaECF Then
       'Não conseguiu comunicação com a impressora
       F_ECFInformacoesFiscais = _
       "Não há comunicação com a impressora!" & Chr(13)
    End If
    
    strExpr = conAbre_Protocolo_Fiscal
    strExpr = strExpr & conINFVariaveis & conVARFlagsFiscais
    strExpr = strExpr & conFecha_Protocolo_Fiscal
    
    If Not F_ComandoECF(strExpr) Then
       F_ECFInformacoesFiscais = F_ECFInformacoesFiscais & _
       "Não executou o comando!"
    End If
    If Not F_FechaPortaECF Then
       Exit Function
    End If
    
    strExpr = F_RetornoArquivo
    
    intRet = CCur(Mid(strExpr, 2))
    strExpr = ""
    If intRet >= 128 Then
       strExpr = strExpr & "Memória fiscal esgotada!" & strLinha
       intRet = intRet - 128
    Else
       strExpr = strExpr & "Memória fiscal normal!" & strLinha
    End If
    If intRet >= 32 Then
      strExpr = strExpr & "É possível cancelar o último cupom!" & strLinha
      intRet = intRet - 32
    Else
      strExpr = strExpr & "Não é possível cancelar cupom!" & strLinha
    End If
    If intRet >= 8 Then
      strExpr = strExpr & "Já houve redução Z no dia!" & strLinha
      intRet = intRet - 8
    Else
      strExpr = strExpr & "Redução Z não efetuada ainda!" & strLinha
    End If
    If intRet >= 4 Then
       strExpr = strExpr & "Horário de verão!" & strLinha
       intRet = intRet - 4
    Else
       strExpr = strExpr & "Horário normal!" & strLinha
    End If
    If intRet >= 2 Then
       strExpr = strExpr & "Fechamento com formas de pagamento iniciado!" & strLinha
       intRet = intRet - 2
    Else
       strExpr = strExpr & "Cupom sem iniciar formas de pagamento!" & strLinha
    End If
    If intRet >= 1 Then
       strExpr = strExpr & "Cupom fiscal aberto!" & strLinha
       intRet = intRet - 1
    Else
       strExpr = strExpr & "Cupom fiscal fechado!" & strLinha
    End If
    
    F_ECFInformacoesFiscais = UCase(strExpr)
End Function
Function F_ECFHorarioVerao() As String
    '    Esta função retorna verdadeiro se a impressora fiscal está
    'setada para o horário de verão
    '    Retorna: "V" - Horário de Verão  "N" - Horário Normal
    
    Dim blnVerao As Boolean
    Dim strExpr As String
    Dim strDataECF As String
    Dim Ret As Integer, ACK As Integer, ST1 As Integer, ST2 As Integer
     
    strDataECF = Format(F_ECFData, "DD/MM/YYYY")
    F_ECFHorarioVerao = False
    If Not F_AbrePortaECF Then
       'Não conseguiu comunicação com a impressora
       F_ECFHorarioVerao = ""  'Retorna Vazio Erro
    End If
    
    'Verifica se está no horário de verão
    blnVerao = InStr(1, F_ECFInformacoesFiscais, "VERÃO") <> 0
    
    
    'Passa informações sobre o horário de verão:
    strExpr = "A impressora fiscal está funcionando no Horário "
    strExpr = strExpr & IIf(blnVerao, "de Verão!", "Normal!") & Chr(13)
    strExpr = strExpr & "Para poder continuar a operação:" & Chr(13)
    strExpr = strExpr & "Não pode ter ocorrido movimento no dia;" & Chr(13)
    strExpr = strExpr & "Ou logo após a Redução Z, se estiver "
    strExpr = strExpr & "saindo do horário normal;" & Chr(13)
    strExpr = strExpr & "Ou Após uma hora efetuada a Redução Z, "
    strExpr = strExpr & "se estiver saindo do horário de verão!"
    strExpr = strExpr & Chr(13) & Chr(13)
    strExpr = strExpr & F_Centra("D E S E J A   C O N T I N U A R ? ", 80)
    
    Beep
    Beep
    Beep
    If MsgBox(strExpr, vbQuestion + vbYesNo) <> vbYes Then
       'Não foi efetuado o horário de verão
       Exit Function
    End If
    
    If blnVerao And Month(strDataECF) > 10 And Month(strDataECF) < 2 Then
       'Está no horário de verão e o mês está Outubro e Fevereiro
       
       strExpr = F_MesExtenso(strDataECF) & " é um mês de horário de verão!"
       strExpr = strExpr & Chr(13) & "Não é permitido alterar!"
       MsgBox strExpr, vbExclamation
       Exit Function
    End If
    
    If Not blnVerao And Month(strDataECF) > 2 And Month(strDataECF) < 10 Then
       'Está no horário normal e o mês está entre Fevereiro e Outubro
       strExpr = F_MesExtenso(strDataECF) & " é um mês de horário normal!"
       strExpr = strExpr & Chr(13) & "Não é permitido alterar!"
       MsgBox strExpr, vbExclamation
       Exit Function
    End If
    
    If Not F_AbrePortaECF Then
       Exit Function
    End If
    
    strExpr = conAbre_Protocolo_Fiscal
    strExpr = strExpr & conCMDPro_Horario_Verao
    strExpr = strExpr & conFecha_Protocolo_Fiscal
    
    If Not F_ComandoECF(strExpr) Then
       Exit Function
    End If
    Ret = Status_Mp20FI(Ret, ACK, ST1, ST2)
    If Not F_Vazio(F_ECFSt1St2(ST1, ST2)) Then
       MsgBox F_ECFSt1St2(ST1, ST2)
    End If
    If Not F_FechaPortaECF() Then
       Exit Function
    End If
    F_ECFHorarioVerao = True
    
    
    Exit Function
    
TrataErro:
    P_MostraErro
End Function


Function F_ECFSt1St2(ST1, ST2) As String
   'Esta função Retorna as descrições de ST1 e ST2
   Dim strMens As String
   
   'Verificando os sets de ST1
   
   strMens = ""
   If ST1 >= 128 Then
      strMens = strMens & "O Papel acabou!" & Chr(13) & Chr(10)
      ST1 = ST1 - 128
   End If
   If ST1 >= 64 Then
      strMens = strMens & "Pouco papel!" & Chr(13) & Chr(10)
      'frmVenda.lblMensagem.caption = "Pouco Papel"
      ST1 = ST1 - 64
   End If
   If ST1 >= 32 Then
      strMens = strMens & "Erro no relógio!" & Chr(13) & Chr(10)
      ST1 = ST1 - 32
   End If
   If ST1 >= 16 Then
         strMens = strMens & "Impressora em erro!" & Chr(13) & Chr(10)
         ST1 = ST1 - 32
   End If
   If ST1 >= 8 Then
         strMens = strMens & "Primeiro dado do comando não é ESC" & Chr(13) & Chr(10)
         ST1 = ST1 - 8
   End If
   If ST1 >= 4 Then
         strMens = strMens & "Comando inexistente!" & Chr(13) & Chr(10)
         ST1 = ST1 - 4
   End If
   If ST1 >= 2 Then
         strMens = strMens & "Cupom aberto!" & Chr(13) & Chr(10)
         ST1 = ST1 - 2
   End If
   If ST1 >= 1 Then
         strMens = strMens & "Número de parâmetros inválidos" & Chr(13) & Chr(10)
         ST1 = ST1 - 1
   End If
   
   
   If ST2 >= 128 Then
      strMens = strMens & "Tipo de parâmetro inválido!" & Chr(13) & Chr(10)
      ST2 = ST2 - 128
   End If
   If ST2 >= 64 Then
      strMens = strMens & "Memória fiscal lotada!" & Chr(13) & Chr(10)
      ST2 = ST2 - 64
   End If
   If ST2 >= 32 Then
      strMens = strMens & "Erro na memória RAM CMOS não volátil!" & Chr(13) & Chr(10)
      ST2 = ST2 - 32
   End If
   If ST2 >= 16 Then
      strMens = strMens & "Alíquota não programada!" & Chr(13) & Chr(10)
      ST2 = ST2 - 32
   End If
   If ST2 >= 8 Then
      strMens = strMens & "Capacidade de alíquotas prgramáveis lotada!" & Chr(13) & Chr(10)
      ST2 = ST2 - 8
   End If
   If ST2 >= 4 Then
      strMens = strMens & "Cancelamento não permitido!" & Chr(13) & Chr(10)
      ST2 = ST2 - 4
   End If
   If ST2 >= 2 Then
      strMens = strMens & "CGC/IE do proprietário não programados!" & Chr(13) & Chr(10)
      ST2 = ST2 - 2
   End If
   If ST2 >= 1 Then
      strMens = strMens & "Comando não executado" & Chr(13) & Chr(10)
      ST2 = ST2 - 1
   End If

F_ECFSt1St2 = strMens
End Function

Sub P_PortaComunicacaoECF()
    Dim intCont As Integer
    
    If MsgBox("Localizar impressora fiscal?", _
    vbQuestion + vbYesNo) = vbYes Then
        P_PreencheECFProtocolo
        For intCont = 1 To 4
            frmSplash.lblAviso.Refresh
            strPorta_Com_Fscl = "Com" & Format(intCont)
            IniPortaStr (strPorta_Com_Fscl) 'Abre a comunicação com o ECF
        
            frmSplash.lblAviso.caption = "Localizando ECF em " & _
            strPorta_Com_Fscl
        
            frmSplash.lblAviso.Refresh
            If FormataTX(Chr(27) & "|35|23|" & Chr(27)) = 0 Then
                'Achou uma porta com comunicação positiva
           
                FechaPorta 'Fecha a comunicação com o ECF
                Kill "Mp20fi.Ret" 'Apaga o arquivo gerado
           
                MsgBox "A Impressora Fiscal encontra-se em " & _
                strPorta_Com_Fscl & "!", vbInformation
           
                Exit Sub
             
             End If
        Next
        'O computador não está com ECF conectado
        MsgBox "Não há impressora fiscal conectada!", vbInformation
    End If
    
    'Zera a variável de comunicação com a impressora fiscal
    strPorta_Com_Fscl = ""
End Sub

Function F_ECFSimboloMoeda() As String
    'Esta função retorna o símbolo da moeda corrente no ECF
    Dim strMens As String
    
    F_ECFSimboloMoeda = ""
    If Not F_AbrePortaECF Then
        Exit Function
    End If
    
    strMens = conAbre_Protocolo_Fiscal
    strMens = strMens & conINFVariaveis & conVARMoeda & conFecha_Protocolo_Fiscal
    
    F_ComandoECF (strMens)
    
    strMens = F_RetornoArquivo
    
    F_FechaPortaECF
    F_ECFSimboloMoeda = strMens
End Function
Function F_Trunca(strExpressao, strCar As String) As String

Dim intgrCont As Integer
Dim strngPedaco As String
'Car Maior que Expressao
If Len(strExpressao) < Len(strCar) Then
   F_Trunca = strExpressao
   Exit Function
End If
intgrCont = 1
Do While intgrCont <= Len(strExpressao)
    strngPedaco = Mid(strExpressao, intgrCont, Len(strCar))
    If strngPedaco = strCar Then
       intgrCont = intgrCont + Len(strCar)
    Else
       intgrCont = intgrCont + 1
       F_Trunca = F_Trunca & strngPedaco
    End If
Loop
End Function

Function F_ConvDOS(strExpr As String) As String
    'Função converte os acentos do DOS para Windows
    Dim Cont As Integer
    F_ConvDOS = strExpr
    For Cont = 1 To Len(strExpr)
        If Mid(strExpr, Cont, 1) = Chr(128) Then
           Mid(F_ConvDOS, Cont, 1) = "Ç"
        ElseIf Mid(F_ConvDOS, Cont, 1) = Chr(129) Then
           Mid(F_ConvDOS, Cont, 1) = "ü"
        ElseIf Mid(strExpr, Cont, 1) = Chr(130) Then
           Mid(F_ConvDOS, Cont, 1) = "é"
        ElseIf Mid(strExpr, Cont, 1) = Chr(131) Then
           Mid(F_ConvDOS, Cont, 1) = "â"
        ElseIf Mid(strExpr, Cont, 1) = Chr(132) Then
           Mid(F_ConvDOS, Cont, 1) = "ä"
        ElseIf Mid(strExpr, Cont, 1) = Chr(133) Then
           Mid(F_ConvDOS, Cont, 1) = "à"
        ElseIf Mid(strExpr, Cont, 1) = Chr(134) Then
           Mid(F_ConvDOS, Cont, 1) = "å"
        ElseIf Mid(strExpr, Cont, 1) = Chr(135) Then
           Mid(F_ConvDOS, Cont, 1) = "ç"
        ElseIf Mid(strExpr, Cont, 1) = Chr(136) Then
           Mid(F_ConvDOS, Cont, 1) = "ê"
        ElseIf Mid(strExpr, Cont, 1) = Chr(137) Then
           Mid(F_ConvDOS, Cont, 1) = "ë"
        ElseIf Mid(strExpr, Cont, 1) = Chr(138) Then
           Mid(F_ConvDOS, Cont, 1) = "è"
        ElseIf Mid(strExpr, Cont, 1) = Chr(139) Then
           Mid(F_ConvDOS, Cont, 1) = "ï"
        ElseIf Mid(strExpr, Cont, 1) = Chr(140) Then
           Mid(F_ConvDOS, Cont, 1) = "î"
        ElseIf Mid(strExpr, Cont, 1) = Chr(141) Then
           Mid(F_ConvDOS, Cont, 1) = "ì"
        ElseIf Mid(strExpr, Cont, 1) = Chr(142) Then
           Mid(F_ConvDOS, Cont, 1) = "Ä"
        ElseIf Mid(strExpr, Cont, 1) = Chr(143) Then
           Mid(F_ConvDOS, Cont, 1) = "Å"
        ElseIf Mid(strExpr, Cont, 1) = Chr(144) Then
           Mid(F_ConvDOS, Cont, 1) = "É"
        ElseIf Mid(strExpr, Cont, 1) = Chr(145) Then
           Mid(F_ConvDOS, Cont, 1) = "æ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(146) Then
           Mid(F_ConvDOS, Cont, 1) = "Æ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(147) Then
           Mid(F_ConvDOS, Cont, 1) = "ô"
        ElseIf Mid(strExpr, Cont, 1) = Chr(148) Then
           Mid(F_ConvDOS, Cont, 1) = "ö"
        ElseIf Mid(strExpr, Cont, 1) = Chr(149) Then
           Mid(F_ConvDOS, Cont, 1) = "ò"
        ElseIf Mid(strExpr, Cont, 1) = Chr(150) Then
           Mid(F_ConvDOS, Cont, 1) = "û"
        ElseIf Mid(strExpr, Cont, 1) = Chr(151) Then
           Mid(F_ConvDOS, Cont, 1) = "ù"
        ElseIf Mid(strExpr, Cont, 1) = Chr(152) Then
           Mid(F_ConvDOS, Cont, 1) = "ÿ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(153) Then
           Mid(F_ConvDOS, Cont, 1) = "Ö"
        ElseIf Mid(strExpr, Cont, 1) = Chr(154) Then
           Mid(F_ConvDOS, Cont, 1) = "Ü"
        ElseIf Mid(strExpr, Cont, 1) = Chr(155) Then
           Mid(F_ConvDOS, Cont, 1) = "ø"
        ElseIf Mid(strExpr, Cont, 1) = Chr(156) Then
           Mid(F_ConvDOS, Cont, 1) = "£"
        ElseIf Mid(strExpr, Cont, 1) = Chr(157) Then
           Mid(F_ConvDOS, Cont, 1) = "Ø"
        ElseIf Mid(strExpr, Cont, 1) = Chr(158) Then
           Mid(F_ConvDOS, Cont, 1) = "×"
        ElseIf Mid(strExpr, Cont, 1) = Chr(159) Then
           Mid(F_ConvDOS, Cont, 1) = "ƒ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(160) Then
           Mid(F_ConvDOS, Cont, 1) = "á"
        ElseIf Mid(strExpr, Cont, 1) = Chr(161) Then
           Mid(F_ConvDOS, Cont, 1) = "í"
        ElseIf Mid(strExpr, Cont, 1) = Chr(162) Then
           Mid(F_ConvDOS, Cont, 1) = "ó"
        ElseIf Mid(strExpr, Cont, 1) = Chr(163) Then
           Mid(F_ConvDOS, Cont, 1) = "ú"
        ElseIf Mid(strExpr, Cont, 1) = Chr(164) Then
           Mid(F_ConvDOS, Cont, 1) = "ñ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(165) Then
           Mid(F_ConvDOS, Cont, 1) = "Ñ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(166) Then
           Mid(F_ConvDOS, Cont, 1) = "ª"
        ElseIf Mid(strExpr, Cont, 1) = Chr(167) Then
           Mid(F_ConvDOS, Cont, 1) = "º"
        ElseIf Mid(strExpr, Cont, 1) = Chr(168) Then
           Mid(F_ConvDOS, Cont, 1) = "¿"
        ElseIf Mid(strExpr, Cont, 1) = Chr(169) Then
           Mid(F_ConvDOS, Cont, 1) = "®"
        ElseIf Mid(strExpr, Cont, 1) = Chr(170) Then
           Mid(F_ConvDOS, Cont, 1) = "¬"
        ElseIf Mid(strExpr, Cont, 1) = Chr(171) Then
           Mid(F_ConvDOS, Cont, 1) = "½"
        ElseIf Mid(strExpr, Cont, 1) = Chr(172) Then
           Mid(F_ConvDOS, Cont, 1) = "¼"
        ElseIf Mid(strExpr, Cont, 1) = Chr(173) Then
           Mid(F_ConvDOS, Cont, 1) = "¡"
        ElseIf Mid(strExpr, Cont, 1) = Chr(174) Then
           Mid(F_ConvDOS, Cont, 1) = "«"
        ElseIf Mid(strExpr, Cont, 1) = Chr(175) Then
           Mid(F_ConvDOS, Cont, 1) = "»"
        ElseIf Mid(strExpr, Cont, 1) = Chr(176) Then
           Mid(F_ConvDOS, Cont, 1) = "°"
        ElseIf Mid(strExpr, Cont, 1) = Chr(177) Then
           Mid(F_ConvDOS, Cont, 1) = "±"
        ElseIf Mid(strExpr, Cont, 1) = Chr(178) Then
           Mid(F_ConvDOS, Cont, 1) = "²"
        ElseIf Mid(strExpr, Cont, 1) = Chr(179) Then
           Mid(F_ConvDOS, Cont, 1) = "¦"
        ElseIf Mid(strExpr, Cont, 1) = Chr(180) Then
           Mid(F_ConvDOS, Cont, 1) = "´"
        ElseIf Mid(strExpr, Cont, 1) = Chr(181) Then
           Mid(F_ConvDOS, Cont, 1) = "Á"
        ElseIf Mid(strExpr, Cont, 1) = Chr(182) Then
           Mid(F_ConvDOS, Cont, 1) = "Â"
        ElseIf Mid(strExpr, Cont, 1) = Chr(183) Then
           Mid(F_ConvDOS, Cont, 1) = "À"
        ElseIf Mid(strExpr, Cont, 1) = Chr(184) Then
           Mid(F_ConvDOS, Cont, 1) = "©"
        ElseIf Mid(strExpr, Cont, 1) = Chr(185) Then
           Mid(F_ConvDOS, Cont, 1) = "¹"
        ElseIf Mid(strExpr, Cont, 1) = Chr(186) Then
           Mid(F_ConvDOS, Cont, 1) = "º"
        ElseIf Mid(strExpr, Cont, 1) = Chr(187) Then
           Mid(F_ConvDOS, Cont, 1) = "»"
        ElseIf Mid(strExpr, Cont, 1) = Chr(188) Then
           Mid(F_ConvDOS, Cont, 1) = "¼"
        ElseIf Mid(strExpr, Cont, 1) = Chr(189) Then
           Mid(F_ConvDOS, Cont, 1) = "¢"
        ElseIf Mid(strExpr, Cont, 1) = Chr(190) Then
           Mid(F_ConvDOS, Cont, 1) = "¥"
        ElseIf Mid(strExpr, Cont, 1) = Chr(191) Then
           Mid(F_ConvDOS, Cont, 1) = "¿"
        ElseIf Mid(strExpr, Cont, 1) = Chr(192) Then
           Mid(F_ConvDOS, Cont, 1) = "À"
        ElseIf Mid(strExpr, Cont, 1) = Chr(193) Then
           Mid(F_ConvDOS, Cont, 1) = "Á"
        ElseIf Mid(strExpr, Cont, 1) = Chr(194) Then
           Mid(F_ConvDOS, Cont, 1) = "Ö"
        ElseIf Mid(strExpr, Cont, 1) = Chr(195) Then
           Mid(F_ConvDOS, Cont, 1) = "+"
        ElseIf Mid(strExpr, Cont, 1) = Chr(196) Then
           Mid(F_ConvDOS, Cont, 1) = "-"
        ElseIf Mid(strExpr, Cont, 1) = Chr(197) Then
           Mid(F_ConvDOS, Cont, 1) = "+"
        ElseIf Mid(strExpr, Cont, 1) = Chr(198) Then
           Mid(F_ConvDOS, Cont, 1) = "ã"
        ElseIf Mid(strExpr, Cont, 1) = Chr(199) Then
           Mid(F_ConvDOS, Cont, 1) = "Ã"
        ElseIf Mid(strExpr, Cont, 1) = Chr(200) Then
           Mid(F_ConvDOS, Cont, 1) = "+"
        ElseIf Mid(strExpr, Cont, 1) = Chr(201) Then
           Mid(F_ConvDOS, Cont, 1) = "+"
        ElseIf Mid(strExpr, Cont, 1) = Chr(202) Then
           Mid(F_ConvDOS, Cont, 1) = "-"
        ElseIf Mid(strExpr, Cont, 1) = Chr(203) Then
           Mid(F_ConvDOS, Cont, 1) = "-"
        ElseIf Mid(strExpr, Cont, 1) = Chr(204) Then
           Mid(F_ConvDOS, Cont, 1) = "¦"
        ElseIf Mid(strExpr, Cont, 1) = Chr(205) Then
           Mid(F_ConvDOS, Cont, 1) = "-"
        ElseIf Mid(strExpr, Cont, 1) = Chr(206) Then
           Mid(F_ConvDOS, Cont, 1) = "+"
        ElseIf Mid(strExpr, Cont, 1) = Chr(207) Then
           Mid(F_ConvDOS, Cont, 1) = "¤"
        ElseIf Mid(strExpr, Cont, 1) = Chr(208) Then
           Mid(F_ConvDOS, Cont, 1) = "ð"
        ElseIf Mid(strExpr, Cont, 1) = Chr(209) Then
           Mid(F_ConvDOS, Cont, 1) = "Ð"
        ElseIf Mid(strExpr, Cont, 1) = Chr(210) Then
           Mid(F_ConvDOS, Cont, 1) = "Ê"
        ElseIf Mid(strExpr, Cont, 1) = Chr(211) Then
           Mid(F_ConvDOS, Cont, 1) = "Ë"
        ElseIf Mid(strExpr, Cont, 1) = Chr(212) Then
           Mid(F_ConvDOS, Cont, 1) = "È"
        ElseIf Mid(strExpr, Cont, 1) = Chr(213) Then
           Mid(F_ConvDOS, Cont, 1) = "i"
        ElseIf Mid(strExpr, Cont, 1) = Chr(214) Then
           Mid(F_ConvDOS, Cont, 1) = "Í"
        ElseIf Mid(strExpr, Cont, 1) = Chr(215) Then
           Mid(F_ConvDOS, Cont, 1) = "Î"
        ElseIf Mid(strExpr, Cont, 1) = Chr(216) Then
           Mid(F_ConvDOS, Cont, 1) = "Ï"
        ElseIf Mid(strExpr, Cont, 1) = Chr(217) Then
           Mid(F_ConvDOS, Cont, 1) = "+"
        ElseIf Mid(strExpr, Cont, 1) = Chr(218) Then
           Mid(F_ConvDOS, Cont, 1) = "+"
        ElseIf Mid(strExpr, Cont, 1) = Chr(219) Then
           Mid(F_ConvDOS, Cont, 1) = "_"
        ElseIf Mid(strExpr, Cont, 1) = Chr(220) Then
           Mid(F_ConvDOS, Cont, 1) = "_"
        ElseIf Mid(strExpr, Cont, 1) = Chr(221) Then
           Mid(F_ConvDOS, Cont, 1) = "¦"
        ElseIf Mid(strExpr, Cont, 1) = Chr(222) Then
           Mid(F_ConvDOS, Cont, 1) = "Ì"
        ElseIf Mid(strExpr, Cont, 1) = Chr(223) Then
           Mid(F_ConvDOS, Cont, 1) = "_"
        ElseIf Mid(strExpr, Cont, 1) = Chr(224) Then
           Mid(F_ConvDOS, Cont, 1) = "Ó"
        ElseIf Mid(strExpr, Cont, 1) = Chr(225) Then
           Mid(F_ConvDOS, Cont, 1) = "ß"
        ElseIf Mid(strExpr, Cont, 1) = Chr(226) Then
           Mid(F_ConvDOS, Cont, 1) = "Ô"
        ElseIf Mid(strExpr, Cont, 1) = Chr(227) Then
           Mid(F_ConvDOS, Cont, 1) = "Ò"
        ElseIf Mid(strExpr, Cont, 1) = Chr(228) Then
           Mid(F_ConvDOS, Cont, 1) = "õ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(229) Then
           Mid(F_ConvDOS, Cont, 1) = "Õ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(230) Then
           Mid(F_ConvDOS, Cont, 1) = "µ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(231) Then
           Mid(F_ConvDOS, Cont, 1) = "þ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(232) Then
           Mid(F_ConvDOS, Cont, 1) = "Þ"
        ElseIf Mid(strExpr, Cont, 1) = Chr(233) Then
           Mid(F_ConvDOS, Cont, 1) = "Ú"
        ElseIf Mid(strExpr, Cont, 1) = Chr(234) Then
           Mid(F_ConvDOS, Cont, 1) = "Û"
        ElseIf Mid(strExpr, Cont, 1) = Chr(235) Then
           Mid(F_ConvDOS, Cont, 1) = "Ù"
        ElseIf Mid(strExpr, Cont, 1) = Chr(236) Then
           Mid(F_ConvDOS, Cont, 1) = "ý"
        ElseIf Mid(strExpr, Cont, 1) = Chr(237) Then
           Mid(F_ConvDOS, Cont, 1) = "Ý"
        ElseIf Mid(strExpr, Cont, 1) = Chr(238) Then
           Mid(F_ConvDOS, Cont, 1) = "¯"
        ElseIf Mid(strExpr, Cont, 1) = Chr(239) Then
           Mid(F_ConvDOS, Cont, 1) = "´"
        ElseIf Mid(strExpr, Cont, 1) = Chr(240) Then
           Mid(F_ConvDOS, Cont, 1) = "­"
        ElseIf Mid(strExpr, Cont, 1) = Chr(241) Then
           Mid(F_ConvDOS, Cont, 1) = "±"
        ElseIf Mid(strExpr, Cont, 1) = Chr(242) Then
           Mid(F_ConvDOS, Cont, 1) = "_"
        ElseIf Mid(strExpr, Cont, 1) = Chr(243) Then
           Mid(F_ConvDOS, Cont, 1) = "¾"
        ElseIf Mid(strExpr, Cont, 1) = Chr(244) Then
           Mid(F_ConvDOS, Cont, 1) = "¶"
        ElseIf Mid(strExpr, Cont, 1) = Chr(245) Then
           Mid(F_ConvDOS, Cont, 1) = "§"
        ElseIf Mid(strExpr, Cont, 1) = Chr(246) Then
           Mid(F_ConvDOS, Cont, 1) = "÷"
        ElseIf Mid(strExpr, Cont, 1) = Chr(247) Then
           Mid(F_ConvDOS, Cont, 1) = "¸"
        ElseIf Mid(strExpr, Cont, 1) = Chr(248) Then
           Mid(F_ConvDOS, Cont, 1) = "°"
        ElseIf Mid(strExpr, Cont, 1) = Chr(249) Then
           Mid(F_ConvDOS, Cont, 1) = "¨"
        ElseIf Mid(strExpr, Cont, 1) = Chr(250) Then
           Mid(F_ConvDOS, Cont, 1) = "·"
        ElseIf Mid(strExpr, Cont, 1) = Chr(251) Then
           Mid(F_ConvDOS, Cont, 1) = "³"
        ElseIf Mid(strExpr, Cont, 1) = Chr(253) Then
           Mid(F_ConvDOS, Cont, 1) = "²"
        ElseIf Mid(strExpr, Cont, 1) = Chr(254) Then
           Mid(F_ConvDOS, Cont, 1) = "_"
        ElseIf Mid(strExpr, Cont, 1) = Chr(255) Then
           Mid(F_ConvDOS, Cont, 1) = " "
        End If
    Next
End Function

Sub P_PreencheECFProtocolo()
    ' este procedimento preenche os protocolos de entrada e saida da ECF
    conFecha_Protocolo_Fiscal = "|" & Chr(27)
    conAbre_Protocolo_Fiscal = Chr(27) & "|"
End Sub

