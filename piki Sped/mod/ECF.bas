Attribute VB_Name = "ECF"
'este módulo serve para colocar os procedimentos
'padrões de chamadas de ecf das impressoras fiscais
Public ecf_cupomImpressoParcial As Boolean 'caso o cupom já tenha sido imprimido na transação TEF,
                            'estará verdadeiro. Quando termina o CUPOM, poe vazio.
Type ultReducao
    modoReducao As String
    GrandeTotal As String
    Cancelamentos As String
    Descontos As String
    Tributos As String
    Totalizadores_Parciais_Tributados As String
    Sangria As String
    Suprimentos As String
    Totalizadores_nao_sujeitos_ao_ICMS As String
    Contadores_dos_TPs_nao_sujeitos_ao_ICMS As String
    Contador_de_ordem_de_operacao As String
    Contador_de_operacoes_nao_sujeitas_ao_ICMS As String
    Numero_de_Aliquotas_cadastradas As String
    Data_do_movimento As String
    Acrescimos As String
    Acrescimo_financeiro As String
    
End Type

Sub ecf_LeituraX()
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_LeituraX
    Else
        i = Bematech_FI_LeituraX
    End If
    If i <> 1 Then
        MsgBox "Erro na impressão da leitura X", vbInformation
    End If
End Sub

Function ecf_CancelaCupom() As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_CancelaCupom()
        
        If i <> 1 Then
            Erro_Daruma i
        End If
        
    Else
        i = Bematech_FI_CancelaCupom()
        If i <> 0 Then
           'F_RetornoVariaveis
        End If
    End If
    ecf_CancelaCupom = i
End Function
Sub ecf_ResetaImpressora()
    Dim i As Integer
    Dim iAck As Integer, ist1 As Integer, ist2 As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_RetornoImpressora(iAck, ist1, ist2)
    
    Else
        i = Bematech_FI_LeituraX
    End If
    If i <> 1 Then
        MsgBox "Erro na execução do comando", vbInformation
    End If
End Sub
Sub ecf_MemoriaFiscal(sDataIni As String, sDataFin As String)
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_LeituraMemoriaFiscalData(sDataIni, sDataFin)
        If i <> 1 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_LeituraMemoriaFiscalData(sDataIni, sDataFin)
    End If
    
End Sub
Sub ecf_ReducaoZ(data As String)
    Dim i As Integer
    Dim d As DataInicialFinal
    d = F_BuscaDataInicial(Date)
    If d.dataI = "" Then
        Exit Sub
    End If
    
    If strImprFiscal = "D" Then
        i = Daruma_FI_ReducaoZ(Format(d.dataI, "ddmmyy"), Format(Time, "hhmmss"))
        If i <> 0 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_ReducaoZ(Format(d.dataI, "ddmmyy"), Format(Time, "hhmmss"))
        If i <> 0 Then
            
        End If
    
    End If
End Sub



Sub Erro_Daruma(iRet As Integer)
    Dim i As Integer
    Dim strErro As String
    strErro = Space(4)
    
    i = Daruma_FI_RetornaErroExtendido(strErro)
    strErro = str(strErro)
If strErro = 0 Then
    MsgBox "00 IF em modo Manutenção. Foi ligada sem o Jumper de Operação"
ElseIf strErro = 1 Then
    MsgBox "01 Método disponível somente em modo Manutenção"
ElseIf strErro = 2 Then
    MsgBox "02 Erro durante a gravação da Memória Fiscal"
ElseIf strErro = 3 Then
    MsgBox "03 Memória Fiscal esgotada"
ElseIf strErro = 4 Then
    MsgBox "04 Erro no relógio interno da IF"
ElseIf strErro = 5 Then
    MsgBox "05 Falha mecânica na IF"
ElseIf strErro = 6 Then
    MsgBox "06 Erro durante a leitura da Memória Fiscal"
ElseIf strErro = 10 Then
    MsgBox "10 Documento sendo emitido"
ElseIf strErro = 11 Then
    MsgBox "11 Documento não foi aberto"
    
ElseIf strErro = 12 Then
    MsgBox "12 Não existe documento a cancelar"
ElseIf strErro = 13 Then
    MsgBox "13 Dígito não numérico não esperado foi encontrado nos Parâmetros do Método:s"
ElseIf strErro = 14 Then
    MsgBox "14 Não há mais memória disponível para esta operação"
ElseIf strErro = 15 Then
    MsgBox "15 Item a cancelar não foi encontrado"
ElseIf strErro = 16 Then
    MsgBox "16 Erro de sintaxe no método"
ElseIf strErro = 17 Then
    MsgBox "17 'Estouro' de capacidade numérica (overflow)"
ElseIf strErro = 18 Then
    MsgBox "18 Selecionado totalizador tributado com alíquota de imposto não definida"
ElseIf strErro = 19 Then
    MsgBox "19 Memória Fiscal vazia"
ElseIf strErro = 20 Then
    MsgBox "20 Não existem campos que requerem atualização"
ElseIf strErro = 21 Then
    MsgBox "21 Detectado proximidade do final da bobina de papel"
ElseIf strErro = 22 Then
    MsgBox "22 Cupom de Redução Z já foi emitido. IF inoperante até 0:00h do próximo dia"
ElseIf strErro = 23 Then
    MsgBox "23 Redução Z do período anterior ainda pendente. IF inoperante"
ElseIf strErro = 24 Then
    MsgBox "24 Valor de desconto ou acréscimo inválido (limitado a 100%)"
ElseIf strErro = 26 Then
    MsgBox "26 Médoto não pode ser executado."
ElseIf strErro = 25 Then
    MsgBox "25 Caráctere inválido foi encontrado nos Parâmetros do Métodos"
ElseIf strErro = 26 Then
    MsgBox "26    Comando não pode ser executado."
ElseIf strErro = 27 Then
    MsgBox "27 Nenhum periférico conectado a interface auxiliar"
ElseIf strErro = 28 Then
    MsgBox "28 Foi encontrado um campo em zero"
ElseIf strErro = 29 Then
    MsgBox "29 Documento anterior não foi Cupom Fiscal. Não pode emitir Cupom Adicional"
ElseIf strErro = 30 Then
    MsgBox "30 Acumulador Não Fiscal selecionado não é válido ou não está disponível"
ElseIf strErro = 31 Then
    MsgBox "31 Não pode autenticar. Excedeu 4 repetições ou não é permitida nesta fase"
ElseIf strErro = 32 Then
    MsgBox "32 Cupom adicional inibido por configuração"
ElseIf strErro = 35 Then
    MsgBox "35 Relógio Interno Inoperante"
ElseIf strErro = 36 Then
    MsgBox "36 Versão do firmware gravada na Memória Fiscal não é a esperada"
ElseIf strErro = 37 Then
    MsgBox "37 Alíquota de imposto informada já está carregada na memória"
ElseIf strErro = 38 Then
    MsgBox "38 Forma de pagamento selecionada não é válida"
ElseIf strErro = 39 Then
    MsgBox "39 Erro na seqüência de fechamento do Cupom Fiscal"
ElseIf strErro = 40 Then
    MsgBox "40 IF em Jornada Fiscal. Alteração da configuração não é permitida"
ElseIf strErro = 41 Then
    MsgBox "41 Data inválida. Data fornecida é inferior à última gravada na Memória Fiscal"
ElseIf strErro = 42 Then
    MsgBox "42 Leitura X inicial ainda não foi emitida"
ElseIf strErro = 43 Then
    MsgBox "43 Não pode emitir Comprovante Vinculado"
ElseIf strErro = 44 Then
    MsgBox "44 Cupom de Orçamento não permitido para este estabelecimento"
ElseIf strErro = 45 Then
    MsgBox "45 Campo obrigatório em branco"
ElseIf strErro = 48 Then
    MsgBox "48 Não pode estornar"
ElseIf strErro = 49 Then
    MsgBox "49 Forma de pagamento indicada não encontrada"
ElseIf strErro = 50 Then
    MsgBox "50 Fim da bobina de papel"
ElseIf strErro = 51 Then
    MsgBox "51 Nenhum usuário cadastrado na MF"
ElseIf strErro = 52 Then
    MsgBox "52 MF não instalada ou não inicializada"
ElseIf strErro = 56 Then
    MsgBox "56 Documento já aberto"
ElseIf strErro = 61 Then
    MsgBox "61 Queda de energia durante a emissão de Cupom Fiscal"
ElseIf strErro = 76 Then
    MsgBox "76 Desconto em ISS não permitido neste ECF (a programação deverá ser feita por meio de intervenção técnica e caso o Estado permita)"
ElseIf strErro = 77 Then
    MsgBox "77 Acréscimo em IOF inibido"
ElseIf strErro = 80 Then
    MsgBox "80 Periférico na interface auxiliar não pode ser reconhecido"
ElseIf strErro = 81 Then
    MsgBox "81 Solicitado preenchimento de cheque de banco desconhecido"
ElseIf strErro = 82 Then
    MsgBox "82 Solicitado transmissão de mensagem nula pela interface auxiliar"
ElseIf strErro = 83 Then
    MsgBox "83 Extenso do cheque não cabe no espaço disponível"
ElseIf strErro = 84 Then
    MsgBox "84 Erro na comunicação com a interface auxiliar"
ElseIf strErro = 85 Then
    MsgBox "85 Erro no dígito verificador durante comunicação com a PertoCheck"
ElseIf strErro = 86 Then
    MsgBox "86 Falha na carga de geometria de folha de cheque"
ElseIf strErro = 87 Then
    MsgBox "87 Parâmetros do Método: invállido para o campo de data do cheque"
    ElseIf strErro = 90 Then
    MsgBox "90 Sequência de validação de número de série inválida"
 End If

End Sub

Sub ecf_HorarioVerao()
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_CfgHorarioVerao("1") '1 pra entrar 0 para sair
        If i <> 0 Then
            
        End If
    Else
        i = Bematech_FI_ProgramaHorarioVerao()
        If i <> 0 Then
            
        End If
    End If
End Sub
Sub ecf_MapaResumo()
    If strImprFiscal = "D" Then
        i = Daruma_FI_MapaResumo()
    Else
        i = Bematech_FI_MapaResumo()
    End If
    Shell "Notepad C:\retorno.txt"
End Sub
Function ecf_NumeroCupom(strNum_Cpm As String) As Integer
    Dim sCupom As String
    Dim i As Integer
    sCupom = Space(6)
    If strImprFiscal = "D" Then
        i = Daruma_FI_NumeroCupom(sCupom)
    Else
        i = Bematech_FI_NumeroCupom(sCupom)
    End If
    ecf_NumeroCupom = i
    strNum_Cpm = sCupom
End Function


Function ecf_MapaResumo1() As Integer
    
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_MapaResumo()
    Else
        i = Bematech_FI_MapaResumo()
    End If
    If i <> 0 Then
        Shell "notepad C:\retorno.txt", vbMaximizedFocus
    End If
    
 
End Function


Function ecf_AbreCupom(strCPF As String) As Integer
    Dim sCupom As String
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_AbreCupom(strCPF)
    Else
        i = Bematech_FI_AbreCupom(strCPF)
    End If
    ecf_AbreCupom = i
 
End Function

Function ecf_VendeItem(cod_pro As String, nome_pro As String, ecf_icms As String, sTipoQtd As String, _
    qtd As String, casas_dec As String, valor_unit As String, _
    tipo_desc As String, desconto As String) As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_VendeItem(cod_pro, Left(nome_pro, 29), ecf_icms, _
        sTipoQtd, qtd, casas_dec, Format(valor_unit, "###,##0.000"), tipo_desc, desconto)
    Else
        i = Bematech_FI_VendeItem(cod_pro, Left(nome_pro, 29), ecf_icms, _
        sTipoQtd, qtd, casas_dec, Format(valor_unit, "###,##0.000"), tipo_desc, desconto)
    End If
    ecf_VendeItem = i
End Function
Function ecf_CancelaItemGenerico(item As String) As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_CancelaItemGenerico(item)
        If i <> 0 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_CancelaItemGenerico(item)
        If i <> 0 Then
           'F_RetornoVariaveis
        End If
    End If
    ecf_CancelaItemGenerico = i
End Function
Function ecf_IniciaFechamentoCupom(acresc_desc As String, tipo_Acresc As String, desc As String) As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_IniciaFechamentoCupom(acresc_desc, tipo_Acresc, desc)
        If i <> 1 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_IniciaFechamentoCupom(acresc_desc, tipo_Acresc, desc)
        If i <> 0 Then
           'F_RetornoVariaveis
        End If
    End If
    ecf_IniciaFechamentoCupom = i
End Function
Function ecf_EfetuaFormaPagamento(formaPag As String, valor As String) As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_EfetuaFormaPagamento(formaPag, valor)
        If i <> 1 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_EfetuaFormaPagamento(formaPag, valor)
        If i <> 0 Then
           'F_RetornoVariaveis
        End If
    End If
    ecf_EfetuaFormaPagamento = i
End Function
Function ecf_TerminaFechamentoCupom(msg As String, Optional respondonde As Integer = -1) As Integer
    Dim i As Integer
    
    
    Retorno = Ecf_ImpressoraNaoResponde
    If Retorno = 0 Then
        respondonde = 0
    ElseIf Retorno = 1 Then
        respondonde = 1
        Exit Function
    End If
    
    sMsg = Left(msg, 320)
    If strImprFiscal = "D" Then
        i = Daruma_FI_TerminaFechamentoCupom(msg)
        If i <> 1 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_TerminaFechamentoCupom(msg)
        If i <> 0 Then
           'F_RetornoVariaveis
        End If
    End If
    ecf_TerminaFechamentoCupom = i
End Function


Function ecf_VerificaImpressoraLigada() As Integer
    Dim iCount As Integer
    If strImprFiscal = "D" Then
        iCount = Daruma_FI_VerificaImpressoraLigada()
    Else
        iCount = Bematech_FI_VerificaImpressoraLigada()
    End If
    ecf_VerificaImpressoraLigada = iCount
End Function
Function ecf_RetornoImpressora(ack As Integer, st1 As Integer, st2 As Integer) As Integer
    If strImprFiscal = "D" Then
        ecf_RetornoImpressora = Daruma_FI_RetornoImpressora(ack, st1, st2)
    Else
        ecf_RetornoImpressora = Bematech_FI_RetornoImpressora(ack, st1, st2)
    End If
     
End Function
Function ecf_verificaRetornoImpressoraMfd(iAck As Integer, ist1 As Integer, ist2 As Integer, ist3 As Integer) As Integer
    Dim iRetorno As Integer
    If strImprFiscal = "D" Then
        iRetorno = Daruma_FI_RetornoImpressoraMFD(iAck, ist1, ist2, ist3)
    Else
        iRetorno = Bematech_FI_RetornoImpressoraMFD(iAck, ist1, ist2, ist3)
    End If
    ecf_verificaRetornoImpressoraMfd = iRetorno
    Select Case ist3
    


        Case 1: MsgBox " Comando INVÁLIDO"
        Case 2:     MsgBox " Erro DESCONHECIDO"
        Case 3: MsgBox " NÚMERO DE PARÂMETRO INVÁLIDO"
        Case 4: MsgBox " TIPO DE PARÂMETRO INVÁLIDO"
        Case 5: MsgBox " TODAS ALÍQUOTAS JÁ PROGRAMADAS"
        Case 6: MsgBox " TOTALIZADOR NÃO FISCAL JÁ PROGRAMADO"
        Case 7: MsgBox " CUPOM FISCAL ABERTO"
        Case 8: MsgBox " CUPOM FISCAL FECHADO"
        Case 9: MsgBox "           ECF OCUPADO"
        Case 10: MsgBox "  IMPRESSORA EM ERRO"
        Case 11: MsgBox "  IMPRESSORA SEM PAPEL"
        Case 12: MsgBox "  IMPRESSORA COM CABEÇA LEVANTADA"
        Case 13: MsgBox "  IMPRESSORA OFF LINE"
        Case 14: MsgBox "  ALÍQUOTA NÃO PROGRAMADA"
        Case 15: MsgBox "  TERMINADOR DE STRING FALTANDO"
        Case 16: MsgBox "  ACRÉSCIMO OU DESCONTO MAIOR QUE O TOTAL DO CUPOM FISCAL"
        Case 17: MsgBox "  CUPOM FISCAL SEM ITEM VENDIDO"
        Case 18: MsgBox "  COMANDO NÃO EFETIVADO"
        Case 19: MsgBox "  SEM ESPAÇO PARA NOVAS FORMAS DE PAGAMENTO"
        Case 20: MsgBox "  FORMA DE PAGAMENTO NÃO PROGRAMADA"
        Case 21: MsgBox "  ÍNDICE MAIOR QUE NÚMERO DE FORMA DE PAGAMENTO"
        Case 22: MsgBox "  FORMAS DE PAGAMENTO ENCERRADAS"
        Case 23: MsgBox "  CUPOM NÃO TOTALIZADO"
        Case 24: MsgBox "  COMANDO MAIOR QUE 7Fh (127d)"
        Case 25: MsgBox "  CUPOM FISCAL ABERTO E SEM ÍTEM"
        Case 26: MsgBox " CANCELAMENTO NÃO IMEDIATAMENTE APÓS"
        Case 27: MsgBox "  CANCELAMENTO JÁ EFETUADO"
        Case 28: MsgBox "  COMPROVANTE DE CRÉDITO OU DÉBITO NÃO PERMITIDO OU JÁ EMITIDO"
        Case 29: MsgBox "  MEIO DE PAGAMENTO NÃO PERMITE TEF"
        Case 30: MsgBox "  SEM COMPROVANTE NÃO FISCAL ABERTO"
        Case 31: MsgBox "  COMPROVANTE DE CRÉDITO OU DÉBITO JÁ ABERTO"
        Case 32: MsgBox "  REIMPRESSÃO NÃO PERMITIDA"
        Case 33: MsgBox "  COMPROVANTE NÃO FISCAL JÁ ABERTO"
        Case 34: MsgBox "  TOTALIZADOR NÃO FISCAL NÃO PROGRAMADO"
        Case 35: MsgBox "  CUPOM NÃO FISCAL SEM ÍTEM VENDIDO"
        Case 36: MsgBox "  ACRÉSCIMO E DESCONTO MAIOR QUE TOTAL CNF"
        Case 37: MsgBox "  MEIO DE PAGAMENTO NÃO INDICADO"
        Case 38: MsgBox "  MEIO DE PAGAMENTO DIFERENTE DO TOTAL DO RECEBIMENTO"
        Case 39: MsgBox "  NÃO PERMITIDO MAIS DE UMA SANGRIA OU SUPRIMENTO"
        Case 40:: MsgBox "  RELATÓRIO GERENCIAL JÁ PROGRAMADO"
        Case 41: MsgBox "  RELATÓRIO GERENCIAL NÃO PROGRAMADO"
        Case 42: MsgBox "  RELATÓRIO GERENCIAL NÃO PERMITIDO"
        Case 43: MsgBox "  MFD NÃO INICIALIZADA"
        Case 44: MsgBox "   MFD AUSENTE"
        Case 45: MsgBox "  MFD SEM NÚMERO DE SÉRIE"
        Case 46: MsgBox "  MFD JÁ INICIALIZADA"
        Case 47:             MsgBox "  MFD LOTADA"
        Case 48: MsgBox "  CUPOM NÃO FISCAL ABERTO"
        Case 49: MsgBox "  MEMÓRIA FISCAL DESCONECTADA"
        Case 50: MsgBox "  MEMÓRIA FISCAL SEM NÚMERO DE SÉRIE DA MFD"
        Case 51: MsgBox "  MEMÓRIA FISCAL LOTADA"
        Case 52: MsgBox "  DATA INICIAL INVÁLIDA"
        Case 53: MsgBox " DATA FINAL INVÁLIDA"
        Case 54: MsgBox "  CONTADOR DE REDUÇÃO Z INICIAL INVÁLIDO"
        Case 55: MsgBox "  CONTADOR DE REDUÇÃO Z FINAL INVÁLIDO"
        Case 56: MsgBox "  ERRO DE ALOCAÇÃO"
        Case 57: MsgBox "  DADOS DO RTC INCORRETOS"
        Case 58: MsgBox "  DATA ANTERIOR AO ÚLTIMO DOCUMENTO EMITIDO"
        Case 59: MsgBox "  FORA DE INTERVENÇÃO TÉCNICA"
        Case 60: MsgBox "  EM INTERVENÇÃO TÉCNICA"
        Case 61: MsgBox "  ERRO NA MEMÓRIA DE TRABALHO"
        Case 62: MsgBox "  JÁ HOUVE MOVIMENTO NO DIA"
        Case 63: MsgBox "  BLOQUEIO POR RZ"
        Case 64: MsgBox "  FORMA DE PAGAMENTO ABERTA"
        Case 65: MsgBox "  AGUARDANDO PRIMEIRO PROPRIETÁRIO"
        Case 66: MsgBox "  AGUARDANDO RZ"
        Case 67: MsgBox "  ECF OU LOJA IGUAL A ZERO"
        Case 68: MsgBox "  CUPOM ADICIONAL NÃO PERMITIDO"
        Case 69: MsgBox "  DESCONTO MAIOR QUE TOTAL VENDIDO EM ICMS"
        Case 70: MsgBox "  RECEBIMENTO NÃO FISCAL NULO NÃO PERMITIDO"
        Case 71: MsgBox " ACRÉSCIMO OU DESCONTO MAIOR QUE TOTAL NÃO FISCAL"
        Case 72: MsgBox " MEMÓRIA FISCAL LOTADA PARA NOVO CARTUCHO"
        Case 73: MsgBox " ERRO DE GRAVAÇÃO NA MF"
        Case 74: MsgBox " ERRO DE GRAVAÇÃO NA MFD"
        Case 75: MsgBox " DADOS DO RTC ANTERIORES AO ÚLTIMO DOC ARMAZENADO"
        Case 76: MsgBox " MEMÓRIA FISCAL SEM ESPAÇO PARA GRAVAR LEITURAS DA MFD"
        Case 77: MsgBox " MEMÓRIA FISCAL SEM ESPAÇO PARA GRAVAR VERSAO DO SB"
        Case 78: MsgBox " DESCRIÇÃO IGUAL A DEFAULT NÃO PERMITIDO"
        Case 79: MsgBox " EXTRAPOLADO NÚMERO DE REPETIÇÕES PERMITIDAS"
        Case 80: MsgBox " SEGUNDA VIA DO COMPROVANTE DE CRÉDITO OU DÉBITO NÃO PERMITIDO"
        Case 81: MsgBox " PARCELAMENTO FORA DA SEQUÊNCIA"
        Case 82: MsgBox " COMPROVANTE DE CRÉDITO OU DÉBITO ABERTO"
        Case 83: MsgBox " TEXTO COM SEQUÊNCIA DE ESC INVÁLIDA"
        Case 84: MsgBox " TEXTO COM SEQUÊNCIA DE ESC INCOMPLETA"
        Case 85: MsgBox " VENDA COM VALOR NULO"
        Case 86: MsgBox " ESTORNO DE VALOR NULO"
        Case 87: MsgBox " FORMA DE PAGAMENTO DIFERENTE DO TOTAL DA SANGRIA"
        Case 88: MsgBox " REDUÇÃO NÃO PERMITIDA EM INTERVENÇÃO TÉCNICA"
        Case 89: MsgBox " AGUARDANDO RZ PARA ENTRADA EM INTERVENÇÃO TÉCNICA"
        Case 90: MsgBox " FORMA DE PAGAMENTO COM VALOR NULO NÃO PERMITIDO"
        Case 91: MsgBox " ACRÉSCIMO E DESCONTO MAIOR QUE VALOR DO ÍTEM"
        Case 92: MsgBox " AUTENTICAÇÃO NÃO PERMITIDA"
        Case 93: MsgBox " TIMEOUT NA VALIDAÇÃO"
        Case 94: MsgBox " COMANDO NÃO EXECUTADO EM IMPRESSORA BILHETE DE PASSAGEM"
        Case 95: MsgBox " COMANDO NÃO EXECUTADO EM IMPRESSORA DE CUPOM FISCAL"
        Case 96: MsgBox " CUPOM NÃO FISCAL FECHADO"
        Case 97: MsgBox " PARÂMETRO NÃO ASCII EM CAMPO ASCII"
        Case 98: MsgBox " PARÂMETRO NÃO ASCII NUMÉRICO EM CAMPO ASCII NUMÉRICO"
        Case 99: MsgBox " TIPO DE TRANSPORTE INVÁLIDO"
        Case 100: MsgBox "  DATA E HORA INVÁLIDA"
        Case 101: MsgBox "  SEM RELATÓRIO GERENCIAL OU COMPROVANTE DE CRÉDITO OU DÉBITO ABERTO"
        Case 102: MsgBox "  NÚMERO DO TOTALIZADOR NÃO FISCAL INVÁLIDO"
        Case 103: MsgBox "  PARÂMETRO DE ACRÉSCIMO OU DESCONTO INVÁLIDO"
        Case 104: MsgBox "  ACRÉSCIMO OU DESCONTO EM SANGRIA OU SUPRIMENTO NÃO PERMITIDO"
        Case 105: MsgBox "  NÚMERO DO RELATÓRIO GERENCIAL INVÁLIDO"
        Case 106: MsgBox "  FORMA DE PAGAMENTO ORIGEM NÃO PROGRAMADA"
        Case 107: MsgBox "  FORMA DE PAGAMENTO DESTINO NÃO PROGRAMADA"
        Case 108: MsgBox "  ESTORNO MAIOR QUE FORMA PAGAMENTO"
        Case 109: MsgBox "  CARACTER NUMÉRICO NA CODIFICAÇÃO GT NÃO PERMITIDO"
        Case 110: MsgBox "  ERRO NA INICIALIZAÇÃO DA MF"
        Case 111: MsgBox "  NOME DO TOTALIZADOR EM BRANCO NÃO PERMITIDO"
        Case 112: MsgBox "  DATA E HORA ANTERIORES AO ÚLTIMO DOC ARMAZENADO"
        Case 113: MsgBox "  PARÂMETRO DE ACRÉSCIMO OU DESCONTO INVÁLIDO"
        Case 114: MsgBox "  ÍTEM ANTERIOR AOS TREZENTOS ÚLTIMOS"
        Case 115: MsgBox "  ÍTEM NÃO EXISTE OU JÁ CANCELADO"
        Case 116: MsgBox "  CÓDIGO COM ESPAÇOS NÃO PERMITIDO"
        Case 117: MsgBox "  DESCRICAO SEM CARACTER ALFABÉTICO NÃO PERMITIDO"
        Case 118: MsgBox "  ACRÉSCIMO MAIOR QUE VALOR DO ÍTEM"
        Case 119: MsgBox "  DESCONTO MAIOR QUE VALOR DO ÍTEM"
        Case 120: MsgBox "  DESCONTO EM ISS NÃO PERMITIDO"
        Case 121: MsgBox "  ACRÉSCIMO EM ÍTEM JÁ EFETUADO"
        Case 122: MsgBox "  DESCONTO EM ÍTEM JÁ EFETUADO"
        Case 123: MsgBox "  ERRO NA MEMÓRIA FISCAL CHAMAR CREDENCIADO"
        Case 124: MsgBox "  AGUARDANDO GRAVAÇÃO NA MEMÓRIA FISCAL"
        Case 125: MsgBox "  CARACTER REPETIDO NA CODIFICAÇÃO DO GT"
        Case 126: MsgBox "  VERSÃO JÁ GRAVADA NA MEMÓRIA FISCAL"
        Case 127: MsgBox "  ESTOURO DE CAPACIDADE NO CHEQUE"
        Case 128: MsgBox "  TIMEOUT NA LEITURA DO CHEQUE"
        Case 129: MsgBox "  MÊS INVÁLIDO"
        Case 130: MsgBox "  COORDENADA INVÁLIDA"
        Case 131:: MsgBox "  SOBREPOSIÇÃO DE TEXTO"
        Case 132: MsgBox "  SOBREPOSIÇÃO DE TEXTO NO VALOR"
        Case 133: MsgBox "  SOBREPOSIÇÃO DE TEXTO NO EXTENSO"
        Case 134: MsgBox "  SOBREPOSIÇÃO DE TEXTO NO FAVORECIDO"
        Case 135: MsgBox "  SOBREPOSIÇÃO DE TEXTO NA LOCALIDADE"
        Case 136: MsgBox "  SOBREPOSIÇÃO DE TEXTO NO OPCIONAL"
        Case 137: MsgBox "  SOBREPOSIÇÃO DE TEXTO NO DIA"
        Case 138: MsgBox "  SOBREPOSIÇÃO DE TEXTO NO MÊS"
        Case 139: MsgBox "  SOBREPOSIÇÃO DE TEXTO NO ANO"
        Case 140: MsgBox "  USANDO MFD DE OUTRO ECF"
        Case 141: MsgBox "  PRIMEIRO DADO DIFERENTE DE ESC OU 1C"
        Case 142: MsgBox "  NÃO PERMITIDO ALTERAR SEM INTERVENÇÃO TÉCNICA"
        Case 143: MsgBox "  DADOS DA ÚLTIMA RZ CORROMPIDOS"
        Case 144: MsgBox "  COMANDO NÃO PERMITIDO NO MODO INICIALIZAÇÃO"
        Case 145:::: MsgBox "  AGUARDANDO ACERTO DE RELÓGIO"
        Case 146::::: MsgBox "  MFD JÁ INICIALIZADA PARA OUTRA MF"
        Case 147: MsgBox "  AGUARDANDO ACERTO DO RELÓGIO OU DESBLOQUEIO PELO TECLADO"
        Case 148: MsgBox "  VALOR FORMA DE PAGAMENTO MAIOR QUE MÁXIMO PERMITIDO"
        Case 149: MsgBox "  RAZÃO SOCIAL EM BRANCO"
        Case 150:: MsgBox "  NOME DE FANTASIA EM BRANCO"
        Case 151: MsgBox "  ENDEREÇO EM BRANCO"
        Case 152: MsgBox "  ESTORNO DE CDC NÃO PERMITIDO"
        Case 153: MsgBox "  DADOS DO PROPRIETÁRIO IGUAIS AO ATUAL"
        Case 154: MsgBox "  ESTORNO DE FORMA DE PAGAMENTO NÃO PERMITIDO"
        Case 155: MsgBox "  DESCRIÇÃO FORMA DE PAGAMENTO IGUAL JÁ PROGRAMADA"
        Case 156::: MsgBox "  ACERTO DE HORÁRIO DE VERÃO SÓ IMEDIATAMENTE APÓS RZ"
        Case 157:: MsgBox "  IT NÃO PERMITIDA MF RESERVADA PARA RZ"
        Case 158: MsgBox "  SENHA CNPJ INVÁLIDA"
        Case 159: MsgBox "  TIMEOUT NA INICIALIZAÇÃO DA NOVA MF"
        Case 160: MsgBox "  NÃO ENCONTRADO DADOS NA MFD"
        Case 161: MsgBox "  SANGRIA OU SUPRIMENTO DEVEM SER ÚNICOS NO CNF"
        Case 162: MsgBox "  ÍNDICE DA FORMA DE PAGAMENTO NULO NÃO PERMITIDO"
        Case 163: MsgBox "  UF DESTINO INVÁLIDA"
        Case 164: MsgBox "  TIPO DE TRANSPORTE INCOMPATÍVEL COM UF DESTINO"
        Case 165: MsgBox "  DESCRIÇÃO DO PRIMEIRO ÍTEM DO BILHETE DE PASSAGEM DIFERENTE DE TARIFA"
        Case 166: MsgBox "  AGUARDANDO IMPRESSÃO DE CHEQUE OU AUTENTICAÇÃO"
        Case 167: MsgBox "  NÃO PERMITIDO PROGRAMAÇAO CNPJ IE COM ESPAÇOS EM BRANCO"
        Case 168: MsgBox "  NÃO PERMITIDO PROGRAMAÇÃO UF COM ESPAÇOS EM BRANCO"
        Case 169: MsgBox "  NÚMERO DE IMPRESSÕES DA FITA DETALHE NESTA INTERVENÇÃO TÉCNICA ESGOTADO"
        Case 170: MsgBox " CF JÁ SUBTOTALIZADO"
        Case 171: MsgBox " CUPOM NÃO SUBTOTALIZADO"
        Case 172:: MsgBox " ACRÉSCIMO EM SUBTOTAL JÁ EFETUADO"
        Case 173:: MsgBox " DESCONTO EM SUBTOTAL JÁ EFETUADO"
        Case 174:: MsgBox " ACRÉSCIMO NULO NÃO PERMITIDO"
        Case 175: MsgBox " DESCONTO NULO NÃO PERMITIDO"
        Case 176: MsgBox " CANCELAMENTO DE ACRÉSCIMO OU DESCONTO EM SUBTOTAL NÃO PERMITIDO"
        Case 177:  MsgBox " Data INVÁLIDA"
        Case 178: MsgBox " VALOR DO CHEQUE NULO NÃO PERMITIDO"
        Case 179: MsgBox " VALOR DO CHEQUE INVÁLIDO"
        Case 180: MsgBox " CHEQUE SEM LOCALIDADE NÃO PERMITIDO"
        Case 181: MsgBox " CANCELAMENTO ACRÉSCIMO EM ÍTEM NÃO PERMITIDO"
        Case 182: MsgBox " CANCELAMENTO DESCONTO EM ÍTEM NÃO PERMITIDO"
        Case 183: MsgBox "NÚMERO MÁXIMO DE ÍTENS ATINGIDO"
        Case 184: MsgBox "NÚMERO DE ÍTEM NULO NÃO PERMITIDO"
        Case 185: MsgBox " MAIS QUE DUAS ALÍQUOTAS DIFERENTES NO BILHETE DE PASSAGEM NÃO PERMITIDO"
        Case 186: MsgBox "ACRÉSCIMO OU DESCONTO EM ITEM NÃO PERMITIDO"
        Case 187: MsgBox "   CANCELAMENTO DE ACRÉSCIMO OU DESCONTO EM ITEM NÃO PERMITIDO"
        Case 188: MsgBox "   CLICHE JÁ IMPRESSO"
        Case 189: MsgBox "   TEXTO OPCIONAL DO CHEQUE EXCEDEU O MÁXIMO PERMITIDO"
        Case 190: MsgBox "   IMPRESSÃO AUTOMÁTICA NO VERSO NÃO PERMITIDO NESTE EQUIPAMENTO"
        Case 191: MsgBox "   TIMEOUT NA INSERÇÃO DO CHEQUE"
        Case 192: MsgBox "   OVERFLOW NA CAPACIDADE DE TEXTO DO COMPROVANTE DE CRÉDITO OU DÉBITO"
        Case 193: MsgBox "   PROGRAMAÇÃO DE ESPAÇOS ENTRE CUPONS MENOR QUE O MÍNIMO PERMITIDO"
        Case 194: MsgBox "   EQUIPAMENTO NÃO POSSUI LEITOR DE CHEQUE"
        Case 195: MsgBox "  PROGRAMAÇÃO DE ALÍQUOTA COM VALOR NULO NÃO PERMITIDO"
        Case 196: MsgBox "  PARÂMETRO BAUD RATE INVÁLIDO"
        Case 197: MsgBox "  CONFIGURAÇÃO PERMITIDA SOMENTE PELA PORTA DOS FISCO"
        Case 198: MsgBox "  VALOR TOTAL DO ITEM EXCEDE 11 DÍGITOS"
        Case 199: MsgBox "  PROGRAMAÇÃO DA MOEDA COM ESPAÇOS EM BRACO NÃO PERMITIDO"
        Case 200: MsgBox "  CASAS DECIMAIS DEVEM SER PROGRAMADAS COM 2 OU 3"
        Case 201: MsgBox "  NÃO PERMITE CADASTRAR USUÁRIOS DIFERENTES NA MESMA MFD"
        Case 202: MsgBox "  IDENTIFICAÇÃO DO CONSUMIDOR NÃO PERMITIDA PARA SANGRIA OU SUPRIMENTO"
        Case 203: MsgBox "  CASAS DECIMAIS EM QUANTIDADE MAIOR DO QUE A PERMITIDA"
        Case 204: MsgBox "  CASAS DECIMAIS DO UNITÁRIO MAIOR DO QUE O PERMITIDA"
        Case 205: MsgBox "  POSIÇÃO RESERVADA PARA ICMS"
        Case 206: MsgBox "  POSIÇÃO RESERVADA PARA ISS"
        Case 207: MsgBox "  TODAS AS ALÍQUOTAS COM A MESMA VINCULAÇÃO NÃO PERMITIDO"
        Case 208: MsgBox "  DATA DE EMBARQUE ANTERIOR A DATA DE EMISSÃO"
        Case 208: MsgBox "  RETORNO_DATA_DE_EMBARQUE_ANTERIOR_A_DATA_DA_EMISSAO"
        Case 209: MsgBox "  RETORNO_ALIQUOTA_DE_ISS_NAO_PERMITIDO_SEM_INSCRICAO_MUNICIPAL"
        Case 210: MsgBox "  RETORNO_PACOTE_CLICHE_FORA_DA_SEQUENCIA"
        Case 211: MsgBox "  RETORNO_ESPACO_PARA_PARA_ARMAZENAMENTO_CLICHE_ESGOTADO"
        Case 212: MsgBox "  RETORNO_CLICHE_GRAFICO_NAO_DISPONIVEL_PARA_CONFIRMACAO"
        Case 213: MsgBox "  RETORNO_CRC_DO_CLICHE_GRAFICO_DIFERENTE_DO_INFORMADO"
    
        
    End Select
    
    
    
End Function


Sub ecf_verificaImpressora(Optional bMsg As Boolean = True)
    Dim iRetorno As Integer
    Dim ack As Integer
    Dim st1 As Integer, st2 As Integer
    Dim sMsg As String
    If strImprFiscal = "D" Then
        iRetorno = Daruma_FI_RetornoImpressora(ack, st1, st2)
    Else
        iRetorno = Bematech_FI_RetornoImpressora(ack, st1, st2)
    
    End If
    If iRetorno = 0 Then
        MsgBox "Erro de comunicação!", vbCritical
        Exit Sub
    ElseIf iRetorno = -8 Then
        MsgBox " Erro ao criar ou gravar no arquivo STATUS.TXT ou RETORNO.TXT. ", vbCritical
        Exit Sub
    End If
    
    If st1 >= 128 Then                  ' // bit 7  7
       st1 = st1 - 128
       sMsg = sMsg & "Fim de Papel"
    End If
    If st1 >= 64 Then                  ' // bit 7  6
       st1 = st1 - 64
       sMsg = sMsg & "Pouco Papel"
       'frmVenda.lblMensagem.caption = "Pouco Papel"
    End If
    If st1 >= 32 Then                  ' // bit 7  5
       st1 = st1 - 32
       sMsg = sMsg & "Erro no Relógio"
    End If
    If st1 >= 16 Then                  ' // bit 7  4
       st1 = st1 - 16
       sMsg = sMsg & "Impressora em Erro"
    End If
    If st1 >= 8 Then                  ' // bit 7  3
    st1 = st1 - 8
        sMsg = sMsg & "Comando não iniciado com ESC"
    End If
    If st1 >= 4 Then                  ' // bit 7  2
       st1 = st1 - 4
       sMsg = sMsg & "Comando Inexistente"
    End If
    If st1 >= 2 Then                  ' // bit 7  1
       st1 = st1 - 2
       sMsg = sMsg & "Cupom Aberto"
    End If
    If st1 >= 1 Then                  ' // bit 7  0
       st1 = st1 - 1
       sMsg = sMsg & "Número de Parâmetro(s) Inválido(s)"
    End If
    
    'Codificando o ST2
    
    If st2 >= 128 Then                  ' // bit 7  7
       st2 = st2 - 128
       sMsg = sMsg & "Tipo de Parâmetro de Comando Inválido"
    End If
    If st2 >= 64 Then                  ' // bit 7  6
       st2 = st2 - 64
       sMsg = sMsg & "Memória Fiscal Lotada"
    End If
    If st2 >= 32 Then                  ' // bit 7  5
       st2 = st2 - 32
       sMsg = sMsg & "Erro na Memória RAM"
    End If
    If st2 >= 16 Then                  ' // bit 7  4
       st2 = st2 - 16
       sMsg = sMsg & "Alíquota Não Programada"
    End If
    If st2 >= 8 Then                  ' // bit 7  3
       st2 = st2 - 8
       sMsg = sMsg & "Capacidade de Alíquotas Lotada"
    End If
    If st2 >= 4 Then                  ' // bit 7  2
       st2 = st2 - 4
       sMsg = sMsg & "Cancelamento Não Permitido"
    End If
    If st2 >= 2 Then                  ' // bit 7  1
       st2 = st2 - 2
       sMsg = sMsg & "CNPJ/IE do Proprietário Não Programado"
    End If
    If st2 >= 1 Then                  ' // bit 7  0
       st2 = st2 - 1
       sMsg = sMsg & "Comando Não Executado"
    End If
    
    If sMsg <> "" Then
        If bMsg Then
            MsgBox sMsg, vbCritical
        Else
            'frmVenda.lblMensagem.caption = bMsg
        End If
    End If
    
End Sub
Function F_TrataRetornoVendeItem(i As Integer) As Boolean
    If i = 0 Then
                MsgBox "Erro de comunicação.", vbCritical
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -2 Then
                MsgBox "Parâmetro inválido na função.!", vbCritical
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -3 Then
                MsgBox "Alíquota não programada."
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -4 Then
                MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows."
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -5 Then
                MsgBox "Erro ao abrir a porta de comunicação."
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -27 Then
                MsgBox "Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2)."
                Bematech_FI_CancelaCupom
                Exit Function
            End If
            F_TrataRetornoVendeItem = True
End Function
Function ecf_abreComprovanteNaoFiscalVinculado(formaP As String, valor As String, cupom As String) As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_AbreComprovanteNaoFiscalVinculado(formaP, valor, cupom)
        If i <> 0 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_AbreComprovanteNaoFiscalVinculado(formaP, valor, cupom)
        If i <> 0 Then
            
        End If
        ecf_abreComprovanteNaoFiscalVinculado = i
    End If
End Function
Function ecf_UsaComprovanteNaoFiscalVinculado(buf As String) As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_UsaComprovanteNaoFiscalVinculado(buf)
        If i <> 0 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_UsaComprovanteNaoFiscalVinculado(buf)
        If i <> 0 Then
            
        End If
    
    End If
    ecf_UsaComprovanteNaoFiscalVinculado = i
End Function


Function ecf_fechaComprovanteNaoFiscalVinculado() As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_FechaComprovanteNaoFiscalVinculado
        If i <> 0 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_FechaComprovanteNaoFiscalVinculado
        
        If i <> 0 Then
                        
        End If
            
    End If
    ecf_fechaComprovanteNaoFiscalVinculado = i
End Function

    
Function ecf_FaltaPapel() As Boolean

    Dim iRetorno As Integer
    Dim ack As Integer
    Dim st1 As Integer, st2 As Integer
    Dim sMsg As String
    ecf_FaltaPapel = False
    If strImprFiscal = "D" Then
        iRetorno = Daruma_FI_RetornoImpressora(ack, st1, st2)
    Else
        iRetorno = Bematech_FI_RetornoImpressora(ack, st1, st2)
    End If
    
    If st1 >= 128 Then                  ' // bit 7  7
       
       ecf_FaltaPapel = True
    End If
End Function

Function F_TrataErroECF() As Boolean

    Dim iRetorno As Integer
    Dim ack As Integer
    Dim st1 As Integer, st2 As Integer
    Dim sMsg As String
    F_TrataErroECF = True
    If strImprFiscal = "D" Then
        iRetorno = Daruma_FI_RetornoImpressora(ack, st1, st2)
    Else
        iRetorno = Bematech_FI_RetornoImpressora(ack, st1, st2)
    End If
    If iRetorno = 0 Then
        MsgBox "Erro de comunicação!", vbCritical
        F_TrataErroECF = False
        Exit Function
    ElseIf iRetorno = -8 Then
        MsgBox " Erro ao criar ou gravar no arquivo STATUS.TXT ou RETORNO.TXT. ", vbCritical
        F_TrataErroECF = False
        Exit Function
    End If
    
    If st1 >= 128 Then                  ' // bit 7  7
       st1 = st1 - 128
       sMsg = sMsg & "Fim de Papel"
       F_TrataErroECF = False
    End If
    If st1 >= 64 Then                  ' // bit 7  6
       st1 = st1 - 64
       sMsg = sMsg & "Pouco Papel"
       F_TrataErroECF = False
       'frmVenda.lblMensagem.caption = "Pouco Papel"
    End If
    If st1 >= 32 Then                  ' // bit 7  5
       st1 = st1 - 32
       sMsg = sMsg & "Erro no Relógio"
       F_TrataErroECF = False
    End If
    If st1 >= 16 Then                  ' // bit 7  4
       st1 = st1 - 16
       sMsg = sMsg & "Impressora em Erro"
       F_TrataErroECF = False
    End If
    If st1 >= 8 Then                  ' // bit 7  3
    st1 = st1 - 8
        sMsg = sMsg & "Comando não iniciado com ESC"
        F_TrataErroECF = False
    End If
    If st1 >= 4 Then                  ' // bit 7  2
       st1 = st1 - 4
       sMsg = sMsg & "Comando Inexistente"
       F_TrataErroECF = False
    End If
    If st1 >= 2 Then                  ' // bit 7  1
       st1 = st1 - 2
       sMsg = sMsg & "Cupom Aberto"
       F_TrataErroECF = False
    End If
    If st1 >= 1 Then                  ' // bit 7  0
       st1 = st1 - 1
       sMsg = sMsg & "Número de Parâmetro(s) Inválido(s)"
       F_TrataErroECF = False
    End If
    
    'Codificando o ST2
    
    If st2 >= 128 Then                  ' // bit 7  7
       st2 = st2 - 128
       sMsg = sMsg & "Tipo de Parâmetro de Comando Inválido"
       F_TrataErroECF = False
    End If
    If st2 >= 64 Then                  ' // bit 7  6
       st2 = st2 - 64
       sMsg = sMsg & "Memória Fiscal Lotada"
       F_TrataErroECF = False
    End If
    If st2 >= 32 Then                  ' // bit 7  5
       st2 = st2 - 32
       sMsg = sMsg & "Erro na Memória RAM"
       F_TrataErroECF = False
    End If
    If st2 >= 16 Then                  ' // bit 7  4
       st2 = st2 - 16
       sMsg = sMsg & "Alíquota Não Programada"
       F_TrataErroECF = False
    End If
    If st2 >= 8 Then                  ' // bit 7  3
       st2 = st2 - 8
       sMsg = sMsg & "Capacidade de Alíquotas Lotada"
       F_TrataErroECF = False
    End If
    If st2 >= 4 Then                  ' // bit 7  2
       st2 = st2 - 4
       sMsg = sMsg & "Cancelamento Não Permitido"
       F_TrataErroECF = False
    End If
    If st2 >= 2 Then                  ' // bit 7  1
       st2 = st2 - 2
       sMsg = sMsg & "CNPJ/IE do Proprietário Não Programado"
       F_TrataErroECF = False
    End If
    If st2 >= 1 Then                  ' // bit 7  0
       st2 = st2 - 1
       sMsg = sMsg & "Comando Não Executado"
       F_TrataErroECF = False
    End If

    If sMsg <> "" Then
        MsgBox sMsg, vbCritical
    End If


End Function
Sub ecf_AcionaGuilhotina()
    If strImprFiscal = "B" Then
        Bematech_FI_AcionaGuilhotinaMFD 0
    ElseIf strImprFiscal = "D" Then
    
    End If
End Sub
Function Ecf_ImpressoraNaoResponde() As Integer
    Dim sMsg As String
    Dim i As Integer
    P_Msg "Conectando com ECF..."
    Ecf_ImpressoraNaoResponde = -1 'estava ligada
    sMsg = "Impressora Não Responde" & Chr(13)
    sMsg = sMsg & "Tentar Novamente?"
    Do While i <> 1
        i = ecf_VerificaImpressoraLigada
        If i <> 1 Then
            If F_EscolheSimNao(sMsg) = 0 Then
                Do Until (i = -1 Or i = 1)
                    i = Ecf_ImpressoraNaoResponde()
                Loop
                If i = -1 Then i = 0   'para o caso de ter ligado, a função recursiva iria retornar -1
                Ecf_ImpressoraNaoResponde = i 'ligou ou teclou não depois de ter teclado sim
                                
                Exit Function
            Else
                Ecf_ImpressoraNaoResponde = 1 'não ligou, e cancelou
                Exit Function
            End If
        End If
    Loop
    
    

End Function
Function ecf_RelatorioGerencial(Texto As String) As Integer
    If strImprFiscal = "D" Then
        ecf_RelatorioGerencial = Daruma_FI_RelatorioGerencial(Texto)
    ElseIf strImprFiscal = "B" Then
        ecf_RelatorioGerencial = Bematech_FI_RelatorioGerencial(Texto)
    End If
End Function
Function ecf_FechaRelatorioGerencial() As Integer
    If strImprFiscal = "D" Then
        ecf_FechaRelatorioGerencial = Daruma_FI_FechaRelatorioGerencial()
    ElseIf strImprFiscal = "B" Then
        ecf_FechaRelatorioGerencial = Bematech_FI_FechaRelatorioGerencial()
    End If
End Function
Function ecf_RecebimentoNaoFiscal(ByVal IndiceTotalizador As String, ByVal valor As String, ByVal FormaPagamento As String) As Integer
    
    If strImprFiscal = "D" Then
        ecf_RecebimentoNaoFiscal = Daruma_FI_RecebimentoNaoFiscal(IndiceTotalizador, valor, FormaPagamento)
    ElseIf strImprFiscal = "E" Then
        '
    Else
        ecf_RecebimentoNaoFiscal = Bematech_FI_RecebimentoNaoFiscal(IndiceTotalizador, valor, FormaPagamento)
    End If
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

Function ecf_VerificaEstadoImpressora(ByRef iAck As Integer, ByRef ist1 As Integer, ByRef ist2 As Integer) As Integer
    If strImprFiscal = "D" Then
        ecf_VerificaEstadoImpressora = Daruma_FI_VerificaEstadoImpressora(iAck, ist1, ist2)
    ElseIf strImprFiscal = "E" Then
        
    Else
        ecf_VerificaEstadoImpressora = Bematech_FI_VerificaEstadoImpressora(iAck, ist1, ist2)
    End If
End Function

Function ecf_DadosUltReducao() As ultReducao
    Dim i As Integer
    Dim dados As String
    Dim c As New Collection
    dados = Space(631)
    If strImprFiscal = "D" Then
        i = Daruma_FI_DadosUltimaReducao(dados)
    ElseIf strImprFiscal = "E" Then
        'i = elgin_DadosUltimaReducao(dados)
    Else
        i = Bematech_FI_DadosUltimaReducao(dados)
    End If
    
    
    Set c = F_SeparaTexto(dados, ",", False)
    ecf_DadosUltReducao.modoReducao = c(1)
    ecf_DadosUltReducao.GrandeTotal = c(2)
    ecf_DadosUltReducao.Cancelamentos = c(3)
    ecf_DadosUltReducao.Descontos = c(4)
    ecf_DadosUltReducao.Tributos = c(5)
    ecf_DadosUltReducao.Totalizadores_Parciais_Tributados = c(6)
    ecf_DadosUltReducao.Sangria = c(7)
    ecf_DadosUltReducao.Suprimentos = c(8)
    ecf_DadosUltReducao.Totalizadores_nao_sujeitos_ao_ICMS = c(9)
    ecf_DadosUltReducao.Contadores_dos_TPs_nao_sujeitos_ao_ICMS = c(10)
    ecf_DadosUltReducao.Contador_de_ordem_de_operacao = c(11)
    ecf_DadosUltReducao.Contador_de_operacoes_nao_sujeitas_ao_ICMS = c(12)
    ecf_DadosUltReducao.Numero_de_Aliquotas_cadastradas = c(13)
    ecf_DadosUltReducao.Data_do_movimento = c(14)
    ecf_DadosUltReducao.Acrescimos = c(15)
    ecf_DadosUltReducao.Acrescimo_financeiro = c(16)
    
    
End Function
