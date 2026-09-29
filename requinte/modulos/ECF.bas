Attribute VB_Name = "ECF"
'este módulo serve para colocar os procedimentos
'padrões de chamadas de ecf das impressoras fiscais

Public Type retorno_impressora
    i As Integer
    ack As Integer
    st1 As Integer
    st2 As Integer
    msg As String
    
    fimpapel As Boolean
    poucopapel As Boolean
    errorelogio As Boolean
    impressoraemerro As Boolean
    primeironaoesc As Boolean
    cmdinexistente As Boolean
    cupomaberto As Boolean
    numeroparamcmdinvalido As Boolean
    
    
    tipoparcmdinvalido As Boolean
    memfislotada As Boolean
    erroramcmos As Boolean
    aliquotanaoprogr As Boolean
    capacidaliqesgotada As Boolean
    cancelamentonaopermitido As Boolean
    cnpjnaoprogramado As Boolean
    comandonaoexecut As Boolean
    
    
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
Function ecf_abreComprovanteNaoFiscalVinculado(ByVal FormaPagamento As String, ByVal valor As String, ByVal NumeroCupom As String) As Integer
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_AbreComprovanteNaoFiscalVinculado(FormaPagamento, valor, NumeroCupom)
    ElseIf strImprFiscal = "E" Then
        i = Elgin_AbreComprovanteNaoFiscalVinculado(FormaPagamento, valor, NumeroCupom)
    ElseIf strImprFiscal = "B" Then
        i = Bematech_FI_AbreComprovanteNaoFiscalVinculado(FormaPagamento, valor, NumeroCupom)
    End If
End Function
Function ecf_fechaComprovanteNaoFiscalVinculado()
    Dim i As Integer
    If strImprFiscal = "D" Then
        i = Daruma_FI_FechaComprovanteNaoFiscalVinculado
    ElseIf strImprFiscal = "E" Then
        i = Elgin_FechaComprovanteNaoFiscalVinculado
    ElseIf strImprFiscal = "B" Then
        i = Bematech_FI_FechaComprovanteNaoFiscalVinculado
    End If
    ecf_fechaComprovanteNaoFiscalVinculado = i
End Function
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
Sub ecf_ReducaoZ()
    Dim i As Integer
    
    If strImprFiscal = "D" Then
        i = Daruma_FI_ReducaoZ(Format(Date, "ddmmyy"), Format(Time, "hhmmss"))
        If i <> 0 Then
            Erro_Daruma i
        End If
    Else
        i = Bematech_FI_ReducaoZ(Format(Date, "ddmmyy"), Format(Time, "hhmmss"))
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
    Shell "Notepad " & App.path & "\retorno.txt"
End Sub
Function ecf_NumeroCupom(Num_Cpm As String) As Integer
    Dim sCupom As String
    Dim i As Integer
    sCupom = Space(6)
    If strImprFiscal = "D" Then
        i = Daruma_FI_NumeroCupom(sCupom)
    Else
        i = Bematech_FI_NumeroCupom(sCupom)
    End If
    ecf_NumeroCupom = i
    Num_Cpm = sCupom
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
Function ecf_TerminaFechamentoCupom(msg As String) As Integer
    Dim i As Integer
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
Sub ecf_verificaImpressora()
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
        MsgBox sMsg, vbCritical
    End If
    
End Sub
Function ecf_RetornoImpressora(ack As Integer, st1 As Integer, st2 As Integer) As Integer
    If strImprFiscal = "D" Then
        ecf_RetornoImpressora = Daruma_FI_RetornoImpressora(ack, st1, st2)
    ElseIf strImprFiscal = "E" Then
         
    Else
        ecf_RetornoImpressora = Bematech_FI_RetornoImpressora(ack, st1, st2)
    End If
     
End Function
Function ecf_retornoImprComandos() As retorno_impressora
    Dim i As Integer
    Dim ack As Integer
    Dim st1 As Integer
    Dim st2 As Integer
    Dim sMsg As String
    i = ecf_VerificaEstadoImpressora(ack, st1, st2)
    ecf_retornoImprComandos.i = i
    ecf_retornoImprComandos.ack = ack
    ecf_retornoImprComandos.st1 = st1
    ecf_retornoImprComandos.st2 = st2
    
    
    
    If st1 >= 128 Then                  ' // bit 7  7
       st1 = st1 - 128
       sMsg = sMsg & "Fim de Papel" & Chr(13)
       ecf_retornoImprComandos.fimpapel = True
    End If
    If st1 >= 64 Then                  ' // bit 7  6
       st1 = st1 - 64
       sMsg = sMsg & "Pouco Papel" & Chr(13)
       ecf_retornoImprComandos.poucopapel = True
       'frmVenda.lblMensagem.caption = "Pouco Papel"
    End If
    If st1 >= 32 Then                  ' // bit 7  5
       st1 = st1 - 32
       sMsg = sMsg & "Erro no Relógio" & Chr(13)
       ecf_retornoImprComandos.errorelogio = True
    End If
    If st1 >= 16 Then                  ' // bit 7  4
       st1 = st1 - 16
       sMsg = sMsg & "Impressora em Erro" & Chr(13)
       ecf_retornoImprComandos.impressoraemerro = True
    End If
    If st1 >= 8 Then                  ' // bit 7  3
    st1 = st1 - 8
        sMsg = sMsg & "Comando não iniciado com ESC" & Chr(13)
        ecf_retornoImprComandos.primeironaoesc = True
    End If
    If st1 >= 4 Then                  ' // bit 7  2
       st1 = st1 - 4
       sMsg = sMsg & "Comando Inexistente" & Chr(13)
       ecf_retornoImprComandos.cmdinexistente = True
    End If
    If st1 >= 2 Then                  ' // bit 7  1
       st1 = st1 - 2
       sMsg = sMsg & "Cupom Aberto" & Chr(13)
       ecf_retornoImprComandos.cupomaberto = True
    End If
    If st1 >= 1 Then                  ' // bit 7  0
       st1 = st1 - 1
       sMsg = sMsg & "Número de Parâmetro(s) Inválido(s)" & Chr(13)
       ecf_retornoImprComandos.numeroparamcmdinvalido = True
    End If
    
    'Codificando o ST2
    
    If st2 >= 128 Then                  ' // bit 7  7
       st2 = st2 - 128
       sMsg = sMsg & "Tipo de Parâmetro de Comando Inválido" & Chr(13)
       ecf_retornoImprComandos.tipoparcmdinvalido = True
    End If
    If st2 >= 64 Then                  ' // bit 7  6
       st2 = st2 - 64
       sMsg = sMsg & "Memória Fiscal Lotada" & Chr(13)
       ecf_retornoImprComandos.memfislotada = True
    End If
    If st2 >= 32 Then                  ' // bit 7  5
       st2 = st2 - 32
       sMsg = sMsg & "Erro na Memória RAM" & Chr(13)
       ecf_retornoImprComandos.erroramcmos = True
    End If
    If st2 >= 16 Then                  ' // bit 7  4
       st2 = st2 - 16
       sMsg = sMsg & "Alíquota Não Programada" & Chr(13)
       ecf_retornoImprComandos.aliquotanaoprogr = True
    End If
    If st2 >= 8 Then                  ' // bit 7  3
       st2 = st2 - 8
       sMsg = sMsg & "Capacidade de Alíquotas Lotada" & Chr(13)
       ecf_retornoImprComandos.capacidaliqesgotada = True
    End If
    If st2 >= 4 Then                  ' // bit 7  2
       st2 = st2 - 4
       sMsg = sMsg & "Cancelamento Não Permitido" & Chr(13)
       ecf_retornoImprComandos.cancelamentonaopermitido = True
    End If
    If st2 >= 2 Then                  ' // bit 7  1
       st2 = st2 - 2
       sMsg = sMsg & "CNPJ/IE do Proprietário Não Programado" & Chr(13)
       ecf_retornoImprComandos.cnpjnaoprogramado = True
    End If
    If st2 >= 1 Then                  ' // bit 7  0
       st2 = st2 - 1
       sMsg = sMsg & "Comando Não Executado" & Chr(13)
       ecf_retornoImprComandos.comandonaoexecut = True
    End If
    ecf_retornoImprComandos.msg = sMsg

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

Function ecf_VerificaEstadoImpressora(ByRef iAck As Integer, ByRef ist1 As Integer, ByRef ist2 As Integer) As Integer
    If strImprFiscal = "D" Then
        ecf_VerificaEstadoImpressora = Daruma_FI_VerificaEstadoImpressora(iAck, ist1, ist2)
    ElseIf strImprFiscal = "E" Then
        
    Else
        ecf_VerificaEstadoImpressora = Bematech_FI_VerificaEstadoImpressora(iAck, ist1, ist2)
    End If
End Function


