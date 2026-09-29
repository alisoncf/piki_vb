Attribute VB_Name = "Dedicado"
Option Explicit
Public Declare Function ConfiguraIntSiTefInterativo Lib "CliSitef32I.dll" (ByVal pEnderecoIP As String, ByVal pCodigoLoja As String, ByVal pNumeroTerminal As String, ByVal ConfiguraResultado As Integer) As Long
Public Declare Function IniciaFuncaoSiTefInterativo Lib "CliSitef32I.dll" (ByVal Funcao As Long, ByVal pValor As String, ByVal pCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String, ByVal pOperador As String, ByVal pParamAdic As String) As Long
Public Declare Sub FinalizaTransacaoSiTefInterativo Lib "CliSitef32I.dll" (ByVal Confirma As Integer, ByVal pNumeroCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String)
Public Declare Sub FinalizaTransacaoSiTefInterativoA Lib "CliSitef32I.dll" (ByVal Confirma As Integer, ByVal pNumeroCuponFiscal As String, ByVal pDataFiscal As String, ByVal pHorario As String)
Public Declare Function ContinuaFuncaoSiTefInterativo Lib "CliSitef32I.dll" (ByRef pProximoComando As Long, ByRef pTipoCampo As Long, ByRef pTamanhoMinimo As Integer, ByRef pTamanhoMaximo As Integer, ByVal pBuffer As String, ByVal TamMaxBuffer As Long, ByVal ContinuaNavegacao As Long) As Long
Dim i As Integer
Global Resultado     As Long
Global bTefDedicado As Boolean 'diz se usa tef dedicado ou não
Global bEfetuouFormaPagamentoCartao As Boolean 'para saber q a forma de pagamento foi efetuada na impressora
Global cTroco As Currency 'caso haja saque, pra trazer a finalizadora
Dim sCab As String
Dim sDataTr As String
Dim sHoraTr As String
Dim Buffer As String * 20000
Dim bkBuf As String
Dim iFinalizouNao As Boolean
Function F_BuscaOpcao(Texto As String, Optional sSepara As String = ":") As String
    
    Dim iCount As Long
    
    Dim sPalavra As String
    Dim sBuf As String
    For iCount = 1 To Len(Texto)
        sBuf = Mid(Texto, iCount, 1)
        If InStr(sSepara, sBuf) > 0 Then
            F_BuscaOpcao = sPalavra
        Else
            sPalavra = sPalavra & sBuf
        End If
    Next
End Function

Function F_EscolheSimNao(pergunta As String) As Integer
'    If MsgBox(pergunta & "?", vbYesNo + vbQuestion) = vbYes Then
'        F_EscolheSimNao = 0
'    Else
'        F_EscolheSimNao = 1
'    End If
On Error Resume Next
    If frmSimNao.tag = "" Then
        frmSimNao.tag = "1"
    End If
    frmSimNao.lblMsg.caption = F_Quebra(pergunta, 60)
    frmSimNao.Show 1
    frmSimNao.Visible = False
    
    F_EscolheSimNao = Int(frmSimNao.tag)
    Unload frmSimNao
    Screen.ActiveForm.Refresh
End Function
Function P_TrataProximoComando(proximocomando As Long, Buffer As String, continua As Long, iLen As Integer, tipocampo As Long) As String
    Dim s As String
        
        
    If Left(Buffer, 1) <> Chr(0) Then
        bkBuf = Buffer
    ElseIf Len(Trim(Buffer)) = 1 Then
        bkBuf = ""
    End If
        
    Buffer = bkBuf
    If proximocomando = 1 Then
            P_Msg Buffer
        ElseIf proximocomando = 2 Then
            
            P_Msg Buffer
        ElseIf proximocomando = 3 Then
            P_Msg bkBuf
            'P_Msg Buffer
            s = Buffer
            Buffer = ""
            Buffer = s
        ElseIf proximocomando = 4 Then
            sCab = Buffer
        ElseIf proximocomando = 11 Then
            'P_Msg ""
            MsgBox "11"
        ElseIf proximocomando = 12 Then
            'P_Msg ""
            MsgBox "12"
        ElseIf proximocomando = 13 Then
            'P_Msg ""
'            MsgBox "13"
            Buffer = ""
        ElseIf proximocomando = 14 Then
            sCab = ""
            Resultado = 0
            Buffer = ""
        ElseIf proximocomando = 15 Then
            P_Msg sCab
        
        ElseIf proximocomando = 16 Then
            sCab = ""
            Buffer = ""
        ElseIf proximocomando = 20 Then
            i = F_EscolheSimNao(Buffer)
            Buffer = i
        ElseIf proximocomando = 21 Then
            i = F_Retorno(sCab, Trim(Buffer), continua)
            
            Buffer = i
            
        ElseIf proximocomando = 22 Then
            P_MsgOk Buffer
            'MsgBox Buffer, vbInformation
                       
        ElseIf proximocomando = 23 Then
        '    MsgBox "23"
            If F_MsgModal(F_AjustaTexto(Buffer), Buffer) = 1 Then
                continua = 1
            End If
        ElseIf proximocomando = 30 Then
        
            Buffer = F_LeValor(Buffer, Buffer, continua, iLen, IIf(tipocampo = 500, True, False))
            'tratar o tammax e min aki!
            
        ElseIf proximocomando = 31 Then
            Buffer = F_LeValor(Buffer, Buffer, continua, iLen)
        ElseIf proximocomando = 34 Then
            Buffer = F_LeValor(Buffer, Buffer, continua, iLen)
        ElseIf proximocomando = 122 Then
            MsgBox "aqui"
            
        ElseIf proximocomando = 35 Then
            MsgBox "35"
        Else
            If proximocomando > 0 Then
                MsgBox proximocomando & "-" & Buffer
            End If
        End If
        P_TrataProximoComando = Buffer
End Function
Function F_Retorno(pergunta As String, sOp As String, continua As Long) As Integer
    frmEscolhe.tag = "1"
    frmEscolhe.Label1.caption = pergunta
    frmEscolhe.sOp = sOp
    frmEscolhe.Show 1
    On Error Resume Next
    If frmEscolhe.tag = "" Then
        F_Retorno = Int(F_BuscaOpcao(frmEscolhe.lst.Text, ":"))
    ElseIf frmEscolhe.tag = "1" Then
        F_Retorno = 1
        continua = 1
    ElseIf frmEscolhe.tag = "-1" Then
        F_Retorno = -1
        continua = -1
    End If
    Unload frmEscolhe
        
End Function
Function F_LeValor(pergunta As String, sOp As String, continua As Long, iLen As Integer, Optional bMask As Boolean = False) As String
    frmLeValor.tag = "1"
    frmLeValor.Label1.caption = pergunta
    frmLeValor.sOp = sOp
    If iLen > 0 Then
        frmLeValor.txtbxValor.MaxLength = iLen
    End If
    If bMask Then
        frmLeValor.txtbxValor.PasswordChar = "*"
    Else
        frmLeValor.txtbxValor.PasswordChar = ""
    End If
    frmLeValor.Show 1
    On Error Resume Next
    
    If frmLeValor.tag = "" Then
        F_LeValor = frmLeValor.txtbxValor.Text & Chr(0)
        
        continua = 0
    ElseIf frmLeValor.tag = "-1" Then
        F_LeValor = "-1"
        continua = -1
    ElseIf frmLeValor.tag = "1" Then
        F_LeValor = "1"
        continua = 1
    End If
    Unload frmLeValor
        
End Function

Sub Gerencial()
    Dim Retorno        As Long
    
    Dim proximocomando As Long
    Dim tipocampo      As Long
    Dim TamanhoMinimo  As Integer
    Dim TamanhoMaximo  As Integer
    Dim ContinuaNavegacao  As Long
    Dim i As Integer
    Dim sCab As String
    Dim iCount As Integer
    Dim c As New Collection
    Dim primeiravia As String
    Dim segundavia As String
    Dim sAutorizacaoCheque As String
    Dim sData As String
    Dim sHora As String
    Configura strIP_TerminalTEF
    ContinuaNavegacao = 0
    sData = Date
    sHora = Time
    cTroco = 0
    iFinalizouNao = False
    'hora que inicia o cancelamento
    Retorno = IniciaFuncaoSiTefInterativo(110, 0, "", Format(sData, "yyyymmdd"), Format(sHora, "hhmmss"), "", "")
    
    Do While Retorno = 10000
        
        
        Retorno = ContinuaFuncaoSiTefInterativo(proximocomando, tipocampo, TamanhoMinimo, TamanhoMaximo, (Buffer), Len(Buffer), ContinuaNavegacao)
        
        If tipocampo = -1 And ContinuaNavegacao = 1 Then
'            MsgBox "-1"
            ContinuaNavegacao = 0
'            MsgBox "zerou"
        End If
        If Retorno <> -2 Then
            P_TrataProximoComando proximocomando, Buffer, ContinuaNavegacao, TamanhoMaximo, tipocampo
        End If
        If ContinuaNavegacao = 1 Then
            'MsgBox "jogou 1" 'apenas para teste
            
        End If
        'ContinuaNavegacao = 0
        
        
        
        If tipocampo = 505 Then 'qtd parcelas
            Buffer = Format(Buffer, "0#")
        
        
        ElseIf tipocampo = 512 Then
            If Len(Trim(Buffer)) - 1 > TamanhoMaximo Then
                MsgBox "Sua digitação possui " & Len(Trim(Buffer)) - 1 & " Dígitos, quando o máximo é " & TamanhoMaximo, vbExclamation
'                ContinuaNavegacao = 1
                'Buffer = 0
            ElseIf Len(Trim(Buffer)) < TamanhoMinimo Then
                MsgBox "Sua digitação possui " & Len(Trim(Buffer)) - 1 & " Dígitos, quando o mínimo é " & TamanhoMinimo, vbExclamation
                'ContinuaNavegacao = 1
                Buffer = 0
            End If
            
        ElseIf tipocampo = 514 Then
        'cód seg cartão
        ElseIf tipocampo = 136 Then
        'Contém as 6 primeiras posições do cartão (bin)
            'MsgBox Buffer
        ElseIf tipocampo = 120 Then
            sAutorizacaoCheque = Buffer
        ElseIf tipocampo = 121 Then
            
            primeiravia = Buffer
            
        ElseIf tipocampo = 122 Then
            segundavia = Buffer
        ElseIf tipocampo = 20 Then
            MsgBox "vinte"
        End If
        
        'If Trim(Buffer) = "" And ContinuaNavegacao = 1 Then
        '    Exit Function
        'End If
    
    Loop
    If Retorno = 0 Then
         
        If tipocampo = -1 Then
            If IsLoad("frmMsg") Then
                Unload frmMsg
            End If
        ElseIf tipocampo = 1 Then
            'ecf_LeituraX
            
            'Set c = F_SeparaTexto(Trim(Buffer))
            'For i = 1 To c.count
            '    iCount = ecf_RelatorioGerencial(c(i))
                
            'Next
            'ecf_FechaRelatorioGerencial
            'sCab = F_AjustaTexto(sCab)
            
            
        End If
    End If
    If primeiravia <> "" Or segundavia <> "" Then
       ' ecf_LeituraX
        If ImprimeGerencialDuasVias("0", (primeiravia), (segundavia), "", "", "") Then
            FinalizaTransacaoSiTefInterativo 1, "", Format(sData, "yyyymmdd"), Format(sHora, "hhmmss")
        Else
            FinalizaTransacaoSiTefInterativo 0, "", Format(sData, "yyyymmdd"), Format(sHora, "hhmmss")
            MsgBox "Transação não foi efetuada, favor reter o Cupom.", vbExclamation
            
        End If
    End If
    
trata_retorno Retorno
    P_Status
  End Sub
Sub trata_retorno(Retorno As Long)
    
    If Retorno > 0 Then
        P_MsgOk "Negada Pelo autorizador"
    ElseIf Retorno = -1 Then
        P_MsgOk "Módulo não inicializado"
    ElseIf Retorno = -2 Then
        P_MsgOk "Cancelado pelo Operador"
    ElseIf Retorno = -3 Then
        P_MsgOk "Cancelado pelo Usuário"
    ElseIf Retorno = -4 Then
        P_MsgOk "Falta de memória pra rodar a função"
    ElseIf Retorno = -5 Then
        P_MsgOk "Sem comunicação com o Sitef"
    ElseIf Retorno = -6 Then
        P_MsgOk "Operação cancelada pelo usuário"
    ElseIf Retorno < 0 Then
        P_MsgOk "Erro desconhecido " & Retorno
    End If
    
    P_Msg ""
    
End Sub
Public Sub Configura(ByVal IP As Variant)

  Dim Retorno        As Long

  'Screen.MousePointer = vbHourglass
''stridestab = "SE000001"

  Retorno = ConfiguraIntSiTefInterativo(IP & Chr(0), strId_Estab & Chr(0), strId_Estab & Chr(0), 0)
  Screen.MousePointer = vbDefault

  If (Retorno = 0) Then
    P_Msg "Sitef Ok"
  Else
    MsgBox "Erro: Retorno -> " & CStr(Retorno)
  End If

End Sub

Public Function EfetuaPagamento(valor As String, cupom As String, atendente As String, nomecli As String, cpfcli As String, enderecocli As String) As Boolean

  Dim Retorno        As Long
  'Dim Buffer         As String * 20000
  
  Dim proximocomando As Long
  Dim tipocampo      As Long
  Dim TamanhoMinimo  As Integer
  Dim TamanhoMaximo  As Integer
  Dim ContinuaNavegacao  As Long
  Dim primeiravia As String
  Dim segundavia As String
  Dim sLinhaCheque As String
  Dim sCab As String
  Dim i As Integer
  Dim c As New Collection
  Dim iCount As Integer
  Dim iVias As Integer
  Dim sData As String
  Dim sHora As String
  Dim bkBuf As String
  Configura strIP_TerminalTEF
  i = ecf_NumeroCupom(strNum_Cpm)
    iFinalizouNao = False
  Screen.MousePointer = vbHourglass
  sDataTr = Date
  sHoraTr = Time
  Retorno = IniciaFuncaoSiTefInterativo(0, Format(valor, "###,##0.00") & Chr(0), cupom & Chr(0), Format(sDataTr, "yyyymmdd") & Chr(0), Format(sHoraTr, "hhmmss") & Chr(0), atendente & Chr(0), Chr(0))
  Screen.MousePointer = vbDefault
    cTroco = 0
  proximocomando = 0
  tipocampo = 0
  TamanhoMinimo = 0
  TamanhoMaximo = 0
  ContinuaNavegacao = 0
  Resultado = 0
  Buffer = String(20000, Chr(0))

  Do

    'Screen.MousePointer = vbHourglass
    'Buffer = Buffer & Chr(0)
    Buffer = Trim(Buffer)
    
    
    
    'If tipocampo = -1 Then
    '    Buffer = String(20000, 0)
    'End If
    
    'Buffer = Buffer
    
    'Retorno = ContinuaFuncaoSiTefInterativo(proximocomando, tipocampo, TamanhoMinimo, TamanhoMaximo, Buffer, Len(Buffer), ContinuaNavegacao)
    Retorno = ContinuaFuncaoSiTefInterativo(proximocomando, tipocampo, TamanhoMinimo, TamanhoMaximo, Buffer, Len(Buffer), ContinuaNavegacao)
    If Left(Buffer, 1) <> Chr(0) Then
        bkBuf = Buffer
    ElseIf Len(Trim(Buffer)) = 1 Then
        bkBuf = ""
    End If
    
    Buffer = bkBuf
    
    Screen.MousePointer = vbDefault

    If tipocampo = -1 And ContinuaNavegacao = 1 Then
'            MsgBox "-1"
            ContinuaNavegacao = 0
'            MsgBox "zerou"
        End If


    If (Retorno = 10000) Then
        If proximocomando = 23 Then
            If F_MsgModal(Buffer, Buffer) = 1 Then
                ContinuaNavegacao = 1
            End If
                        
        
        ElseIf proximocomando = 0 Then
            P_Msg "Emitindo Comprovante"
            If tipocampo = 100 Then
                Buffer = "Transacao Ok!"
                P_Msg Trim(Buffer)
                Sleep 5000
            End If
        Else
            P_Msg Buffer
        End If
        
        If ContinuaNavegacao = 1 Then
                'MsgBox "pare"
        End If
        P_TrataProximoComando proximocomando, Buffer, ContinuaNavegacao, TamanhoMaximo, tipocampo
        If tipocampo = 505 Then 'qtd parcelas
            Buffer = Format(Buffer, "0#")
        ElseIf tipocampo = 512 Then
            If Len(Trim(Buffer)) - 1 > TamanhoMaximo Then
                MsgBox "Sua digitação possui " & Len(Trim(Buffer)) - 1 & " Dígitos, quando o máximo é " & TamanhoMaximo, vbExclamation
'                ContinuaNavegacao = 1
                'Buffer = 0
            ElseIf Len(Trim(Buffer)) < TamanhoMinimo Then
                MsgBox "Sua digitação possui " & Len(Trim(Buffer)) - 1 & " Dígitos, quando o mínimo é " & TamanhoMinimo, vbExclamation
                'ContinuaNavegacao = 1
                Buffer = 0
            End If
        ElseIf tipocampo = 514 Then
        'cód seg cartão
        ElseIf tipocampo = 136 Then
        'Contém as 6 primeiras posições do cartão (bin)
            'MsgBox Buffer
        ElseIf tipocampo = 120 Then
            sLinhaCheque = Buffer
        ElseIf tipocampo = 121 Then
            primeiravia = Buffer 'primeira via do cnfv
        ElseIf tipocampo = 122 Then
            segundavia = Buffer 'segunda via do cnfv
        ElseIf tipocampo = 130 Then
            If Left(Buffer, 1) = Chr(0) Then
                cTroco = 0
            Else
                cTroco = CCur(Buffer)
            End If
        ElseIf tipocampo = 20 Then
            MsgBox "vinte"
        End If
        
        'If Trim(Buffer) = "" And ContinuaNavegacao = 1 Then
        '    Exit Function
        'End If
    End If

  Loop Until Not (Retorno = 10000)
    
  If (Retorno = 0) Then
    P_Msg "Gravando Transação..."
    P_GravaTransacaoTef cupom, valor
    P_Msg "Imprimindo Comprovante..."
    If ImprimeComprovanteDuasVias(valor, F_SubstituiChar(primeiravia, Chr(10), Chr(13) + Chr(10)), _
        F_SubstituiChar(segundavia, Chr(10), Chr(13) + Chr(10)), nomecli, cpfcli, enderecocli) Then
        FinalizaTransacaoSiTefInterativo 1, cupom, Format(Date, "yyyymmdd"), Format(Time, "hhmmss")
        If sLinhaCheque <> "" Then
            P_MsgOk Buffer, "Por favor, anote o Nº de Autorização"
        End If
        P_Msg "Aguardando Transação TEF"
            EfetuaPagamento = True
        
        P_Msg ""
        P_DeletaTransacaoTef cupom, strIP_Usu
    Else
        FinalizaTransacaoSiTefInterativo 0, cupom, Format(Date, "yyyymmdd"), Format(Time, "hhmmss")
        P_MsgOk "Transação Não Efetuada Favor Reter Cupom..."
      '  Sleep 2000
        EfetuaPagamento = False
        P_Msg ""
        P_DeletaTransacaoTef strNum_Cpm, strIP_Usu
    End If
    
  Else
    
'    'MsgBox "Erro:" & Buffer & " Retorno -> " & CStr(Retorno)
'    If Retorno > 0 Then
'        MsgBox "Negada Pelo autorizador"
'    ElseIf Retorno = -1 Then
'        MsgBox "Módulo não inicializado"
'    ElseIf Retorno = -2 Then
'        MsgBox "Cancelado pelo Operador"
'    ElseIf Retorno = -3 Then
'        MsgBox "Cancelado pelo Usuário"
'    ElseIf Retorno = -4 Then
'        MsgBox "Falta de memória pra rodar a função"
'    ElseIf Retorno = -5 Then
'        MsgBox "Sem comunicação com o Sitef"
'    ElseIf Retorno = -6 Then
'        MsgBox "Operação cancelada pelo usuário"
'    ElseIf Retorno < 0 Then
'        MsgBox "Erro desconhecido " & Retorno
'    End If

    EfetuaPagamento = False
    trata_retorno Retorno
  End If

End Function

Function ImprimeComprovanteDuasVias(valor As String, prim_via As String, seg_via As String, _
    Optional nomecli As String = "", Optional cpfcli As String = "", Optional enderecocli As String) As Boolean
    Dim c As New Collection
    Dim iCount As Integer
    Dim responde As Integer
    Dim Retorno As Integer
    Dim sArqTmp As String
    Dim sBuf As String
    Dim i As Integer
    Dim Flag As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    Close #1
    
    sArqTmp = App.Path & "\cnvtmp.txt"
    responde = -1
    
    responde = Ecf_ImpressoraNaoResponde
    
    If responde = 1 Then
        P_MsgOk "Transação não foi efetuada. Favor reter o Cupom"
        ImprimeComprovanteDuasVias = False
        Exit Function
    ElseIf responde = 0 Then
        'ecf_LeituraX
        ImprimeComprovanteDuasVias = ImprimeComprovanteDuasVias(valor, prim_via, seg_via, nomecli, cpfcli, enderecocli)
        Exit Function
    End If
    
    
    
    
    
    i = 0
    
    
    
    
    Do While i <> 1
    
    
    
    i = ecf_EfetuaFormaPagamento("Cartao", FormatVal(CCur(valor) + cTroco))
        If i = 1 Then
            bEfetuouFormaPagamentoCartao = True
            
            strSQL = F_BuscaUmCampo("msgecf_par", TAB_PARAMETROS, "")
            strSQL = strSQL & Chr(13) & nomecli & "-" & cpfcli
            strSQL = strSQL & "-" & enderecocli
            strSQL = strSQL & Chr(13) + Chr(10) & "Caixa: " & strCod_Cai & " Atendente: " & strNome_Usu
            ecf_TerminaFechamentoCupom strSQL, responde
            
            
        Else
            'MsgBox "Não efetuou Forma de Pagamento!", vbExclamation
        End If
    Loop
    
        ImprimeComprovanteDuasVias = False
        i = ecf_fechaComprovanteNaoFiscalVinculado
        i = ecf_abreComprovanteNaoFiscalVinculado("Cartao", Format(CCur(valor) + CCur(cTroco), "###,##0.00"), strNum_Cpm)
        
    For iCount = 1 To 2
        P_Msg "Abrindo Arquivo Temporário"
        Open sArqTmp For Binary As #1 'ponto de parada aqui!
        If iCount = 1 Then
            Put 1, , prim_via
        Else
            
            For i = 1 To 4
                ecf_UsaComprovanteNaoFiscalVinculado Chr(13) + Chr(10)
            Next
            'ecf_AcionaGuilhotina
            'Bematech_FI_LinhasEntreCupons (2)
            Put 1, , seg_via
        End If
        Close #1
        Open sArqTmp For Input As #1
        
        Do While Not EOF(1)
            
            Line Input #1, sBuf
    '        MsgBox sBuf
            i = ecf_UsaComprovanteNaoFiscalVinculado(sBuf) 'desligar a impressora para testar
            i = ecf_UsaComprovanteNaoFiscalVinculado(Chr(13) + Chr(10)) 'para a impressora não economizar o papel
            
            Retorno = Ecf_ImpressoraNaoResponde
            
            If Retorno = 0 Then
                'ecf_LeituraX
                ImprimeComprovanteDuasVias = ImprimeGerencialDuasVias(valor, prim_via, seg_via, nomecli, cpfcli, enderecocli)
                Exit Function
            ElseIf Retorno = 1 Then
                ImprimeComprovanteDuasVias = False
                Exit Function
            End If
            Do While ecf_FaltaPapel = True
                If iFinalizouNao = True Then
                    
                    Exit Function
                End If
                If F_EscolheSimNao("Falta Papel" & Chr(13) & _
                    "Impressora Não Responde. Tentar Novamente?") = 1 Then
                    P_MsgOk "Transação não foi efetuada. Favor reter o Cupom"
                    ImprimeComprovanteDuasVias = False
                    
                    Exit Function
                Else
                End If
                    ecf_fechaComprovanteNaoFiscalVinculado
                    'ecf_TerminaFechamentoCupom "Obrigado! Volte Sempre!"
                    
                    strSQL = F_BuscaUmCampo("msgecf_par", TAB_PARAMETROS, "")
                    strSQL = strSQL & Chr(13) & nomecli & "-" & cpfcli
                    strSQL = strSQL & "-" & enderecocli
                    strSQL = strSQL & Chr(13) + Chr(10) & "Caixa: " & strCod_Cai & " Atendente: " & strNome_Usu
                    ecf_TerminaFechamentoCupom strSQL, responde
                    
                    'ecf_LeituraX
                    ecf_verificaImpressora
                    Flag = ImprimeGerencialDuasVias(valor, prim_via, seg_via, nomecli, cpfcli, enderecocli)
                    If Flag = True Then
                        ImprimeComprovanteDuasVias = True
                        Exit For
                    End If
                    
            Loop
        Loop
        Close #1
        
    Next
    
    responde = Ecf_ImpressoraNaoResponde 'parada pro teste 49 'aqui tinha um ponto de parada
    ''sim desligar a impressora acima, depois ligar respondendo sim
    '
    If responde = 1 Then 'aqui tinha um ponto de parada
        P_MsgOk "Transação não foi efetuada. Favor reter o Cupom"
        ImprimeComprovanteDuasVias = False
        Exit Function
    ElseIf responde = 0 Then
        'ecf_LeituraX
        ImprimeComprovanteDuasVias = ImprimeGerencialDuasVias(valor, prim_via, seg_via, nomecli, cpfcli, enderecocli)
        Exit Function
    End If
    
    ecf_fechaComprovanteNaoFiscalVinculado
    
    ImprimeComprovanteDuasVias = True
    Exit Function
    
trata_erro:
    MostraErro
End Function



Function ImprimeGerencialDuasVias(valor As String, prim_via As String, seg_via As String, nomecli As String, cpfcli As String, enderecocli As String) As Boolean
    Dim c As New Collection
    Dim iCount As Integer
    Dim responde As Integer
    Dim Retorno As Integer
    Dim iVias As Integer
    Dim sTexto As String
    Dim iImpr As Boolean
    Dim strSQL As String
    
    responde = -1
    
    
    responde = Ecf_ImpressoraNaoResponde
    If responde = 1 Then
        P_MsgOk "Transação não foi efetuada. Favor reter o Cupom"
        ImprimeGerencialDuasVias = False
        Exit Function
    ElseIf responde = 0 Then
        'ecf_LeituraX
        ImprimeGerencialDuasVias = ImprimeGerencialDuasVias(valor, prim_via, seg_via, nomecli, cpfcli, enderecocli)
        Exit Function
    End If
    
    
    
    
    ImprimeGerencialDuasVias = False

    
    ecf_FechaRelatorioGerencial
    For iVias = 1 To 2
        If iVias = 1 Then
            sTexto = prim_via
        Else
            sTexto = seg_via
        End If
        Set c = F_SeparaTexto(sTexto, Chr(10), False) 'impressão da primeira via
        For i = 1 To c.Count
            If iFinalizouNao = True Then
                GoTo termino_nao_conf
            End If
            
            iCount = ecf_RelatorioGerencial(c(i)) 'desligar a impressora para testar
            iCount = ecf_RelatorioGerencial(Chr(13) + Chr(10))
            
            Retorno = Ecf_ImpressoraNaoResponde
            
            If Retorno = 0 Then
                'ecf_LeituraX
                ImprimeGerencialDuasVias = ImprimeGerencialDuasVias(valor, prim_via, seg_via, nomecli, cpfcli, enderecocli)
                Exit Function
            ElseIf Retorno = 1 Then
                ImprimeGerencialDuasVias = False
                Exit Function
            End If
            Do While ecf_FaltaPapel = True And iFinalizouNao = False
                If F_EscolheSimNao("Falta Papel" & Chr(13) & _
                    "Impressora Não Responde. Tentar Novamente?") = 1 Then
                    iFinalizouNao = True
                    P_MsgOk "Transação não foi efetuada. Favor reter o Cupom"
                    
                    ImprimeGerencialDuasVias = False
                    
                    GoTo termino_nao_conf
                    Exit Function
                Else
                End If
                    If iFinalizouNao = True Then
                        GoTo termino_nao_conf
                    End If
                    
                    
                    ecf_fechaComprovanteNaoFiscalVinculado
                    'ecf_TerminaFechamentoCupom "Obrigado! Volte Sempre!"
                    strSQL = F_BuscaUmCampo("msgecf_par", TAB_PARAMETROS, "")
                    strSQL = strSQL & Chr(13) & nomecli & "-" & cpfcli
                    strSQL = strSQL & "-" & enderecocli
                    strSQL = strSQL & Chr(13) + Chr(10) & "Caixa: " & strCod_Cai & " Atendente: " & strNome_Usu
                    ecf_TerminaFechamentoCupom strSQL, responde
                    
                    
                    
                    'ecf_LeituraX
                    ecf_verificaImpressora
                    
                    iImpr = ImprimeGerencialDuasVias(valor, prim_via, seg_via, nomecli, cpfcli, enderecocli)
                    ImprimeGerencialDuasVias = iImpr
                    If iImpr = True Then
                        GoTo termino
                    End If
                    
                                
                    
            Loop
        Next
        If iVias = 1 And prim_via <> "" And seg_via <> "" Then
            ecf_UsaComprovanteNaoFiscalVinculado (" " & Chr(13) + Chr(10))
            ecf_UsaComprovanteNaoFiscalVinculado (" " & Chr(13) + Chr(10))
            'ecf_AcionaGuilhotina
            ecf_UsaComprovanteNaoFiscalVinculado (" " & Chr(13) + Chr(10))
        End If
    Next
termino:
    ecf_fechaComprovanteNaoFiscalVinculado
    
    ImprimeGerencialDuasVias = True
   Exit Function
termino_nao_conf:
ImprimeGerencialDuasVias = False
Exit Function
    
End Function

Function ImprimeGerencialDuasVias1(valor As String, prim_via As String, seg_via As String) As Boolean
    Dim c As New Collection
    Dim iCount As Integer
    Dim responde As Integer
    Dim Retorno As Integer
    Dim iVias As Integer
    Dim sTexto As String
    Dim iImpr As Boolean
    
    
    responde = -1
    
    
    responde = Ecf_ImpressoraNaoResponde
    If responde = 1 Then
        P_MsgOk "Transação não foi efetuada. Favor reter o Cupom"
        ImprimeGerencialDuasVias1 = False
        Exit Function
    ElseIf responde = 0 Then
        'ecf_LeituraX
        ImprimeGerencialDuasVias1 = ImprimeGerencialDuasVias1(valor, prim_via, seg_via)
        Exit Function
    End If
    
    
    
    
    ImprimeGerencialDuasVias1 = False

    
    ecf_FechaRelatorioGerencial
    For iVias = 1 To 2
        If iVias = 1 Then
            sTexto = prim_via
        Else
            sTexto = seg_via
        End If
        Set c = F_SeparaTexto(sTexto, Chr(10), False) 'impressão da primeira via
        For i = 1 To c.Count
            iCount = ecf_RelatorioGerencial(c(i)) 'desligar a impressora para testar
            iCount = ecf_RelatorioGerencial(Chr(13) + Chr(10))
            
            Retorno = Ecf_ImpressoraNaoResponde
            
            If Retorno = 0 Then
                'ecf_LeituraX
                ImprimeGerencialDuasVias1 = ImprimeGerencialDuasVias1(valor, prim_via, seg_via)
                Exit Function
            ElseIf Retorno = 1 Then
                ImprimeGerencialDuasVias1 = False
                Exit Function
            End If
            Do While ecf_FaltaPapel = True Or iFinalizouNao = True
                If F_EscolheSimNao("Falta Papel" & Chr(13) & _
                    "Impressora Não Responde. Tentar Novamente?") = 1 Then
                    iFinalizouNao = True
                    P_MsgOk "Transação não foi efetuada. Favor reter o Cupom"
                    
                    ImprimeGerencialDuasVias1 = False
                    
                    GoTo termino_nao_conf
                    Exit Function
                Else
                End If
                    ecf_fechaComprovanteNaoFiscalVinculado
                    'ecf_TerminaFechamentoCupom "Obrigado! Volte Sempre!"
                    'ecf_LeituraX
                    ecf_verificaImpressora
                    
                    iImpr = ImprimeGerencialDuasVias1(valor, prim_via, seg_via)
                    ImprimeGerencialDuasVias1 = iImpr
                    If iImpr = True Then
                        GoTo termino
                    End If
                    
                                
                    
            Loop
        Next
        If iVias = 1 And prim_via <> "" And seg_via <> "" Then
            ecf_UsaComprovanteNaoFiscalVinculado (" " & Chr(13) + Chr(10))
            ecf_UsaComprovanteNaoFiscalVinculado (" " & Chr(13) + Chr(10))
            'ecf_AcionaGuilhotina
            ecf_UsaComprovanteNaoFiscalVinculado (" " & Chr(13) + Chr(10))
        End If
    Next
termino:
    ecf_fechaComprovanteNaoFiscalVinculado
    
    ImprimeGerencialDuasVias1 = True
   Exit Function
termino_nao_conf:
ImprimeGerencialDuasVias1 = False
Exit Function
'    Exit Function
'
'    Set c = F_SeparaTexto(seg_via, Chr(10), False)
'    For i = 1 To c.count
'
'        iCount = ecf_UsaComprovanteNaoFiscalVinculado(c(i))
'        iCount = ecf_RelatorioGerencial(Chr(13) + Chr(10))
'        Retorno = Ecf_ImpressoraNaoResponde
'        If Retorno = 0 Then
'            ImprimeGerencialDuasVias1 = ImprimeGerencialDuasVias1(valor, prim_via, seg_via)
'        ElseIf Retorno = 1 Then
'            ImprimeGerencialDuasVias1 = False
'            Exit Function
'        End If
'
'
'    Next
'    ecf_FechaComprovanteNaoFiscalVinculado
'
'    ImprimeGerencialDuasVias1 = True
    
End Function

Sub ColetaCampo(ByVal Comando As Integer, _
                ByVal tipocampo As Long, _
                ByVal TamanhoMinimo As Integer, _
                ByVal TamanhoMaximo As Integer, _
                ByRef pDadosComando As String, _
                ByRef pCampo As String)

  Dim strComando       As String
  Dim strTipoCampo     As String
  Dim strTamanhoMinimo As String
  Dim strTamanhoMaximo As String

  DoEvents

  strComando = "Comando      :" & CStr(Comando)
  strTipoCampo = "TipoCampo    :" & CStr(tipocampo)
  strTamanhoMinimo = "TamanhoMinimo:" & CStr(TamanhoMinimo)
  strTamanhoMaximo = "TamanhoMaximo:" & CStr(TamanhoMaximo)

End Sub

Sub RecebeResultado(ByVal tipocampo As Long, _
                    ByRef pResultado As String)

  MsgBox "TipoCampo :" & CStr(tipocampo)
  MsgBox "pResultado:" & CStr(pResultado)

End Sub


Function efetua_venda(vlr_cartao As String, vlr_avista As String, vlr_aprazo As String, Optional desconto_ven As String = "0", _
    Optional nomecli As String = "", Optional cpfcli As String = "", Optional enderecocli As String = "") As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim i As Integer
    Dim Resposta As Integer
    Dim Retorno As Integer
        Retorno = 0
        i = F_RetornoVariaveis()
        If i = 2 Or i = 4 Then
            MsgBox "O Cupom fiscal não se encontra aberto!", vbExclamation
            efetua_venda = False
        End If
        bEfetuouFormaPagamentoCartao = False
        
        Retorno = Ecf_ImpressoraNaoResponde 'desligar 1ª vez
        If Retorno = 0 Then
            'MsgBox "Ligou e funcionou"
        ElseIf Retorno = 1 Then
'            MsgBox "Disse que não quer"
            efetua_venda = False
            Exit Function
        End If
        
        P_Msg "Iniciando fechamento do Cupom"
        
        i = ecf_IniciaFechamentoCupom("D", "$", desconto_ven)  'desconto sempre em dinheiro
        
        
        
        If vlr_avista > 0 Then
            P_Status "Efetuando pagamento A Vista"
            ecf_EfetuaFormaPagamento "Dinheiro", Format(vlr_avista, "###,##0.00")
        End If
        If vlr_aprazo > 0 Then
            P_Status "Efetuando pagamento A Prazo"
            ecf_EfetuaFormaPagamento "A Prazo", Format(vlr_aprazo, "###,##0.00")
        End If
        
        
        P_Msg "Efetuando pagamento cartão"
        
        'ecf_EfetuaFormaPagamento "Cartao", Format(vlr_cartao, "###,##0.00")
            
        
        'bEfetuouFormaPagamentoCartao = True
        
        P_Msg "Iniciando Transação TEF"
        
        
        Retorno = Ecf_ImpressoraNaoResponde 'desligar segunda vez
        If Retorno = 0 Then
        
        ElseIf Retorno = 1 Then
            'escolheu não
            efetua_venda = False
            FinalizaTransacaoSiTefInterativo 1, strNum_Cpm, Format(Date, "yyyymmdd"), Format(Time, "hhmmss")
            If bEfetuouFormaPagamentoCartao = True Then
                ecf_CancelaCupom
            End If
            Exit Function
            
        End If
        
        
        Do While ecf_FaltaPapel = True
            If F_EscolheSimNao("Falta Papel" & Chr(13) & "Tentar Novamente?") = 1 Then
                efetua_venda = False
                Exit Function
            End If
        Loop
        'aqui vai começar a chamada do cartão
        If EfetuaPagamento(vlr_cartao, strNum_Cpm, strCod_Usu, nomecli, cpfcli, enderecocli) = True Then
            efetua_venda = True
            'FinalizaTransacaoSiTefInterativo 0, strNum_Cpm, Format(Date, "yyyymmdd"), Format(Time, "hhmmss")
            
        Else
            efetua_venda = False
            'FinalizaTransacaoSiTefInterativo 1, strNum_Cpm, Format(Date, "yyyymmdd"), Format(Time, "hhmmss")
            Exit Function
        End If
        
        
        
        P_Msg "Processo Concluído"
        P_Msg ""
        
        P_Status ""
        
    
    efetua_venda = True
    
End Function


Sub P_GravaTransacaoTef(num_cpm As String, valor_tra As String)
    Dim strSQL As String
    'On Error Resume Next
    strSQL = "insert into tab_trtef(num_cpm,data_tra,ip_cai,valor_tra)values("
    strSQL = strSQL & "'" & num_cpm & "',"
    strSQL = strSQL & F_GravaDataTempo(sDataTr & " " & sHoraTr) & ","
    strSQL = strSQL & "'" & strIP_Usu & "',"
    strSQL = strSQL & F_GravaDec(valor_tra) & ")"
    conn.Execute strSQL
    
End Sub

Sub P_DeletaTransacaoTef(num_cpm As String, ip_cai As String)
    Dim strSQL As String
    On Error Resume Next
    strSQL = "delete from tab_trtef where num_cpm='" & num_cpm & "' "
    strSQL = strSQL & " and ip_cai='" & ip_cai & "'"
    conn.Execute strSQL
End Sub
Sub P_VerificaTefPendente()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sNum As String
    Dim sHora As String
    Dim sData As String
    Dim l As Long
    strSQL = "select * from tab_trtef where ip_cai='" & strIP_Usu & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            P_MsgOk "A Transação Efetuada. Favor re-imprimir último cupom", "Transação Tef"
            sData = Format(!data_tra, "yyyymmdd")
            sHora = Format(!data_tra, "hhmmss")
            sNum = !num_cpm
            
            'Do While l <> 10000
             '   l = IniciaFuncaoSiTefInterativo(110, !valor_tra, sNum, Format(Date, "yyyymmdd"), _
            '    Format(Time, "hhmmss"), strNome_Usu, "")
            'Loop
            FinalizaTransacaoSiTefInterativo 1, sNum, _
                Format(Date, "yyyymmdd"), Format(Time, "hhmmss")
                
                
            P_DeletaTransacaoTef !num_cpm, strIP_Usu
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
