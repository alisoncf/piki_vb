Attribute VB_Name = "tef_nfe"
Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

'constantes de diretório
Public Const DiretorioTEFreq = "C:\tef_dial\req\"
Public Const DiretorioTEFresp = "C:\tef_dial\resp\"
'Public Const DiretorioTEFreq = "C:\hipertef\req\"
'Public Const DiretorioTEFresp = "C:\hipertef\resp\"
Public strREDE As String
Public strAUTORIZACAO As String

Public buffer As String ' Buffer Generico

Public ValorTef As String  'Valor do Cartao a ser passado para o TEF
Public NumeroCupomTef As String  'Numero do Cupom Fiscal para a Compra do TEF
Public CnpjTransacaoTEF As String 'cnpj da bandeira
Public CodigaoAutorizacaoTEF As String 'cnpj da bandeira
Public RedeAutorizacaoTef As String

'constantes de diretório
'Declaracaoes da Dll de comunicacao Bematech   - BemaFi32.dll
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpAplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public UltimoCupomTEF As String





Public Enum ETEF_TipoParcelamento
    qualquer = 0
    a_vista = 1
    parcelado_emisso = 2
    parcelado_Estabelecimento = 3
    predatado = 4
    predatado_forcado = 5
    
End Enum





Type TransacaoTefResposta
    Comando As String '000-000
    identificacao As String '001-000
    documento_fiscal As String '002-000
    valor_total As String '003-000
    moeda As String '004-00
    Entidade_Cliente As String '006-000
    Identificador_Cliente As String '007-000
    
    status As String '009-000
    rede As String '010
    tipo_transacao As String '011-000
    nsu As String '012
    codigo_autorizacao As String '013-000
    Tipo_parcelamento As String '732
    
    parcelas As String '018-000
    dataComprovante As String '022-000
    Hora_comprovante As String '023-000
    tamanho_via As String '028
    via As String '028-xxx
    NomeCartaoOuAdm As String  '040
    CapacidadesAutomacao As String '706
    ValorOriginal707 As String '707
    ValorTroco708 As String
    ValorDesconto709 As String
    RegistroFinalizador As String
    mensagem As String '030-000
    
End Type

Type TransacaoTefRequest
    Comando As String '000-000
    identificacao As String '001-000
    documento_fiscal As String '002-000
    valor_total As String '003-000
    moeda As String '004-00
    Entidade_Cliente As String '006-000
    Identificador_Cliente As String '007-000
    rede As String '010
    nsu As String '012
    codigo_autorizacao As String '013-000
    qtd_parcelas As String '018-000
    dataComprovante As String '022-000
    Hora_comprovante As String '023-000
    data_preDatado As String '024-000
    codigo_controle As String '027-000
    indice_estabelecimento As String '702
    CapacidadesAutomacao As String '706
    desconto As String '709
    EmpresaAutomacao As String '716
    DataHoraFiscal As String '717
    DadosAdd1 As String '722
    DadosAdd2 As String '723
    DadosAdd3 As String '724
    DadosAdd4 As String '725
    IdiomaCliente As String '726
    TaxaServico As String '727
    TaxaEmbarque As String '728
    operacao As String '730
    Tipo_Cartao As String '731
    
    Tipo_Financiamento As String '732
    versao_interface As String '733-000
    nome_automacao As String '735
    versao_automacao As String '736
    registro_certificao As String '738
    indice_rede_adquirente As String '739
    RegistroFinalizador As String
    
    
    
    
End Type

Global TransacaoTEFAtual As TransacaoTefResposta
Global RequisicaoTEFAtual As TransacaoTefRequest


'-------------------------------------
'Funcao para Criar arquivo para o GP e Copiar para o GP
Public Function CriaArquivo(sBufferArquivo As String, sPath As String) As Boolean
On Error GoTo trata_Erro
'Cria temp para TEF
Dim arq As String

arq = App.path & "\intpos.tmp"



FechaArquivoExterno 1

If Dir(DiretorioTEFresp & "intpos.001") <> "" Then
'    MsgBox "Já existe um arquivo no servidor de destino do TEF", vbExclamation
'    CriaArquivo = False
'    Exit Function
End If

If Dir(arq) <> "" Then
    Kill arq
End If
If Dir(sPath & "INTPOS.001") <> "" Then
    Kill sPath & "INTPOS.001"
End If



Open arq For Binary As #1
    Put #1, , sBufferArquivo
Close #1
'Copia para o Diretorio do GP
'Open sPath & "INTPOS.001" For Binary As #1
'    Put #1, , sBufferArquivo
'Close #1

'On Error Resume Next

FileCopy App.path & "\intpos.tmp", sPath & "INTPOS.001"
CriaArquivo = True
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
    

     Do While ((Dir(DiretorioTEFresp & sNome)) = "")
If Minutos = 1 Then
            TimeIni = Minute(Time)
            If (TimeIni >= TimeFinal) Then: EsperaArquivo = -1: Exit Function 'Retorna Time Out
        Else
            EsperaArquivo = -1: Exit Function 'Retorna Time Out
        End If
        
    Loop

    EsperaArquivo = 1
End Function
Function ContadorTEF()
    UltimoCupomTEF = ChaveSQL("tab_tef", "contador_tef", "0######")
    ContadorTEF = UltimoCupomTEF
End Function






Public Function PreparaVenda(sValor As String, sNumeroCupom As String, _
    Optional rede As String = "", Optional operacao As String = "", Optional Tipo_Cartao As String = "", _
    Optional Tipo_Financiamento As String = "", Optional parcelas As String, Optional moeda As String = "0", _
    Optional capacidades As String = "") As Boolean
    Dim s As String
    If moeda = "" Then
        moeada = "0"
    End If
    P_Status "Enviando dados TEF"
    
    RequisicaoTEFAtual.Comando = "CRT"
    RequisicaoTEFAtual.identificacao = ContadorTEF
    RequisicaoTEFAtual.valor_total = sValor
    RequisicaoTEFAtual.moeda = moeda
    RequisicaoTEFAtual.rede = rede
    
    RequisicaoTEFAtual.qtd_parcelas = parcelas
    RequisicaoTEFAtual.CapacidadesAutomacao = capacidades
    RequisicaoTEFAtual.valor_total = sValor
    RequisicaoTEFAtual.operacao = operacao
    RequisicaoTEFAtual.Tipo_Cartao = Tipo_Cartao
    RequisicaoTEFAtual.Tipo_Financiamento = Tipo_Financiamento
    
    PreparaVenda = tef_nfe.EnviaComandoRequisicaoTEFVenda(RequisicaoTEFAtual)
    
    
    Exit Function
    s = "000-000 = CRT" + Chr(13) + Chr(10)
    s = s & "001-000 = " + ContadorTEF + Chr(13) + Chr(10)
    s = s & "002-000 = " + F_SoNumero(sNumeroCupom) + Chr(13) + Chr(10)
    s = s & "003-000 = " + FormatValTef(Trim(sValor)) + Chr(13) + Chr(10)
    s = s & "004-000 = " & moeda + Chr(13) + Chr(10)
    If rede <> "" Then
        s = s & "010-000 = " & rede + Chr(13) + Chr(10)
    End If
    If parcelas <> "" Then
        s = s & "018-000 = " & parcelas + Chr(13) + Chr(10)
    End If
    If capacidades <> "" Then
        s = s & "706-000 = " & capacidade + Chr(13) + Chr(10)
    End If
    
    
    
    If operacao <> "" Then
        s = s & "730-000 = " & operacao + Chr(13) + Chr(10)
    End If
    If Tipo_Cartao <> "" Then
        s = s & "731-000 = " & Tipo_Cartao + Chr(13) + Chr(10)
    End If
    If Tipo_Financiamento <> "" Then
        s = s & "732-000 = " & Tipo_Financiamento + Chr(13) + Chr(10)
    End If
    
    s = s & "733-000 = " & aspas & "V2.20" & aspas + CrLf
    s = s & "735-000 = " & aspas & "Piki" & aspas + CrLf
    s = s & "736-000 = " & aspas & "2023" & aspas + CrLf
    s = s & "738-000 = " + CrLf
    
    s = s & "999-999 = 0"
    P_Status "Criando arquivo para transação"
    If CriaArquivo(s, DiretorioTEFreq) Then
    
        buffer = s
        PreparaVenda = True
    Else
        PreparaVenda = False
    End If
    P_Status ""
End Function
Public Function FinalizaTef(valor As String) As Boolean
    Dim str As String
    Dim iCount As Integer
    Dim retorno As Integer
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
        buffer = ""
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
                buffer = buffer + Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)  'Nao pego os campos e nao pego as aspas
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
    Open App.path & "\TEFTEMP1.TXT" For Binary As #1
    Put #1, , buffer
    Close #1

'Abrir o Arquivo de BackUp e vamos jogar linha a linha na Impressora
'aqui devo buscar se a transação foi ou não efetuada
    
    
    
    
    
    retorno = 0
    Do While retorno <> 1
        Trava_Teclado
        retorno = ecf_VerificaImpressoraLigada
        If retorno <> 1 Then
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
        Open App.path & "\TEFTEMP1.TXT" For Input As #1
        
        Do While Not EOF(1)
        
    
    
            Line Input #1, sBufArq
            retorno = 0
            Do While retorno <> 1 Or ist1 >= 128
                
                
                
                retorno = ecf_RetornoImpressora(iAck, ist1, ist2)
                
                
                If retorno <> 1 Or ist1 >= 128 Then
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
    Dim retorno As Integer
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
        buffer = ""
        Do While Not EOF(1)
    
            Line Input #1, sBufArq
            If (Mid(sBufArq, 1, 3) = "030") Then
                
                Exit Do
            End If
            If (Mid(sBufArq, 1, 3) = "009") Then
                s009 = Trim(Mid(sBufArq, 10))
            End If
            If (Mid(sBufArq, 1, 3) = "028") Then
                
                sMsg = sMsg + Mid(sBufArq, 10) + Chr(13) + Chr(10)
            End If
            If (Mid(sBufArq, 1, 3) = "029") Then
                buffer = buffer + Mid(sBufArq, 12, Len(sBufArq) - 12) + Chr(13) + Chr(10)  'Nao pego os campos e nao pego as aspas
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
    Open App.path & "\TEFTEMP1.TXT" For Binary As #1
    Put #1, , buffer
    Close #1

    
    
    
    
    
    If CInt(s009) = 0 Then
    
    
        ConfirmaVenda
    
        ImprimeComprovanteNaoFiscalTEF buffer, True
        FinalizaTefNaoFiscal = True
    Else
        FinalizaTefNaoFiscal = False
    
    End If
    
    
    
    Habilita_Teclado
    
    
    
    
    
    
    FinalizaTefNaoFiscal = True
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
Dim buffer As String
On Error Resume Next
Open DiretorioTEFresp & "INTPOS.001" For Input As #1
P_Status "Confirmando transação"
buffer = ""
Buffer2 = ""
   Do While Not EOF(1)
      Line Input #1, sBufArq
      If Mid(sBufArq, 1, 3) = "001" Or _
         Mid(sBufArq, 1, 3) = "010" Or _
         Mid(sBufArq, 1, 3) = "012" Or _
         Mid(sBufArq, 1, 3) = "027" Or _
         Mid(sBufArq, 1, 3) = "999" Then
            buffer = buffer + sBufArq & Chr(13) & Chr(10)
      End If
   Loop
  
   Close #1
   
   Buffer2 = "000-000 = CNF" & Chr(13) & Chr(10) + buffer
   
   CriaArquivo Buffer2, DiretorioTEFreq
   

killer:
   Sleep (500)
   
    If ((Dir(DiretorioTEFresp & "INTPOS.001")) <> "") Then
        Kill DiretorioTEFresp & "INTPOS.001"
    Else
        GoTo killer
    End If
   On Error Resume Next
   Kill App.path & "\TEFTEMP1.TXT"
   Kill App.path & "\intpos.tmp"
   killIntPosResp
  'MsgBox "Autorização Efetivada! Venda Confirmada", vbInformation
   P_Status ""
End Sub

Public Sub NaoConfirmaVenda()
Dim buffer As String
Dim Buffer2 As String
Dim sNsu As String
Dim sRede As String
Dim sValor As String
Dim sMsg As String
Dim sOp As String

Open DiretorioTEFresp & "INTPOS.001" For Input As #1
   
   buffer = ""
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
            
            
            
            buffer = buffer + sBufArq & Chr(13) & Chr(10)
      End If
   Loop
  
   Close #1
   
   Buffer2 = "000-000 = NCN" & Chr(13) & Chr(10) + buffer
   
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
    Kill App.path & "\TEFTEMP1.TXT"
    Kill App.path & "\intpos.tmp"
    'Bematech_FI_FechaRelatorioGerencial
    
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
        'Open "C:\\TEFTEMP1.TXT" For Input As #1
        Open DiretorioTEFresp & "intpos.001" For Input As #1
        
        Do While Not EOF(1)
        
            Line Input #1, sBufArq
            If Left(sBufArq, 3) = "029" Then
                Bematech_FI_RelatorioGerencial Mid(sBufArq, 12, Len(sBufArq) - 12)
            End If
            retorno = Bematech_FI_VerificaImpressoraLigada() ' ARRUMEI AKI
            Do While retorno <> 1
                retorno = Bematech_FI_VerificaImpressoraLigada()
                If retorno <> 1 Then
                    Habilita_Teclado
                    If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                        strOK = False
                        Close #1
                        Exit Function
                    End If
                Else
                        Trava_Teclado
                        Close #1
                        ImprimeGerencial
                        strOK = True
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
    
    Dim s As String
    Trava_Teclado
    
    
        
                
        Open DiretorioTEFresp & "Intpos.001" For Input As #1
        
        Do While Not EOF(1)
            Trava_Teclado
            Line Input #1, sBufArq
            
            s = s & sBufArq & CrLf
            
        Loop
        
        Close #1
    
    
    
    ImprimeComprovanteNaoFiscalTEF s, True
    
    Habilita_Teclado
    killIntPosResp
    ImprimeGerencialADM = True
    
End Function

Function F_VerificaGP() As Boolean
    On Error GoTo trata_Erro
    'P_ContadorTEF
    P_Status "Criando arquivo..."
    buffer = "000-000 = ATV" + Chr(13) + Chr(10) + _
             "001-000 =" + ContadorTEF + Chr(13) + Chr(10) + _
             "999-999 = 0"
             
    CriaArquivo buffer, DiretorioTEFreq
    P_Status "Aguardando Retorno..."
    retorno = EsperaArquivo("INTPOS.STS", 6, 0)
    P_Status ""
    If (retorno = -1) Then
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
    Dim retorno As Integer
    Dim buffer As String
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
    
    
    Dim Cont As String
    Cont = ContadorTEF
    P_Status "Enviando dados para o Gerenciador Padrão..."
    buffer = "000-000 = ADM" + Chr(13) + Chr(10) + _
             "001-000 = " + Cont + Chr(13) + Chr(10) + _
             "999-999 = 0"
             
    CriaArquivo buffer, DiretorioTEFreq
    retorno = EsperaArquivo("Intpos.001", 5, 1)
    P_Status ""
    bOk = True
    sText = ""
    If (retorno = 1) Then
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
            ImprimeComprovanteNaoFiscalTEF sText
            ConfirmaVenda
            
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
        On Error Resume Next
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


Function F_RecebNaoFiscalTef(sIndice As String, sValor As String) As Boolean
    'Esta função emite o cupom fiscal da nota
    
    Dim sBufArq As String
    Dim buffer As String
    Dim bOk As Boolean
    Dim sMsg As String
    Dim retorno As Integer
    
    
    Do While retorno <> 1
        
        retorno = Bematech_FI_VerificaImpressoraLigada()
        If retorno <> 1 Then
            Habilita_Teclado
            If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                F_RecebNaoFiscalTef = False
                Exit Function
            End If
        End If
    Loop
    
    retorno = Bematech_FI_RecebimentoNaoFiscal(sIndice, sValor, "cartao")
    If retorno <> 1 Then
        MsgBox "O Recebimento não foi efetuado!", vbCritical
        F_RecebNaoFiscalTef = False
        Exit Function
    End If
    retorno = 0 'colou aki
    
    
    NumeroCupomTef = Space(14)
    Bematech_FI_NumeroCupom (NumeroCupomTef)
    ValorTef = sValor

    PreparaVenda Format(ValorTef / 100, "###,##0.00"), NumeroCupomTef
    
    DoEvents
    
    retorno = EsperaArquivo("INTPOS.001", 5, 1) 'baixar para sete segundos
    
    If retorno = -1 Then
        
        MsgBox "Gerenciador padrão não está ativo!", vbCritical + vbSystemModal
        Exit Function
    End If

    Trava_Teclado
    If (retorno = 1 Or -1) Then
        Open DiretorioTEFresp & "INTPOS.001" For Input As #1
            buffer = ""
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
            retorno = 0
            Do While retorno <> 1
                Trava_Teclado
                retorno = Bematech_FI_VerificaImpressoraLigada()
                If retorno <> 1 Then
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
            retorno = 0
            Do While retorno <> 1
                Trava_Teclado
                retorno = Bematech_FI_VerificaImpressoraLigada()
                If retorno <> 1 Then
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
            Kill App.path & "\TEFTEMP1.TXT"
            Kill App.path & "\intpos.tmp"
            
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
    Dim buffer As String
    Dim bOk As Boolean
    Dim sMsg As String
    Dim retorno As Integer
    Dim ack As Integer
    Dim st1 As Integer
    Dim st2 As Integer
    'Dim rImp As retorno_impressora
    On Error GoTo trata_Erro
    killIntPosResp
    If strNum_Cpm = "" Then
        MsgBox "Não há um cupom fiscal Gerado!", vbExclamation
        Exit Function
    End If
    
    Do While retorno <> 1
        P_Msg "Conectando com ECF", "Transferência Eletrônica de Fundos"
        retorno = ecf_VerificaImpressoraLigada
        If retorno <> 1 Then
            Habilita_Teclado
            If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                P_Msg "", ""
                F_VendaCupomTEF = False
                Exit Function
            End If
        End If
    Loop
    
        
    retorno = 0 'colou aki
    Do While retorno <> 1
        P_Msg "Conectando com ECF", "Transferência Eletrônica de Fundos"
        retorno = ecf_VerificaImpressoraLigada()
        If retorno <> 1 Then
            
            If MsgBox("Impressora não responde! tentar novamente?", vbYesNo + vbQuestion + vbSystemModal) = vbNo Then
                F_VendaCupomTEF = False
                P_Msg "", ""
                Exit Function
            End If
        End If
    Loop
    If operacao = "B" Then
        'rImp = ecf_retornoImprComandos()
        
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
    retorno = EsperaArquivo("INTPOS.001", 5, 1) 'baixar para sete segundos
    
    If retorno = -1 Then
        P_Msg
            MsgBox "Gerenciador padrão não está ativo!", vbCritical + vbSystemModal
        Exit Function
    End If
    
    Trava_Teclado
    If (retorno = 1 Or -1) Then
        Open DiretorioTEFresp & "INTPOS.001" For Input As #1
            buffer = ""
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
            retorno = 0
            Do While retorno <> 1
                Trava_Teclado
                P_Msg "Conectando ECF", "Transferência Eletrônica de Fundos"
                retorno = ecf_VerificaImpressoraLigada()
                If retorno <> 1 Then
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
            
            retorno = 0
            Do While retorno <> 1
                Trava_Teclado
                retorno = ecf_VerificaImpressoraLigada()
                If retorno <> 1 Then
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
            ''ecf_TerminaFechamentoCupom constMsg_Usu1
            
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
            Kill App.path & "\TEFTEMP1.TXT"
            Kill App.path & "\intpos.tmp"
            
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
Function TefNaoFiscalSemECF(sValor As String, Optional rede As String, Optional operacao As String, _
    Optional Tipo_Cartao As String, Optional Tipo_Financiamento As String, Optional parcelas As String, _
    Optional moeda As String = "0", Optional capacidade As String = "", Optional BImprimeComprovante As Boolean = False) As Boolean
    'Esta função emite o cupom fiscal da nota
    
    Dim sBufArq As String
    Dim buffer As String
    Dim bOk As Boolean
    Dim sMsg As String
    Dim retorno As Integer
    Dim Resposta As String
    Dim NumeroCupomTef As String
    NumeroCupomTef = ContadorTEF
        
        
    ValorTef = FormatVal(sValor)
    

    If Not PreparaVenda(ValorTef, NumeroCupomTef, rede, operacao, Tipo_Cartao, _
        Tipo_Financiamento, parcelas, moeda, capacidade) Then
        Habilita_Teclado
        Exit Function
    End If
    
    'DoEvents
    
    
    TransacaoTEFAtual = BuscaRetornoTef
    
    If TransacaoTEFAtual.status <> "0" Then
        MsgBox "Erro na transação", vbExclamation
        NaoConfirmaVenda
        Exit Function
        
    Else
        TefNaoFiscalSemECF = True
        If BImprimeComprovante Then
            ImprimeComprovanteNaoFiscalTEF TransacaoTEFAtual.via, True
                        
            TefNaoFiscalSemECF = True
        
        End If
        ConfirmaVenda
    End If
    
    
    
    
    
    
    Exit Function
    
    
    
    
    retorno = EsperaArquivo("INTPOS.001", 5, 1) 'baixar para sete segundos
    
    If retorno = -1 Then
        
        MsgBox "Gerenciador padrão não está ativo!", vbCritical + vbSystemModal
        Exit Function
    End If

    
    
    
        
    
    
    Trava_Teclado
    Sleep (500)
    
    If (retorno = 1 Or -1) Then
        Open DiretorioTEFresp & "INTPOS.001" For Input As #1
            buffer = ""
            bOk = True
            Do While Not EOF(1)
    
                Line Input #1, sBufArq
                If (Mid(sBufArq, 1, 3) = "009") Then
                    If Right(sBufArq, 1) <> "0" Then
                        'MsgBox "Não foi possível concluir a transação!", vbInformation
                        bOk = False
                    End If
                    Resposta = Trim(Mid(sBufArq, 10))
                    TransacaoTEFAtual.status = Resposta
                    
                    
                End If
                If (Mid(sBufArq, 1, 3) = "030") Then
                    sMsg = Mid(sBufArq, 10, Len(sBufArq) - 9)
                    TransacaoTEFAtual.mensagem = sMsg
                End If
                If (Mid(sBufArq, 1, 3) = "013") Then
                    CodigaoAutorizacaoTEF = Trim(Mid(sBufArq, 10))
                    TransacaoTEFAtual.codigo_autorizacao = CodigaoAutorizacaoTEF
                End If
                If (Mid(sBufArq, 1, 3) = "010") Then
                    RedeAutorizacaoTef = Trim(Mid(sBufArq, 10))
                    TransacaoTEFAtual.rede = RedeAutorizacaoTef
                
                End If
            Loop
            
        Close #1
            
        
        If Resposta <> "0" Then
            TefNaoFiscalSemECF = False
            
            MsgBox sMsg, vbExclamation
            Kill DiretorioTEFresp & "intpos.001"
        
            
            
            Exit Function
        End If
        
            
            
            
        If FinalizaTefNaoFiscal() Then ' abre a emissão dos comprovantes TEF
                
                TefNaoFiscalSemECF = True
            
            Habilita_Teclado
        Else
            Habilita_Teclado
            strSQL = "Não foi possível concluir a transação!" & Chr(13)
            strSQL = strSQL & "Mensagem: '" & sMsg & "'" & Chr(13)
            
            MsgBox strSQL, vbCritical + vbSystemModal
            'Bematech_FI_CancelaCupom
            'NaoConfirmaVenda ' - so mandar ncn qd. desl. micro\imp.
            TefNaoFiscalSemECF = False
            Habilita_Teclado
            If ((Dir(DiretorioTEFresp & "INTPOS.001")) <> "") Then
                Kill DiretorioTEFresp & "INTPOS.001"
            End If
            On Error Resume Next
            
            Kill App.path & "\TEFTEMP1.TXT"
            Kill App.path & "\intpos.tmp"
            
            
            
            
        End If
    End If
    
Habilita_Teclado
End Function

Function BuscaRetornoTef() As TransacaoTefResposta
    Dim s As String
    Dim sBufArq As String
    Dim buffer As String
    Dim retorno As Integer
    Dim ret_com As String
    P_Status "Buscando arquivo de retorno TEF... "
    retorno = EsperaArquivo("INTPOS.001", 5, 1) 'baixar para sete segundos
    
    If retorno = -1 Then
        
        MsgBox "Gerenciador padrão não está ativo!", vbCritical + vbSystemModal
        Exit Function
    End If

    
    
    Trava_Teclado
    Sleep (500)
    
    If (retorno = 1 Or -1) Then
        
        ret_com = LeArquivoExternoUTF8(DiretorioTEFresp & "INTPOS.001")
        Open DiretorioTEFresp & "INTPOS.001" For Input As #1
            buffer = ""
            
            Do While Not EOF(1)
    
                Line Input #1, sBufArq
                If (Mid(sBufArq, 1, 3) = "000") Then
                    BuscaRetornoTef.Comando = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "001") Then
                    BuscaRetornoTef.identificacao = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "002") Then
                    BuscaRetornoTef.documento_fiscal = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "003") Then
                    BuscaRetornoTef.valor_total = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "004") Then
                    BuscaRetornoTef.moeda = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "006") Then
                    BuscaRetornoTef.Entidade_Cliente = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "007") Then
                    BuscaRetornoTef.Identificador_Cliente = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "009") Then
                    BuscaRetornoTef.status = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "010") Then
                    BuscaRetornoTef.rede = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "011") Then
                    BuscaRetornoTef.tipo_transacao = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "012") Then
                    BuscaRetornoTef.nsu = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "013") Then
                    BuscaRetornoTef.codigo_autorizacao = Trim(Mid(sBufArq, 10))
                    CodigaoAutorizacaoTEF = BuscaRetornoTef.codigo_autorizacao
                    
                End If
                If (Mid(sBufArq, 1, 3) = "017") Then
                    BuscaRetornoTef.Tipo_parcelamento = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "018") Then
                    BuscaRetornoTef.parcelas = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "022") Then
                    BuscaRetornoTef.dataComprovante = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "023") Then
                    BuscaRetornoTef.Hora_comprovante = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "028") Then
                    BuscaRetornoTef.tamanho_via = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "029") Then
                    BuscaRetornoTef.via = BuscaRetornoTef.via & Trim(Mid(sBufArq, 10)) + CrLf
                End If
                If (Mid(sBufArq, 1, 3) = "040") Then
                    BuscaRetornoTef.NomeCartaoOuAdm = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "707") Then
                    BuscaRetornoTef.ValorOriginal707 = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "708") Then
                    BuscaRetornoTef.ValorTroco708 = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "709") Then
                    BuscaRetornoTef.ValorDesconto709 = Trim(Mid(sBufArq, 10))
                End If
                If (Mid(sBufArq, 1, 3) = "999") Then
                    BuscaRetornoTef.RegistroFinalizador = Trim(Mid(sBufArq, 10))
                End If
                
            Loop
            
        Close #1
        BuscaRetornoTef.via = mudarCaractere(BuscaRetornoTef.via, aspas, "")

End If
P_Status ""
End Function


Function RetornaConteudoSemTagsTEF(conteudo As String, Optional tag As String = "029") As String
    Dim i As Integer
    Dim inicio As Integer
    Dim fim As Integer
    Dim s As String
    Dim iAux As Integer
    iAux = 1
    
    Do While iAux <> 0
        inicio = InStr(1, conteudo, tag) & Format(i, "00#")
    Loop
    
    
    
    

End Function


Sub ImprimeComprovanteNaoFiscalTEF(Arquivo As String, Optional visualiza As Boolean = True, Optional limpaArquivo As Boolean = False)
    AbreArquivoExterno Arquivo, 1
    Dim obj As Object
    Dim pPold As Integer
    pOld = BuscaImpressoraPadrao
    Set Printer = Printers(ImpressoraPadraoCupomNFCe.Indice)
    Set obj = frmVisualiza.Pic
inicio:
    If visualiza = True Then
        P_CentraImpressao obj, "Comprovante de Transação"
        
        
        
        obj.Print Arquivo
        frmVisualiza.Show 1
        
    Else
        
        obj.Print Arquivo
        obj.Print "__________________________________"
        obj.Print ""
        obj.Print "2a Via Cliente."
        
        obj.Print Arquivo
        Printer.EndDoc
        frmVisualiza.tag = ""
    End If
    
    If frmVisualiza.tag <> "" Then
        Set obj = Printer
        visualiza = False
        GoTo inicio
    End If
    
    Unload frmVisualiza
    
    
    FechaArquivoExterno 1
    
    Set Printer = Printers(pOld)
    
    
End Sub

Function EnviaComandoRequisicaoTEFVenda(Comando As TransacaoTefRequest)
    
    Dim s As String
    Dim CrLf As String
    CrLf = Chr(13) + Chr(10)
    
    If Comando.qtd_parcelas <> "" Then
        If CInt(Comando.qtd_parcelas) > 1 And Comando.Tipo_Financiamento = "" Then
            Comando.Tipo_Financiamento = "3" 'financiamento pelo lojista
        End If
    End If
    
    s = "000-000 = CRT" + CrLf
    s = s & "001-000 = " + Comando.identificacao + CrLf
    If Comando.documento_fiscal <> "" Then
        s = s & "002-000 = " + Comando.documento_fiscal + CrLf
    End If
    
    s = s & "003-000 = " + FormatValTef(Comando.valor_total) + CrLf
    
    s = s & "004-000 = " & Comando.moeda + CrLf
    If Comando.Entidade_Cliente <> "" Then
        s = s & "006-000 = " & Comando.Entidade_Cliente + CrLf
    End If
    If Comando.Identificador_Cliente <> "" Then
        s = s & "007-000 = " & Comando.Identificador_Cliente + CrLf
    End If
    If Comando.rede <> "" Then
        s = s & "010-000 = " & Comando.rede + CrLf
    End If
    If Comando.qtd_parcelas <> "" Then
        s = s & "018-000 = " & Comando.qtd_parcelas + CrLf
    End If
    If Comando.data_preDatado <> "" Then
        s = s & "024-000 = " & Comando.data_preDatado + CrLf
    End If
    If Comando.indice_estabelecimento <> "" Then
        s = s & "702-000 = " & Comando.indice_estabelecimento + CrLf
    End If
    If Comando.desconto <> "" Then
        s = s & "709-000 = " & Comando.desconto + CrLf
    End If
    If Comando.CapacidadesAutomacao <> "" Then
        s = s & "706-000 = " & Comando.CapacidadesAutomacao + CrLf
    End If
    If Comando.EmpresaAutomacao <> "" Then
        s = s & "716-000 = " & Comando.EmpresaAutomacao + CrLf
    End If
    If Comando.DataHoraFiscal <> "" Then
        s = s & "717-000 = " & Comando.DataHoraFiscal + CrLf
    End If
    If Comando.DadosAdd1 <> "" Then
        s = s & "722-000 = " & Comando.DadosAdd1 + CrLf
    End If
    If Comando.DadosAdd2 <> "" Then
        s = s & "723-000 = " & Comando.DadosAdd2 + CrLf
    End If
    If Comando.DadosAdd3 <> "" Then
        s = s & "724-000 = " & Comando.DadosAdd3 + CrLf
    End If
    If Comando.DadosAdd4 <> "" Then
        s = s & "725-000 = " & Comando.DadosAdd4 + CrLf
    End If
    If Comando.IdiomaCliente <> "" Then
        s = s & "726-000 = " & Comando.IdiomaCliente + CrLf
    End If
    If Comando.TaxaServico <> "" Then
        s = s & "727-000 = " & Comando.TaxaServico + CrLf
    End If
    If Comando.TaxaEmbarque <> "" Then
        s = s & "728-000 = " & Comando.TaxaEmbarque + CrLf
    End If
    If Comando.operacao <> "" Then
        s = s & "730-000 = " & Comando.operacao + CrLf
    End If
    
    If Comando.Tipo_Cartao <> "" Then
        s = s & "731-000 = " & Comando.Tipo_Cartao + CrLf
    End If
    If Comando.Tipo_Financiamento <> "" Then
        s = s & "732-000 = " & Comando.Tipo_Financiamento + CrLf
    End If
    If Comando.versao_interface <> "" Then
        s = s & "733-000 = " & Comando.versao_interface + CrLf
    End If
    If Comando.nome_automacao <> "" Then
        s = s & "735-000 = " & Comando.nome_automacao + CrLf
    End If
    If Comando.versao_automacao <> "" Then
        s = s & "736-000 = " & Comando.versao_automacao + CrLf
    End If
    
    If Comando.registro_certificao <> "" Then
        s = s & "738-000 = " & Comando.registro_certificao + CrLf
    End If
    If Comando.indice_rede_adquirente <> "" Then
        s = s & "739-000 = " & Comando.indice_rede_adquirente + CrLf
    End If
    s = s & "999-999 = 0"
    P_Status "Criando arquivo para transação"
    If CriaArquivo(s, DiretorioTEFreq) Then
    
        buffer = s
        EnviaComandoRequisicaoTEFVenda = True
    Else
        EnviaComandoRequisicaoTEFVenda = False
    End If
    P_Status ""
    
End Function


Sub ImprimeSegundaViaCartao(num_ven As String, Optional vias As Integer = 1)
    Dim comprovante As String
    comprovante = F_BuscaUmCampo("mensagem", "tab_venfin", "num_Ven='" & num_ven & "'")
    
    ImprimeComprovanteNaoFiscalTEF comprovante
    
End Sub
