Attribute VB_Name = "Modulo"


Global bUsaTEF As Boolean
Global bUsaECF As Boolean


Global trata_Erro As label
Global Banco As database




Global bDedicado As database

Global sDriver As String
Global sServer As String
Global sUid As String
Global sDataBase As String
Global sPwd As String
Global sOption As String

Global strNomeSis As String
Global strNum_Cpm As String
Global intCUPOM As Integer
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

Global strCancelarItem_Usu As String
Global strCancelaVenda_Usu As String


Global strIP_Usu As String 'o ip da mákina
Global strHost_Usu As String 'nome do computador
Global strCod_Cai As String 'o código do caixa buscado pelo ip
Global strCod_Tur As String
Global strSerie_Nfe As String 'serie do caixa
Global crcCpm As Currency 'o valor da venda Tj blue
Global strNome_Usu As String
Global strNivel_Usu As String
Global strCod_Usu As String

'empresa
Global strCorFundoUsu As String
Global strNome_Emp As String
Global strCod_Emp As String
Global strCaminho_Logo As String


Global strFantasia_Emp As String
Global strEndereco_Emp As String
Global strBairro_Emp As String
Global strCidade_Emp As String
Global strComplemento_Emp As String
Global strNumero_Emp As String
Global strLogradouro_Emp  As String
'sped e nfe
Global strCodigoMunicipioIBGE As String
Global strCodigoEstadoIBGE As String
Global strCodigoPaisIBGE As String
Global strPais As String

Global strUF_Emp As String
Global strCEP_Emp As String
Global strFone_Emp As String
Global strFax_Emp As String
Global strCNPJ_Emp As String

Global strIE_Emp As String
Global strCNAE_Emp As String
Global strEmail_emp As String
Global strIEST_Emp As String
Global strIM_Emp As String
Global strCRT_emp As CRT
Global strCor_Emp As String
Global strTipo_IcmsNFe As String
Global strCST_Padrao As String
Global strCST_IPI As String
Global strCST_PIS As String
Global strCST_COFINS As String
Global strFormaPagamentoPadrao As String
Global strNomeResp_Emp As String
Global strCRCResp_Emp As String
Global strNomeMoeda_Par As String
Global strsimbMoeda_Par As String
Global strPluraMoeda_Par As String
Global crcTxFederativa_Par As Currency
Global crcTxEsdadual_Par As Currency
Global crcTxMargemSeg_Par As Currency
Global crcTxAdm As Currency
Global crcTxDiversas As Currency
Global crcComissVen As Currency 'comissão sobre as vendas
Global crcDesconto_Avista As Currency
Global strArquivoServidor As String

Global strHora As String 'teste
Global strNomeBanco As String
Global strLocalBanco As String
Global strUserBanco As String
Global strEmpresaLogada As String
Global strCnpjEmpresaLogada As String
Global strFantasiaEmpresaLogada As String
Global strUsaComissaoAcerto As String
Global bUsaEstoqueSeparado As Boolean

Global strIP_Servidor As String
Global strIP_TerminalTEF As String
Global strId_TerminalTef As String
Global strId_Estab As String

Global picVis(9)  As PictureBox

Global strPlaVendasEfetuadas As String
Global strPlaRecebDupl As String
Global strPlaPagDupl As String
Global strPlaNaoClassD As String
Global strPlaNaoClassC As String
Global CrLf As String



Public Const constContaCaixa = "001" 'CÓDIGO DA CONTA CAIXA PADRÃO

Public Enum caso
    nada_muda = 0
    Minusculo = 1
    primeira_maiscula = 2
    Maiusculo = 3
End Enum

Public Enum operacao_um_campo
    cst_icms = 1
    CST_IPI = 2
    cst_pis = 3
    
End Enum


Type DataInicialFinal
    dataI As String
    dataF As String
End Type





Public Type impressao

    PagIni As Integer
    PagFin As Integer
    Qualidade As Integer
End Type

Public Type valChart ' para abrir o grafico valor, caption
    Value As Long
    caption As String
End Type

Public Type BlChart ' para carregar o valor de cada coluna, e quantas colunas tem
    val() As valChart
    Count As Long
End Type






Sub P_LocalizaOperacao(sOp As String, obj As Object)
    Dim iCount As Integer
    Dim iLen As Integer
    iLen = Len(sOp)
    On Error Resume Next
    With obj
        For iCount = 0 To .ListCount - 1
            obj.ListIndex = iCount
            If Left(obj.Text, iLen) = sOp Then
                Exit Sub
            End If
        Next
        'MsgBox "Operação não encontrada!", vbExclamation
        obj.ListIndex = -1
        obj.SetFocus
    End With
End Sub
Function SoLetraENumero(Texto As String) As String
    Dim s As String
    Dim i As Integer
    Dim s2 As String
    
    
    For i = 1 To Len(Texto)
        s = Mid(Texto, i, 1)
        If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWYZ", UCase(s)) > 0 Then
            s2 = s2 & s
        End If
    Next
    SoLetraENumero = s2
End Function

Function AjustaPlaca(placa As String) As String
    AjustaPlaca = Format(SoLetraENumero(UCase(placa)))
End Function
Sub P_AjustaData(objData, Optional sFormato As String = "dd/mm/yyyy")

    On Error GoTo TrataErro
    If objData.Text = "" Then Exit Sub
     
    objData.Text = Left(objData.Text, 10)
    If InStr(1, objData.Text, "/") <> 0 Then
        objData.Text = Format(objData.Text, sFormato)
    Else
        If Len(objData.Text) <> 8 Then
            MsgBox "Sua digitação da Data está Errada!", vbExclamation
            objData.SetFocus
            Exit Sub
        Else
            objData.Text = Left(objData.Text, 2) & "/" & Mid(objData.Text, 3, 2) & "/" & Right(objData.Text, 4)
        End If
    End If
    
    If Len(objData.Text) = 7 Then
    Else
    If Not F_Vazio(objData.Text) And _
        F_ValidaData(objData.Text) = "" Then
       MsgBox "Data inválida!", vbExclamation
       objData.SetFocus
       Exit Sub
    End If
    End If
    
    If CDate(objData.Text) < "31/12/1900" Then
        
        MsgBox "Data inválida!,Verifique o Texto Digitado", vbExclamation
        objData.SetFocus
    End If
    Exit Sub

TrataErro:

    MsgBox "Data inválida! Verifique o texto Digitado.", vbExclamation
    objData.SetFocus
End Sub
Function F_Vazio(strExpress As String) As Boolean
    F_Vazio = (RTrim(LTrim(strExpress)) = "")
End Function

Function F_ValidaData(strExpr As String) As String


Dim strDia, strMes, StrAno As String

strExpr = F_SoNumero(strExpr)

If val(strExpr) < 1 Or Len(strExpr) <> 8 Then
   F_ValidaData = ""
   Exit Function
End If
strDia = Left(strExpr, 2)
strMes = Mid(strExpr, 3, 2)
StrAno = Right(strExpr, 4)
If val(strDia) > 31 Or val(strDia) < 1 Then
   F_ValidaData = ""
ElseIf val(strDia) > 30 And InStr(1, _
"04_06_09_11", strMes) > 0 Then
   F_ValidaData = ""
ElseIf val(strMes) > 12 Or val(strMes) < 1 Then
   F_ValidaData = ""
ElseIf strMes = "02" And _
val(strDia) > IIf(val(StrAno) Mod 4 > 0, 28, 29) Then
   F_ValidaData = ""
Else
   F_ValidaData = strDia + "/" + strMes + "/" + StrAno
End If
End Function
Sub P_CentraForm(obj As Object, Optional iRecuo As Integer = 500)
    On Error Resume Next
        
        obj.Top = (mdiFrm.height / 2) - (obj.height / 2) - iRecuo
        
        
        obj.Left = ((mdiFrm.width - 1320) / 2) - (obj.width / 2)
        
            
        
End Sub
Sub P_CentraObejct(obj As Object, frm As Object, _
        Optional bTop As Boolean = True, Optional bLeft As Boolean = True)
    On Error Resume Next
        If bTop Then
            obj.Top = (frm.height / 2) - (obj.height / 2) - 200
        End If
        If bLeft Then
            obj.Left = (frm.width / 2) - (obj.width / 2) - 50
        End If
End Sub

Function F_GeraDig(numero As String) As String
Dim Chave As String
Dim a, total As Integer
Chave = "123456789123456789"
total = 0
For a = 1 To Len(Trim(numero))
   total = total + (val(Mid(numero, a, 1)) * _
   val(Mid(Chave, a, 1)))
Next a
F_GeraDig = numero & F_ZerEsq(total Mod 9, 1, 0)
End Function
Function F_ZerEsq(Num, tam As Integer, _
Optional Dec As Integer = 0) As String
'Função que retorna Zeros à esquerda de um número
'Sintaxe.: F_ZerEsq( Num, Tam, Dec )
'Entrada.: Num - Número que se deseja formatar
'          Tam - Tamanho do número formatado
'          Dec - Tamanho das casas decimais
'Saída...: Número formato com zeros à esquerda
'OBS.....: Se o número for maior que Tam a saída será representada por "*"(Asterisco)
'          Tam - representa o inteiro do número

Dim aux As String, Red As String, Acr As String
Acr = "F"
aux = IIf(Dec > 0, "." + String(Dec, "0"), "")
If InStr(1, str(Num), ".") <> 0 Then
   
   Red = Trim(str(Num))
   aux = Mid(Red, InStr(1, Red, "."))
   
   If Len(aux) > Dec + 1 Then
      Red = Mid(aux, Dec + 2, 1)
      If val(Red) > 4 Then
         
         Red = str(1 / val("1" + String(Dec, "0")))
         aux = Trim(str(val(aux) + val(Red)))
         
         If val(aux) >= 1 Then
            Acr = "T"
            aux = IIf(Dec > 0, "." + String(Dec, "0"), "")
         End If
      End If
   End If
   
   aux = Left(aux + "000000", Dec + 1)

End If
Red = aux
If Acr <> "F" Then Num = Num + 1
aux = IIf(Dec > 0, Trim(str(Fix(Num))), Trim(str(Int(Num))))
aux = Right(String(tam, "0") + aux, tam) + Red
Red = Trim(str(Int(Num)))

If Acr <> "F" Then Num = Num - 1

F_ZerEsq = IIf(Len(Red) > Len(aux), String(tam, "*") _
+ IIf(Dec > 0, "." + String(Dec, "*"), ""), aux)

End Function



Sub P_BuscaIni()
    Dim sNomeArq As String
    Dim sPath As String
    Dim s As String
    sNomeArq = "piki.ini"
    
    If Dir("c:\windows\system32\" & sNomeArq) <> "" Then
        sPath = "c:\windows\system32\" & sNomeArq
    ElseIf Dir(App.path & "\" & sNomeArq) <> "" Then
        sPath = App.path & "\" & sNomeArq
    End If
    If sPath = "" Then
        MsgBox "O Arquivo " & sNomeArq & " não foi encontrado para inicialização"
        
    Else
        Open sPath For Input As #1
        Do While Not EOF(1)
            Line Input #1, s
            s = LCase(s)
            If Left(Trim(s), 1) <> "/" Or Left(Trim(s), 1) <> "[" Or Left(Trim(s), 1) <> "]" Then
                
                
                If Left(Trim(s), 6) = "driver" Then
                    sDriver = Trim(Right(s, Len(s) - 7))
                
                End If
                If Left(Trim(s), 6) = "server" Then
                    sServer = Trim(Right(s, Len(s) - 7))
                End If
                If Left(Trim(s), 3) = "uid" Then
                    sUid = Trim(Right(s, Len(s) - 4))
                End If
                If Left(Trim(s), 3) = "pwd" Then
                    sPwd = Trim(Right(s, Len(s) - 4))
                End If
                If Left(Trim(s), 4) = "data" Then
                    sDataBase = Trim(Right(s, Len(s) - 9))
                End If
                If Left(Trim(s), 6) = "option" Then
                    sOption = Trim(Right(s, Len(s) - 7))
                End If
            End If
        Loop
        Close #1
    End If
End Sub


Function F_AbreBanco(Optional bBase_Teste As Boolean = False) As Boolean
    
    On Error GoTo erro
    DBEngine.DefaultType = dbUseODBC
   
    strConnect = "Driver=" & sDriver & ";"
    
    strConnect = strConnect & "Server=" & sServer & ";Database=" & sDataBase & ";Uid=" & sUid & ";port=3306;pwd=" & sPwd & ";option=" & sOption
    
    'strConnect = "market"
   conn.Open strConnect
   conn_client.CursorLocation = adUseClient
   conn_client.Open strConnect
   
   conn_client.CursorLocation = adUseClient
        'Set Banco = OpenDatabase(Trim(sDataBase), "", "", strConnect)
'    End If
   
    F_AbreBanco = True
    Exit Function
erro:
    


MsgBox Err.Description & Chr(13) & strConnect, vbCritical, conSigla_Sistema
    F_AbreBanco = False
End Function
Sub P_BuscaNomeBanco()
    Dim sNomeArq As String
    sNomeArq = "piki.ini"
    If Dir("c:\windows\system32\" & sNomeArq) = "" Then
        If Dir("c:\windows\system\" & sNomeArq) = "" Then
            If Dir(App.path & "\" & sNomeArq) = "" Then
                strNomeBanco = "forte"
            Else
                Open App.path & "\" & sNomeArq For Input As #1
        
                Line Input #1, strNomeBanco
                Close #1
            End If
        Else
            Open "c:\windows\system\" & sNomeArq For Input As #1
        
            Line Input #1, strNomeBanco
            Close #1
        End If
    Else
        Open "c:\windows\system32\" & sNomeArq For Input As #1
        
        Line Input #1, strNomeBanco
        Close #1
    End If
    If strNomeBanco = "" Then
        MsgBox "Houve um erro na busca do arquivo piki.ini, conecte com 'market'", vbCritical
        strNomeBanco = "market"
    End If
End Sub
Sub Main()
    On Error GoTo erro
    Dim agora As String
    Dim i
    CrLf = Chr(13) + Chr(10)
    If App.PrevInstance = True Then
        MsgBox "Este Programa já Está sendo executado!", vbExclamation
        End
    End If
   
    bTefDedicado = False
    P_BuscaIni
    tabelas.iniciaTabelas
    If Not F_AbreBanco() Then
        End
    End If
    CarregarParametros
    
    
    agora = HoraServidor
    i = DateDiff("n", agora, Now)
    If i > 2 Or i < -2 Then
        MsgBox "A hora do sistema está diferente da hora do servidor em " & i & " minutos", vbCritical
    End If
    
    
    
        PrimeiroAcesso
        
        iniciaVariaveis
        frmLogin.Show 1
        If strNome_Usu = "" Then
            P_Fim
        End If
        iniciaVariaveis
        Unload frmLogin
        
        
        
    
        mdiFrm.Show 0
    
   
    Exit Sub
erro:
    P_MostraErro
    P_Fim
 
End Sub

Sub P_Fim()
    Dim i As Integer
    On Error Resume Next
    For i = 1 To Forms.Count - 1
        Unload Forms(i)
        
    Next
    Banco.Close
    FechaBancoAdo
    End
End Sub

    
    
    
    
Sub P_AjustaFone(obj As Object)
    obj.Text = F_AjustaTelefone(obj.Text)
End Sub
Function f_ajustatelefoneNovo(fone As String) As String
    
End Function
    
    
    
Function F_AjustaTelefone(strExpr As String) As String
   '  Esta função ajusta uma expressão digitada
   'a formato de telefone
   F_AjustaTelefone = strExpr
   Exit Function
   
   F_AjustaTelefone = F_SoNumero(strExpr, True)
   strExpr = F_SoNumero(strExpr)
   If Left(strExpr, 1) = "0" Then
      'Retira o 0(zero) inicial
      strExpr = Right(strExpr, Len(strExpr) - 1)
   End If
   '0800-55-1668
   If Len(strExpr) = 7 Then
      'Telefone local
      F_AjustaTelefone = Left(strExpr, 3) & "-" & Right(strExpr, 4)
      
      F_AjustaTelefone = Right(Space(16) & F_AjustaTelefone, 16)
   
   ElseIf Len(strExpr) = 8 Then
      'Telefone local de São-Paulo
      F_AjustaTelefone = Left(strExpr, 4) & "-" & Right(strExpr, 4)
      
      F_AjustaTelefone = Right(Space(16) & F_AjustaTelefone, 16)
      
   ElseIf Len(strExpr) = 9 And _
   Left(strExpr, 3) <> "800" Then
      
      'DDD
      F_AjustaTelefone = "(0.." & Left(strExpr, 2) & ") " & _
       Mid(strExpr, 3, 3) & "-" & Right(strExpr, 4)
   ElseIf Len(strExpr) = 9 And Left(strExpr, 3) = "800" Then
      'DDG
      F_AjustaTelefone = "(0800) -" & Mid(strExpr, 5, 2) & _
       "- " & Right(strExpr, 4)
   
   ElseIf Len(strExpr) = 10 Then
      'DDD São Paulo
      F_AjustaTelefone = "(0.." & Left(strExpr, 2) & ")" & _
       Mid(strExpr, 3, 4) & "-" & Right(strExpr, 4)
   End If
End Function
Function FormataTelefone(strExpr As String) As String
   '  Esta função ajusta uma expressão digitada
   'a formato de telefone
   FormataTelefone = F_SoNumero(strExpr, True)
   strExpr = F_SoNumero(strExpr)
   If Left(strExpr, 1) = "0" Then
      'Retira o 0(zero) inicial
      strExpr = Right(strExpr, Len(strExpr) - 1)
   End If
   '0800-55-1668
   If Len(strExpr) = 7 Then
      'Telefone local
      FormataTelefone = Left(strExpr, 3) & "-" & Right(strExpr, 4)
      
      FormataTelefone = Right(Space(16) & FormataTelefone, 16)
   
   ElseIf Len(strExpr) = 8 Then
      'Telefone local de São-Paulo
      FormataTelefone = Left(strExpr, 4) & "-" & Right(strExpr, 4)
      
      FormataTelefone = Right(Space(16) & FormataTelefone, 16)
      
   ElseIf Len(strExpr) = 9 And _
   Left(strExpr, 3) <> "800" Then
      
      'DDD
      FormataTelefone = "(" & Left(strExpr, 2) & ") " & _
       Mid(strExpr, 3, 3) & "-" & Right(strExpr, 4)
   ElseIf Len(strExpr) = 9 And Left(strExpr, 3) = "800" Then
      'DDG
      FormataTelefone = "(0800) -" & Mid(strExpr, 5, 2) & _
       "- " & Right(strExpr, 4)
   
   ElseIf Len(strExpr) = 10 Then
      'DDD São Paulo
      FormataTelefone = "(" & Left(strExpr, 2) & ")" & _
       Mid(strExpr, 3, 4) & "-" & Right(strExpr, 4)
   End If
End Function

Function F_SoNumero(strExpr As String, Optional blnCar As Boolean = False) As String
   Dim intCont As Integer
   Dim strValidos As String
   
   strExpr = UCase(strExpr)
   F_SoNumero = ""
   strValidos = IIf(blnCar, "0123456789()./X,", "0123456789")
   For intCont = 1 To Len(strExpr)
       If InStr(1, strValidos, Mid(strExpr, intCont, 1)) > 0 Then
          F_SoNumero = F_SoNumero & Mid(strExpr, intCont, 1)
       End If
   Next
   
End Function

Function F_FormataCEP(strExpr As String) As String
  'Esta função formata o CEP
  'Entrada cep a formatar
  strExpr = F_SoNumero(strExpr)
  strExpr = Format(strExpr, "0#\.###\-###")
  If strExpr = "00.000-000" Then
     F_FormataCEP = ""
     Exit Function
  End If
  F_FormataCEP = strExpr
  
End Function
Function F_PesqData(strData) As String
    On Error Resume Next
    F_PesqData = Format(strData, "yyyy-mm-dd")
    F_PesqData = "'" & F_PesqData & "'"
    If F_PesqData = "" Then
        F_PesqData = "0000-00-00"
    End If

End Function
Function SenhaConfereUsu(senha As String) As Boolean
    Dim strSQL As String
    Dim senhaUsu As String
    strSQL = SenhaMySQL(senha)
    senhaUsu = F_BuscaUmCampo("senha_usu", "tab_usuarios", "cod_fun='" & strCod_Usu & "'")
    If senhaUsu = strSQL Then
        SenhaConfereUsu = True
    Else
        SenhaConfereUsu = False
    End If




End Function
Function SenhaMySQL(senha As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select password('" & senha & "') as senha"
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        SenhaMySQL = rPesq!senha
    End With
    

End Function
Function F_GeraPW(sExpr As String) As String
    Dim intCont As Integer ' by Son
    Dim intAux As Integer
    Dim strAux As String
    F_GeraPW = ""
    For intCont = 1 To Len(sExpr)
        strAux = Mid(sExpr, intCont, 1)
        intAux = Asc(strAux)
        F_GeraPW = F_GeraPW + ChrW$(intAux + 1 + intCont)
    Next
End Function
Function F_RetornaPW(sExpr As String) As String
    Dim intCont As Integer
    Dim intAux As Integer
    Dim strAux As String
    F_RetornaPW = ""
    For intCont = 1 To Len(sExpr)
        strAux = Mid(sExpr, intCont, 1)
        intAux = Asc(strAux)
        F_RetornaPW = F_RetornaPW + ChrW$(intAux - 1 - intCont)
    Next
End Function
Sub P_AjustaTexto(ByRef objText As Object, Optional caso As caso = caso.Maiusculo)
    AjustaTexto objText, caso
End Sub
Sub AjustaTexto(ByRef objText As Object, Optional caso As caso = caso.Maiusculo)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    'Torna o texto maiúsculo
    
    If caso = 3 Then
        strAux = UCase(objText.Text)
    ElseIf caso = 2 Then
        strAux = StrConv(objText.Text, vbProperCase)
    
    ElseIf caso = 1 Then
        strAux = LCase(objText.Text)
    Else
        strAux = objText.Text
        
    End If
    For intCont = 1 To Len(strAux)
        strAux2 = Mid(strAux, intCont, 1)
        If strAux2 = "'" Then
           Mid(strAux, intCont, 1) = "`" 'Trocando ' por `
        ElseIf strAux2 = "[" Then
           Mid(strAux, intCont, 1) = "{" 'Trocando [ por {
        ElseIf strAux2 = "]" Then
           Mid(strAux, intCont, 1) = "}" 'trocando ] por }
        ElseIf strAux2 = "|" Then
           Mid(strAux, intCont, 1) = "¦" 'trocando | por ¦
        End If
    Next
    objText.Text = strAux
End Sub
Function F_AjustaTexto(objText As String)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    'Torna o texto maiúsculo
    strAux = RTrim((objText))
    
    For intCont = 1 To Len(strAux)
        strAux2 = Mid(strAux, intCont, 1)
'        MsgBox Asc(strAux2) & " - " & strAux2
        If strAux2 = "'" Then
           Mid(strAux, intCont, 1) = "`" 'Trocando ' por `
        ElseIf strAux2 = "[" Then
           Mid(strAux, intCont, 1) = "{" 'Trocando [ por {
        ElseIf strAux2 = "]" Then
           Mid(strAux, intCont, 1) = "}" 'trocando ] por }
        ElseIf strAux2 = "|" Then
           Mid(strAux, intCont, 1) = "¦" 'trocando | por ¦
        ElseIf strAux2 = "\" Then
           Mid(strAux, intCont, 1) = "/" 'trocando | por ¦
        ElseIf Asc(strAux2) = 0 Then
            Mid(strAux, intCont, 1) = " "
        ElseIf Asc(strAux2) = 32 Then
            Mid(strAux, intCont, 1) = " "
        ElseIf Asc(strAux2) = 10 Then
            Mid(strAux, intCont, 1) = Chr(13)
            
        ElseIf Asc(strAux2) = 13 Then
            'Mid(strAux, intCont, 1) = " "
        End If
    Next
    F_AjustaTexto = strAux
End Function


Function F_SubstituiChar(objText As String, chrEncontrar As String, _
    chrSubstituto As String, Optional trmDireita As Boolean = False)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    'Torna o texto maiúsculo
    If trmDireita Then
        strAux = RTrim((objText))
    Else
        strAux = objText
    End If
    
    For intCont = 1 To Len(strAux)
        If Mid(strAux, intCont, 1) = chrEncontrar Then
            strAux2 = strAux2 + chrSubstituto
        Else
            strAux2 = strAux2 + Mid(strAux, intCont, 1)
        End If
'
    Next
    F_SubstituiChar = strAux2
End Function


Function F_ColocaAspa(objText As String)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    
    strAux = objText
    
    For intCont = 1 To Len(strAux)
        strAux2 = Mid(strAux, intCont, 1)
        If strAux2 = "`" Then
           Mid(strAux, intCont, 1) = "'" 'Trocando ' por `
        
        End If
    Next
    F_ColocaAspa = strAux
End Function

Sub P_SelecExpr(ByRef objctObjeto As Object, strExpr As String)
   'Este procedimento localiza dentro de um combo ou list box a
   'expressão contida em strexpr, selecionando a expressão
   Dim intCont As Integer, intAtual As Integer
   intAtual = objctObjeto.ListIndex
   strExpr = Trim(strExpr)
   For intCont = 0 To objctObjeto.ListCount - 1
       objctObjeto.ListIndex = intCont
       If Left(objctObjeto.Text, Len(strExpr)) = strExpr Then
          Exit Sub
       End If
   Next
   objctObjeto.ListIndex = intAtual
End Sub

Function F_SoLetra(strExpr As String) As String
    'Esta função retorna somente as letras contidas em uma expressão
    'Ex: A-E_i,O-c = AEiOc
    Dim intCont As Integer
    
    F_SoLetra = ""
    
    For intCont = 1 To Len(strExpr)
        If (Mid(strExpr, intCont, 1) >= "A" And _
        Mid(strExpr, intCont, 1) <= "Z") Or _
        (Mid(strExpr, intCont, 1) >= "a" And _
        Mid(strExpr, intCont, 1) <= "z") Then
        
            F_SoLetra = F_SoLetra & Mid(strExpr, intCont, 1)
            
        End If
    Next
    
End Function
Sub P_ValorLostFocus(ByRef objValor As Object, _
Optional strFormat As String = "###,##0.00")
    PerdeFoco objValor, strFormat
End Sub

Sub PerdeFoco(ByRef objValor As Object, _
Optional strFormat As String = "###,##0.00")
    'Procedimento de perda de foco sobre valores
    Dim strExpr As String
    On Error GoTo trata_Erro
    strExpr = F_SoNumero(objValor.Text, True)
    If F_Vazio(objValor.Text) Then
        strExpr = "0"
    End If
    If Right(strExpr, 1) = "." Then
       strExpr = "0" & strExpr
       Mid(strExpr, Len(strExpr)) = ","
    ElseIf Left(Right(strExpr, 2), 1) = "." Then
       strExpr = "0" & strExpr
       Mid(strExpr, Len(strExpr) - 1) = ","
    ElseIf Left(Right(strExpr, 3), 1) = "." Then
       strExpr = "0" & strExpr
       Mid(strExpr, Len(strExpr) - 2) = ","
    End If
    objValor.Text = strExpr
    'objValor.Text = (CCur(F_AjustaValor(objValor.Text)))
    objValor.Text = Format(objValor.Text, strFormat)
    
    If CCur(F_SoNumero(objValor.Text)) < 0 Then
       'Se o valor for menor que zero, vira zero
       objValor.Text = "0,00"
    End If
Exit Sub
trata_Erro:
objValor.Text = "0,00"
End Sub

Sub P_gotfocus(objText As Object)
    ganhaFoco objText
End Sub
Sub ganhaFoco(objText As Object)

    objText.SelStart = 0
    objText.SelLength = Len(RTrim(objText.Text))
End Sub
Function F_CGC(CGC As String) As Boolean
'Esta função permite validar o CGC
'Sintaxe.: F_CGC( CGC )
'Entrada.: CGC - Nº do CGC a validar
'Saída...: Verdadeiro se exister o CGC, senão Falso

Dim M_VetNum(13), M_VetCGC(13), M_Num As String
Dim D_12, D_13, X, Y, M_D1, M_D2 As Integer
CGC = F_SoNumero(CGC)
If val(CGC) < 1 Then
   F_CGC = False
   Exit Function
End If
M_Num = "6543298765432"
CGC = F_ZerEsq(val(CGC), 14, 0)
For X = 1 To 13
        M_VetNum(X) = val(Mid(M_Num, X, 1))
Next
For X = 1 To 12
        M_VetCGC(X) = Mid(CGC, X, 1)
Next
D_12 = 0
For Y = 1 To 12
        D_12 = D_12 + val(M_VetCGC(Y)) * M_VetNum(Y + 1)
Next
M_D1 = 11 - (D_12 - Int(D_12 / 11) * 11)
If M_D1 >= 10 Then
   M_D1 = 0
End If
M_VetCGC(13) = Trim(str(M_D1))
For Y = 1 To 13
        D_13 = D_13 + val(M_VetCGC(Y)) * M_VetNum(Y)
Next
M_D2 = 11 - (D_13 - Int(D_13 / 11) * 11)
If M_D2 >= 10 Then
   M_D2 = 0
End If
If Right(CGC, 2) <> Trim(str(M_D1)) & Trim(str(M_D2)) Then
   F_CGC = False
   Exit Function
End If
F_CGC = True
End Function
Function F_CPF(cpf As String) As Boolean

'Função que permite validar CPF
'Sintaxe.: F_CPF( CPF )
'Entrada.: CPF - número do CPF
'Saída...: Verdadeiro se o CPF existir, senão falso

Dim controle As String
Dim ContIni, ContiFim, soma As Integer
Dim ii As Integer
Dim i As Integer
Dim digito As Integer

cpf = F_SoNumero(cpf)
cpf = Left(F_ZerEsq(val(cpf), 11, 0), 11)
ContIni = 2
ContiFim = 10
If val(cpf) < 1 Then
   F_CPF = False
   Exit Function
End If

For ii = 1 To 2
    soma = 0
    For i = ContIni To ContiFim
        soma = soma + (val(Mid(cpf, i - ii, 1)) * _
        (ContiFim + 1 + ii - i))
    Next i
    If ii = 2 Then
       soma = soma + (2 * digito)
    End If
    digito = (soma * 10) Mod 11
    If digito = 10 Then digito = 0
    
    controle = RTrim(controle) + _
    F_ZerEsq(CDbl(digito), 1, 0)
    
    ContIni = 3
    ContiFim = 11
Next ii
F_CPF = IIf(Trim(controle) = Right(cpf, 2), True, False)
End Function

Sub P_CEPLostFocus(objCEP As Object)
    'Procedimento a ser usado quando o
    'objeto que contém CEP perde o foco
    objCEP.MaxLength = 10
    objCEP.Text = F_SoNumero(objCEP.Text)
    If Len(objCEP.Text) <> 8 Then
       Beep
       MsgBox "Informe o número do CEP corretamente!", _
       vbExclamation, conSigla_Sistema
    End If
    objCEP.Text = Format(objCEP.Text, "0#\.###\-###")
End Sub
Sub P_FormataCGC(objCGC As Object)
    objCGC.Text = Format(objCGC.Text, "0#\.###\.###\/####\-##")
End Sub
Sub P_FormataCPF(objCPF As Object)
    objCPF.Text = Format(objCPF.Text, "0##\.###\.###\-##")
End Sub
Function F_FormataCPF(objCPF As String)
    F_FormataCPF = Format(objCPF, "0##\.###\.###\-##")
End Function
Function F_FormataCNPJ(objCPF As String)
    F_FormataCNPJ = Format(objCPF, "0#\.###\.###\/####\-##")
End Function

Sub P_PreencheUF(ByRef objObjeto As Object)
    Dim strngEstados As String
    Dim intCont As Integer
    strngEstados = "AC_AL_AM_AP_BA_CE_DF_ES_GO_MA_MG_MS_MT" & _
    "_PA_PB_PE_PI_PR_RJ_RN_RO_RR_RS_SC_SE_SP_TO"
    
    objObjeto.Clear
    For intCont = 1 To Len(strngEstados) Step 3
        objObjeto.AddItem Mid(strngEstados, intCont, 2)
    Next
End Sub

Sub P_PreencheOperacoes(cmbxOBJ As ComboBox, sTipo_Op As String)
    cmbxOBJ.Clear 'E-Entrada S-Saida
    If sTipo_Op = "E" Then
        cmbxOBJ.AddItem "1.12 -Compras Revenda Estadual"
        cmbxOBJ.AddItem "2.12 -Compras Revenda InterEstadual"
        cmbxOBJ.AddItem "1.22 -Transf. Comercial Estadual"
        cmbxOBJ.AddItem "2.22 -Transf. Comercial InterEstadual"
        cmbxOBJ.AddItem "1.32 -Devolução Estadual"
        cmbxOBJ.AddItem "2.32 -Devolução InterEstadual"
        cmbxOBJ.AddItem "1.91 -Compras Ativo Imob.Estadual"
        cmbxOBJ.AddItem "2.91 -Compras Ativo Imob.InterEstadual"
        cmbxOBJ.AddItem "1.92 -Transf. Ativo Imob.Estadual"
        cmbxOBJ.AddItem "2.92 -Transf. Ativo Imob.InterEstadual"
        cmbxOBJ.AddItem "1.93 -Mat.Consumo Estadual"
        cmbxOBJ.AddItem "2.93 -Mat.Consumo InterEstadual"
        cmbxOBJ.AddItem "1.99D-Demonstração Estadual"
        cmbxOBJ.AddItem "2.99D-Demonstração InterEstadual"
        cmbxOBJ.AddItem "1.99 -Outras Entradas Estadual"
        cmbxOBJ.AddItem "2.99 -Outras Entradas InterEstadual"
    Else
        cmbxOBJ.AddItem "5.12-Vendas Merc.Estadual"
        cmbxOBJ.AddItem "6.12-Vendas Merc.InterEstadual"
        cmbxOBJ.AddItem "5.32-Devolução Estadual"
        cmbxOBJ.AddItem "6.32-Devolução InterEstadual"
        cmbxOBJ.AddItem "5.22-Transf. Merc. Estadual"
        cmbxOBJ.AddItem "6.22-Transf. Merc InterEstadual"
        cmbxOBJ.AddItem "5.73-Venda Merc.Retida Estadual"
        cmbxOBJ.AddItem "5.99-Outras Saídas Estadual"
        cmbxOBJ.AddItem "6.99-Outras Saídas InterEstadual"
        cmbxOBJ.AddItem "5.93-Trans Mat.Consumo Estadual"
        cmbxOBJ.AddItem "6.93-Trans Mat.Consumo InterEstadual"
    End If
End Sub
Function F_CharValido(cChar As String) As Boolean
    Dim iChar As Long
    
    On Error GoTo trata_Erro
    
    iChar = Asc(cChar)
    If iChar = vbKeyReturn Or iChar = vbKeyBack Or iChar = vbKeyUp _
                Or iChar = vbKeyDown Or iChar = vbKeyLeft Or iChar = vbKeyRight _
                Or iChar = vbKeyPageDown Or iChar = vbKeyPageUp Or _
                iChar = vbKeyHome Or iChar = vbKeyDelete _
                Or iChar = vbKeyEnd Or iChar = vbKeyInsert Or iChar = vbKeyTab Or _
                iChar = vbKeyPause Or iChar = vbKeyPrint Or iChar = vbKeyScrollLock _
                Or iChar = vbKeyF1 Or iChar = vbKeyF2 Or iChar = vbKeyF3 _
                Or iChar = vbKeyF4 Or iChar = vbKeyF5 Or iChar = vbKeyF6 _
                Or iChar = vbKeyF7 Or iChar = vbKeyF8 Or iChar = vbKeyF9 _
                Or iChar = vbKeyF10 Or iChar = vbKeyF11 Or iChar = vbKeyF12 _
                Or iChar = vbKeyCancel Or iChar = vbKeyCapital _
                Or iChar = vbKeyClear Or iChar = vbKeyControl _
                Or iChar = vbKeyDelete Or iChar = vbKeyEnd Or iChar = vbKeyEscape _
                Or iChar = vbKeyNumlock Then
        F_CharValido = False
    Else
        F_CharValido = True
                
                
    End If
                
    
    
    Exit Function
trata_Erro:
        F_CharValido = False
End Function
Sub P_Status(Optional sMsg As String = "", Optional iPanel As Integer = 3, Optional bEspere As Boolean = True)
    On Error GoTo trata_Erro
    Dim iCount As Integer
    mdiFrm.lblMsg.caption = sMsg
    
    mdiFrm.lblMsg.Refresh
    If Trim(sMsg) <> "" And bEspere Then
        For iCount = 0 To Forms.Count - 1
            Forms(iCount).MousePointer = 11
        Next
    Else
        For iCount = 0 To Forms.Count - 1
            Forms(iCount).MousePointer = 0
        Next
    End If
    
    Exit Sub
trata_Erro:
On Error Resume Next
mdiFrm.lblMsg.caption = ""


End Sub

Sub P_PreencheMes(obj As Object)
    obj.Clear
    obj.AddItem "Janeiro"
    obj.AddItem "Fevereiro"
    obj.AddItem "Março"
    obj.AddItem "Abril"
    obj.AddItem "Maio"
    obj.AddItem "Junho"
    obj.AddItem "Julho"
    obj.AddItem "Agosto"
    obj.AddItem "Setembro"
    obj.AddItem "Outubro"
    obj.AddItem "Novembro"
    obj.AddItem "Dezembro"
End Sub
Function F_TipoDoc(iTipo As Integer, Optional bAbrev As Boolean = False) As String
    On Error GoTo trata_Erro
    If bAbrev = True Then
        If iTipo = 1 Then
            F_TipoDoc = "NF"
        ElseIf iTipo = 2 Then
            F_TipoDoc = "CF"
        ElseIf iTipo = 3 Then
            F_TipoDoc = "CT"
        End If
    Else
        If iTipo = 1 Then
            F_TipoDoc = "Nota Fiscal"
        ElseIf iTipo = 2 Then
            F_TipoDoc = "Cupom Fiscal"
        ElseIf iTipo = 3 Then
            F_TipoDoc = "Conhecimento de Transporte"
        End If
    End If
    Exit Function
trata_Erro:
    MsgBox Err.Description, vbCritical, constSigla_sis & "- F_tipodoc"
End Function
Function FormatVal(sValor, Optional sFormat As String = "###,###,##0.00") As String
    On Error GoTo trata_Erro
    FormatVal = Format(sValor, sFormat)
    If FormatVal = "" Then FormatVal = "0,00"
    Exit Function
    
trata_Erro:
    FormatVal = "0,00"
End Function

Sub callobjP()
    On Error Resume Next
    frmTestPrint.Show 1
End Sub
Function F_OpenLpt(Optional sLpt As String = "lpt1") As Boolean
    On Error GoTo TrataErro
    Open sLpt For Output As #1
    Print #1, Chr(27) + "@";
    F_OpenLpt = True
    Exit Function
TrataErro:
    If Err = 55 Then
        MsgBox " A Porta já está aberta ", vbExclamation, "Erro na Abertura da Porta de impressão"
        If Not F_CloseLpt Then Exit Function
    Else:
        MsgBox "Ocorreu um erro na abertura da porta " & sLpt & " verifique se a Impressora está disponível!", vbExclamation, "Erro na Abertura da Porta de Impressão"
    End If
    F_OpenLpt = False
End Function
Function F_CloseLpt()
    On Error Resume Next
    Close #1
    
End Function





Sub P_NegritoLpt1(bNegrito As Boolean)
    On Error GoTo TrataErro
    If bNegrito Then
        Print #1, Chr(27) + "E";
    Else
        Print #1, Chr(27) + "F";
    End If
    Exit Sub
TrataErro:
    
End Sub
Sub P_SublLpt1(bSubl As Boolean)
    On Error GoTo TrataErro
    If bSubl Then
        Print #1, Chr(27) + "-" + "1";
    Else
        Print #1, Chr(27) + "-" + "0";
    End If
    Exit Sub
TrataErro:
End Sub

Sub P_PrintLpt(sExpr As String, Optional bySPC As Integer = 0, Optional bContinua As Boolean = True, Optional sFont As String = "M", Optional byAllign As Byte = 1, Optional estilo As String = "5")
    estilo = CStr(estilo)
    
    P_FontLpt sFont, estilo
    If byAllign = 3 Then
        bySPC = bySPC - CByte(Len(sExpr))
    ElseIf byAllign = 2 Then
        bySPC = (bySPC - CByte(Len(sExpr))) / 2
    End If
    If bContinua Then
        Print #1, Chr(27) + "@";
        P_FontLpt sFont, estilo
        Print #1, Spc(bySPC); sExpr;
    Else
        P_FontLpt sFont, estilo
        Print #1, Spc(bySPC); sExpr
        Print #1, Chr(27) + "@";
    End If
    P_FontLpt sFont, estilo
End Sub
Sub P_FontLpt(Optional sFonte As String = "P", Optional sEstilo As String = "F")

    If sFonte = "P" Then 'E = negrito 'M = draft 'P = normal
        Print #1, Chr(27) + "P";
    ElseIf sFonte = "E" Then
        Print #1, Chr(27) + "E";
    ElseIf sFonte = "F" Then
        Print #1, Chr(27) + "F";
    ElseIf sFonte = "M" Then
        Print #1, Chr(27) + "M";
    ElseIf sFonte = "15" Then
        Print #1, Chr(27) + Chr(15);
    ElseIf sFonte = "18" Then
        Print #1, Chr(18);
    ElseIf sFonte = "14" Then
        Print #1, Chr(27) + Chr(14);
    ElseIf sFonte = "20" Then
        Print #1, Chr(20);
    ElseIf sFonte = "'4'" Then
        Print #1, Chr(27) + "4";
    ElseIf sFonte = "'5'" Then
        Print #1, Chr(27) + "5";
    ElseIf sFonte = "-1" Then
        Print #1, Chr(27) + "-" + "1"
    ElseIf sFonte = "-0" Then
       Print #1, Chr(27) + "-" + "1"
    End If
    
    If sEstilo = "F" Then
    
        Print #1, Chr(27) + "F";
    ElseIf sEstilo = "E" Then
        Print #1, Chr(27) + "E";
    ElseIf sEstilo = "4" Then
        Print #1, Chr(27) + "4";
    ElseIf sEstilo = "5" Then
        Print #1, Chr(27) + "5";
    ElseIf sEstilo = "-1" Then
        Print #1, Chr(27) + "-" + "1"
    ElseIf sEstilo = "-0" Then
       Print #1, Chr(27) + "-" + "0"
    End If
End Sub

Sub P_FeedLpt1(Optional byQuantidade As Byte = 1)
    Dim iCount As Integer
    For iCount = 1 To byQuantidade
        Print #1,
    Next
End Sub


Sub P_DireitaImpressao(objP As Object, strMens As String)
    objP.CurrentX = objP.ScaleWidth - objP.TextWidth(strMens)
    objP.Print strMens
End Sub





Sub CallView()
    P_Status "Gerando Visualização... Aguarde."
    
    frmVisualiza.Show 0
    
End Sub
Function F_DataMaior(dtData1 As String, dtdata2 As String) As Boolean
    On Error GoTo trata_Erro
    If CDate(dtData1) > CDate(dtdata2) Then
        F_DataMaior = True
    Else
        F_DataMaior = False
    End If
trata_Erro:
End Function

'Function F_buscacodbarra_ent(sCodbar As String, _
'cIpi_Max As Currency, cValor_Max As Currency, dInd As Double) As Produto
'    Dim strSQL As String
'    Dim rPesq As New adodb.Recordset
'    Dim p As Produto
'
'
'        p = F_BuscaProdSaida(sCodbar)
'        If p.cod_pro = "" Then
'            Exit Function
'        End If
'        F_buscacodbarra_ent.cod_cat = p.cod_cat
'           F_buscacodbarra_ent.cod_mar = p.cod_mar
'           F_buscacodbarra_ent.cod_pro = p.cod_pro
'           F_buscacodbarra_ent.codbarra_pro = p.codbarra_pro
'           F_buscacodbarra_ent.c1_pro = p.c1_pro
'           F_buscacodbarra_ent.cr_pro = p.cr_pro
'           F_buscacodbarra_ent.ef_pro = p.ef_pro
'           F_buscacodbarra_ent.er_pro = p.er_pro
'           F_buscacodbarra_ent.s_pro = p.s_pro
'           F_buscacodbarra_ent.nome_cat = p.nome_cat
'           F_buscacodbarra_ent.nome_mar = p.nome_mar
'           F_buscacodbarra_ent.nome_pro = p.nome_pro
'           F_buscacodbarra_ent.p1_pro = p.p1_pro
'           F_buscacodbarra_ent.pv_pro = p.pv_pro
'
'           F_buscacodbarra_ent.r_pro = p.r_pro
'
'            On Error Resume Next
'
'            frmValoresEntr.lblCod_Pro.caption = p.cod_pro
'            frmValoresEntr.lblNome_Pro.caption = p.nome_pro
'            frmValoresEntr.txtbxCusto.Text = FormatVal(p.cr_pro)
'            frmValoresEntr.txtbxPreco.Text = FormatVal(p.pv_pro)
'            frmValoresEntr.lblCat.caption = p.nome_cat
'
'            frmValoresEntr.lblMar.caption = p.nome_mar
'            frmValoresEntr.txtbxIpi.Visible = True
'            frmValoresEntr.lblIPI.Visible = True
'            frmValoresEntr.txtbxPreco.Text = "0,00"
'            frmValoresEntr.txtbxPreco.Enabled = True
'            frmValoresEntr.txtbxPreco.Tag = cValor_Max
'            frmValoresEntr.lblRestante.caption = "Vlr Rest:(R$): " & FormatVal(cValor_Max)
'            frmValoresEntr.txtbxIpi.Tag = cIpi_Max
'            frmValoresEntr.lblIPI_Rest.caption = "IPI Rest:(R$): " & FormatVal(cIpi_Max)
'
'            frmValoresEntr.Show 1
'
'           F_buscacodbarra_ent.qtd_pro = CCur(frmValoresEntr.txtbxQtd.Text)
'
'           F_buscacodbarra_ent.preco = CCur(frmValoresEntr.txtbxPreco.Text)
'           F_buscacodbarra_ent.ipi = CCur(frmValoresEntr.txtbxIpi.Text)
'
'            Unload frmValoresEntr
'
'
'
'
'End Function

Sub LockForm(opcao As Boolean, frm As Form)
    'opcao true loca, false desloca
    Dim iCount As Integer
    For iCount = 0 To Forms.Count - 1
        If opcao = True Then
            Forms(iCount).Enabled = False
        Else
            Forms(iCount).Enabled = True
        End If
        If Forms(iCount).Name = frm.Name Then
            Forms(iCount).Enabled = True
        End If
    Next
End Sub



Function IsLoad(Nomefrm As String) As Boolean
    Dim i As Integer
    Dim sText As String
    For i = 0 To Forms.Count - 1
        If UCase(Forms(i).Name) = UCase(Nomefrm) Then
            IsLoad = True
            Exit Function
        End If
    Next
    IsLoad = False
End Function
Sub RefreshDataPosition(dt As Object)
    Dim iPos As Integer
    On Error GoTo trata_Erro
    
    If dt.Recordset.RecordCount = 0 Then Exit Sub
    On Error Resume Next
    iPos = dt.Recordset.AbsolutePosition
    dt.Refresh
    dt.Recordset.Move iPos
    Exit Sub
trata_Erro:
    On Error Resume Next
    'dt.Refresh
    'MsgBox Err.Description & " erro em refreshdataposition", vbExclamation
End Sub
Sub RefreshAdoPosition(dt As Object)
    Dim iPos As Integer
    On Error GoTo trata_Erro
        
    If dt.RecordCount = 0 Then Exit Sub
    On Error Resume Next
    iPos = dt.AbsolutePosition
    dt.Requery
    dt.Move iPos - 1
    Exit Sub
trata_Erro:
    On Error Resume Next
    'dt.Refresh
    'MsgBox Err.Description & " erro em refreshdataposition", vbExclamation
End Sub




Sub MudaAparencia(frm As Form)
    Dim strSQL As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim rPesq As New ADODB.Recordset
    Dim sObj As String
    On Error Resume Next
    strSQL = "select * from tab_aparencia where cod_usu = '" & strCod_Usu & "' "
    rPesq.Open strSQL, conn_client
    If rPesq.RecordCount = 0 Then Exit Sub
    For iCount = 0 To frm.Controls.Count - 1
        sObj = LCase(TypeName(frm.Controls(iCount)))
        If sObj = "label" Then
            If frm.Controls(iCount).tag <> "!" Then 'if ! not change
                frm.Controls(iCount).Appearance = rPesq!appearance_par
                frm.Controls(iCount).ForeColor = rPesq!corfonte_par
                frm.Controls(iCount).BackStyle = 0
            Else
                If frm.Controls(iCount).Name = "lblCaptionFrame" Then
                    frm.Controls(iCount).caption = frm.caption
                End If
            End If
        ElseIf sObj = "textbox" Or sObj = "combobox" Then
            If frm.Controls(iCount).tag <> "!" Then 'if ! not change
                frm.Controls(iCount).ForeColor = rPesq!corfontecaixa_par
                frm.Controls(iCount).BackColor = rPesq!corcaixa_par
                frm.Controls(iCount).Appearance = rPesq!appearance_par
            End If
        ElseIf sObj = "msflexgrid" Or sObj = "dbgrid" Then
            frm.Controls(iCount).BackColor = rPesq!corcaixa_par
            frm.Controls(iCount).ForeColor = rPesq!corfontecaixa_par
            If sObj = "dbgrid" Then
                For iCol = 0 To frm.Controls(iCount).Columns.Count - 1
                    frm.Controls(iCount).Columns(iCol).DividerStyle = 6
                Next
                frm.Controls(iCount).RowDividerStyle = 6
                frm.Controls(iCount).MarqueeStyle = 0
            End If
            
            frm.Controls(iCount).Appearance = rPesq!appearance_par
        ElseIf sObj = "frame" Or sObj = "sstab" Then
            If frm.Controls(iCount).tag <> "!" Then
                frm.Controls(iCount).BackColor = rPesq!corframe_par
                frm.Controls(iCount).ForeColor = rPesq!corfonte_par
            End If
        ElseIf sObj = "optionbutton" Or sObj = "checkbox" Then
            frm.Controls(iCount).BackColor = rPesq!corframe_par
            frm.Controls(iCount).ForeColor = rPesq!corfonte_par
        Else
            
        End If
        
    Next
    frm.BackColor = rPesq!corform_par
    frm.ForeColor = rPesq!corfonte_par
    rPesq.Close

End Sub
Function ContraBarraPipe(Text As String) As String
    Dim iCount As Long
    Dim s As String
    Dim strAux As String
    For intCont = 1 To Len(Text)
        strAux = Mid(Text, intCont, 1)
        If strAux = "\" Then
          strAux = "|"
        ElseIf strAux = "|" Then
           strAux = "\"  'trocando | por ¦
        End If
        s = s & strAux
    Next
    ContraBarraPipe = s
End Function

Function F_Cumprimenta() As String
'---Função cumprimenta
'---Retorna Bom Dia, Boa Tarde, Boa Noite,
'cfe hora do Sistema
If val(Left(Time, 2)) >= 0 And val(Left(Time, 2)) _
< 13 Then
   F_Cumprimenta = "Bom dia!!!"
ElseIf val(Left(Time, 2)) >= 12 And val(Left(Time, 2)) _
< 18 Then
   F_Cumprimenta = "Boa tarde!!!"
Else
   F_Cumprimenta = "Boa noite!!!"
End If
End Function

Sub MostraErro(Optional strProc As String = "Erro ")
  'Procedimento com a finalidade de comentar erro
  'Sintaxe.: MostraErro()
  'Se for executado algum erro desconhecido fornece a mensagem de erro desconhecido
  ''On Error  Resume Next
  Dim strErro As String
  'P_Status ""  'Colocando barra de Status em modo normal
  If Err = 0 Then
     'O procedimento foi passado sem que algum erro ocorresse
     'Exit Sub
     'Err = Err.Number
  End If
  Select Case Err.Number
      Case 7, 31001
           strErro = "Há pouca memória disponível, feche alguns aplicativos abertos!"
      Case 11
           strErro = "Divisão por zero!!!"
      Case 48
           strErro = "Não é possível executar DLL (Biblioteca Dinâmica)"
      Case 49
           strErro = "Argumentos não foram passados corretamente à DLL"
      Case 71
           strErro = "Erro ao ler a unidade de disco!"
      Case 93
           strErro = "Expressão inválida!"
      Case 3021
           strErro = "Não Há Registro Atual"
      Case 3029
           strErro = "Senha ou usuários inválidos para o banco de dados!"
      Case 3037
           strErro = "Não é possível abrir mais tabelas ou consultas!"
      Case 3043
           strErro = "A Rede ou o Banco de Dados estão desativados!"
      Case 3008, 3261, 3189
           strErro = "O Banco de Dados está aberto em modo exclusivo por outro usuário"
      Case 3009, 3211, 3212
           strErro = "Não é possível executar a ação"
      Case 3046, 3158, 3186
           strErro = "Não é possível bloquear o registro"
      Case 3187
           strErro = "Não é possível ler o registro especificado"
      Case 3218, 3260, 3188
           strErro = "Não é possível atualizar os dados"
      Case 3146, 3028
           strErro = "A conexão com o banco de dados falhou!" & Chr(13)
           strErro = strErro & "Tente novamente!"
      Case 3155
           strErro = "Erro duplicidade de código!" & Chr(13)
           strErro = strErro & "Tente novamente!" & Chr(13)
           strErro = strErro & Err.Description
      Case Else
           strErro = "ERRO DESCONHECIDO!!!!" & Chr(13) & Chr(13)
           strErro = strErro & "Módulo: " & Err.Source & Chr(13)
           strErro = strErro & "Código do erro:" & str(Err) & Chr(13) & Chr(13)
           strErro = strErro & Err.Description & Chr(13) & Chr(13) & _
            Format(Now, "dd \de mmmm \de yyyy    hh:mm:ss")
  End Select
 
  MsgBox strErro, vbCritical + vbSystemModal, strProc
  P_Status ""
  
End Sub


Function F_Centra(strMens As String, intTam As Integer, _
Optional strCar As String = " ") As String
'   Esta função centraliza um texto no tamanho indicado
On Error Resume Next
    strMens = Left(strMens, intTam)

    strMens = strMens & String(Int((intTam - Len(strMens)) / 2), strCar)
    F_Centra = String(intTam - Len(strMens), strCar) & strMens

End Function
Function F_MesExtenso(Optional datData As String) As String
'Funcão..: F_MesExtenso( Data )
'Função que retorna a Data por Extenso
'Entrada.: Data
'Saída...: Mês por extenso
If IsEmpty(datData) Then
   datData = Date
End If
Dim mes(12) As String
mes(1) = "Janeiro"
mes(2) = "Fevereiro"
mes(3) = "Março"
mes(4) = "Abril"
mes(5) = "Maio"
mes(6) = "Junho"
mes(7) = "Agosto"
mes(8) = "Setembro"
mes(9) = "Outubro"
mes(10) = "Novembro"
mes(11) = "Dezembro"
mes(12) = "Fevereiro"
F_MesExtenso = mes(Month(datData))
End Function



Function F_ValToText(crcValor As Currency, Optional strFormato As String = "###,##0.00") As String
    'Converte um número para texto formatado
    F_ValToText = Format(crcValor, strFormato)
End Function
Sub P_ValorGotFocus(objValor As Object)
    'Procedimento de ajuste de foco
    ganhaFoco objValor
End Sub
Function F_DataInicialFinal(Optional DatIniFin As Boolean = True, _
Optional strmensagem As String = "", Optional strData As String = "") As String
    ' Esta função invoca um formulário de entrada de dados com data
    'inicial e data final. Retornando uma string com as duas datas
    ' Se o parâmetro DatIniFin for falso ela retorna uma só data
    If strmensagem <> "" Then
       frmDataInicialFinal.caption = strmensagem
    End If
    If F_Vazio(strData) Then
       strData = Date
    End If
    With frmDataInicialFinal
        .txtbxDataInicial = strData
        .txtbxDataFinal = DateAdd("m", 1, strData)
    End With
    
    If DatIniFin Then
       'Chama o formulário para as datas iniciais e finais
       frmDataInicialFinal.Show 1
    Else
       'Chama o formulário para retornar uma data somente
       With frmDataInicialFinal
            .tag = "Simples"
            .Show 1
       End With
    End If
    F_DataInicialFinal = frmDataInicialFinal.tag 'Iguala com o tag
    Unload frmDataInicialFinal 'Fecha o formulário
    
End Function
Function F_Arredonda(vrnNumero As Variant) As Currency
    'Esta função arredonda valores de orçamento pela constante
    Dim strNumero As String
    strNumero = Format(Int(vrnNumero * 1000))
    If Right(strNumero, 1) <> 0 Then
         strNumero = Format(strNumero + 10)
    End If
    F_Arredonda = CCur(Format(Int(CCur(strNumero) / 10), _
    "##############.00")) / 100
    
End Function


Function F_TextToVal(strNumero As String) As Currency
   On Error GoTo TrataErro
    'Converte um texto em currency
    F_TextToVal = CCur(strNumero)
Exit Function
TrataErro:
    F_TextToVal = 0
End Function

Function F_AjustaValor(ByRef strValor As String, Optional sFormat As String = "########0.00") As String
On Error GoTo TrataErro
    If Left(Trim(strValor), 1) = "," Or Left(Trim(strValor), 1) = "." Then
        strValor = "0" & Trim(strValor)
    End If
    If strValor = "" Then strValor = "0,00"
    strValor = Format(strValor, sFormat)
    If Left(strValor, 1) = "." Or Left(strValor, 1) _
    = "," Then
       Mid(strValor, 1) = ","
       strValor = "0" + strValor
    ElseIf Left(strValor, 1) = "0" Then
       If InStr(1, strValor, ".") > 0 Then
          Mid(strValor, InStr(1, strValor, "."), 1) = ","
       End If
    ElseIf CCur(strValor) = 0 Then
       F_AjustaValor = "0.00"
       Exit Function
    End If
    Mid(strValor, InStr(1, strValor, ","), 1) = "."
    F_AjustaValor = strValor
    Exit Function
TrataErro:
    F_AjustaValor = "0.00"
        
End Function


Function F_CalcICMSSubst(crcValor As Currency, crcICMS As Currency, crcIvaInt As Currency, _
crcAliquotaInt As Currency) As Currency
 F_CalcICMSSubst = crcValor * (1 + crcIvaInt) 'Valor c/despesa
 F_CalcICMSSubst = F_CalcICMSSubst * crcAliquotaInt
 F_CalcICMSSubst = F_CalcICMSSubst - crcICMS  'Dimuindo o crédito ICMS
 F_CalcICMSSubst = F_CalcICMSSubst + crcValor

End Function


Sub P_PreencheNotaSaida(cmbxOBJ As Object)
    'Preenche o tipo de operação da Nota Fiscal
    cmbxOBJ.Clear
    cmbxOBJ.AddItem "5.12-Vendas Merc.Estadual"
    cmbxOBJ.AddItem "6.12-Vendas Merc.InterEstadual"
    cmbxOBJ.AddItem "5.32-Devolução Estadual"
    cmbxOBJ.AddItem "6.32-Devolução InterEstadual"
    cmbxOBJ.AddItem "5.22-Transf. Merc. Estadual"
    cmbxOBJ.AddItem "6.22-Transf. Merc InterEstadual"
    cmbxOBJ.AddItem "5.73-Venda Merc.Retida Estadual"
    cmbxOBJ.AddItem "5.99-Outras Saídas Estadual"
    cmbxOBJ.AddItem "6.99-Outras Saídas InterEstadual"
    cmbxOBJ.AddItem "5.93-Trans Mat.Consumo Estadual"
    cmbxOBJ.AddItem "6.93-Trans Mat.Consumo InterEstadual"
End Sub

Sub P_PreencheSituacaoTributaria(cmbxOBJ As ComboBox)
    'Carrega um combobox com as situações da tributação do ICMS
    cmbxOBJ.Clear
    cmbxOBJ.AddItem "00-Tributada Integralmente"
    cmbxOBJ.AddItem "01-Tributada e C/Cobr do ICMS por Subst"
    cmbxOBJ.AddItem "02-Com Redução de Base de Cálculo"
    cmbxOBJ.AddItem "03-Isenta ou não tributada C/Cobr ICMS Subst"
    cmbxOBJ.AddItem "04-Isenta ou não tributada"
    cmbxOBJ.AddItem "05-Com Suspensão ou diferimento"
    cmbxOBJ.AddItem "06-ICMS Cobrado Ant. por Subst Trib"
    cmbxOBJ.AddItem "07-C/Redução De Base de Calc C/Cobr ICMS Subst"
    cmbxOBJ.AddItem "09-Outras"
End Sub

Sub P_PreencheTaxasICMS(cmbxOBJ As ComboBox)
    cmbxOBJ.Clear
    cmbxOBJ.AddItem "00 %"
    cmbxOBJ.AddItem "05 %"
    cmbxOBJ.AddItem "07 %"
    cmbxOBJ.AddItem "12 %"
    cmbxOBJ.AddItem "17 %"
    cmbxOBJ.AddItem "18 %"
End Sub
Sub P_AjustaNota(ByRef objNota As TextBox, _
Optional intTam As Integer = 8)
    'Formata um objeto com zeros na frente do número, que deve ser maior que 0
    On Error GoTo TrataErro:
    objNota = F_SoNumero(objNota.Text)
    If CCur(objNota.Text) < 1 Then
       objNota.SetFocus
       Beep
       MsgBox "Número inválido!", vbExclamation, conSigla_Sistema
       objNota.SetFocus
       Exit Sub
    End If
    objNota.Text = Format(objNota.Text, "0" & String(intTam - 1, "#"))
    Exit Sub

TrataErro:
    Exit Sub
    
End Sub



Sub P_AjustaEstoqueReservado(cod_pro As String)
        
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim total As Currency
    
    strSQL = "select ti.cod_emp,sum(ti.qtd_ite) as total from " & TAB_VENDAS & " as tv inner join  " & TAB_ITVENDAS & " as ti "
    strSQL = strSQL & " on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " and situacao_Ven<4 "
    strSQL = strSQL & " group by cod_emp"
    
    
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        strSQL = "Update estoque set reservado=0 where cod_pro='" & cod_pro & "'"
        conn.Execute strSQL
        Do While Not .EOF
            
        
            
            strSQL = "update estoque set reservado=" & F_GravaDec(!total)
            strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
            strSQL = strSQL & " and cod_emp='" & !cod_emp & "' "
        
            conn.Execute strSQL
            total = total + !total
            .MoveNext
        Loop
        strSQL = "update " & TAB_PRODUTOS & " set reservado_pro=" & F_GravaDec(total)
        strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
        conn.Execute strSQL
        
        MsgBox "Reservado: " & total, vbInformation
    End With
    
    
    
    
    
End Sub
Sub ajustaEstoque()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ti.cod_pro,ti.cod_emp,sum(qtd_ite) as total,reservado from " & TAB_ITVENDAS
    strSQL = strSQL & " as ti inner join " & TAB_VENDAS & " as tv on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join estoque as e on e.cod_pro=ti.cod_pro and e.cod_emp=ti.cod_emp"
    strSQL = strSQL & " where situacao_ven<4 group by ti.cod_pro,ti.cod_emp having total <> reservado"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update estoque set reservado=" & F_GravaDec(!total)
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
            strSQL = strSQL & " and cod_emp='" & !cod_emp & "'"
            conn.Execute strSQL
            
            .MoveNext
            
        Loop
        .Close
    End With
    
    P_Status

 strSQL = "UPDATE produtos SET er_pro = "
    strSQL = strSQL & " (SELECT SUM(er) AS total FROM estoque AS e"
    strSQL = strSQL & " WHERE produtos.cod_pro=e.cod_Pro)"
    strSQL = strSQL & "WHERE cod_pro IN(SELECT cod_pro FROM estoque WHERE er>0)"
    conn.Execute strSQL
    
    strSQL = "UPDATE produtos SET reservado_pro = "
    strSQL = strSQL & " (SELECT SUM(reservado) AS total FROM estoque AS e"
    strSQL = strSQL & " WHERE produtos.cod_pro=e.cod_Pro )"
    strSQL = strSQL & "WHERE cod_pro IN(SELECT cod_pro FROM estoque WHERE reservado>0)"
    conn.Execute strSQL
    


End Sub

Function F_SenhaCodBar() As String
    On Error Resume Next
    With frmSenhaCodBar
        .Show 1
        
        F_SenhaCodBar = .tag
        
        Unload frmSenhaCodBar
    End With
End Function
Function F_senha(Optional bLockUser As Boolean = True) As Boolean
    Dim sSenha As String
    Dim strSQL As String
    
        frmLogin.tag = "senha"
        On Error Resume Next
        frmLogin.txtbxNome_Usu.Text = strNome_Usu
        If bLockUser = True Then
            frmLogin.txtbxNome_Usu.Locked = True
        Else
            frmLogin.txtbxNome_Usu.Locked = False
        End If
        frmLogin.BorderStyle = 1
        TrocaTabKey 9
        frmLogin.Show 1
        If frmLogin.tag = "" Then
               
            F_senha = False
        Else
            F_senha = True
        End If
        Unload frmLogin
End Function
Sub P_ImprxEstiquetas(obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from " & TAB_PRODUTOS & " order by nome_pro "

    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            'obj.PaintPicture mdiFrm.ImageListPrint.ListImages(1).Picture, -10, 5, 500, 400
            obj.fontsize = 14
            P_ImprX obj, strNome_Emp, 10
            
            'obj.Print
            If !pei_pro <> "" Then
                P_ImprX obj, "PEI: " & !pei_pro, 50, False
                            
            Else
                P_ImprX obj, "", 50, False
            End If
            
            P_ImprX obj, Format(!cod_pro, "00#\.###\-#"), 0
            P_ImprX obj, "R$ " & FormatVal(!preco_pro), 30, False
            obj.fontsize = 8
            P_ImprX obj, !nome_pro, 0, False
            
            obj.Line (0, 0)-(200, 0)
            obj.Line (60, 0)-(60, 600)
            obj.Line (120, 0)-(120, 600)
            obj.Line (180, 0)-(180, 600)
            .MoveNext
        Loop
    End With
    
End Sub
Function F_Troco(num_ven As String) As String
    'On Error Resume Next
    frmTroco.tag = num_ven
    
    frmTroco.Show 1 ', mdiFrm
    'frmTroco.SetFocus
    'Do While frmTroco.Visible = True
    '
    '    DoEvents
    'Loop
    '
    
    'On Error Resume Next
    'Unload frmMsg
    F_Troco = frmTroco.tag
    Unload frmTroco
    
End Function
Function F_UsuLogadoCaixa(sCod_Cai As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select tf.guerra_fun from " & TAB_TURNCAI & " as tt, "
    strSQL = strSQL & " tab_Funcionarios as tf where cod_cai='" & sCod_Cai & "' "
    strSQL = strSQL & " and cod_Tur='" & strCod_Tur & "' "
    strSQL = strSQL & " and data_tur=" & F_GravaData(Date)
    strSQL = strSQL & " and tf.cod_fun=tt.login_usu "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then
            F_UsuLogadoCaixa = !guerra_fun
        End If
        .Close
    End With
End Function

Function QuebraLinha(Texto As String, tamanho As Integer) As String
    Dim iCount As Integer
    Dim iAux As Integer
    Dim sAux As String
    Dim sPalavra As String
    Dim sFinal As String
    For iCount = 1 To Len(Texto)
        iAux = iAux + 1
        sAux = Mid(Texto, iCount, 1)
        If sAux = " " Or sAux = "-" Or sAux = "/" Or sAux = "_" Then
            
            If iAux - 1 + Len(sPalavra) >= tamanho + 1 Then
                iAux = 0
                
                sFinal = sFinal & Chr(13) & sPalavra & sAux
                
            Else
                sFinal = sFinal & sPalavra & sAux
            End If
            sPalavra = ""
        Else
            sPalavra = sPalavra & sAux
        End If
    Next
    sFinal = sFinal & sPalavra
    QuebraLinha = sFinal
End Function
Function FormatCodigo(sCod As String) As String
    On Error Resume Next
    FormatCodigo = Format(sCod, "00#\.####\-#")
End Function
Function F_EscolheVendedor() As String
    'On Error Resume Next
    frmEscolheVendedor.Show 1
    F_EscolheVendedor = frmEscolheVendedor.tag
    Unload frmEscolheVendedor
End Function
Function F_Vendedor() As String
    On Error Resume Next
    frmVendedor.Show 1
    F_Vendedor = frmVendedor.tag

    Unload frmVendedor
End Function

Function F_InList(msf As ListBox, Text As String) As Integer
    Dim iCount As Integer
    Dim bOk As Boolean
    With msf
        bOk = False
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            
            If UCase(.Text) = UCase(Text) Then
                bOk = True
                F_InList = iCount
                Exit For
            End If
        Next
        If bOk = False Then
            F_InList = -1
        End If
        
    End With
End Function

Function F_InFlxgrd(msf As Object, Text As String, colum As Integer, Optional contains As Boolean = False) As Integer
    Dim iCount As Integer
    Dim bOk As Boolean
    With msf
        bOk = False
        For iCount = 0 To .rows - 1
            .row = iCount
            .col = colum
            
            If contains = False Then
                If UCase(.Text) = UCase(Text) Then
                    bOk = True
                    F_InFlxgrd = iCount
                    Exit For
                End If
            Else
                If InStr(UCase(.Text), UCase(Text)) > -1 Then
                    bOk = True
                    F_InFlxgrd = iCount
                    Exit For
                End If
            End If
            
        Next
        If bOk = False Then
            F_InFlxgrd = -1
        End If
        
    End With
End Function

Sub P_mensagem(msg As String, frm As Form, Optional acao As Boolean = True)
    On Error Resume Next
    If acao = True Then
        frmMsg.lblMsg.caption = msg
        frmMsg.Show 0, frm
        
    Else
        Unload frmMsg
    End If
End Sub
Function F_senhaReturnUsu(Optional bLock As Boolean = False, Optional novoUsu As Boolean = False) As String
    Dim sSenha As String
    Dim strSQL As String
    Dim sCod As String
    Dim sNome As String
    sCod = strCod_Usu
    sNome = strNome_Usu
        frmLogin.tag = "senha"
        On Error Resume Next
        
        frmLogin.txtbxNome_Usu.Text = strNome_Usu
        frmLogin.txtbxNome_Usu.Locked = bLock
        frmLogin.BorderStyle = 1
        TrocaTabKey 13
        frmLogin.Show 1
        
            If frmLogin.tag <> "" Then
                F_senhaReturnUsu = frmLogin.tag
            End If
        If novoUsu = False Then
            strNome_Usu = sNome
            strCod_Usu = sCod
            strNivel_Usu = F_BuscaUmCampo("nivel_usu", "tab_usuarios", "cod_fun='" & strCod_Usu & "'")
        End If
        
        Unload frmLogin
End Function

Sub P_Msg(Optional msg As String, Optional titulo As String)
   On Error Resume Next
   Dim s As String
   'On Error GoTo trata_erro:
   'modo=
   Dim c As New Collection
   Set c = F_SeparaTexto(msg, Chr(0))
   
    If msg <> "" Then
        If Trim(msg) = "" Then
            Exit Sub
        End If
    End If
    
    If msg <> "" Then
        frmMsg.caption = titulo
        frmMsg.lblMsg.caption = F_AjustaTexto(c(1))
        frmMsg.cmndbtnCancela.Visible = False
        If Not IsLoad("frmMsg") Then
                
                
            
                
        End If
        DoEvents
        frmMsg.Show 0, mdiFrm
                frmMsg.lblMsg.Refresh
                frmMsg.SetFocus
        
    
        
        
        'frmMsg.Refresh
    Else
        Unload frmMsg
    End If
    Screen.MousePointer = vbDefault
    Screen.ActiveForm.Refresh
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Function F_MsgModal(Optional msg As String, Optional titulo As String) As Integer
   On Error Resume Next
   On Error Resume Next
   'modo=
   
    msg = Trim(msg)

'    MsgBox msg & "-" & Len(msg)
    
    Screen.MousePointer = vbHourglass
    
    If msg <> "" Then
        If Trim(msg) = "" Then
            Exit Function
        End If
    End If
    
        
        frmMsg.caption = titulo
        frmMsg.lblMsg.caption = (msg)
        frmMsg.cmndbtnCancela.Visible = True
        'frmMsg.MDIChild = False
        If Not IsLoad("frmMsg") Then
            frmMsg.Show 0
        End If
        
        If frmMsg.tag = "1" Then
'            MsgBox "aqui "
            F_MsgModal = 1
            Unload frmMsg
        End If
        
        
       DoEvents
        
    
        Screen.ActiveForm.Refresh
        
        
        
    
    Exit Function
trata_Erro:
    MostraErro
End Function

Sub P_MsgOk(Optional msg As String, Optional titulo As String)
   'On Error GoTo trata_erro
   On Error Resume Next
    If msg <> "" Then
        If Trim(msg) = "" Then
            Exit Sub
        End If
    End If
'lembrar de arrumar aqui
'para rodar do começo ao fim da string
'e quebrar quando encontrar o chr(0)
'        frmMsgOk.lblMsg.Alignment = 0
        frmMsgOk.lblMsg.FontName = "courier"
        frmMsgOk.lblMsg.fontsize = 10
        frmMsgOk.caption = titulo
        frmMsgOk.lblMsg.caption = (F_AjustaTexto(msg))
        frmMsgOk.Picture1.align = 0
        frmMsgOk.lblMsg.Alignment = 0
        frmMsgOk.lblMsg.BackStyle = 1
        frmMsgOk.lblMsg.AutoSize = False
        frmMsgOk.lblMsg.width = frmMsgOk.Picture1.width - 300
        'frmMsgOk.lblMsg.BackColor = vbYellow
        'frmMsgOk.lblMsg.FontName = "Courier New"
        
        
        frmMsgOk.Show 1
        'frmMsgOk.Refresh
        Screen.ActiveForm.Refresh
    Exit Sub
trata_Erro:
    MostraErro
End Sub

Function F_PrimDiaMes(Optional sData As String = "") As String
    If sData = "" Then
        F_PrimDiaMes = "01/" & Right(Format(Date, "dd/mm/yyyy"), 7)
    Else
        sData = Format(sData, "dd/mm/yyyy")
        F_PrimDiaMes = "01/" & Right(Format(sData, "dd/mm/yyyy"), 7)
    End If
    F_PrimDiaMes = CDate(F_PrimDiaMes)
End Function
Function F_UltDiaMes(Optional sData As String = "") As String
    Dim sAux As String
    
        
        sAux = DateAdd("M", 1, F_PrimDiaMes(sData))
        sAux = DateAdd("D", -1, sAux)
        F_UltDiaMes = sAux
        
        
    
    F_UltDiaMes = CDate(F_UltDiaMes)
End Function
Sub ZebraGrid_Old(grid As MSFlexGrid, linhaIni As Integer, Optional cor = &HD3E7EB)



    On Error Resume Next
    Dim iCount As Integer
    Dim iRow As Integer
    With grid
        For iRow = linhaIni To grid.rows - 1 Step 2
        .row = iRow
            For iCount = 0 To .Cols - 1
                .col = iCount
                .CellBackColor = cor
            
            Next
        Next
    End With
End Sub
Sub ColoreLinhaGrid_Old(grid As MSFlexGrid, linha As Integer, Optional cor = &HD3E7EB, _
    Optional negrito As Boolean = False, Optional foreCor = "")



    On Error Resume Next
    Dim iCount As Integer
    Dim iRow As Integer
    With grid
        .row = linha
        For iCount = 0 To .Cols - 1
            .col = iCount
            If cor <> "" Then
                .CellBackColor = cor
            End If
            .CellFontBold = negrito
            If foreCor <> "" Then
                .CellForeColor = foreCor
            End If
        Next
        
    End With
End Sub
Sub P_ajustaGridWid(msflxgrd As MSFlexGrid, Optional iPrimLinha As Boolean = False)
    Dim iCount As Integer
    Dim iRow As Integer
    With msflxgrd
        msflxgrd.Visible = False
        For iRow = 0 To .rows - 1
            .row = iRow
            For iCount = 0 To .Cols - 1
                
                .col = iCount
                If .CellWidth < Printer.TextWidth(.Text) Then
                    .ColWidth(.col) = Printer.TextWidth(.Text)
                End If
            Next
        Next
        msflxgrd.Visible = True
    End With
End Sub
Sub P_EspelhoNota(seq_not As String)
'    Dim fr As New frmEspelhoNota
'    fr.Tag = seq_Not
'    fr.Show 0
End Sub
Sub P_cancelaNotaFiscalSaida(nota_sai As String, serie_sai As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sMsg As String
    strSQL = "select tp.nome_pro,tp.ef_pro,ti.* from tab_itsainf as ti, " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and ti.seq_sai='" & nota_sai & "' "
    
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "update " & TAB_PRODUTOS & " set ef_pro=ef_pro+" & F_GravaDec(!qtd_its)
            strSQL = strSQL & " where cod_pro ='" & !cod_pro & "' "
            conn.Execute strSQL
            
                sMsg = sMsg & !cod_pro & " - " & !nome_pro & " Anterior: " & !ef_pro & " Atual: " & !ef_pro + !qtd_its & Chr(13)
            
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update tab_saidasnf set situacao_sai='3', "
    strSQL = strSQL & " cod_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where seq_sai='" & nota_sai & "' "
    
    conn.Execute strSQL
    MsgBox "Nota Fiscal Cancelada!" & Chr(13) & sMsg, vbInformation
End Sub
Sub P_cancelaCupon(seq_cup As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_itcup where seq_cup='" & seq_cup & "' "
    
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "update " & TAB_PRODUTOS & " set ef_pro=ef_pro+" & F_GravaDec(!qtd_its)
            strSQL = strSQL & " where cod_pro ='" & !cod_pro & "' "
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update cupons set cancelado_cup='S' "
    strSQL = strSQL & " where seq_cup='" & seq_cup & "' "
    
    conn.Execute strSQL
    MsgBox "Nota Fiscal Cancelada!", vbInformation
End Sub

Sub P_AjustaVenda()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sMsg As String
    strSQL = "select * from " & TAB_VENDAS & " where situacao_ven in('2','3') "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("num_ven", " " & TAB_VENFEC & " ", "num_ven='" & !num_ven & "'") = "" Then
                strSQL = "Update " & TAB_VENDAS & " set situacao_ven='1' "
                strSQL = strSQL & " where num_ven='" & !num_ven & "' "
                conn.Execute strSQL
                sMsg = sMsg & "Venda: " & !num_ven & " - " & !nomecli_ven & " Valor: " & FormatVal(!valor_ven)
               End If
            .MoveNext
        Loop
        .Close
        If sMsg <> "" Then
            strSQL = "O Sistema encontrou vendas com inconsistências e, as reabriu,"
            strSQL = strSQL & " por gentileza, informe o reponsável para que as mesmas"
            strSQL = strSQL & " possam ser fechadas novamente."
            strSQL = strSQL & Chr(13) & sMsg
            MsgBox strSQL, vbCritical
        End If
    End With
 
End Sub
Function EsticaTexto(Texto As String, tam As Integer, Optional align As Integer = 0, _
Optional Carac As String = " ", Optional soNum As Boolean = False) As String
    If soNum Then
        Texto = F_SoNumero(Texto)
    End If
    
    
    If Len(Texto) >= tam Then
        EsticaTexto = Left(Texto, tam)
    Else
        If align = 0 Then
            EsticaTexto = Texto & String(tam - Len(Texto), Carac)
        ElseIf align = 1 Then
            EsticaTexto = String((tam - Len(Texto)) / 2, Carac) & Texto & String((tam - Len(Texto)) / 2, Carac)
        Else
            EsticaTexto = String(tam - Len(Texto), Carac) & Texto
        End If
    End If
End Function

Function F_EsticaTexto(Texto As String, tam As Integer, Optional align As Integer = 0) As String
    If tam <= 0 Then Exit Function
    If Len(Texto) >= tam Then
        F_EsticaTexto = Left(Texto, tam)
    Else
        
        If align = 0 Then
            F_EsticaTexto = Texto & String(tam - Len(Texto), " ")
        ElseIf align = 1 Then
            F_EsticaTexto = String((tam - Len(Texto)) / 2, " ") & Texto & String((tam - Len(Texto)) / 2, " ")
        Else
            F_EsticaTexto = String(tam - Len(Texto), " ") & Texto
        End If
    End If
End Function


Function verificaVendaAberta(cod_cai As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    'On Error Resume Next
    strSQL = "select tf.* from  " & TAB_VENDAS & " as tv, "
    strSQL = strSQL & "  " & TAB_VENFEC & "  as tf "
    strSQL = strSQL & " where tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & cod_cai & "' "
    strSQL = strSQL & " and situacao_ven=1 "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then
            verificaVendaAberta = True
        End If
        .Close
    End With
End Function

Sub VerificaAtualizacao()
    Dim strSQL As String
    On Error GoTo trata_Erro
    Dim sDest As String
    Dim sOrig As String
    'Exit Sub
    
    If App.LogMode = 0 Then
      '  Exit Sub
    End If
        
    sDest = App.path & "\" & App.EXEName & ".exe"
    sOrig = strArquivoServidor & App.EXEName & ".exe"
    
    If FileDateTime(sOrig) > FileDateTime(sDest) Then
        strSQL = "Existe uma versão mais nova para atualização."
        strSQL = strSQL & "Assim que possível saia do  " & strNomeSis & " e execute o programa de instalação."
        
            Open "c:\atlz_pk.txt" For Output As #1
            Print #1, sOrig
            Print #1, sDest
        
        Close #1
        
        If MsgBox(strSQL, vbYesNo + vbQuestion) = vbYes Then
                
            Shell App.path & "\atlz_pk.exe", vbNormalFocus
            P_Fim
        End If
        
        
                
        
    End If
    Exit Sub
trata_Erro:
    
    
End Sub
Function F_RetornaPrimEspaco(sText As String, iLocal As Integer) As Integer
    Dim sAux As String
    Dim iAux As Integer
    Dim iCount As Integer
    sAux = Left(sText, iLocal)
    F_RetornaPrimEspaco = iLocal
    For iCount = iLocal To 1 Step -1
        If InStr(1, " -;.!_+|=:,<>", Mid(sAux, iCount, 1)) Then
            F_RetornaPrimEspaco = iCount
            Exit Function
        End If
    Next
    
End Function

Sub P_notaSEPD(num_ven As String, strCUPOM As String)
    
    frmGeraNota.tag = num_ven
    frmGeraNota.txtbxVenda.Text = num_ven
    If strCUPOM <> "" Then
        frmGeraNota.lblCUPOM.caption = strCUPOM
        frmGeraNota.lblCUPOM.tag = strCUPOM
    End If
    frmGeraNota.Show 1
End Sub
Function F_CalculaVenc(database As String, parcelas As Integer, intervalo As Integer) As Date
    If parcelas = 1 Then
        F_CalculaVenc = Date
    Else
        F_CalculaVenc = DateAdd("D", intervalo * (parcelas - 1), database)
    End If
    Exit Function
trata_Erro:
    F_CalculaVenc = Date
End Function

Sub cancela_Cupom()
     Dim strSQL As String
    Dim sNumVen As String
    Dim strNumCupom As String
    strNum_Cpm = "      "
    
    Dim strNota_Sai As String
    strNota_Sai = "      "
    i = ecf_NumeroCupom(strNota_Sai)
    
    If MsgBox("Confirma cancelamento do Último Cupom Fiscal? Nº" & strNota_Sai, vbYesNo + vbCritical) = vbYes Then
        
        
        ecf_CancelaCupom
                
        sNumVen = F_BuscaUmCampo("num_ven", "cupons", "num_cup='" & strNota_Sai & "' and cod_cai='" & strCod_Cai & "'")
        
        
        If sNumVen = "" Then
            MsgBox "Não foi encontrada a venda do cupom fiscal nº " & strNota_Sai, vbExclamation
            Exit Sub
        End If
        P_RetornaEstFiscalCancelamento strNota_Sai, strCod_Cai
        
        If sNumVen <> "" Then
            If MsgBox("Confirma Cancelamento da venda " & sNumVen, vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
        End If
                
        If Not F_CancelaVendaAposCupom(sNumVen) Then
            MsgBox "Falha ao cancelar a venda!", vbCritical
        Else
            MsgBox "Venda Cancelada! Estoque Reposto! Lançamento em Caixa Concluído!", vbInformation
        End If
        
        strSQL = "update tab_saidas set serie_sai='CANC' where num_ven='" & sNumVen & "' "
        strSQL = strSQL & " and nota_Sai='" & strNota_Sai & "'"
        conn.Execute strSQL
    
        
    
    
    
    End If
End Sub
Sub P_MostraErro(Optional strProc As String = "Erro ")
    MostraErro strProc
End Sub
Function F_Quebra(Text As String, qtd_por_linha As Integer) As String
    On Error Resume Next
    Dim iCount As Integer
    Dim strAux As String
    Dim iAux As Integer
    Dim jCount As Integer
    Dim strBuf As String
    Dim strRest As String
    For iCount = 1 To Len(Text)
        iAux = iAux + 1
        strAux = strAux + Mid(Text, iCount, 1)
        If iAux >= qtd_por_linha Then
            If InStr(1, " -;.!_+|=:,<>", Mid(Text, iCount, 1)) Then
                
                strAux = strAux + Chr(13) + Chr(10)
                strBuf = strAux
                iAux = 0
                strest = ""
                
            End If
        
        End If
    Next
    F_Quebra = strAux
End Function
Function F_CFOP() As String
    frmCFOP.Show 1
    F_CFOP = frmCFOP.tag
    Unload frmCFOP
End Function
Sub resizeForm(obj As Form)
    obj.height = mdiFrm.ScaleHeight - 50
    obj.width = mdiFrm.ScaleWidth - 100
    
    If obj.MDIChild = True Then
    
        obj.Top = 50
        obj.Left = 50
    Else
        obj.Top = buscaTop
        obj.Left = buscaLeft
        'P_CentraForm obj, 0
        
    End If
End Sub
Sub posicionaAoFim(obj As Object, container As Object, Optional recuo As Integer = 100)
    On Error Resume Next
    obj.Top = container.height - obj.height - recuo
End Sub
Sub posicionaADireita(obj As Object, container As Object, Optional recuo As Integer = 100)
    On Error Resume Next
    obj.Left = container.width - obj.width - recuo
End Sub

Sub PosicionaGradeWidth(grade As Object, frm As Object, Optional recuo As Integer = 200)
    On Error Resume Next
    grade.width = frm.width - grade.Left - recuo
End Sub
Sub PosicionaGradeLeft(grade As Object, frm As Object, Optional recuo As Integer = 300)
    On Error Resume Next
    grade.width = frm.Left - grade.Left - recuo
End Sub

Sub PosicionaGradeHeigth(grade As Object, frm As Object, Optional recuo As Integer = 200, Optional bContainer As Boolean = False)
    On Error Resume Next
    If Not bContainer Then
        grade.height = frm.Top - grade.Top - recuo
    Else
        grade.height = frm.height - grade.Top - recuo
    End If
End Sub
Sub PosicionaGradeTop(grade As Object, frm As Object, Optional recuo As Integer = 200)
    On Error Resume Next
    
    grade.height = frm.Top - grade.Top - recuo
    
End Sub

Sub PosicionaGrade(grade As Object, objW As Object, objH As Object, Optional recuo As Integer = 200, Optional bContainer As Boolean = False)
    On Error Resume Next
    PosicionaGradeWidth grade, objW
    PosicionaGradeHeigth grade, objH, recuo, bContainer
    
End Sub

Function buscaTop() As Single
    buscaTop = 650
End Function
Function buscaLeft()
    If mdiFrm.pcMenu.Visible = True Then
        buscaLeft = mdiFrm.pcMenu.width + 100
    Else
        buscaLeft = 50
    End If
    
End Function
Sub P_VisualizaContempla(cod_cli As String, seq_pre As String)
    frmVisualizaContempla.lblNomeCli.tag = cod_cli
    frmVisualizaContempla.tag = seq_pre
    frmVisualizaContempla.Show 1
        
End Sub

Sub P_cancelanotaFechada(seq_not As String, Optional tipo As String = "F")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTab As String
    Dim sTabIte As String
    Dim sEst As String
    
    sTab = "sped_nf"
    sTabIte = "sped_nf_itens"
    sEst = "er_pro"
    
    
    strSQL = "select tc.*, ti.* from sped_nf as tc inner join sped_nf_itens as ti on tc.id=ti.id_nf"
    strSQL = strSQL & " where tc.id='" & seq_not & "'"
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update " & TAB_PRODUTOS & " set er_pro=er_pro-" & F_GravaDec(!qtd_convert)
            If !tipo_entrada = "01" Then
                strSQL = strSQL & ",ef_pro=ef_pro-" & F_GravaDec(!qtd_convert)
            End If
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            conn.Execute strSQL
            
            strSQL = "update estoque set er=er-" & F_GravaDec(!qtd_convert)
            If !tipo_entrada = "01" Then
                strSQL = strSQL & ",ef=ef-" & F_GravaDec(!qtd_convert)
            End If
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and cod_emp='" & !cod_emp & "'"
            conn.Execute strSQL
            
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update " & sTab & " set status=5 where id='" & seq_not & "' "
    
    conn.Execute strSQL
    
    MsgBox "Nota cancelada, Estoque reposto!", vbExclamation
    
End Sub

Function F_GravaClass(seq_fin As String, cod_pla As String, valor_cla As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = strSQL & "insert into tab_class(seq_fin,cod_pla,valor_cla,cod_usu)values("
    strSQL = strSQL & "'" & seq_fin & "',"
    strSQL = strSQL & "'" & cod_pla & "',"
    strSQL = strSQL & F_GravaDec(valor_cla) & ","
    strSQL = strSQL & "'" & strCod_Usu & "') "
    conn.Execute strSQL
    F_GravaClass = True
    Exit Function
trata_Erro:
    F_GravaClass = False
End Function
Function F_SelectPlanoContas(expr As String, bCred As Boolean) As PlanoContas
    frmSelect.sSelect = "select cod_pla as cod,nome_pla as nome,tipo_pla as tipo from " & TAB_PLANOCONTAS
    frmSelect.cCriterio.Add " where nome_pla like '" & expr & "%' and tipo_pla='" & IIf(bCred, "C", "D") & "' "
    
    frmSelect.sOrd = " order by nome_pla "
    
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        F_SelectPlanoContas.cod_pla = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1
        
        F_SelectPlanoContas.nome_pla = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.col = 2
        
        F_SelectPlanoContas.tipo_pla = frmSelect.msflxgrd
        
        
    End If
    Unload frmSelect
    
End Function

Public Function Extenso(ByVal valor As _
       Double, Optional ByVal MoedaPlural As _
       String = "Reais", Optional ByVal MoedaSingular As _
       String = "Real") As String
  Dim strValor As String, Negativo As Boolean
  Dim buf As String, Parcial As Integer
  Dim Posicao As Integer, Unidades
  Dim Dezenas, Centenas, PotenciasSingular
  Dim PotenciasPlural

  Negativo = (valor < 0)
  valor = Abs(CDec(valor))
  If valor Then
    Unidades = Array(vbNullString, "Um", "Dois", _
               "Três", "Quatro", "Cinco", _
               "Seis", "Sete", "Oito", "Nove", _
               "Dez", "Onze", "Doze", "Treze", _
               "Quatorze", "Quinze", "Dezesseis", _
               "Dezessete", "Dezoito", "Dezenove")
    Dezenas = Array(vbNullString, vbNullString, _
              "Vinte", "Trinta", "Quarenta", _
              "Cinqüenta", "Sessenta", "Setenta", _
              "Oitenta", "Noventa")
    Centenas = Array(vbNullString, "Cento", _
               "Duzentos", "Trezentos", _
               "Quatrocentos", "Quinhentos", _
               "Seiscentos", "Setecentos", _
               "Oitocentos", "Novecentos")
    PotenciasSingular = Array(vbNullString, " Mil", _
                        " Milhão", " Bilhão", _
                        " Trilhão", " Quatrilhão")
    PotenciasPlural = Array(vbNullString, " Mil", _
                      " Milhões", " Bilhões", _
                      " Trilhões", " Quatrilhões")

    strValor = Left(Format(valor, String(18, "0") & _
               ".000"), 18)
    For Posicao = 1 To 18 Step 3
      Parcial = val(Mid(strValor, Posicao, 3))
      If Parcial Then
        If Parcial = 1 Then
          buf = "Um" & PotenciasSingular((18 - _
                Posicao) \ 3)
        ElseIf Parcial = 100 Then
          buf = "Cem" & PotenciasSingular((18 - _
                Posicao) \ 3)
        Else
          buf = Centenas(Parcial \ 100)
          Parcial = Parcial Mod 100
          If Parcial <> 0 And buf <> vbNullString Then
            buf = buf & " e "
          End If
          If Parcial < 20 Then
            buf = buf & Unidades(Parcial)
          Else
            buf = buf & Dezenas(Parcial \ 10)
            Parcial = Parcial Mod 10
            If Parcial <> 0 And buf <> vbNullString Then
              buf = buf & " e "
            End If
            buf = buf & Unidades(Parcial)
          End If
          buf = buf & PotenciasPlural((18 - Posicao) \ 3)
        End If
        If buf <> vbNullString Then
          If Extenso <> vbNullString Then
            Parcial = val(Mid(strValor, Posicao, 3))
            If Posicao = 16 And (Parcial < 100 Or _
                (Parcial Mod 100) = 0) Then
              Extenso = Extenso & " e "
            Else
              Extenso = Extenso & ", "
            End If
          End If
          Extenso = Extenso & buf
        End If
      End If
    Next
    If Extenso <> vbNullString Then
      If Negativo Then
        Extenso = "Menos " & Extenso
      End If
      If Int(valor) = 1 Then
        Extenso = Extenso & " " & MoedaSingular
      Else
        Extenso = Extenso & " " & MoedaPlural
      End If
    End If
    Parcial = Int((valor - Int(valor)) * _
              100 + 0.1)
    If Parcial Then
      buf = Extenso(Parcial, "Centavos", _
            "Centavo")
      If Extenso <> vbNullString Then
        Extenso = Extenso & " e "
      End If
      Extenso = Extenso & buf
    End If
  End If
End Function


Function F_BuscaDataInicialFinal(Optional dataI As String = "", Optional dataF As String = "") As DataInicialFinal
    frmDataInicialFinal.txtbxDataFinal.Text = dataF
    frmDataInicialFinal.txtbxDataInicial.Text = dataI
    frmDataInicialFinal.Show 1
    F_BuscaDataInicialFinal.dataF = frmDataInicialFinal.txtbxDataFinal.Text
    F_BuscaDataInicialFinal.dataI = frmDataInicialFinal.txtbxDataInicial.Text
    Unload frmDataInicialFinal
End Function

Function F_BuscaDataInicial(dataI As String) As DataInicialFinal
    frmDataInicialFinal.txtbxDataFinal.Enabled = False
    frmDataInicialFinal.txtbxDataFinal.Text = dataI
    frmDataInicialFinal.txtbxDataInicial.Text = dataI
    frmDataInicialFinal.Show 1
    F_BuscaDataInicial.dataF = frmDataInicialFinal.txtbxDataFinal.Text
    F_BuscaDataInicial.dataI = frmDataInicialFinal.txtbxDataInicial.Text
    Unload frmDataInicialFinal
End Function

Function F_SelectCupom(num_cup As String) As String
    frmSelect.sSelect = "select tc.num_ven as venda,tc.num_cup as cupom,tv.nomecli_ven as cliente,tc.cod_cai as caixa,tc.cod_tur as turno, tc.data_cup as emissao from cupons as tc, " & TAB_VENDAS & " as tv "
    
    frmSelect.cCriterio.Add " where num_cup='" & num_cup & "' and tv.nuM_ven=tc.nuM_ven"
    
    
    frmSelect.sOrd = " order by data_cup "
    
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        F_SelectCupom = frmSelect.msflxgrd.Text
        
        
    End If
    Unload frmSelect
    
End Function

Sub P_msgProduto(Optional msg As String = "")
    If msg = "" Then
        Unload frmMsgProduto
    Else
        frmMsgProduto.lblMsg.caption = msg
        frmMsgProduto.Show 0
        
        frmMsgProduto.SetFocus
    End If
    
End Sub

Sub P_CarregaInvent(cmbxInvent As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select distinct datahora_inv from tab_invent order by datahora_inv desc"
        rPesq.Open strSQL, conn_client
    With rPesq
        cmbxInvent.Clear
        Do While Not .EOF
            cmbxInvent.AddItem !DataHora_Inv
            .MoveNext
        Loop
        .Close
    End With
End Sub
Function mudarCaractere(expressao As String, encontrar As String, trocar As String) As String
    mudarCaractere = Replace(expressao, encontrar, trocar)
End Function


Sub FechaObjeto(obj As Object)
    On Error Resume Next
    obj.Close
End Sub


Sub abreObjetoAdo(obj As Object, strSQL As String, Optional dataVinculado As Object = Null)
    On Error GoTo trata_Erro
    
    FechaObjeto obj
    obj.Open strSQL, conn_client
    
    
    
        
        
    Set dataVinculado.DataSource = obj
        
Exit Sub
trata_Erro:

End Sub

Sub TrocaTabKey(Optional key As Integer = 13)
    On Error Resume Next
    If key = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Sub P_Preenche(obj As Object, tabela As String, campo As String, Optional Where As String = "", Optional sOrderBy As String = "", Optional emBranco As Boolean = False)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select " & campo & " as campo from " & tabela
    If Where <> "" Then
        strSQL = strSQL & " where " & Where
    
    End If
    If sOrderBy = "" Then
        strSQL = strSQL & " order by " & campo
    Else
        strSQL = strSQL & " order by " & sOrderBy
    End If
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    If emBranco Then
        obj.AddItem ""
    End If
    With rPesq
        Do While Not .EOF
            obj.AddItem TiraNull(rPesq!campo)
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheFormasPag(cmbx As Object, Optional cod_tip As String = "")
    Dim sWh As String
    
    If cod_tip <> "" Then
        sWh = "cod_tip='" & cod_tip & "'"
    End If
    
End Sub

Sub P_BuscaDefUsu()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    If strCod_Usu = "" Then
        Exit Sub
    End If
    
    strSQL = "select * from " & TAB_DEFUSU & " where cod_usu='" & strCod_Usu & "'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        On Error Resume Next
        strAbrevenda_usu = !abrevenda_usu
        strAbrenota_usu = !abrenota_usu
        strEfetnota_usu = !efetnota_usu
        strAbresaida_usu = !abresaida_usu
        strLancaixa_usu = !lancaixa_usu
        strLandoc_usu = !landoc_usu
        strLanche_usu = !lanche_usu
        strBaixadoc_usu = !baixadoc_usu
        strBaixache_usu = !baixache_usu
        strEdicai_usu = !edicai_usu
        strExclcai_usu = !exclcai_usu
        strExcldoc_usu = !excldoc_usu
        strCadprod_usu = !cadprod_usu
        strAltprod_usu = !altprod_usu
        strAltpreco_usu = !altpreco_usu
        strLiberarCli_usu = !liberarcliente_usu
        '
        
        strVisaoCaixa_usu = !visaocaixa_usu
        strVisaoConsignado_usu = !visaoconsignado_usu
        strVisaoEstoque_usu = !visaoestoque_usu
        strVisaoCrediario_usu = !visaocrediario_usu
        strVisaoFinanceiro_usu = !visaofinanceiro_usu
        strVisaoGeral_usu = !visaogeral_usu
        strVisaoPacote_usu = !visaopacote_usu
        strVisaoVendas_usu = !visaovendas_usu
        
        
        
        strAlterarVlrParc_Usu = !alterarvlrparcela_usu
        strExclVenda_usu = !exclvenda_usu
        strProrrogarPar_Usu = !prorrogarparcela_usu
        strAlterarTerceiroDoc_Usu = !alterarterceirodoc_usu
        strBaixarcheque_Usu = !baixarcheque_usu
        strExcluircheque_Usu = !excluircheque_usu
        strEfetivarVenda_Usu = !efetivarvenda_usu
        strCreditario_Usu = !crediario_usu
        strVisualizarvenda_Usu = !visualizarvenda_usu
        strLancarFin_Usu = !lancarfin_usu
        strEditarFin_Usu = !editarfin_usu
        strExcluirFin_Usu = !excluirfin_usu
        strTransferenciaInterna_usu = !transfinterna_usu
        strEfetuarDevCons_Usu = !efetuardevcons_usu
        strEfetuarAcerto_usu = !efetuaracerto_usu
        strEfetuarTransfCons_Usu = !efetuartransfcons_usu
        strExcluirAcerto_usu = !excluiracerto_usu
        strAlterarComRepr_Usu = !alterarcomrepr_usu
        strBaixarComRepr_usu = !baixarcomrepr_usu
        strCadastrarRepr_Usu = !cadastrarrepr_usu
        strCadFun_Usu = !cadastrarfun_usu
        strCadfor_Usu = !cadastrarfor_usu
        
        '
        
        strCadusu_usu = !cadusu_usu
        strExclVenda_usu = !exclvenda_usu
        strCadcli_usu = !cadcli_usu
        strCadcco_usu = !cadcco_usu
        strCadformaspag_usu = !cadformaspag_usu
        strAltformaspag_usu = !altformaspag_usu
        strViscusto_usu = !viscusto_usu
        strVisrelest_usu = !visrelest_usu
        strVisrelfin_usu = !visrelfin_usu
        strFechacai_usu = !fechacai_usu
        strFechames_usu = !fechames_usu
        strContrpre_usu = !contrpre_usu
        crcDesconto_usu = !desconto_usu
        crcDescontoBaixa_usu = !descontobaixa_usu
        strCorrigeEstCons_Usu = !corrigeestcons_usu
        
        strCancelaVenda_Usu = !cancelarvenda_usu
        strCancelarItem_Usu = !cancelaritem_usu
        
    End With
End Sub

Function RetornaMes(MesExtenso As String) As Integer
    On Error GoTo trata_Erro
    MesExtenso = LCase(MesExtenso)
    If Left(MesExtenso, 3) = "jan" Then
        RetornaMes = 1
    ElseIf Left(MesExtenso, 3) = "fev" Or Left(MesExtenso, 3) = "feb" Then
        RetornaMes = 2
    ElseIf Left(MesExtenso, 3) = "mar" Then
        RetornaMes = 3
    ElseIf Left(MesExtenso, 3) = "abr" Or Left(MesExtenso, 3) = "apr" Then
        RetornaMes = 4
    ElseIf Left(MesExtenso, 3) = "mai" Or Left(MesExtenso, 3) = "may" Then
        RetornaMes = 5
    ElseIf Left(MesExtenso, 3) = "jun" Then
        RetornaMes = 6
    ElseIf Left(MesExtenso, 3) = "jul" Then
        RetornaMes = 7
    ElseIf Left(MesExtenso, 3) = "ago" Or Left(MesExtenso, 3) = "aug" Then
        RetornaMes = 8
    ElseIf Left(MesExtenso, 3) = "set" Or Left(MesExtenso, 3) = "sep" Then
        RetornaMes = 9
    ElseIf Left(MesExtenso, 3) = "out" Or Left(MesExtenso, 3) = "oct" Then
        RetornaMes = 10
    ElseIf Left(MesExtenso, 3) = "nov" Then
        RetornaMes = 11
    ElseIf Left(MesExtenso, 3) = "dez" Or Left(MesExtenso, 3) = "dec" Then
        RetornaMes = 12
    End If
trata_Erro:
End Function

Function F_IsCaixa() As Boolean
    Dim strSQL As String
    Dim strHost As String
    On Error GoTo trata_Erro
    
    strSQL = F_BuscaUmCampo("COD_CAI", "TAB_CAIXASIP", "IP_CAI='" & _
        strIP_Usu & "'")
    If strSQL = "" Then
        strSQL = F_BuscaUmCampo("COD_CAI", "TAB_CAIXASIP", "IP_CAI='" & _
        strHost_Usu & "'")
    End If
    
    If strSQL = "" Then
        F_IsCaixa = False
    Else
        F_IsCaixa = True
    End If
    Exit Function
trata_Erro:
    F_IsCaixa = False
End Function

Function IsCaixa(IP As String) As Boolean
    Dim strSQL As String
    Dim strHost As String
    On Error GoTo trata_Erro
    
    strSQL = F_BuscaUmCampo("ip_cai", "TAB_CAIXASIP", "IP_CAI='" & _
        IP & "'")
    
    
    If strSQL = "" Then
        IsCaixa = False
    Else
        IsCaixa = True
    End If
    Exit Function
trata_Erro:
    IsCaixa = False
End Function


Sub PrimeiroAcesso()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select * from " & TAB_PARAMETROS
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        rPesq.Close
        Exit Sub
    End If
    strSQL = "select * from " & TAB_FUNCIONARIOS
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        rPesq.Close
        Exit Sub
    End If
    
    
    frmParametros.Show 1
    strSQL = "select * from " & TAB_PARAMETROS
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        MsgBox "A empresa não foi cadastrada!", vbExclamation
        P_Fim
    End If
    
    
    
End Sub
