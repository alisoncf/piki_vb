Attribute VB_Name = "Modulo"
 Public sBanco As String


''''Variáveis globais
Option Explicit
Global bUsaPrecoAtacado As Boolean
Global sCampoPrecoConsign As String
Global buscarAtlz As Boolean
Global CrLf As String

Global bUsaTEF As Boolean
Global bUsaECF As Boolean
Global trata_Erro As label
Global Conn As New adodb.Connection
Global Conn_Client As New adodb.Connection
Global Banco As database
Global StrConn As String
Global strConn_Old As String

Global strNum_Cpm As String

Global strCartaoCred As String
Global strDinheiro As String
Global strCartaoDeb As String
Global strTipoPeriodoPar As String 'M= Mês D =dia
Global sDriver As String
Global sserver As String
Global sUid As String
Global sPwd As String
Global sDataBase As String
Global sOption As String

Global strCodigoMunicipioIBGE As String
Global strCodigoEstadoIBGE As String
Global strCodigoPaisIBGE As String
Global strPais As String
Global strCor_Emp As String
Global strCaminhoQrCodeDoc As String

Global strCNAE_Emp As String
Global strEmail_emp As String
Global strIEST_Emp As String
Global strIM_Emp As String
Global strCRT_emp As CRT

Global strTipo_IcmsNFe As String


Public Type tImpressao

    PagIni As Integer
    PagFin As Integer
    Qualidade As Integer
End Type
Type DataIniFin
    dataI As String
    dataF As String
End Type

Public Type valChart ' para abrir o grafico valor, caption
    Value As Long
    caption As String
End Type

Public Type BlChart ' para carregar o valor de cada coluna, e quantas colunas tem
    val() As valChart
    Count As Long
End Type

Type Pedido
    cod_pro As String
    nome_pro As String
    qtd_pro As Currency
    unit_pro As Currency
    loja As String
End Type
Type Prevenda
    cod_pro As String
    nome_pro As String
    qtd_pro As Currency
    unit_pro As Currency
    
End Type


Type parcela
    vencimento As String
    valor As Currency
    boleto As String
    complemento As String
End Type
Sub P_LocalizaOperacao(sOp As String, obj As Object)
    Dim iCount As Integer
    Dim iLen As Integer
    iLen = Len(sOp)
    On Error Resume Next
    With obj
        For iCount = 0 To .ListCount - 1
            obj.ListIndex = iCount
            If Left(obj.text, iLen) = sOp Then
                Exit Sub
            End If
        Next
        'MsgBox "Operação não encontrada!", vbExclamation
        obj.ListIndex = -1
        obj.SetFocus
    End With
End Sub


Sub P_AjustaData(objData, Optional sFormato As String = "dd/mm/yyyy")

    On Error GoTo TrataErro
    If objData.text = "" Then Exit Sub
     
    objData.text = Left(objData.text, 10)
    
    
    
    If InStr(1, objData.text, "/") <> 0 Then
        
        
        If Month(Date) < 5 And Len(objData.text) <= 5 Then
            objData.text = Format(objData.text, sFormato)
            If Month(objData.text) > 5 Then
                objData.text = DateAdd("yyyy", -1, objData.text)
            End If
        Else
            objData.text = Format(objData.text, sFormato)
        End If
    Else
        If Len(objData.text) = 6 Then
            objData.text = Left(objData.text, 2) & "/" & Mid(objData.text, 3, 2) & "/" & Right(objData.text, 2)
            objData.text = Format(objData.text, "dd/mm/yyyy")
        ElseIf Len(objData.text) = 8 Then
            objData.text = Left(objData.text, 2) & "/" & Mid(objData.text, 3, 2) & "/" & Right(objData.text, 4)
            
        Else
            MsgBox "Sua digitação da Data está Errada!", vbExclamation
            objData.SetFocus
            Exit Sub
        End If
    End If
    
    If Len(objData.text) = 7 Then
    Else
    If Not F_Vazio(objData.text) And _
        F_ValidaData(objData.text) = "" Then
       MsgBox "Data inválida!", vbExclamation
       objData.SetFocus
       Exit Sub
    End If
    End If
    
    If CDate(objData.text) < "31/12/1900" Then
        
        MsgBox "Data inválida!,Verifique o Texto Digitado", vbExclamation
        objData.SetFocus
    End If
    Exit Sub

TrataErro:
    On Error Resume Next
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
        
        obj.top = (mdiFrm.height / 2) - (obj.height / 2) - iRecuo
        
        
        obj.Left = ((mdiFrm.width - 1320) / 2) - (obj.width / 2)
        
            
        
End Sub
Sub P_CentraObejct(obj As Object, frm As Object, _
        Optional bTop As Boolean = True, Optional bLeft As Boolean = True)
    On Error Resume Next
        If bTop Then
            obj.top = (frm.height / 2) - (obj.height / 2) - 200
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
Function F_ZerEsq(Num, Tam As Integer, _
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
aux = Right(String(Tam, "0") + aux, Tam) + Red
Red = Trim(str(Int(Num)))

If Acr <> "F" Then Num = Num - 1

F_ZerEsq = IIf(Len(Red) > Len(aux), String(Tam, "*") _
+ IIf(Dec > 0, "." + String(Dec, "*"), ""), aux)

End Function
Function F_AbreBancoODBC(b As database, nome As String) As Boolean
    
    On Error GoTo erro
    DBEngine.DefaultType = dbUseODBC
    'Set bTopa = OpenDatabase("topatudo", , False, "ODBC;DATABASE=topatudo;")
    
    
    
    Set b = OpenDatabase(nome, , False, "ODBC;DATABASE=" & nome & ";")
    
   
    F_AbreBancoODBC = True
    Exit Function
erro:
    MsgBox "Erro na Abertura do Banco!", vbExclamation, App.EXEName
    F_AbreBancoODBC = False
End Function
Function AbreBanco(conexao As adodb.Connection, database As String, Optional sUid As String = "root", _
Optional sserver As String = "localhost", _
Optional sPwd As String = "", _
Optional driver As String = "Driver={MySQL ODBC 3.51 Driver};") As Boolean
    On Error Resume Next
    
    conexao.Close
    
    
    On Error GoTo erro
    Dim strCon As String
    DBEngine.DefaultType = dbUseODBC
    strCon = "Driver={MySQL ODBC 3.51 Driver};" & _
    "Server=" & sserver & ";" & _
    "Port=3306;" & _
    "Database=" & database & ";" & _
    "Uid=" & Trim(sUid) & ";" & _
    "Pwd=" & Trim(sPwd) & ";" & _
    "Option=1 2 8 32 2048 16384;" '131072
    
    
    
    
    
    
    
    conexao.Open strCon
    
    conexao.CursorLocation = adUseClient
AbreBanco = True
    Exit Function
erro:
    


MsgBox Err.Description & Chr(13) & strCon, vbCritical, App.EXEName
    P_Fim
    AbreBanco = False
End Function


Function F_AbreBanco1(Optional conexao As String) As Boolean
    On Error Resume Next
    Dim conec As String
    Conn.Close
    
    Conn_Client.Close
    On Error GoTo erro
    Dim strCon As String
    DBEngine.DefaultType = dbUseODBC
    If conexao = "" Then
        strCon = "Driver={MySQL ODBC 3.51 Driver};" & _
        "Server=" & sserver & ";" & _
        "Port=3306;" & _
        "Database=" & sDataBase & ";" & _
        "Uid=" & Trim(sUid) & ";" & _
        "Pwd=" & Trim(sPwd) & ";" & _
        "Option=1 2 8 32 2048 16384;" '131072
    Else
        strCon = conexao
    End If
    
    
    
    
    
    
    Conn.Open strCon
    Conn_Client.Open strCon
    Conn_Client.CursorLocation = adUseClient
    
    strConn_Old = "requinte"
    conexao = "ODBC;DATABASE=" & sDataBase & ";port=3306;uid=root;server=10.1.1.10"
    'Set Banco = OpenDatabase(sDataBase)
 '   Set banco = OpenDatabase(sDataBase, , , conexao)
   
    F_AbreBanco1 = True
    Exit Function
erro:
    


MsgBox Err.Description & Chr(13) & strCon, vbCritical, App.EXEName
    P_Fim
    F_AbreBanco1 = False
End Function
Function F_AbreBanco(Optional nomeBanco As String = "") As Boolean
    Dim strConnect As String
    On Error GoTo erro
    DBEngine.DefaultType = dbUseODBC
   
    strConnect = "Driver=" & sDriver & ";"
    
    strConnect = strConnect & "Server=" & sserver & ";Database=" & sDataBase & ";Uid=" & sUid & ";port=3306;pwd=" & sPwd & ";option=" & sOption
    
    'strConnect = "market"
   Conn.Open strConnect
   Conn_Client.CursorLocation = adUseClient
   Conn_Client.Open strConnect
   
   Conn_Client.CursorLocation = adUseClient
        'Set Banco = OpenDatabase(Trim(sDataBase), "", "", strConnect)
'    End If
   
    F_AbreBanco = True
    sBanco = Conn.DefaultDatabase
    Exit Function
erro:
    


MsgBox Err.Description & Chr(13) & strConnect, vbCritical
    F_AbreBanco = False
End Function
Sub P_BuscaNomeBanco()
    Dim sNomeArq As String
    sNomeArq = "piki.ini"
    If Dir("c:\windows\system32\" & sNomeArq) = "" Then
        If Dir("c:\windows\system\" & sNomeArq) = "" Then
            If Dir(App.path & "\" & sNomeArq) = "" Then
                strNomeBanco = "requinte"
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
        Open "c:\windows\system32\piki.ini" For Input As #1
        
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
    If App.PrevInstance = True Then
        MsgBox "Este Programa já Está sendo executado!", vbExclamation
        End
    End If
    strDinheiro = "01"
    strCartaoCred = "05"
    strCartaoDeb = "06"
    
    buscarAtlz = True
    P_BuscaIni
    
    
    If Not F_AbreBanco() Then
        End
    End If
    
    
        
    frmLogin.Show 1
    If strNome_Usu = "" Then
        P_Fim
    End If
    Unload frmLogin
    
    IniciaVariaveis
        
     frmEscolherVisao.Show 1
    
    
        mdiFrm.Show 0
    
    'frmMenuVendas.Show 0, mdiFrm
    'mdiFrm.Visible = False
    Exit Sub
erro:

    MsgBox Err.Description
    P_Fim
 
End Sub
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
                    sserver = Trim(Right(s, Len(s) - 7))
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


Sub P_Fim()
    Dim i As Integer
    On Error Resume Next
    For i = 1 To Forms.Count - 1
        Unload Forms(i)
        
    Next
    Conn.Close
    End
End Sub

    
    
    
    
    
    
    
    
Function F_AjustaTelefone(strExpr As String) As String
   '  Esta função ajusta uma expressão digitada
   'a formato de telefone
   F_AjustaTelefone = F_SoNumero(strExpr, True)
   strExpr = F_SoNumero(strExpr)
   If Left(strExpr, 1) = "0" Then
      'Retira o 0(zero) inicial
      strExpr = Right(strExpr, Len(strExpr) - 1)
   End If
   '0800-55-1668
   If Len(strExpr) = 7 Then
      
      F_AjustaTelefone = Left(strExpr, 3) & "-" & Right(strExpr, 4)
      
      F_AjustaTelefone = Right(Space(16) & F_AjustaTelefone, 16)
   
   ElseIf Len(strExpr) = 8 Then
      
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
   F_AjustaTelefone = Trim(F_AjustaTelefone)
End Function
Function F_SoNumero(strExpr As String, Optional blnCar As Boolean = False) As String
   Dim intCont As Integer
   Dim strValidos As String
   
   strExpr = UCase(strExpr)
   F_SoNumero = ""
   strValidos = IIf(blnCar, "0123456789()./X-,", "0123456789")
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
Sub AjustaTexto(ByRef objText As TextBox, Optional opcao As Integer = 0)
    P_AjustaTexto objText, opcao

End Sub
Sub P_AjustaTexto(ByRef objText As TextBox, Optional opcao As Integer = 0)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    'Torna o texto maiúsculo
    If opcao = 0 Then
        strAux = UCase(objText.text)
    ElseIf opcao = 1 Then
        strAux = StrConv(objText.text, vbProperCase)
    ElseIf opcao = 2 Then
        strAux = LCase(objText.text)
    Else
        strAux = (objText.text)
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
    objText.text = strAux
End Sub
Function F_AjustaTexto(objText As String, Optional caixa_alta As Boolean = True)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    'Torna o texto maiúsculo
    If caixa_alta Then
        strAux = UCase(objText)
    Else
        strAux = (objText)
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
        ElseIf strAux2 = "\" Then
           Mid(strAux, intCont, 1) = "/" 'trocando | por ¦
        End If
    Next
    F_AjustaTexto = strAux
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
       If Left(objctObjeto.text, Len(strExpr)) = strExpr Then
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
Function F_CGC(CGC As String) As Boolean
'Esta função permite validar o CGC
'Sintaxe.: F_CGC( CGC )
'Entrada.: CGC - Nº do CGC a validar
'Saída...: Verdadeiro se exister o CGC, senão Falso

Dim M_VetNum(13), M_VetCGC(13), M_Num As String
Dim D_12, D_13, x, y, M_D1, M_D2 As Integer
CGC = F_SoNumero(CGC)
If val(CGC) < 1 Then
   F_CGC = False
   Exit Function
End If
M_Num = "6543298765432"
CGC = F_ZerEsq(val(CGC), 14, 0)
For x = 1 To 13
        M_VetNum(x) = val(Mid(M_Num, x, 1))
Next
For x = 1 To 12
        M_VetCGC(x) = Mid(CGC, x, 1)
Next
D_12 = 0
For y = 1 To 12
        D_12 = D_12 + val(M_VetCGC(y)) * M_VetNum(y + 1)
Next
M_D1 = 11 - (D_12 - Int(D_12 / 11) * 11)
If M_D1 >= 10 Then
   M_D1 = 0
End If
M_VetCGC(13) = Trim(str(M_D1))
For y = 1 To 13
        D_13 = D_13 + val(M_VetCGC(y)) * M_VetNum(y)
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
    objCEP.text = F_SoNumero(objCEP.text)
    If Len(objCEP.text) <> 8 Then
       Beep
       MsgBox "Informe o número do CEP corretamente!", _
       vbExclamation, App.EXEName
    End If
    objCEP.text = Format(objCEP.text, "0#\.###\-###")
End Sub
Sub P_FormataCGC(objCGC As Object)
    objCGC.text = Format(objCGC.text, "0#\.###\.###\/####\-##")
End Sub
Sub P_FormataCPF(objCPF As Object)
    objCPF.text = Format(objCPF.text, "0##\.###\.###\-##")
End Sub
Function F_FormataCPF(objCPF As String)
    
    F_FormataCPF = Format(objCPF, "0##\.###\.###\-##")
    
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


Function F_FormataCGC(objCPF As String)
    
    F_FormataCGC = Format(objCPF, "0#\.###\.###\/####\-##")
    
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
    On Error Resume Next
    Dim iCount As Integer
    mdiFrm.lblMsg.caption = sMsg
    
    If Trim(sMsg) <> "" And bEspere Then
        For iCount = 0 To Forms.Count - 1
            Forms(iCount).MousePointer = 11
        Next
        frmMsg.lblMsg.caption = sMsg
        frmMsg.Show 0, mdiFrm
        frmMsg.SetFocus
        mdiFrm.lblMsg.Refresh
        frmMsg.lblMsg.Refresh
        'DoEvents
    Else
        For iCount = 0 To Forms.Count - 1
            Forms(iCount).MousePointer = 0
        Next
        Unload frmMsg
        'DoEvents
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





Function F_DadosCheque(bVen As Boolean, Optional ven As String = "", Optional nome As String = "", Optional Num As String = "", _
            Optional cco As String = "", Optional age As String = "", Optional bco As String = "", _
            Optional valor As String = "0,00", Optional sCPF As String = "", Optional cod_loc As String) As cheque
    
    
    With frmDadosChe
        .txtbxAgencia.text = age
        .txtbxContaCorrente.text = cco
        .txtbxNome_Che.text = nome
        .txtbxNum_Che.text = Num
        .txtbxValor_Che.text = valor
        .txtbxVencimento.text = ven
        .txtbxCPF.text = sCPF
        .txtbxLocal.tag = cod_loc
        .txtbxLocal.text = F_BuscaUmCampo("nome_loc", "tab_localche", "cod_loc='" & cod_loc & "'")
        If bco <> "" Then
            .txtbxBanco.tag = bco
            .txtbxBanco.text = bco
        End If
        If bVen = False Then
            .txtbxVencimento.Enabled = False
        End If
        On Error Resume Next
        
        .Show 1
        If .tag <> "" Then
            F_DadosCheque.age_che = .txtbxAgencia.text
            F_DadosCheque.cco_che = .txtbxContaCorrente.text
            F_DadosCheque.nome_che = .txtbxNome_Che.text
            F_DadosCheque.num_che = .txtbxNum_Che.text
            F_DadosCheque.Valor_che = .txtbxValor_Che.text
            F_DadosCheque.Venc_che = .txtbxVencimento.text
            F_DadosCheque.cod_bco = .txtbxBanco.tag
            F_DadosCheque.cpf_che = .txtbxCPF.text
            F_DadosCheque.cod_loc = .txtbxLocal.tag
            F_DadosCheque.nome_loc = .txtbxLocal.text
        End If
    End With
        Unload frmDadosChe
End Function
Function F_DadosChePag(Optional ven As String = "", Optional Num As String = "", _
            Optional cod_con As String = "", Optional sVal As String = "", Optional lLock As Boolean = False) As chePag
    
    With frmDadosChPag
        
            .txtbxVencimento.text = ven
            .txtbxNum_Che.text = Num
            .txtbxValor_Che.text = sVal
        
        If cod_con <> "" Then
            P_LocalizaOperacao cod_con, .cmbxContas
        End If
        
        If lLock = True Then
            .txtbxVencimento.Enabled = False
            .txtbxValor_Che.Enabled = False
            .cmbxContas.Enabled = False
        End If
        On Error Resume Next
        
        .Show 1
        
            F_DadosChePag.cod_con = Left(.cmbxContas.text, 3)
            F_DadosChePag.num_che = .txtbxNum_Che.text
            F_DadosChePag.Valor_che = CCur(.txtbxValor_Che.text)
            F_DadosChePag.Venc_che = .txtbxVencimento.text
        
    End With
        Unload frmDadosChPag
        
    
    
End Function

Function F_DadosPar(Optional ven As String = "", _
    Optional valor As String = "0,00", Optional boleto As String = "", Optional complemento = "") As parcela
    
    
    With frmDadosPar
        .txtbxValor_Che.text = valor
        .txtbxVencimento.text = ven
        .txtbxBoleto.text = boleto
        .txtcomplemento.text = complemento
        On Error Resume Next
        
        .Show 1
        If .tag <> "" Then
            
            F_DadosPar.valor = .txtbxValor_Che.text
            F_DadosPar.vencimento = .txtbxVencimento.text
            F_DadosPar.boleto = .txtbxBoleto.text
            F_DadosPar.complemento = .txtcomplemento.text
        End If
    End With
        Unload frmDadosPar
        
    
    
End Function


Sub MudaAparencia(frm As Form)
    Dim strSQL As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim rPesq As New adodb.Recordset
    Dim sObj As String
    On Error Resume Next
    strSQL = "select * from tab_aparencia where cod_usu = '" & strCod_Usu & "' "
    abreObjetoAdo rPesq, strSQL
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
Function ContraBarraPipe(text As String) As String
    Dim iCount As Long
    Dim s As String
    Dim strAux As String
    Dim intCont As Integer
    For intCont = 1 To Len(text)
        strAux = Mid(text, intCont, 1)
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
Sub MostraErro(Optional strProc As String = "Erro em Tempo de Execução")
    P_MostraErro strProc
End Sub

Sub P_MostraErro(Optional strProc As String = "Erro em Tempo de Execução")
  'Procedimento com a finalidade de comentar erro
  'Sintaxe.: P_MostraErro()
  'Se for executado algum erro desconhecido fornece a mensagem de erro desconhecido
  ''On Error  Resume Next
  Dim strErro As String
  'P_Status ""  'Colocando barra de Status em modo normal
  If Err = 0 Then
     'O procedimento foi passado sem que algum erro ocorresse
     Exit Sub
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
           strErro = strErro & Err.Description & Chr(13)
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

Function F_BuscaDatas(Optional dataI As String = "", Optional dataF As String = "") As DataIniFin
    frmDataInicialFinal.txtbxDataInicial.text = dataI
    frmDataInicialFinal.txtbxDataFinal.text = dataF
    frmDataInicialFinal.Show 1
    F_BuscaDatas.dataF = frmDataInicialFinal.txtbxDataFinal.text
    F_BuscaDatas.dataI = frmDataInicialFinal.txtbxDataInicial.text
    Unload frmDataInicialFinal
End Function
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
        frmLogin.txtbxNome_Usu.text = strNome_Usu
        If bLockUser = True Then
            frmLogin.txtbxNome_Usu.Locked = True
        Else
            frmLogin.txtbxNome_Usu.Locked = False
        End If
        frmLogin.BorderStyle = 1
        Sendkeys Chr(9)
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
    Dim rPesq As New adodb.Recordset
    strSQL = "select * from tab_produtos order by nome_pro "

    abreObjetoAdo rPesq, strSQL
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
    
    If F_IsCaixa = False Then
        MsgBox "Este computador não é um caixa", vbExclamation
        Exit Function
    End If
    frmTroco.tag = num_ven
    frmTroco.caption = "Finalizando venda: " & num_ven
    
    
    P_GravaVendaAberta num_ven
    
    frmTroco.Show 1 ', mdiFrm
    
    F_Troco = frmTroco.tag
    Unload frmTroco
    P_DeletaVendaAberta num_ven
    
End Function
Function F_UsuLogadoCaixa(sCod_Cai As String) As String
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    strSQL = "select tf.guerra_fun from tab_turncai as tt, "
    strSQL = strSQL & " tab_Funcionarios as tf where cod_cai='" & sCod_Cai & "' "
    strSQL = strSQL & " and cod_Tur='" & strCod_Tur & "' "
    strSQL = strSQL & " and data_tur=" & F_GravaData(Date)
    strSQL = strSQL & " and tf.cod_fun=tt.login_usu "
    abreObjetoAdo rPesq, strSQL
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
Function F_EscolheVendedor() As String
    'On Error Resume Next
    frmEscolheVendedor.Show 1
    F_EscolheVendedor = frmEscolheVendedor.tag
    Unload frmEscolheVendedor
End Function

Sub P_Graph(bl() As valChart, qtd As Integer, Optional tipo As Long = 1, _
Optional visLeg As Boolean = True, Optional Tit As String = "", _
    Optional Bottom As String = "", Optional toolVis As Boolean = False, _
    Optional rows As Integer = 1, Optional porcentagem As Boolean = True)
    
''0   No graph               12  Buble
''1   2D pie                 13  Tape
''2   3D pie                 14  3D area
''3   2D bar                 15  Log/log
''4   3D bar (default)       16  Lin/log
''5   Gantt                  17  Box-whisker
''6   Line                   18  Open-high-low-close
''7   Log/lin                19  Candlestick
''8   2D area                20  True3D surface
''9   2D scatter             21  True3D scatter
''10  Polar                  22  Time series
''11  High-low-close
'    Dim iRow As Integer
'    Dim iCount As Integer
'    With frmGraph.grph
'        '.NumSets = rows
'        .GraphTitle = Tit
'        .BottomTitle = Bottom
'        .Labels = 1
'        .NumPoints = qtd
'
'        If porcentagem Then
'            .GraphStyle = 4
'        Else
'            .GraphStyle = 0
'        End If
'        If toolVis = True Then
'            .Toolbar = graphRunTime
'        End If
'        For iRow = 1 To rows
'            For iCount = 1 To qtd
'                .ThisSet = iRow
'                .Data(iCount) = bl(iCount - 1).Value
'                If visLeg = True Then
'
'                    .Legend(iCount) = bl(iCount - 1).caption
'                    If .Legend(iCount) = "" Then
'                        .Legend(iCount) = "---"
'                    End If
'                    .LegendStyle = 0
'                End If
'
'            Next
'        Next
'    .GraphType = tipo
'    .DrawMode = 3
'    End With
'    frmGraph.Show 0
End Sub




Sub P_mensagem(msg As String, frm As Form, Optional acao As Boolean = True)
    On Error Resume Next
    If acao = True Then
        frmMsg.lblMsg.caption = msg
        frmMsg.Show 0, frm
        
    Else
        Unload frmMsg
    End If
End Sub
Function F_BuscaDesconto(cod_usu As String, cod_sub As String, cod_For As String) As String
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("desconto_usu", "tab_descontos", "cod_usu='" & cod_usu & _
    "' and cod_sub='" & cod_sub & "' and cod_For='" & cod_For & "'")
    If strSQL = "" Then
        strSQL = "0"
    End If
    F_BuscaDesconto = strSQL
End Function

Function F_senhaReturnUsu(Optional LockUsu As Boolean = False, _
    Optional bNovoUsu As Boolean = True, Optional NomeUsu As String, Optional returnUsu As Boolean = True) As String
    Dim sSenha As String
    Dim strSQL As String
    
        frmLogin.tag = "senha"
        On Error Resume Next
        If NomeUsu = "" Then
            frmLogin.txtbxNome_Usu.text = strNome_Usu
        Else
            frmLogin.txtbxNome_Usu.text = NomeUsu
        End If
        frmLogin.txtbxNome_Usu.Locked = LockUsu
        frmLogin.BorderStyle = 1
        Sendkeys Chr(9)
                    frmLogin.Show 1
        
               
        F_senhaReturnUsu = frmLogin.tag
        If F_senhaReturnUsu <> "" Then
            If returnUsu Then
                strCod_Usu = F_senhaReturnUsu
                strNome_Usu = F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & strCod_Usu & "'")
                P_BuscaDefUsu
            End If
        End If
        
        Unload frmLogin
End Function
Function F_SenhaUsuario(Optional LockUsu As Boolean = False, Optional NovoUsuAssume As Boolean = True) As usuario
    Dim sSenha As String
    Dim strSQL As String
    'NovoUsuAssume -> de agora pra frente, ele é o usuário do sistema
        frmLogin.tag = "senha"
        On Error Resume Next
        
        frmLogin.txtbxNome_Usu.text = strNome_Usu
        
        
        frmLogin.txtbxNome_Usu.Locked = LockUsu
        frmLogin.BorderStyle = 1
        Sendkeys Chr(9)
        frmLogin.Show 1
        
               
        
        If frmLogin.tag <> "" Then
            If NovoUsuAssume Then
                strCod_Usu = frmLogin.tag
                strNome_Usu = UCase(frmLogin.txtbxNome_Usu.text)
            End If
            F_SenhaUsuario.cod_usu = frmLogin.tag
            F_SenhaUsuario.nome_usu = UCase(frmLogin.txtbxNome_Usu.text)
        End If
        
        Unload frmLogin
End Function


Sub P_InfoEntradasPro(cod_pro As String)
    Dim fr As New frmInfoEnt
    fr.tag = cod_pro
    fr.Show 1
End Sub
Sub P_InfoSaidasPro(cod_pro As String)
    Dim fr As New frmInfSai
    fr.tag = cod_pro
    fr.Show 1
End Sub

Sub P_Msg(Optional msg As String, Optional titulo As String)
    On Error GoTo trata_Erro
    If msg <> "" Then
        frmMsg.caption = titulo
        frmMsg.lblMsg.caption = msg
        DoEvents
        frmMsg.Show 0
        
        frmMsg.Refresh
        
        
    Else
        Unload frmMsg
    End If
    Exit Sub
trata_Erro:
'P_MostraErro
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

Sub P_EspelhoNota(seq_not As String)
    'Dim fr As New frmEspelhoNota
    'fr.tag = seq_not
    'fr.Show 0
End Sub
Sub P_cancelaNotaFiscalSaida(nota_sai As String, serie_sai As String)
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    

    strSQL = "select * from tab_itsai where nota_sai='" & nota_sai & "' "
    strSQL = strSQL & " and serie_Sai='" & serie_sai & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_produtos set ef_pro=ef_pro+" & F_GravaDec(!qtd_its)
            strSQL = strSQL & " where cod_pro ='" & !cod_pro & "' "
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update tab_saidas set serie_sai='CANC' "
    strSQL = strSQL & " where nota_sai='" & nota_sai & "' "
    strSQL = strSQL & " and serie_sai='" & serie_sai & "' "
    Conn.Execute strSQL
    MsgBox "Nota Fiscal Cancelada!", vbInformation
End Sub
Sub P_cancelaCupon(seq_cup As String)
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    

    strSQL = "select * from tab_itcup where seq_cup='" & seq_cup & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_produtos set ef_pro=ef_pro+" & F_GravaDec(!qtd_its)
            strSQL = strSQL & " where cod_pro ='" & !cod_pro & "' "
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update tab_cupons set cancelado_cup='S' "
    strSQL = strSQL & " where seq_cup='" & seq_cup & "' "
    
    Conn.Execute strSQL
    MsgBox "Nota Fiscal Cancelada!", vbInformation
End Sub


Sub P_AjustaVenda()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim sMsg As String
    strSQL = "select * from tab_vendas where situacao_ven in('2','3') "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("num_ven", "tab_venfec", "num_ven='" & !num_ven & "'") = "" Then
                strSQL = "Update tab_vendas set situacao_ven='1' "
                strSQL = strSQL & " where num_ven='" & !num_ven & "' "
                Conn.Execute strSQL
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
Function F_DevolveChe(seq_che As String) As Boolean
    frmDevChe.tag = seq_che
    frmDevChe.Show 1
    If frmDevChe.tag <> "" Then
        F_DevolveChe = True
    Else
        F_DevolveChe = False
    End If
    
    
End Function
Sub P_HistCheCPF(cpf_cli As String)
    frmHistoricoChe.tag = cpf_cli
    frmHistoricoChe.txtbxCPF.text = cpf_cli
    frmHistoricoChe.Show 1
    
    
    
End Sub
Function F_HistCheCPF(cpf As String) As cheque

    Dim stSQL As String
    Dim rPesq As New adodb.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "select distinct cco_che,cod_bco,age_che,cpf_che,nome_che from tab_cheques where cpf_che='" & cpf & "' "
    strSQL = strSQL & " order by seq_che "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 1 Then
            F_HistCheCPF.age_che = !age_che
            F_HistCheCPF.cco_che = !cco_che
            F_HistCheCPF.cod_bco = !cod_bco
            
            F_HistCheCPF.cpf_che = !cpf_che
            F_HistCheCPF.nome_che = !nome_che
            
        End If
        If .RecordCount > 1 Then
            Do While Not .EOF
                strSQL = !cod_bco & Chr(9)
                strSQL = strSQL & F_BuscaUmCampo("nome_bco", "tab_bancos", "cod_bco='" & !cod_bco & "'") & Chr(9)
                strSQL = strSQL & !age_che & Chr(9)
                strSQL = strSQL & !cco_che & Chr(9)
                strSQL = strSQL & !nome_che
                frmBuscaCheCPF.msflxgrd.AddItem strSQL
                
                .MoveNext
            Loop
            On Error Resume Next
            frmBuscaCheCPF.Show 1
            frmBuscaCheCPF.msflxgrd.col = 0
            
            F_HistCheCPF.cod_bco = frmBuscaCheCPF.msflxgrd.text
            frmBuscaCheCPF.msflxgrd.col = 2
            F_HistCheCPF.age_che = frmBuscaCheCPF.msflxgrd.text
            frmBuscaCheCPF.msflxgrd.col = 3
            F_HistCheCPF.cco_che = frmBuscaCheCPF.msflxgrd.text
            
            frmBuscaCheCPF.msflxgrd.col = 4
            F_HistCheCPF.nome_che = frmBuscaCheCPF.msflxgrd.text
            Unload frmBuscaCheCPF
            
        End If
    End With
    
Exit Function
trata_Erro:
P_MostraErro

End Function
Function F_EsticaTexto(Texto As String, Tam As Integer, Optional align As Integer = 0) As String
    If Len(Texto) >= Tam Then
        F_EsticaTexto = Left(Texto, Tam)
    Else
        If align = 0 Then
            F_EsticaTexto = Texto & String(Tam - Len(Texto), " ")
        ElseIf align = 1 Then
            F_EsticaTexto = String((Tam - Len(Texto)) / 2, " ") & Texto & String((Tam - Len(Texto)) / 2, " ")
        Else
            F_EsticaTexto = String(Tam - Len(Texto), " ") & Texto
        End If
    End If
End Function
Function EsticaTexto(Texto As String, Tam As Integer, Optional align As Integer = 0, _
Optional Carac As String = " ", Optional soNum As Boolean = False) As String
    If soNum Then
        Texto = F_SoNumero(Texto)
    End If
    
    
    If Len(Texto) >= Tam Then
        EsticaTexto = Left(Texto, Tam)
    Else
        If align = 0 Then
            EsticaTexto = Texto & String(Tam - Len(Texto), Carac)
        ElseIf align = 1 Then
            EsticaTexto = String((Tam - Len(Texto)) / 2, Carac) & Texto & String((Tam - Len(Texto)) / 2, Carac)
        Else
            EsticaTexto = String(Tam - Len(Texto), Carac) & Texto
        End If
    End If
End Function

Sub P_AltPreco(cod_pro As String, novo_preco As Currency, antigo As Currency)
    Dim strSQL As String
    On Error GoTo trata_Erro
    If novo_preco = antigo Then
        Exit Sub
    End If
    strSQL = "insert into tab_altpreco(cod_pro,antigo_pro,novo_pro,data_alt,cod_usu)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(antigo) & ","
    strSQL = strSQL & F_GravaDec(novo_preco) & ","
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    Conn.Execute strSQL
    Exit Sub
trata_Erro:
End Sub

Function VerificaVendaAbertaNoCaixa(cod_cai As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    'On Error Resume Next
    strSQL = "select tf.* from  tab_vendas as tv, "
    strSQL = strSQL & " tab_venfec as tf "
    strSQL = strSQL & " where tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & cod_cai & "' "
    strSQL = strSQL & " and situacao_ven=1 "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            VerificaVendaAbertaNoCaixa = True
        End If
        .Close
    End With
End Function

 Sub VerificaAtualizacao()
    Dim strSQL As String
    On Error GoTo trata_Erro
    Dim sDest As String
    Dim sOrig As String
    
    If VerificaDataBanco = False Then
        P_Fim
        
        
        
        
    End If
    If isDesignTime Then
       ' MsgBox "Design time", vbInformation
        Exit Sub
    End If
    
    'Exit Sub
    If buscarAtlz = False Then
        Exit Sub
    End If
    sDest = App.path & "\" & App.EXEName & ".exe"
    sOrig = strArquivoServidor & App.EXEName & ".exe"
    
    If FileDateTime(sOrig) > FileDateTime(sDest) Then
        strSQL = "Existe uma versão mais nova para atualização."
        strSQL = strSQL & "Deseja atualizar agora?"
        If MsgBox(strSQL, vbYesNo + vbQuestion) = vbYes Then
            Open App.path & "\atlz_pk.txt" For Output As #1
            Print #1, sOrig
            Print #1, sDest
        End If
        Close #1
        
        shell App.path & "\atlz_pk.exe", vbNormalFocus
        End
        
    End If

    Exit Sub
trata_Erro:
   ' If MsgBox("O arquivo para atualização não foi localizado. Deseja cancelar a busca por atualização?", vbYesNo + vbQuestion) Then
   '     buscarAtlz = False
   ' End If
    
    
    
    
End Sub
Sub VerificaAtlzPic()
    If Dir(App.path & "\pic", vbDirectory) = "" Then
        fileSystem.MkDir App.path & "\pic"
    End If
    
    
    
    
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
    Dim i As Integer
    Dim strNota_Sai As String
    strNota_Sai = "      "
    i = ecf_NumeroCupom(strNota_Sai)
    
    If MsgBox("Confirma cancelamento do Último Cupom Fiscal? Nº" & strNota_Sai, vbYesNo + vbCritical) = vbYes Then
        
        
        ecf_CancelaCupom
                
        sNumVen = F_BuscaUmCampo("num_ven", "tab_saidas", "nota_sai='" & strNota_Sai & "' and cod_cai='" & strCod_Cai & "' and serie_sai='CUPM'")
        
        'strSQL = "update tab_saidas set serie_sai='CANC' where num_ven='" & sNumVen & "' "
        'strSQL = strSQL & " and nota_Sai='" & F_SoNumero(strNota_Sai) & "'"
        
    
        'conn.execute strSQL
        If sNumVen = "" Then
            MsgBox "Não foi encontrada a venda do cupom fiscal nº " & strNota_Sai, vbExclamation
            Exit Sub
        End If
        
        If sNumVen <> "" Then
            If MsgBox("Confirma Cancelamento da venda " & sNumVen, vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
        End If
                
                
                
                      
                
        'If Not F_CancelaVendaAposCupom(sNumVen) Then
        '    MsgBox "Falha ao cancelar a venda!", vbCritical
        'Else
        '    MsgBox "Venda Cancelada! Estoque Reposto! Lançamento em Caixa Concluído!", vbInformation
        'End If
    
    
        
    
    
    
    End If
End Sub
Function F_Quebra(text As String, qtd_por_linha As Integer) As String
    On Error Resume Next
    Dim iCount As Integer
    Dim strAux As String
    Dim iAux As Integer
    Dim jCount As Integer
    Dim strBuf As String
    Dim strRest As String
    For iCount = 1 To Len(text)
        iAux = iAux + 1
        strAux = strAux + Mid(text, iCount, 1)
        If iAux >= qtd_por_linha Then
            If InStr(1, " -;.!_+|=:,<>", Mid(text, iCount, 1)) Then
                
                strAux = strAux + Chr(13) + Chr(10)
                strBuf = strAux
                iAux = 0
                strRest = ""
                
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

Sub P_PreenchePermGru(cmbx As Object)
    P_Preenche cmbx, "tab_permgru", "nome_gru"
End Sub
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


Sub P_PreencheTipoPar(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    strSQL = "select * from tab_tipopar order by cod_tip"
    abreObjetoAdo rPesq, strSQL
    
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem Int(rPesq!cod_tip) & " - " & rPesq!nome_tip
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub VerificaAtualizacao1()
    Dim strSQL As String
    On Error GoTo trata_Erro
    Dim sDest As String
    Dim sOrig As String
    'Exit Sub
    sDest = App.path & "\" & App.EXEName & ".exe"
    sOrig = strArquivoServidor & App.EXEName & ".exe"
    
    If FileDateTime(sOrig) > FileDateTime(sDest) Then
        strSQL = "Existe uma versão mais nova para atualização."
        strSQL = strSQL & "Assim que possível saia do  " & constNome_sis & " e execute o programa de instalação."
        MsgBox strSQL, vbInformation
            Open App.path & "\atlz_pk.txt" For Output As #1
            Print #1, sOrig
            Print #1, sDest
        
        Close #1
        
        shell App.path & "\atlz_pk.exe", vbNormalFocus
        
        
    End If
    Exit Sub
trata_Erro:
    
    
End Sub

Sub P_ImprQuebra1(text As String, obj As Object, x1 As Single, X2 As Single)
    Dim c As New Collection
    Dim iCount As Integer
    Dim sH As Single
    obj.ScaleMode = vbMillimeters
    Set c = F_SeparaTexto(text)
    obj.CurrentX = x1
    For iCount = 1 To c.Count
        sH = obj.TextWidth(c(iCount))
        If (sH + obj.CurrentX) > X2 Then
            obj.Print
            obj.CurrentX = x1
        End If
        P_ImprX obj, c(iCount), obj.CurrentX
    Next
End Sub
Sub P_ImprQuebraTexto(text As String, obj As Object, x1 As Single, X2 As Single)
    Dim c As Double
    Dim iCount As Integer
    Dim sH As Single
    Dim sBuf As String
    obj.ScaleMode = vbMillimeters
    
    c = Len(text)
    obj.CurrentX = x1
    For iCount = 1 To c
        sBuf = Mid(text, iCount, 1)
        sH = obj.TextWidth(sBuf)
        If (sH + obj.CurrentX) > X2 Then
            obj.Print
            obj.CurrentX = x1
        End If
        P_ImprX obj, sBuf, obj.CurrentX
    Next
End Sub

Function CollectionToNotIn(c As Collection) As String
    Dim i As Integer
    Dim s As String
    For i = 1 To c.Count
        s = s & c(i)
        If i <> c.Count Then
            s = s & ","
        End If
        
    Next
    CollectionToNotIn = s
End Function


Function F_SeparaTexto(Texto As String, Optional sSepara As String = "-;| /\_.:", _
    Optional bArmazenaSeparador As Boolean = True, Optional formato As String = "", _
    Optional aspas As String = "") As Collection
    Dim cTexto As New Collection
    Dim iCount As Integer
    Set cTexto = Nothing
    Set cTexto = New Collection
    Dim sPalavra As String
    Dim sBuf As String
    For iCount = 1 To Len(Texto)
        sBuf = Mid(Texto, iCount, 1)
        If sBuf = "-" Then
'            MsgBox ""
        End If
        If InStr(sSepara, sBuf) > 0 Then
            
            If formato <> "" Then
                 sPalavra = Format(sPalavra, formato)
            End If
            If aspas <> "" Then
                sPalavra = aspas & sPalavra & aspas
            End If
            If bArmazenaSeparador Then
                If Trim(sPalavra) <> "" Then
                    
                
                 
                        cTexto.Add sPalavra & sBuf
                
                ElseIf (sBuf) <> "" Then
                       cTexto.Add sPalavra & sBuf
                End If
                sPalavra = ""
            Else
                If sPalavra <> "" Then
                    cTexto.Add sPalavra
                End If
                sPalavra = ""
            End If
        Else
            
            sPalavra = sPalavra & sBuf
            
        End If
    Next
    If sPalavra <> "" Then
    
        If formato <> "" Then
                 sPalavra = Format(sPalavra, formato)
            End If
            If aspas <> "" Then
                sPalavra = aspas & sPalavra & aspas
            End If
    
        cTexto.Add sPalavra
    End If
    Set F_SeparaTexto = cTexto
End Function

Sub P_HistProd(cod_pro As String)
    frmHistProd.tag = cod_pro
    frmHistProd.txtProduto.text = F_BuscaUmCampo("nome_pro", TAB_PRODUTOS, "cod_pro='" & cod_pro & "'")
    frmHistProd.txtProduto.tag = cod_pro
    frmHistProd.Show 1

End Sub
Sub P_histCli(cod_cli As String)
    frmParcelasTer.txtbxNomeCli.tag = cod_cli
    
    frmParcelasTer.tag = cod_cli
    
    frmParcelasTer.Show 0
End Sub


Sub P_shell(s As String)
    Dim f As String
    
    
    
    
        
    
    If s = "vendas" Then
        frmEdiVendas.Show 0, mdiFrm
    ElseIf s = "posvenda" Then
        frmRelPosVendaFaixa.Show 1
    ElseIf s = "mesclarorgao" Or s = "mesclaorgao" Or s = "mesclarorgaos" Then
        frmMesclarOrgao.Show 1
        
    ElseIf s = "bxestoque" Then
        frmBaixarEstoque.Show 1
    ElseIf s = "nova_migra" Then
        
        frmNovaMigracao.Show 1
    ElseIf s = "deletarva" Then
        frmDeletarVendas.Show 1
    ElseIf s = "ajuste" Then
        frmAjuste.Show 1
        
    ElseIf s = "migra" Then
        frmMigraDados.Show 1
    ElseIf s = "migracons" Then
        'frmMigraConsignado.Show 1
    ElseIf s = "forcaabertos" Or s = "forçaabertos" Then
        frmForcaAbertos.Show 1
    ElseIf s = "fechacx" Then
        frmMovDiarioCx.Show 1
    ElseIf s = "ecf" Then
        frmECF.Show 0, mdiFrm
    ElseIf s = "ajuste" Then
    ElseIf s = "vendasnaocaixa" Then
        'frmVendasNaoCaixa.Show 1
    ElseIf s = "resumodoc" Then
        'frmRelDocRecResumo.Show 0
    ElseIf s = "queima" Then
        frmQueima.Show 1
    ElseIf s = "pacote" Then
        frmPacote.Show 0, mdiFrm
    ElseIf s = "crediario" Or s = "crediário" Then
        frmCrediario.Show 0
    ElseIf s = "permissoes" Or s = "permissões" Or s = "permissão" Or s = "permissao" Then
        frmEdiPerm.Show 1
    ElseIf s = "migra_ana" Or s = "migraana" Or s = "migralucas" Or s = "migra_lucas" Then
        
        migra_lucas
        
    ElseIf s = "bloqueartodos" Then
        s = InputBox("Informe a data")
        If CDate(s) > Date Then
            MsgBox "Data inválida"
            Exit Sub
        End If
        If s = "" Then
            Exit Sub
        End If
        BloquearTodos s
    ElseIf Left(s, 14) = "migra_produtos" Then
        Dim cod As String
        Dim iCod As Integer
        iCod = InStr(1, s, "=")
        If iCod > 0 Then
            cod = Mid(s, iCod + 1)
        End If
        GeraArquivoMigracaoProduto cod
    ElseIf s = "importa_produtos" Then
        importarProdutos "", True
    ElseIf s = "enviar_produtos" Then
        APIEnviarProdutos
    ElseIf s = "atualizar_produtos" Then
        s = InputBox("Informe a lista de produtos ou deixe em branco para todos")
        ApiAtualizarProdutos s
        
    ElseIf s = "etiqueta" Then
    
        frmGeraEtiquetas.Show 0
    ElseIf s = "xmlutil" Or s = "json" Or s = "xml" Then
        frmXmlUtil.Show 1
    ElseIf s = "migrajson" Then
        frmMigraJSON.Show 1
    ElseIf s = "migrar" Or s = "migrarbd" And strNome_Usu = "ALISON" Then
        frmMigracaoBDUnico.Show 1
    ElseIf s = "parametros" Then
        
        frmParametros.Show 1
    ElseIf s = "parametro_estrutura" Or s = "parametro_regra" Then

        frmParametrosRegras.Show 1
    ElseIf Left(s, 8) = "fila_doc" Then
        fila_doc s
    ElseIf Left(s, 9) = "fila_prod" Then
        fila_prod s
        
    
    ElseIf Left(s, 16) = "sincronizar_docs" Then
        
        Dim inicio As String
        Dim fim As String
        If InStr(1, s, "inicio:") > 0 Then
            inicio = Mid(s, InStr(1, s, "inicio:") + 7, 10)
        End If
        If InStr(1, s, "fim:") > 0 Then
            fim = Mid(s, InStr(1, s, "fim:") + 4, 10)
        End If
        P_Status "sincronizando docs"
        sincronizar_docs inicio, fim
        P_Status ""
    ElseIf s = "fila" Then
        frmFila.Show 0, mdiFrm
    ElseIf Left(s, 15) = "sincronizar_cli" Then
        sincronizar_clientes Mid(s, 21)
    Else
        MsgBox "Ainda não aprendi " & s, vbExclamation
    End If
    
    
End Sub
Private Sub fila_doc(msg As String)
    Dim data_ini As String
    Dim data_fim As String
    Dim reman As String
    Dim cliente As String
    
    data_ini = Mid(msg, InStr(1, msg, "inicio") + 7, 10)
    data_fim = Mid(msg, InStr(1, msg, "fim") + 4, 10)
    reman = Mid(msg, InStr(1, msg, "reman") + 6, 1)
    cliente = Mid(msg, InStr(1, msg, "cliente") + 7, 1)
    
    gerar_fila_doc data_ini, data_fim, IIf(reman <> "" And reman <> "0", True, False)
    frmFila.Show 0, mdiFrm
    
End Sub

Private Sub fila_prod(msg As String)
    Dim data_ini As String
    Dim data_fim As String
    Dim reman As String
    Dim cliente As String
    
    
    
    gerar_fila_prod
    'executa_fila "3"
    frmFila.Show 0, mdiFrm
    
End Sub


Sub migra_lucas()
    Dim strSQL As String
    P_Status "Migrando..."
    strSQL = "replace into requinteana.tab_produtos select * from requinte.tab_produtos"
    Conn.Execute strSQL
    MsgBox "Migração concluída em Lucas!"
    
    P_Status ""
    
End Sub
Function F_CalcIdade(datanasc) As String
    On Error Resume Next
    Dim iMeses As Integer
    
    iMeses = DateDiff("M", datanasc, Date)
    F_CalcIdade = Int(iMeses / 12)
End Function

Function F_TeclasValidas() As String
    F_TeclasValidas = "abcdefg"
End Function
Function F_IsCaixa() As Boolean
    Dim strSQL As String
    Dim strHost As String
    On Error GoTo trata_Erro
    
    strSQL = F_BuscaUmCampo("COD_CAI", "TAB_CAIXASIP", "IP_CAI='" & _
        strHost_Usu & "'")
    If strSQL = "" Then
        'strSQL = F_BuscaUmCampo("COD_CAI", "TAB_CAIXASIP", "IP_CAI='" & _
        'strHost_Usu & "'")
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


Sub P_ImprQuebraTurbo1(text As String, obj As Object, x1 As Single, X2 As Single, _
    Optional bAumenta As Boolean = False, Optional align As Integer = 1)
    On Error Resume Next
    Dim c As New Collection
    Dim sLinha As New Collection
    Dim iCount As Integer
    Dim jCount As Integer
    Dim sH As Single
    Dim sAcumula As String
    Dim pc As PictureBox
    Dim cDif As Single
    Dim iCXAux As Single 'simula o currentx
    Dim cTamPalavra As Single
    'Dim l As linha
    'obj.ScaleMode = vbMillimeters
    If TypeName(obj) <> "Printer" Then
        Set pc = obj
    Else
'        Set pc = Printer
    End If
    Set c = F_SeparaTexto(text)
    iCXAux = x1
    Set sLinha = Nothing
    For iCount = 1 To c.Count
        If obj.CurrentY >= obj.ScaleHeight - 2 Then
            obj.height = obj.height * 1.2
        End If
        sH = obj.TextWidth(c(iCount))
        sH = obj.ScaleX(sH, obj.ScaleMode, vbMillimeters)
        If (sH + iCXAux) > X2 Then
            
            iCXAux = x1
            sLinha.Add sAcumula
            sAcumula = ""
        End If
        If c(iCount) <> Chr(13) Then
            cTamPalavra = obj.TextWidth(c(iCount))
            cTamPalavra = obj.ScaleX(cTamPalavra, obj.ScaleMode, vbMillimeters)
            
            iCXAux = iCXAux + cTamPalavra
            sAcumula = sAcumula + c(iCount)
        Else
            'P_ImprX pc, "", 0, False
        End If
        
        'pc = Nothing
        
        
    Next
    If sAcumula <> "" Then
        sLinha.Add sAcumula
    End If
    obj.CurrentX = x1
    If align = 1 Then ' se for esquerda imprime direto
        For iCount = 1 To sLinha.Count
            P_ImprX obj, sLinha(iCount), CInt(x1), False
        Next
    ElseIf align = 2 Then
        For iCount = 1 To sLinha.Count
            P_CentraImpressao obj, sLinha(iCount)
        Next
    ElseIf align = 3 Then
        For iCount = 1 To sLinha.Count
            P_DireitaImpressao obj, sLinha(iCount), X2 + 4
        Next
    Else 'justificado
        
        For iCount = 1 To sLinha.Count
            
                Set c = Nothing
                Set c = F_SeparaTexto(sLinha(iCount))
                obj.CurrentX = x1
                cDif = obj.TextWidth(sLinha(iCount))
                cDif = obj.ScaleX(cDif, obj.ScaleMode, vbMillimeters)
                cDif = (X2 - 3) - cDif
                For jCount = 1 To c.Count
                    If jCount = c.Count Then
                        If iCount <> sLinha.Count Then
                            P_DireitaImpressao obj, c(jCount), X2 + 4
                        Else
                            P_ImprX obj, c(jCount), obj.CurrentX
                            obj.Print ""
                        End If
                    Else
                    
                        P_ImprX obj, c(jCount), obj.CurrentX
                    End If
                    If iCount <> sLinha.Count Then  'ultima linha não justifica
                        obj.CurrentX = obj.CurrentX + (cDif / (c.Count - 1))
                    
                    End If
                Next
                'obj.Print ""
        Next
    End If
End Sub


Sub P_BuscaDefUsu()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    On Error GoTo trata_Erro
    strSQL = "select tu.nivel_usu,td.* from tab_defusu as td inner join tab_usuarios as tu on tu.cod_fun=td.cod_usu where cod_usu='" & strCod_Usu & "'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        
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
        strNivel_Usu = !nivel_usu
        
        
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
        strDesbloquearVenda_usu = !desbloquearvenda_usu
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
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_AjustaHora(obj As TextBox)
    Dim strHora As String
    Dim sAux As String
    Dim iCount As Integer
    Dim sLef As String
    Dim sRig As String
    
    On Error GoTo trata_Erro
    If Len(obj.text) = 5 Then
        obj.text = Left(obj.text, 2) & ":" & Right(obj.text, 2)
    End If
    If Len(obj.text) = 4 Then
        obj.text = Left(obj.text, 2) & ":" & Right(obj.text, 2)
    End If
    If Len(obj.text) = 3 Then
        obj.text = Left(obj.text, 1) & ":" & Right(obj.text, 2)
    End If
    If Len(obj.text) = 2 Then
        obj.text = Left(obj.text, 2) & ":00"
    End If
    If Len(obj.text) = 1 Then
        obj.text = "0" & obj.text & ":00"
    End If
    obj.text = Format(Trim(obj.text), "hh:mm")
    strHora = obj.text
    For iCount = 1 To Len(strHora)
        sAux = Mid(strHora, iCount, 1)
        If InStr(1, ".;-:,/\| ", sAux) > 0 Then
            sRig = Format(F_SoNumero(Right(strHora, Len(strHora) - iCount)), "0#")
            sLef = Format(F_SoNumero(sLef), "0#")
            Exit For
        Else
            sLef = sLef & sAux
        End If
        
    Next
    If sLef = "" Then
        sLef = "00"
    End If
    If sRig = "" Then
        sRig = "00"
    End If
    If Int(sLef) > 23 Then
        MsgBox "Hora Inválida!", vbCritical
        sLef = "00"
    End If
    If Int(sRig) > 59 Then
        MsgBox "Minuto Inválida!", vbCritical
        sRig = "00"
    End If
    obj.text = sLef & ":" & sRig
    Exit Sub
trata_Erro:
End Sub




Sub P_Preenche(obj As Object, tabela As String, campo As String, Optional Where As String = "", Optional sOrderBy As String = "", Optional emBranco As Boolean = False)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
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

Sub P_PreencheEspCai(cmbx As Object)
    Dim sWh As String
    
    
    P_Preenche cmbx, "tab_espcai", "concat(sigla_cai,'-',descricao_cai)", "", "", True
    
End Sub
Sub P_PreencheFormasPag(cmbx As Object, Optional tipo_cai As String = "DUP")
    Dim sWh As String
    
    
    P_Preenche cmbx, "tab_formaspag", "nome_for", "tipo_cai='" & tipo_cai & "'"
    
End Sub

Sub FechaObjeto(obj As Object)
    On Error Resume Next
    obj.Close
End Sub
Function abreObjetoAdo(obj As Object, strSQL As String, Optional dataVinculado As Object = Null) As Boolean
    On Error Resume Next
    FechaObjeto obj
    
    Conn_Client.Execute "set @rank=0"
    On Error GoTo trata_Erro
    obj.Open strSQL, Conn_Client
    
    
    
    
    Set dataVinculado.DataSource = obj
    abreObjetoAdo = True
    Exit Function
trata_Erro:

End Function

Function AjustaPlaca(placa As String) As String
    'AjustaPlaca = format(placa,"
    Dim letra As String
    Dim numero As String
    letra = F_SoLetra(placa)
    numero = F_SoNumero(placa)
    AjustaPlaca = letra & "-" & numero
End Function
Function BuscaListaDoc() As String
    frmListaDocumentos.Show 1
    BuscaListaDoc = frmListaDocumentos.lista
    
    
End Function
Function mudarCaractere(expressao As String, encontrar As String, trocar As String) As String
    mudarCaractere = Replace(expressao, encontrar, trocar)
End Function

Sub IniciaVariaveis()
    
    Dim rPesq As New adodb.Recordset
    Dim strSQL As String
    iniciaTabelas
    IniciaConfiguracoesNFe
     
    CrLf = Chr(13) = Chr(10)
    
    
        
        caminho_contratos = UnAjustaCaminho(BuscaParametroEstrutura("caminho_contrato").valor)
    
    If strNome_Usu = "ALISON" Then
      '  frmEscolheBanco.Show 1
    End If
    If bUsaPrecoAtacado = True Then
        sCampoPrecoConsign = "pa_pro"
    Else
        sCampoPrecoConsign = "pv_pro"
    End If
    
    If Left(Conn.DefaultDatabase, 8) = "requinte" Then
        modular.bUsaCaixaSeparado = True
        modular.bUsaCobranca = True
        modular.bUsaConsignado = True
        modular.bUsaCrediario = True
        modular.bUsaPacote = True
        strNomedaPromocao = "Queima"
    Else
        modular.bUsaCaixaSeparado = False
        modular.bUsaCobranca = True
        modular.bUsaConsignado = False
        
        modular.bUsaCrediario = False
        modular.bUsaPacote = False
        strNomedaPromocao = "Promoção"
    End If
    
    If Left(Conn.DefaultDatabase, 7) = "estoque" Then
        bUsaEstoqueSeparado = True
    Else
        bUsaEstoqueSeparado = False
    End If
    
    P_Status "Configurando Parâmetros do Sistema..."
    strSQL = "select * from tab_parametros "
    'strImprFiscal = "e" 'elgin
    abreObjetoAdo rPesq, strSQL
    
    If rPesq.RecordCount = 0 Then
        
        
    Else
        With rPesq
           ' strCod_Emp = F_BuscaUmCampo("cod_emp","sped_empresa","cnpj_
            strCod_Emp = !cod_emp
            strNome_Emp = !nome_emp
            strEndereco_Emp = !endereco_Emp
            strBairro_Emp = !bairro_emp
            strComplemento_Emp = !complemento_emp
            strLogradouro_Emp = !Logradouro_emp
            strNumero_Emp = !numero_emp
            strCidade_Emp = !Cidade_Emp
            strIP_Servidor = !ipserver_par
            strUF_Emp = !uf_emp
            strCEP_Emp = !cep_emp
            strFone_Emp = !Fone_Emp
            strFax_Emp = !Fax_Emp
            strCNPJ_Emp = F_SoNumero(!CNPJ_Emp)
            strIE_Emp = !Ie_Emp
            strCRT_emp = !CRT
            crcComissVen = !comissven_par
            If !usacrediario_par = "S" Then
                bUsaCrediario = True
            Else
                bUsaCrediario = False
            End If
            strUsaComissaoAcerto = Format(!usacomissaoacerto_par)
            If strUsaComissaoAcerto = "S" Then
                modular.bUsaComissaoAcerto = True
            End If
            strFantasia_Emp = !fantasia_emp
            strNomeResp_Emp = !NomeResp_Emp
            strNomeMoeda_Par = !NomeMoeda_Par
            strsimbMoeda_Par = !simbMoeda_Par
            strPluraMoeda_Par = !PLuralMoeda_PAr
            crcTxFederativa_Par = !TXFEDERATIVA_PAR
            crcTxAdm = !txadm_par
            
            crcTxMargemSeg_Par = !margemseg_par
            
            bUsaTEF = False
            bUsaECF = False
            strArquivoServidor = !caminhoexec_par
            sFoto_Temp = !fototemp_par
            sGravaFoto_Temp = !gravafototemp_par
            
            .Close
        End With
    End If
    InicializaEmpresa strCod_Emp
    P_BuscaDefUsu
    'On Error Resume Next
    strCorFundoUsu = buscaCorFundoUso(strCod_Usu)
    crcDesconto_Avista = CCur(F_BuscaUmCampo("desconto_for", "tab_formaspag", "cod_for='01'"))
'    P_AjustaVenda
   ' mdiFrm.lblUsu.caption = "" & strNome_Usu
   ' mdiFrm.lblCodUsu.caption = strCod_Usu
    'Bematech_FI_AbrePortaSerial
    If UCase(App.title) = "PIKIBOX_CX" Then
'        P_PortaComunicacaoECF
    End If
    P_BuscaDefUsu
    
    IniciaAcbr
    
    
    
    CrLf = Chr(13) + Chr(10)
    P_Status ""

End Sub

Function F_FormataCNPJ(objCPF As String)
    F_FormataCNPJ = Format(objCPF, "0#\.###\.###\/####\-##")
End Function

