Attribute VB_Name = "Modulo"


Global bUsaTEF As Boolean



Global trata_Erro As Label
Global Banco As database

Global bDedicado As database

Global strNum_Cpm As String
Global intCUPOM As Integer
Global strNome_Emp As String
Global strFantasia_Emp As String
Global strEndereco_Emp As String
Global strBairro_Emp As String
Global strCidade_Emp As String
Global strUF_Emp As String
Global strCEP_Emp As String
Global strFone_Emp As String
Global strFax_Emp As String
Global strCNPJ_Emp As String
Global strIE_Emp As String
Global strNomeResp_Emp As String
Global strCRCResp_Emp As String
Global strNomeMoeda_Par As String
Global strsimbMoeda_Par As String
Global strPluraMoeda_Par As String
Global crcTxFederativa_Par As Currency
Global crcTxMargemSeg_Par As Currency
Global crcTxAdm As Currency
Global crcTxDiversas As Currency
Global crcComissVen As Currency 'comissão sobre as vendas
Global strCod_Usu As String
Global strNome_Usu As String 'o nome
Global strHora As String 'teste
Global strNomeBanco As String
Global strNomeBancoReplica As String

Global strArquivoServidor As String
Public strImprFiscal As String

Global strAbrevenda_usu As String
Global strAbrenota_usu As String
Global strEfetnota_usu As String
Global strAbresaida_usu As String
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
Global strCadusu_usu As String
Global strCadcli_usu As String
Global strCadcco_usu As String
Global strExclVenda_usu As String
Global strCadformaspag_usu As String
Global strAltformaspag_usu As String
Global strViscusto_usu As String
Global strVisrelest_usu As String
Global strVisrelfin_usu As String
Global strFechacai_usu As String
Global strFechames_usu As String
Global strContrpre_usu As String
Global crcDesconto_usu As Currency
Global strIP_Usu As String 'o ip da mákina
Global strCod_Cai As String 'o código do caixa buscado pelo ip
Global strCod_Tur As String
Global crcCpm As Currency 'o valor da venda Tj blue
Global strIP_Servidor As String
Global picVis(9)  As PictureBox

Global strPlaVendasEfetuadas As String
Global strPlaRecebDupl As String
Global strPlaPagDupl As String
Global strPlaNaoClassD As String
Global strPlaNaoClassC As String


Public Const constSigla_sis = "Piki"
Public Const constNome_sis = "Piki Box "
Public Const constVersao_Sis = "V 2005"
Public Const ConstNome_Cria = "xPert Softwares"
Public Const ConstSlogan_Cria = "Soluções Inteligentes para Empresas Inteligentes"
Public Const ConstFone_Cria = "(..62) 3098-3005"
Public Const constMsg_Usu1 = "Obrigado! Volte Sempre!"
Public Const constMsg_Usu2 = "Pense Positivo!"
Public Const constContaCaixa = "001" 'CÓDIGO DA CONTA CAIXA PADRÃO

Type tBanco
    cod_bco As String
    nome_bco As String
End Type

Type planoContas
    cod_pla As String
    tipo_pla As String
    nome_pla As String
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
    count As Long
End Type

Type Produto
    cod_pro As String
    codbarra_pro As String
    nome_pro As String
    e1_pro As Currency
    e2_pro As Currency
    r_pro As Currency
    p1_pro As Currency
    p2_pro As Currency
    p3_pro As Currency 'preco promocional
    prom_pro As String
    c1_pro As Currency
    c2_pro As Currency
    cod_cat As String
    cod_mar As String
    nome_cat As String
    nome_mar As String
    qtd_pro As Currency
    Preco As Currency
    ipi As Currency
    s_pro As String
    lucro_pro As Currency
    codfab_pro As String
    icms As Currency
    ecf_icms As String
    balanca_pro As String
    
    totalbalanca_pro As Currency
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


Type tCliente
    cod_cli As String
    nome_cli As String
    endereco_cli As String
    cnpj_Cli As String
    ie_Cli As String
    telefone_cli As String
    limite_cli As Currency
    uf_cli As String
    bairro_cli As String
    cidade_cli As String
    cep_cli As String
    
End Type
Type cheque
    num_che As String
    cod_bco As String
    cco_che As String
    age_che As String
    nome_che As String
    cnpj_che As String
    Venc_che As String
    Valor_che As Currency
    cpf_che As String
End Type
Type chePag
    num_che As String
    cod_Con As String
    Venc_che As String
    Valor_che As Currency
End Type
Type parcela
    vencimento As String
    valor As Currency
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
        
        obj.Top = (mdiFrm.Height / 2) - (obj.Height / 2) - iRecuo
        
        
        obj.Left = ((mdiFrm.Width - 1320) / 2) - (obj.Width / 2)
        
            
        
End Sub
Sub P_CentraObejct(obj As Object, frm As Object, _
        Optional bTop As Boolean = True, Optional bLeft As Boolean = True)
    On Error Resume Next
        If bTop Then
            obj.Top = (frm.Height / 2) - (obj.Height / 2) - 200
        End If
        If bLeft Then
            obj.Left = (frm.Width / 2) - (obj.Width / 2) - 50
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

Dim Aux As String, Red As String, Acr As String
Acr = "F"
Aux = IIf(Dec > 0, "." + String(Dec, "0"), "")
If InStr(1, str(Num), ".") <> 0 Then
   
   Red = Trim(str(Num))
   Aux = Mid(Red, InStr(1, Red, "."))
   
   If Len(Aux) > Dec + 1 Then
      Red = Mid(Aux, Dec + 2, 1)
      If val(Red) > 4 Then
         
         Red = str(1 / val("1" + String(Dec, "0")))
         Aux = Trim(str(val(Aux) + val(Red)))
         
         If val(Aux) >= 1 Then
            Acr = "T"
            Aux = IIf(Dec > 0, "." + String(Dec, "0"), "")
         End If
      End If
   End If
   
   Aux = Left(Aux + "000000", Dec + 1)

End If
Red = Aux
If Acr <> "F" Then Num = Num + 1
Aux = IIf(Dec > 0, Trim(str(Fix(Num))), Trim(str(Int(Num))))
Aux = Right(String(Tam, "0") + Aux, Tam) + Red
Red = Trim(str(Int(Num)))

If Acr <> "F" Then Num = Num - 1

F_ZerEsq = IIf(Len(Red) > Len(Aux), String(Tam, "*") _
+ IIf(Dec > 0, "." + String(Dec, "*"), ""), Aux)

End Function
Function F_AbreBancoODBC(b As database, Nome As String) As Boolean
    
    On Error GoTo Erro
    DBEngine.DefaultType = dbUseODBC
    'Set bTopa = OpenDatabase("topatudo", , False, "ODBC;DATABASE=topatudo;")
    
    
    
    Set b = OpenDatabase(Nome, , False, "ODBC;DATABASE=" & Nome & ";")
    
   
    F_AbreBancoODBC = True
    Exit Function
Erro:
    MsgBox "Erro na Abertura do Banco!", vbExclamation, conSigla_Sistema
    F_AbreBancoODBC = False
End Function

Function F_AbreBanco(Optional bBase_Teste As Boolean = False) As Boolean
    
'    On Error GoTo Erro
    DBEngine.DefaultType = dbUseODBC
   
    
    If bBase_Teste = True Then
        Set Banco = OpenDatabase("market_teste", , False, "ODBC;DATABASE=market_teste;port=3306;")
    Else
        Set Banco = OpenDatabase(strNomeBanco, , False, "ODBC;DATABASE=" & strNomeBanco & ";port=3306;")
    End If
   
    F_AbreBanco = True
    Exit Function
Erro:
    


MsgBox Err.Description, vbCritical, conSigla_Sistema
    F_AbreBanco = False
End Function
Sub P_BuscaNomeBanco()
    Dim sNomeArq As String
    sNomeArq = "piki.ini"
    If Dir("c:\windows\system32\" & sNomeArq) = "" Then
        If Dir("c:\windows\system\" & sNomeArq) = "" Then
            If Dir(App.Path & "\" & sNomeArq) = "" Then
                strNomeBanco = "market"
            Else
                Open App.Path & "\" & sNomeArq For Input As #1
        
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
    'On Error GoTo Erro
    If App.PrevInstance = True Then
        MsgBox "Este Programa já Está sendo executado!", vbExclamation
        End
    End If
    strCaminhoPkXp = "C:\desenv\TEF Dedicado\Piki Express\executaveis\pikiexpress.exe"
    bTefDedicado = False
    P_BuscaNomeBanco
    If Not F_AbreBanco() Then
        End
    End If
    If UCase(App.Title) = "PIKIBOX_PESQ" Then
        
        
        frmConsPreco.Show 1
    Else
        
        frmLogin.Show 1
        If strNome_Usu = "" Then
            P_Fim
        End If
        Unload frmLogin
        frmSplash.Show 1
    

    
        
        
    
        mdiFrm.Show 0
    End If
    'frmMenuVendas.Show 0, mdiFrm
    'mdiFrm.Visible = False
    Exit Sub
Erro:
    P_Fim
 
End Sub

Sub P_Fim()
    Dim i As Integer
    On Error Resume Next
    For i = 1 To Forms.count - 1
        Unload Forms(i)
        
    Next
    Banco.Close
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
Sub P_AjustaTexto(ByRef objText As TextBox)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    'Torna o texto maiúsculo
    strAux = UCase(objText.Text)
    
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
    strAux = UCase(objText)
    
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

Sub P_SelecTexto(objText As Object)
    'Este procedimento seleciona o texto
    'do objeto em questão
    objText.SelStart = 0
    objText.SelLength = Len(RTrim(objText.Text))
End Sub
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
Function F_CPF(CPF As String) As Boolean

'Função que permite validar CPF
'Sintaxe.: F_CPF( CPF )
'Entrada.: CPF - número do CPF
'Saída...: Verdadeiro se o CPF existir, senão falso

Dim controle As String
Dim ContIni, ContiFim, soma As Integer
Dim ii As Integer
Dim i As Integer
Dim Digito As Integer

CPF = F_SoNumero(CPF)
CPF = Left(F_ZerEsq(val(CPF), 11, 0), 11)
ContIni = 2
ContiFim = 10
If val(CPF) < 1 Then
   F_CPF = False
   Exit Function
End If

For ii = 1 To 2
    soma = 0
    For i = ContIni To ContiFim
        soma = soma + (val(Mid(CPF, i - ii, 1)) * _
        (ContiFim + 1 + ii - i))
    Next i
    If ii = 2 Then
       soma = soma + (2 * Digito)
    End If
    Digito = (soma * 10) Mod 11
    If Digito = 10 Then Digito = 0
    
    controle = RTrim(controle) + _
    F_ZerEsq(CDbl(Digito), 1, 0)
    
    ContIni = 3
    ContiFim = 11
Next ii
F_CPF = IIf(Trim(controle) = Right(CPF, 2), True, False)
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
    mdiFrm.StatusBar1.Panels(iPanel).Text = sMsg
    
    If Trim(sMsg) <> "" And bEspere Then
        For iCount = 0 To Forms.count - 1
            Forms(iCount).MousePointer = 11
        Next
    Else
        For iCount = 0 To Forms.count - 1
            Forms(iCount).MousePointer = 0
        Next
    End If
    Exit Sub
trata_Erro:
On Error Resume Next
mdiFrm.StatusBar1.Panels(iPanel).Text = ""
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

Function callprinter(Optional bPag As Boolean = False, Optional bQuali As Boolean = False, Optional padrao As Boolean = False) As impressao
    On Error Resume Next
    Dim iPrinter As Integer
    If bPag = True Then
        frmTestPrint.optnDE.Enabled = True
        frmTestPrint.optnSelect.Enabled = True
        frmTestPrint.optnTUdo.Enabled = True
        frmTestPrint.Frame1.Enabled = True
        frmTestPrint.Frame2.Enabled = True
    End If
    If bQuali = True Then
        frmTestPrint.cmbxQuali.Enabled = True
    End If
    frmTestPrint.Show 1
    callprinter.PagFin = val(frmTestPrint.txtbxFin.Text)
    callprinter.PagIni = val(frmTestPrint.txtbxIni.Text)
    callprinter.Qualidade = frmTestPrint.cmbxQuali.ListIndex
    Unload frmTestPrint
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
Sub P_ImprLinha(obj As Object, Optional sChar As String = "-", Optional IQuant As Integer = 300)
    P_CentraImpressao obj, String(IQuant, sChar)
End Sub
Sub P_Line(obj As Object, x0 As Single, y0 As Single, x1 As Single, y1 As Single, _
    Optional drwStyle As DrawStyleConstants = DrawStyleConstants.vbSolid)
    Dim dr As Integer
    Dim x As Single
    Dim y As Single
    'on error resume next
    x = obj.CurrentX
    y = obj.CurrentY
    dr = obj.DrawStyle
    obj.DrawStyle = drwStyle
    obj.Line (x0, y0)-(x1, y1)
    obj.DrawStyle = dr
    obj.CurrentY = y
    obj.CurrentX = x
    End Sub
Sub P_line1(obj As Object, Optional drwStyle As DrawStyleConstants = DrawStyleConstants.vbSolid)
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth - 1, obj.CurrentY, drwStyle
End Sub


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
Sub P_PrintXY(objP As Object, sExpr As String, x As Single, y As Single, Optional bContinuaX = True, Optional btAllign As Byte = 1)
    Dim sY As Single, sX As Single '1-> Esquerda 2-Centra 3- Direita
    objP.ScaleMode = vbMillimeters
    objP.CurrentY = y
    If btAllign = 2 Then
        objP.CurrentX = (objP.ScaleWidth - objP.TextWidth(sExpr)) / 2
    ElseIf btAllign = 3 Then
        objP.CurrentX = x - objP.TextWidth(sExpr)
    Else
        objP.CurrentX = x
    End If
    objP.Print sExpr
    
    If bContinuaX = True Then objP.CurrentY = y
End Sub
Sub P_ImprX(objP As Object, strExpressao As String, CurrentX As Integer, Optional blnCont As Boolean = True, _
            Optional blnDireita As Boolean = False, Optional grade As Integer = -1, Optional dist As Single = 1.3)
    
    Dim sngX1 As Single
    Dim sngY1 As Single
    Dim HeiExp As Single ' altura da expressao
    Dim WidExp As Single ' Tamanho da expressao
    Dim xRetorno As Single ' o x que retornará para o objeto
    Dim yRetorno As Single ' o x que retornará para o objeto
    Dim xIni As Single
    Dim drSt As DrawStyleConstants
    objP.ScaleMode = vbMillimeters
    sngY1 = objP.CurrentY 'o y inicial
    sngX1 = objP.CurrentX ' o x inicial
    xIni = CurrentX ' o x inicial em que a expressão será impressa
    HeiExp = objP.TextHeight(strExpressao)
    WidExp = objP.TextWidth(strExpressao)
    
    objP.CurrentX = CurrentX
    
    If blnDireita Then
        objP.CurrentX = objP.CurrentX - objP.TextWidth(strExpressao)
    End If
    If blnCont Then
        objP.Print strExpressao;
        objP.CurrentY = sngY1
    Else
        objP.Print strExpressao
    End If
    xRetorno = objP.CurrentX
    yRetorno = objP.CurrentY
    If grade = 0 Then
        drSt = vbSolid
    Else
        drSt = vbDot
    End If
    If grade > -1 Then
        If blnDireita = False Then
            'P_Line objP, xIni - 0.1, sngY1, WidExp + xIni + 0.1, sngY1, vbSolid
             P_Line objP, xIni - dist, sngY1, xIni - dist, sngY1 + HeiExp, drSt
            'P_Line objP, WidExp + xIni + 0.1, sngY1, WidExp + xIni + 0.1, sngY1 + HeiExp, vbSolid
            'P_Line objP, xIni - 0.1, sngY1 + HeiExp, WidExp + xIni + 0.1, sngY1 + HeiExp, vbSolid
        Else
            xIni = xIni - WidExp - dist
            dist = dist + 1
            'P_Line objP, xIni - 0.1, sngY1, WidExp + xIni + 0.1, sngY1, vbSolid
            'P_Line objP, xIni - 0.1, sngY1, xIni - 0.1, sngY1 + HeiExp, vbDot
            P_Line objP, WidExp + xIni + dist, sngY1, WidExp + xIni + dist, sngY1 + HeiExp, drSt
            'P_Line objP, xIni - 0.1, sngY1 + HeiExp, WidExp + xIni + 0.1, sngY1 + HeiExp, vbSolid
        End If
    End If
    objP.CurrentY = yRetorno
    objP.CurrentX = xRetorno
End Sub
Sub P_CentraImpressao(objP As Object, strMens As String)

    objP.CurrentX = (objP.ScaleWidth - objP.TextWidth(strMens)) / 2
    objP.Print strMens
End Sub

Sub P_DireitaImpressao(objP As Object, strMens As String)
    objP.CurrentX = objP.ScaleWidth - objP.TextWidth(strMens)
    objP.Print strMens
End Sub
Sub P_CabRelTimbr(sMsg As String, obj As Object)
    Dim iScale As Integer
    Dim wid As Single
    Dim hei As Single
    Dim iY As Single
    iScale = obj.ScaleMode
    Dim ifont As Integer
    On Error Resume Next
    obj.ScaleMode = 1
    obj.fontsize = 14
    obj.FontName = "arial"
    ifont = obj.fontsize
    P_CentraImpressao obj, strNome_Emp
    obj.fontsize = 10
    P_CentraImpressao obj, strEndereco_Emp & " " & strBairro_Emp & " " & strCidade_Emp & "/" & strUF_Emp
    
    
    
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
    
    obj.fontsize = 12
    obj.FontUnderline = True
    iY = obj.CurrentY
    P_CentraImpressao obj, sMsg
    obj.fontsize = 8
    obj.CurrentY = iY
    P_ImprX obj, Date & "-" & Time, 0, False
    obj.FontUnderline = False
    obj.fontsize = 10
    obj.Print ""
    
    
    'obj.CurrentY = obj.CurrentY + 1
End Sub
Sub P_Timbr(sMsg As String, obj As Object)
    obj.ScaleMode = 1
    obj.PaintPicture mdiFrm.ImageListPrint.ListImages(1).Picture, _
        obj.CurrentX - 10, obj.CurrentY - 5, 800, 700
    
    obj.fontsize = 18
    obj.FontName = "Arial"
    P_PrintXY obj, strNome_Emp, 15, obj.CurrentY
    obj.fontsize = 10
    P_ImprX obj, strEndereco_Emp & "   Fone: " & _
        strFone_Emp, 70, False
    P_PrintXY obj, strFantasia_Emp, 15, Printer.CurrentY + 2
    obj.fontsize = 10
    P_ImprX obj, sMsg & "  ", 70, False
    
    'obj.Line (68, obj.CurrentY - 12)-(68, obj.CurrentY)
    'obj.Line (0, obj.CurrentY + 2)-(200, obj.CurrentY + 2)
    'obj.Line (168, 0)-(168, 12)
    'obj.CurrentY = obj.CurrentY + 1
End Sub

Sub P_Rodape(Optional strMens As String = "", Optional blnNovaPagina = True, Optional msgDireita As String = "")
    Printer.FontName = "Times New Roman"
    If F_Vazio(strMens) Then
        strMens = ConstEmpresa
    End If
    Printer.FontName = "arial"
    Printer.fontsize = 10
    Printer.CurrentY = Printer.ScaleHeight - 15
    P_ImprX Printer, String(200, "_"), 0, False
    Printer.fontsize = 10
    Printer.FontBold = True
    
    P_ImprX Printer, msgDireita, 200, True, True
    P_CentraImpressao Printer, strMens
    Printer.fontsize = 8
    Printer.fontsize = 10
    Printer.FontBold = False
    Printer.Print "" & String(200, "_") & ""
    If blnNovaPagina Then
        Printer.NewPage
    End If
End Sub
Sub P_FimRel(obj As Object, Optional strMens As String = "", Optional blnNovaPagina = True, Optional msgDireita As String = "")
    Dim y As Single
    obj.FontName = "Times New Roman"
    If F_Vazio(strMens) Then
        strMens = ConstEmpresa
    End If
    obj.FontName = "arial"
    obj.fontsize = 8
    
    P_ImprX obj, String(200, "_"), 0, False
    obj.fontsize = 10
    obj.FontBold = False
    
    P_ImprX obj, msgDireita, 200, True, True
    y = obj.CurrentY
    obj.fontsize = 5
    P_ImprX obj, constNome_sis, 0, False
    P_ImprX obj, ConstNome_Cria & " - Anápolis/GO - " & ConstFone_Cria, 0, False
    P_ImprX obj, ConstSlogan_Cria, 0, False
    obj.CurrentY = y
    obj.fontsize = 10
    P_CentraImpressao obj, strMens
    obj.fontsize = 8
    obj.fontsize = 10
    obj.FontBold = False
    obj.CurrentY = obj.CurrentY + 0.5
    obj.Print "" & String(200, "-") & ""
    On Error GoTo trata_Erro
    
    obj.NewPage
    
trata_Erro:
End Sub

Sub P_RodapeVis(obj As Object, Optional strMens As String = "", Optional msgDireita As String = "")
    On Error Resume Next
    obj.FontName = "Times New Roman"
    If F_Vazio(strMens) Then
        strMens = ConstEmpresa
    End If
    obj.FontName = "arial"
    obj.fontsize = 10
    obj.CurrentY = obj.ScaleHeight - 15
    P_ImprX obj, String(200, "_"), 0, False
    obj.fontsize = 10
    obj.FontBold = True
    
    P_ImprX obj, msgDireita, 200, True, True
    P_CentraImpressao obj, strMens
    obj.fontsize = 8
    obj.fontsize = 10
    obj.FontBold = False
    obj.Print "" & String(200, "_") & ""
    obj.NewPage
    obj.endoc
    
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

Function F_BuscaProd(Optional busca As String = "N", _
                 Optional sIni As String = "") As Produto
    
    On Error Resume Next
    frmSelectProd.txtbxExpr.Text = sIni
    frmSelectProd.Show 1
    
    If frmSelectProd.Tag = "" Then
        Unload frmSelectProd
        Exit Function
    Else
        With frmSelectProd.dtPro.Recordset
            
            F_BuscaProd.cod_cat = !cod_cat
            F_BuscaProd.cod_mar = !cod_mar
            
            F_BuscaProd.cod_pro = !cod_pro
            'F_BuscaProd.custo = !custo_pro
            'F_BuscaProd.estoque = !estoque_pro
            F_BuscaProd.nome_cat = !nome_cat
            F_BuscaProd.nome_mar = !nome_mar
            
            F_BuscaProd.nome_pro = !nome_pro
            F_BuscaProd.Preco = !preco_pro
            'F_BuscaProd.reservado = !reservado_pro
            'F_BuscaProd.pei = !pei_pro
            'F_BuscaProd.tipo = !tipo_pro
                    
            Unload frmSelectProd
        End With
    End If
    
End Function
Function F_BuscaProdExt() As Pedido
    frmSelectProd.Tag = "E"
    On Error Resume Next
    
    frmSelectProd.Show 1
    
    If frmSelectProd.Tag = "" Then
        Unload frmSelectProd
        Exit Function
    Else
        With frmSelectProd.dtPro.Recordset
            
                       
            F_BuscaProdExt.cod_pro = !cod_pro
               
            
            
            
            F_BuscaProdExt.nome_pro = !nome_pro
            F_BuscaProdExt.unit_pro = !preco_pro
            
            On Error Resume Next
            
            frmValoresExt.lblCod_Pro.caption = !cod_pro
            frmValoresExt.lblNome_pro.caption = !nome_pro
            
            frmValoresExt.txtbxPreco.Text = FormatVal(!preco_pro)
            
                       
            frmValoresExt.Show 1
            
            F_BuscaProdExt.qtd_pro = CCur(frmValoresExt.txtbxQtd.Text)
            F_BuscaProdExt.loja = Left(frmValoresExt.cmbxLoja.Text, 3)
            
            Unload frmValoresExt
            Unload frmSelectProd
        End With
    End If
    
End Function
Function F_BuscaProdPreVen() As Prevenda
    frmSelectProd.Tag = "E"
    On Error Resume Next
    
    frmSelectProd.Show 1
    
    If frmSelectProd.Tag = "" Then
        Unload frmSelectProd
        Exit Function
    Else
        With frmSelectProd.dtPro.Recordset
            
                       
            F_BuscaProdPreVen.cod_pro = !cod_pro
               
            
            
            
            F_BuscaProdPreVen.nome_pro = !nome_pro
            F_BuscaProdPreVen.unit_pro = !preco_pro
            
            On Error Resume Next
            
            frmValoresPReVEn.lblCod_Pro.caption = !cod_pro
            frmValoresPReVEn.lblNome_pro.caption = !nome_pro
            
            frmValoresPReVEn.txtbxPreco.Text = FormatVal(!preco_pro)
            
                       
            frmValoresPReVEn.Show 1
            
            F_BuscaProdPreVen.qtd_pro = CCur(frmValoresPReVEn.txtbxQtd.Text)
            
            
            Unload frmValoresPReVEn
            Unload frmSelectProd
        End With
    End If
    
End Function
Function F_buscacodbarra_ent(sCodbar As String, _
cIpi_Max As Currency, cValor_Max As Currency, dInd As Double) As Produto
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim p As Produto
    

        p = F_BuscaProdSaida(sCodbar, "B")
        If p.cod_pro = "" Then
            Exit Function
        End If
        F_buscacodbarra_ent.cod_cat = p.cod_cat
           F_buscacodbarra_ent.cod_mar = p.cod_mar
           F_buscacodbarra_ent.cod_pro = p.cod_pro
           F_buscacodbarra_ent.codbarra_pro = p.codbarra_pro
           F_buscacodbarra_ent.c1_pro = p.c1_pro
           F_buscacodbarra_ent.c2_pro = p.c2_pro
           F_buscacodbarra_ent.e1_pro = p.e1_pro
           F_buscacodbarra_ent.e2_pro = p.e2_pro
           F_buscacodbarra_ent.s_pro = p.s_pro
           F_buscacodbarra_ent.nome_cat = p.nome_cat
           F_buscacodbarra_ent.nome_mar = p.nome_mar
           F_buscacodbarra_ent.nome_pro = p.nome_pro
           F_buscacodbarra_ent.p1_pro = p.p1_pro
           F_buscacodbarra_ent.p2_pro = p.p2_pro
            
           F_buscacodbarra_ent.r_pro = p.r_pro
            
            On Error Resume Next
            
            frmValoresEntr.lblCod_Pro.caption = p.cod_pro
            frmValoresEntr.lblNome_pro.caption = p.nome_pro
            frmValoresEntr.txtbxCusto.Text = FormatVal(p.c2_pro)
            frmValoresEntr.txtbxPreco.Text = FormatVal(p.p2_pro)
            frmValoresEntr.lblCat.caption = p.nome_cat
            
            frmValoresEntr.lblMar.caption = p.nome_mar
            frmValoresEntr.txtbxIpi.Visible = True
            frmValoresEntr.lblIPI.Visible = True
            frmValoresEntr.txtbxPreco.Text = "0,00"
            frmValoresEntr.txtbxPreco.Enabled = True
            frmValoresEntr.txtbxPreco.Tag = cValor_Max
            frmValoresEntr.lblRestante.caption = "Vlr Rest:(R$): " & FormatVal(cValor_Max)
            frmValoresEntr.txtbxIpi.Tag = cIpi_Max
            frmValoresEntr.lblIPI_Rest.caption = "IPI Rest:(R$): " & FormatVal(cIpi_Max)
                
            frmValoresEntr.Show 1
            
           F_buscacodbarra_ent.qtd_pro = CCur(frmValoresEntr.txtbxQtd.Text)
            
           F_buscacodbarra_ent.Preco = CCur(frmValoresEntr.txtbxPreco.Text)
           F_buscacodbarra_ent.ipi = CCur(frmValoresEntr.txtbxIpi.Text)
            
            Unload frmValoresEntr
            

    
    
End Function

Sub LockForm(opcao As Boolean, frm As Form)
    'opcao true loca, false desloca
    Dim iCount As Integer
    For iCount = 0 To Forms.count - 1
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
    For i = 0 To Forms.count - 1
        If UCase(Forms(i).Name) = UCase(Nomefrm) Then
            IsLoad = True
            Exit Function
        End If
    Next
    IsLoad = False
End Function
Sub RefreshDataPosition(dt As Data)
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
Function F_DadosCheque(bVen As Boolean, Optional ven As String = "", Optional Nome As String = "", Optional Num As String = "", _
            Optional cco As String = "", Optional age As String = "", Optional bco As String = "", _
            Optional valor As String = "0,00", Optional sCPF As String = "", Optional JAM As Currency = 0) As cheque
    
    
    With frmDadosChe
        .txtbxAgencia.Text = age
        .txtbxContaCorrente.Text = cco
        .txtbxNome_Che.Text = Nome
        .txtbxNum_Che.Text = Num
        .txtbxValor_Che.Text = valor
        .txtbxVencimento.Text = ven
        .txtbxCPF.Text = sCPF
        .lblValor.caption = FormatVal(valor)
        .lblJam.caption = FormatVal(JAM)
        If bco <> "" Then
            P_LocalizaOperacao bco, .cmbxBanco
        End If
        If bVen = False Then
            .txtbxVencimento.Enabled = False
        End If
        On Error Resume Next
        
        .Show 1
        If .Tag <> "" Then
            F_DadosCheque.age_che = .txtbxAgencia.Text
            F_DadosCheque.cco_che = .txtbxContaCorrente.Text
            F_DadosCheque.nome_che = .txtbxNome_Che.Text
            F_DadosCheque.num_che = .txtbxNum_Che.Text
            F_DadosCheque.Valor_che = .txtbxValor_Che.Text
            F_DadosCheque.Venc_che = .txtbxVencimento.Text
            F_DadosCheque.cod_bco = Left(.cmbxBanco.Text, 3)
            F_DadosCheque.cpf_che = .txtbxCPF.Text
            
        End If
    End With
        Unload frmDadosChe
End Function
Function F_DadosChePag(Optional ven As String = "", Optional Num As String = "", _
            Optional cod_Con As String = "", Optional sVal As String = "", Optional lLock As Boolean = False) As chePag
    
    With frmDadosChPag
        
            .txtbxVencimento.Text = ven
            .txtbxNum_Che.Text = Num
            .txtbxValor_Che.Text = sVal
        
        If cod_Con <> "" Then
            P_LocalizaOperacao cod_Con, .cmbxContas
        End If
        
        If lLock = True Then
            .txtbxVencimento.Enabled = False
            .txtbxValor_Che.Enabled = False
            .cmbxContas.Enabled = False
        End If
        On Error Resume Next
        
        .Show 1
        
            F_DadosChePag.cod_Con = Left(.cmbxContas.Text, 3)
            F_DadosChePag.num_che = .txtbxNum_Che.Text
            F_DadosChePag.Valor_che = CCur(.txtbxValor_Che.Text)
            F_DadosChePag.Venc_che = .txtbxVencimento.Text
        
    End With
        Unload frmDadosChPag
        
    
    
End Function

Function F_DadosPar(Optional ven As String = "", _
    Optional valor As String = "0,00") As parcela
    
    
    With frmDadosPar
        .txtbxValor_Che.Text = valor
        .txtbxVencimento.Text = ven
                
        On Error Resume Next
        
        .Show 1
        If .Tag <> "" Then
            
            F_DadosPar.valor = .txtbxValor_Che.Text
            F_DadosPar.vencimento = .txtbxVencimento.Text
            
        End If
    End With
        Unload frmDadosPar
        
    
    
End Function


Sub MudaAparencia(frm As Form)
    Dim strSQL As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim rPesq As Recordset
    Dim sObj As String
    On Error Resume Next
    strSQL = "select * from tab_aparencia where cod_usu = '" & strCod_Usu & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rPesq.RecordCount = 0 Then Exit Sub
    For iCount = 0 To frm.Controls.count - 1
        sObj = LCase(TypeName(frm.Controls(iCount)))
        If sObj = "label" Then
            If frm.Controls(iCount).Tag <> "!" Then 'if ! not change
                frm.Controls(iCount).Appearance = rPesq!appearance_par
                frm.Controls(iCount).ForeColor = rPesq!corfonte_par
                frm.Controls(iCount).BackStyle = 0
            Else
                If frm.Controls(iCount).Name = "lblCaptionFrame" Then
                    frm.Controls(iCount).caption = frm.caption
                End If
            End If
        ElseIf sObj = "textbox" Or sObj = "combobox" Then
            If frm.Controls(iCount).Tag <> "!" Then 'if ! not change
                frm.Controls(iCount).ForeColor = rPesq!corfontecaixa_par
                frm.Controls(iCount).BackColor = rPesq!corcaixa_par
                frm.Controls(iCount).Appearance = rPesq!appearance_par
            End If
        ElseIf sObj = "msflexgrid" Or sObj = "dbgrid" Then
            frm.Controls(iCount).BackColor = rPesq!corcaixa_par
            frm.Controls(iCount).ForeColor = rPesq!corfontecaixa_par
            If sObj = "dbgrid" Then
                For iCol = 0 To frm.Controls(iCount).Columns.count - 1
                    frm.Controls(iCount).Columns(iCol).DividerStyle = 6
                Next
                frm.Controls(iCount).RowDividerStyle = 6
                frm.Controls(iCount).MarqueeStyle = 0
            End If
            
            frm.Controls(iCount).Appearance = rPesq!appearance_par
        ElseIf sObj = "frame" Or sObj = "sstab" Then
            If frm.Controls(iCount).Tag <> "!" Then
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

Sub P_MostraErro(Optional strProc As String = "Erro ")
  'Procedimento com a finalidade de comentar erro
  'Sintaxe.: P_MostraErro()
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
  Beep
  Beep
  Beep
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
    P_SelecTexto objValor
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
            .Tag = "Simples"
            .Show 1
       End With
    End If
    F_DataInicialFinal = frmDataInicialFinal.Tag 'Iguala com o tag
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


Function F_BuscaProdSaida(Optional busca As String = "", Optional tipo_busca As String = "N") As Produto
    On Error Resume Next
    Dim rPesq As Recordset
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    On Error GoTo trata_Erro
    For iCount = 1 To 5
        strSQL = "select distinct tc.nome_Cat, tm.nome_mar,"
        strSQL = strSQL & "tp.*,(tp.er_pro) as estoque "
        strSQL = strSQL & " from tab_produtos as tp,tab_categorias as tc, "
        strSQL = strSQL & " tab_codbarras as tb, "
        strSQL = strSQL & " tab_marcas as tm"
    
        If iCount = 1 Then
            strSQL = strSQL & " where tb.codbarra_PRO like '"
            strSQL = strSQL & busca & "' "
            tipo_busca = "B"
        ElseIf iCount = 2 Then
            strSQL = strSQL & " where tp.COD_PRO like '"
            strSQL = strSQL & busca & "%' "
            tipo_busca = "C"
        ElseIf iCount = 4 Then
            strSQL = strSQL & " where tp.CODFAB_PRO like '"
            strSQL = strSQL & busca & "%' "
            tipo_busca = "R"
        ElseIf iCount = 5 Then
            strSQL = strSQL & " where tb.CODbarra_PRO = '"
            strSQL = strSQL & Left(busca, 7) & "' "
            strSQL = strSQL & " and tp.balanca_pro='S' "
            tipo_busca = "B"
        Else
            strSQL = strSQL & " where NOME_PRO like '"
            strSQL = strSQL & busca & "%' "
            tipo_busca = "N"
        End If
    
        strSQL = strSQL & " and tp.cod_pro=tb.cod_pro "
        strSQL = strSQL & " and tc.cod_cat=tp.cod_Cat and tm.cod_mar=tp.cod_mar "
        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
        
    
        
    
        With rPesq
            If .RecordCount = 1 Then
                F_BuscaProdSaida.cod_cat = !cod_cat
                F_BuscaProdSaida.cod_mar = !cod_mar
                F_BuscaProdSaida.cod_pro = !cod_pro
                F_BuscaProdSaida.codbarra_pro = !codbarra_pro
                F_BuscaProdSaida.e1_pro = !er_pro
                F_BuscaProdSaida.e2_pro = !er_pro
                F_BuscaProdSaida.nome_cat = !nome_cat
                F_BuscaProdSaida.nome_mar = !nome_mar
                F_BuscaProdSaida.c2_pro = !cr_pro
                F_BuscaProdSaida.nome_pro = !nome_pro
                F_BuscaProdSaida.p1_pro = !pV_pro
                F_BuscaProdSaida.codfab_pro = !codfab_pro
                F_BuscaProdSaida.lucro_pro = !lucro_pro
                F_BuscaProdSaida.p2_pro = !pV_pro
                F_BuscaProdSaida.p3_pro = !pV_pro
                F_BuscaProdSaida.s_pro = !st_pro
                F_BuscaProdSaida.ecf_icms = F_IcmsProd(!cod_pro, 1)
                
                F_BuscaProdSaida.codbarra_pro = F_BuscaUmCampo("codbarra_pro", "tab_codbarras", "cod_pro='" & !cod_pro & "'")
                F_BuscaProdSaida.balanca_pro = !balanca_pro
                If iCount = 5 Then
                    F_BuscaProdSaida.totalbalanca_pro = CCur(Mid(busca, 8, 5)) / 100
                    'F_BuscaProdSaida.qtd_pro = ccur(
                End If
                
                Exit Function
            End If
        End With
    
    Next
    
    
    
    frmProdutoSaida.Tag = busca
    
    frmProdutoSaida.dbgrd.Tag = tipo_busca 'código de barra
    frmProdutoSaida.Show 1
    
    If frmProdutoSaida.Tag = "" Then
        Unload frmProdutoSaida
        Exit Function
    Else
        With frmProdutoSaida.dtPro.Recordset
            
            F_BuscaProdSaida.cod_cat = !cod_cat
            F_BuscaProdSaida.cod_mar = !cod_mar
            
            F_BuscaProdSaida.cod_pro = !cod_pro
            F_BuscaProdSaida.codbarra_pro = !codbarra_pro
            F_BuscaProdSaida.e1_pro = !ef_pro
            F_BuscaProdSaida.e2_pro = !er_pro
            F_BuscaProdSaida.nome_cat = !nome_cat
            F_BuscaProdSaida.nome_mar = !nome_mar
            F_BuscaProdSaida.c2_pro = !cr_pro
            F_BuscaProdSaida.nome_pro = !nome_pro
            F_BuscaProdSaida.p1_pro = !pV_pro
            F_BuscaProdSaida.codfab_pro = !codfab_pro
            
            F_BuscaProdSaida.p2_pro = !pV_pro
            F_BuscaProdSaida.lucro_pro = !lucro_pro
            
            F_BuscaProdSaida.p3_pro = !pV_pro
            F_BuscaProdSaida.s_pro = !st_pro
            F_BuscaProdSaida.ecf_icms = F_IcmsProd(!cod_pro, 3)
            F_BuscaProdSaida.codbarra_pro = F_BuscaUmCampo("codbarra_pro", "tab_codbarras", "cod_pro='" & !cod_pro & "'")
            F_BuscaProdSaida.balanca_pro = !balanca_pro
            Unload frmProdutoSaida
        End With
    End If
    Exit Function
trata_Erro:
    
End Function

Sub P_AjustaEstoqueReservado()
        
    
    
    
    
    
    
    
    
    
    
End Sub
Function F_SenhaCodBar() As String
    On Error Resume Next
    With frmSenhaCodBar
        .Show 1
        
        F_SenhaCodBar = .Tag
        
        Unload frmSenhaCodBar
    End With
End Function
Function F_senha(Optional bLockUser As Boolean = True) As Boolean
    Dim sSenha As String
    Dim strSQL As String
    
        frmLogin.Tag = "senha"
        On Error Resume Next
        frmLogin.txtbxNome_Usu.Text = strNome_Usu
        If bLockUser = True Then
            frmLogin.txtbxNome_Usu.Locked = True
        Else
            frmLogin.txtbxNome_Usu.Locked = False
        End If
        frmLogin.BorderStyle = 1
        SendKeys Chr(9)
        frmLogin.Show 1
        If frmLogin.Tag = "" Then
               
            F_senha = False
        Else
            F_senha = True
        End If
        Unload frmLogin
End Function
Sub P_ImprxEstiquetas(obj As Object)
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_produtos order by nome_pro "

    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            obj.PaintPicture mdiFrm.ImageListPrint.ListImages(1).Picture, -10, 5, 500, 400
            obj.fontsize = 14
            P_ImprX obj, strFantasia_Emp, 10
            
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
Function F_Troco(valor As Currency, num_ven As String, Cod_Ven As String, _
    cod_cli As String, nome_cli As String, operacao As String) As String
    'On Error Resume Next
    frmTroco.Tag = num_ven
    frmTroco.lblValor.caption = FormatVal(valor)
    frmTroco.lblValor.Tag = FormatVal(valor)
    frmTroco.lblNome_ven.Tag = Cod_Ven
    frmTroco.lblNome_ven.caption = F_BuscaUmCampo("guerra_fun", _
        "tab_funcionarios", "cod_fun='" & Cod_Ven & "'")
    frmTroco.operacao = operacao
    frmTroco.txtbxNome_Cli.Tag = cod_cli
    frmTroco.txtbxNome_Cli.Text = nome_cli
    frmVenda.Enabled = False
    frmTroco.Show 0, mdiFrm
    frmTroco.SetFocus
    Do While frmTroco.Visible = True
        
        DoEvents
    Loop
    F_Troco = frmTroco.Tag
    
    On Error Resume Next
    Unload frmMsg
    Unload frmTroco
    frmVenda.Enabled = True
End Function
Function F_UsuLogadoCaixa(sCod_Cai As String) As String
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select tf.guerra_fun from tab_turncai as tt, "
    strSQL = strSQL & " tab_Funcionarios as tf where cod_cai='" & sCod_Cai & "' "
    strSQL = strSQL & " and cod_Tur='" & strCod_Tur & "' "
    strSQL = strSQL & " and data_tur=" & F_GravaData(Date)
    strSQL = strSQL & " and tf.cod_fun=tt.login_usu "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
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
    F_EscolheVendedor = frmEscolheVendedor.Tag
    Unload frmEscolheVendedor
End Function
Function F_Vendedor() As String
    On Error Resume Next
    frmVendedor.Show 1
    F_Vendedor = frmVendedor.Tag

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

Function F_InFlxgrd(msf As MSFlexGrid, Text As String, colum As Integer) As Integer
    Dim iCount As Integer
    Dim bOk As Boolean
    With msf
        bOk = False
        For iCount = 0 To .rows - 1
            .Row = iCount
            .Col = colum
            If UCase(.Text) = UCase(Text) Then
                bOk = True
                F_InFlxgrd = iCount
                Exit For
            End If
        Next
        If bOk = False Then
            F_InFlxgrd = -1
        End If
        
    End With
End Function

Sub P_Mensagem(msg As String, frm As Form, Optional acao As Boolean = True)
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
        frmLogin.Tag = "senha"
        On Error Resume Next
        
        frmLogin.txtbxNome_Usu.Text = strNome_Usu
        frmLogin.txtbxNome_Usu.Locked = bLock
        frmLogin.BorderStyle = 1
        SendKeys Chr(9)
        frmLogin.Show 1
        
            If frmLogin.Tag <> "" Then
                F_senhaReturnUsu = frmLogin.Tag
            End If
        If novoUsu = False Then
            strNome_Usu = sNome
            strCod_Usu = sCod
        End If
        
        Unload frmLogin
End Function

Sub P_InfoEntradasPro(cod_pro As String)
    Dim fr As New frmInfoEnt
    fr.Tag = cod_pro
    fr.Show 1
End Sub
Sub P_InfoSaidasPro(cod_pro As String)
    Dim fr As New frmInfSai
    fr.Tag = cod_pro
    fr.Show 1
End Sub

Sub P_Msg(Optional msg As String, Optional titulo As String)
    On Error Resume Next
    If msg <> "" Then
        frmMsg.caption = titulo
        frmMsg.lblMsg.caption = msg
    
        frmMsg.Show 0, mdiFrm
        frmMsg.SetFocus
        frmMsg.Refresh
    Else
        Unload frmMsg
    End If
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
Sub zebraGrid(grid As MSFlexGrid, linhaIni As Integer, Optional cor = &HD3E7EB)



    On Error Resume Next
    Dim iCount As Integer
    Dim iRow As Integer
    With grid
        For iRow = linhaIni To grid.rows - 1 Step 2
        .Row = iRow
            For iCount = 0 To .Cols - 1
                .Col = iCount
                .CellBackColor = cor
            
            Next
        Next
    End With
End Sub
Sub ColoreLinhaGrid(grid As MSFlexGrid, Linha As Integer, Optional cor = &HD3E7EB, _
    Optional negrito As Boolean = False, Optional foreCor = "")



    On Error Resume Next
    Dim iCount As Integer
    Dim iRow As Integer
    With grid
        .Row = Linha
        For iCount = 0 To .Cols - 1
            .Col = iCount
            .CellBackColor = cor
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
            .Row = iRow
            For iCount = 0 To .Cols - 1
                
                .Col = iCount
                If .CellWidth < Printer.TextWidth(.Text) Then
                    .ColWidth(.Col) = Printer.TextWidth(.Text)
                End If
            Next
        Next
        msflxgrd.Visible = True
    End With
End Sub
Sub P_EspelhoNota(seq_Not As String)
    Dim fr As New frmEspelhoNota
    fr.Tag = seq_Not
    fr.Show 0
End Sub
Function F_BuscaCliente(tipo_ter As String, busca As String) As tCliente
    On Error Resume Next
    frmSelectCli.txtbxNome.Tag = tipo_ter
    frmSelectCli.txtbxNome.Text = busca
    frmSelectCli.Show 1
    If frmSelectCli.Tag <> "" Then
        F_BuscaCliente.cod_cli = frmSelectCli.dtBusca.Recordset!cod
        F_BuscaCliente.nome_cli = frmSelectCli.dtBusca.Recordset!Nome
        F_BuscaCliente.telefone_cli = frmSelectCli.dtBusca.Recordset!telefone
        F_BuscaCliente.endereco_cli = frmSelectCli.dtBusca.Recordset!Endereco
        F_BuscaCliente.ie_Cli = frmSelectCli.dtBusca.Recordset!rg
        F_BuscaCliente.cnpj_Cli = frmSelectCli.dtBusca.Recordset!CPF
        F_BuscaCliente.limite_cli = frmSelectCli.dtBusca.Recordset!limite_cli
        F_BuscaCliente.bairro_cli = frmSelectCli.dtBusca.Recordset!bairro_cli
        F_BuscaCliente.cidade_cli = frmSelectCli.dtBusca.Recordset!cidade_cli
        F_BuscaCliente.cep_cli = frmSelectCli.dtBusca.Recordset!cep_cli
        F_BuscaCliente.uf_cli = frmSelectCli.dtBusca.Recordset!uf_cli
        
    End If
    Unload frmSelectCli
End Function
Sub P_cancelaNotaFiscalSaida(nota_sai As String, serie_sai As String)
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_itsai where nota_sai='" & nota_sai & "' "
    strSQL = strSQL & " and serie_Sai='" & serie_sai & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_produtos set ef_pro=ef_pro+" & F_GravaDec(!qtd_its)
            strSQL = strSQL & " where cod_pro ='" & !cod_pro & "' "
            Banco.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update tab_saidas set serie_sai='CANC' "
    strSQL = strSQL & " where nota_sai='" & nota_sai & "' "
    strSQL = strSQL & " and serie_sai='" & serie_sai & "' "
    Banco.Execute strSQL
    MsgBox "Nota Fiscal Cancelada!", vbInformation
End Sub

Sub P_AjustaVenda()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sMsg As String
    strSQL = "select * from tab_vendas where situacao_ven in('2','3') "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("num_ven", "tab_venfec", "num_ven='" & !num_ven & "'") = "" Then
                strSQL = "Update tab_vendas set situacao_ven='1' "
                strSQL = strSQL & " where num_ven='" & !num_ven & "' "
                Banco.Execute strSQL
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
Function F_BuscaBanco(expr As String) As tBanco
    On Error Resume Next
    frmBuscaBan.txtbxExpr.Text = expr
    frmBuscaBan.Show 1
    
    If frmBuscaBan.Tag <> "" Then
        F_BuscaBanco.cod_bco = frmBuscaBan.dtServ.Recordset!cod_bco
        F_BuscaBanco.nome_bco = frmBuscaBan.dtServ.Recordset!nome_bco
    End If
    Unload frmBuscaBan
End Function
Function F_DevolveChe(seq_che As String) As Boolean
    frmDevChe.Tag = seq_che
    frmDevChe.Show 1
    If frmDevChe.Tag <> "" Then
        F_DevolveChe = True
    Else
        F_DevolveChe = False
    End If
    
    
End Function
Sub P_HistCheCPF(cpf_cli As String)
    frmHistoricoChe.Tag = cpf_cli
    frmHistoricoChe.txtbxCPF.Text = cpf_cli
    frmHistoricoChe.Show 1
    
    
    
End Sub
Function F_HistCheCPF(CPF As String) As cheque

    Dim stSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    strSQL = "select distinct cco_che,cod_bco,age_che,cpf_che,nome_che from tab_cheques1 where cpf_che='" & CPF & "' "
    strSQL = strSQL & " order by seq_che "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
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
            frmBuscaCheCPF.msflxgrd.Col = 0
            
            F_HistCheCPF.cod_bco = frmBuscaCheCPF.msflxgrd.Text
            frmBuscaCheCPF.msflxgrd.Col = 2
            F_HistCheCPF.age_che = frmBuscaCheCPF.msflxgrd.Text
            frmBuscaCheCPF.msflxgrd.Col = 3
            F_HistCheCPF.cco_che = frmBuscaCheCPF.msflxgrd.Text
            
            frmBuscaCheCPF.msflxgrd.Col = 4
            F_HistCheCPF.nome_che = frmBuscaCheCPF.msflxgrd.Text
            Unload frmBuscaCheCPF
            
        End If
        .Close
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
    Banco.Execute strSQL
    Exit Sub
trata_Erro:
End Sub

Function verificaVendaAberta(Cod_Cai As String) As Boolean
    Dim strSQL As String
    Dim rPesq As Recordset
    'On Error Resume Next
    strSQL = "select tf.* from  tab_vendas as tv, "
    strSQL = strSQL & " tab_venfec as tf "
    strSQL = strSQL & " where tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & Cod_Cai & "' "
    strSQL = strSQL & " and situacao_ven=1 "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
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
    sDest = App.Path & "\" & App.EXEName & ".exe"
    sOrig = strArquivoServidor & App.EXEName & ".exe"
    
    If FileDateTime(sOrig) > FileDateTime(sDest) Then
        strSQL = "Existe uma versão mais nova para atualização."
        strSQL = strSQL & "Assim que possível saia do  " & constNome_sis & " e execute o programa de instalação."
        MsgBox strSQL, vbInformation
            Open "c:\atlz_pk.txt" For Output As #1
            Print #1, sOrig
            Print #1, sDest
        
        Close #1
        
        Shell App.Path & "\atlz_pk.exe", vbNormalFocus
        
        
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
    callprinter
    frmGeraNota.Tag = num_ven
    frmGeraNota.txtbxVenda.Text = num_ven
    If strCUPOM <> "" Then
        frmGeraNota.lblCUPOM.caption = strCUPOM
        frmGeraNota.lblCUPOM.Tag = strCUPOM
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
                
        sNumVen = F_BuscaUmCampo("num_ven", "tab_saidas", "nota_sai='" & strNota_Sai & "' and cod_cai='" & strCod_Cai & "'")
        
        
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
                
        If Not F_CancelaVendaAposCupom(sNumVen, True) Then
            MsgBox "Falha ao cancelar a venda!", vbCritical
        Else
            MsgBox "Venda Cancelada! Estoque Reposto! Lançamento em Caixa Concluído!", vbInformation
        End If
        
        strSQL = "update tab_saidas set serie_sai='CANC' where num_ven='" & sNumVen & "' "
        strSQL = strSQL & " and nota_Sai='" & strNota_Sai & "'"
        Banco.Execute strSQL
    
        
    
    
    
    End If
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
    F_CFOP = frmCFOP.Tag
    Unload frmCFOP
End Function
Sub resizeForm(obj As Form)
    obj.Height = mdiFrm.ScaleHeight - mdiFrm.StatusBar1.Height - 50
    obj.Width = mdiFrm.ScaleWidth - 100
    obj.Top = 50
    obj.Left = 50
End Sub
Sub P_VisualizaContempla(cod_cli As String, seq_pre As String)
    frmVisualizaContempla.lblNomeCli.Tag = cod_cli
    frmVisualizaContempla.Tag = seq_pre
    frmVisualizaContempla.Show 1
        
End Sub

Sub P_cancelanotaFechada(seq_Not As String, Optional tipo As String = "F")
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sTab As String
    Dim sTabIte As String
    Dim sEst As String
    If tipo = "F" Then
        sTab = "tab_compras"
        sTabIte = "tab_itcompras"
        sEst = "er_pro"
    Else
        sTab = "tab_cmp"
        sTabIte = "tab_itcmp"
        sEst = "ef_pro"
    End If
    
    strSQL = "select ti.* from " & sTab & " as tc, " & sTabIte & " as ti "
    strSQL = strSQL & " where ti.seq_not=tc.seq_not "
    strSQL = strSQL & " and ti.seq_not='" & seq_Not & "'"
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_produtos set " & sEst & "=" & sEst & "-" & !qtd_ite
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Banco.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update " & sTab & " set status_not=5 where seq_not='" & seq_Not & "' "
    
    Banco.Execute strSQL
    
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
    Banco.Execute strSQL
    F_GravaClass = True
    Exit Function
trata_Erro:
    F_GravaClass = False
End Function
Function F_SelectPlanoContas(expr As String, bCred As Boolean) As planoContas
    frmSelect.sSelect = "select cod_pla as cod,nome_pla as nome,tipo_pla as tipo from tab_planocontas"
    frmSelect.cCriterio.Add " where nome_pla like '" & expr & "%' and tipo_pla='" & IIf(bCred, "C", "D") & "' "
    
    frmSelect.sOrd = " order by nome_pla "
    
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    If frmSelect.Tag <> "" Then
        frmSelect.msflxgrd.Col = 0
        
        F_SelectPlanoContas.cod_pla = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.Col = 1
        
        F_SelectPlanoContas.nome_pla = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.Col = 2
        
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


