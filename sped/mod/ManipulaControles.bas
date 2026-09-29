Attribute VB_Name = "ManipulaControles"
Function EsticaTexto(texto As String, Optional Tam As Integer = 0, Optional align As Integer = 0, _
Optional Carac As String = " ", Optional soNum As Boolean = False) As String
    If soNum Then
        texto = F_SoNumero(texto)
    End If
    
    
    If Len(texto) >= Tam Then
        EsticaTexto = Left(texto, Tam)
    Else
        If align = 0 Then
            EsticaTexto = texto & String(Tam - Len(texto), Carac)
        ElseIf align = 1 Then
            EsticaTexto = String((Tam - Len(texto)) / 2, Carac) & texto & String((Tam - Len(texto)) / 2, Carac)
        Else
            EsticaTexto = String(Tam - Len(texto), Carac) & texto
        End If
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
Sub P_AjustaTexto(ByRef objText As TextBox, Optional opcao As Integer = 0)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    'Torna o texto maiúsculo
    If opcao = 0 Then
        strAux = UCase(objText.Text)
    ElseIf opcao = 1 Then
        strAux = StrConv(objText.Text, vbProperCase)
    ElseIf opcao = 2 Then
        strAux = LCase(objText.Text)
    Else
        strAux = (objText.Text)
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
Function F_CGC(CGC As String) As Boolean
'Esta função permite validar o CGC
'Sintaxe.: F_CGC( CGC )
'Entrada.: CGC - Nº do CGC a validar
'Saída...: Verdadeiro se exister o CGC, senão Falso

Dim M_VetNum(13), M_VetCGC(13), M_Num As String
Dim D_12, D_13, X, Y, M_D1, M_D2 As Integer
CGC = F_SoNumero(CGC)
If Val(CGC) < 1 Then
   F_CGC = False
   Exit Function
End If
M_Num = "6543298765432"
CGC = F_ZerEsq(Val(CGC), 14, 0)
For X = 1 To 13
        M_VetNum(X) = Val(Mid(M_Num, X, 1))
Next
For X = 1 To 12
        M_VetCGC(X) = Mid(CGC, X, 1)
Next
D_12 = 0
For Y = 1 To 12
        D_12 = D_12 + Val(M_VetCGC(Y)) * M_VetNum(Y + 1)
Next
M_D1 = 11 - (D_12 - Int(D_12 / 11) * 11)
If M_D1 >= 10 Then
   M_D1 = 0
End If
M_VetCGC(13) = Trim(Str(M_D1))
For Y = 1 To 13
        D_13 = D_13 + Val(M_VetCGC(Y)) * M_VetNum(Y)
Next
M_D2 = 11 - (D_13 - Int(D_13 / 11) * 11)
If M_D2 >= 10 Then
   M_D2 = 0
End If
If Right(CGC, 2) <> Trim(Str(M_D1)) & Trim(Str(M_D2)) Then
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
CPF = Left(F_ZerEsq(Val(CPF), 11, 0), 11)
ContIni = 2
ContiFim = 10
If Val(CPF) < 1 Then
   F_CPF = False
   Exit Function
End If

For ii = 1 To 2
    soma = 0
    For i = ContIni To ContiFim
        soma = soma + (Val(Mid(CPF, i - ii, 1)) * _
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
If InStr(1, Str(Num), ".") <> 0 Then
   
   Red = Trim(Str(Num))
   Aux = Mid(Red, InStr(1, Red, "."))
   
   If Len(Aux) > Dec + 1 Then
      Red = Mid(Aux, Dec + 2, 1)
      If Val(Red) > 4 Then
         
         Red = Str(1 / Val("1" + String(Dec, "0")))
         Aux = Trim(Str(Val(Aux) + Val(Red)))
         
         If Val(Aux) >= 1 Then
            Acr = "T"
            Aux = IIf(Dec > 0, "." + String(Dec, "0"), "")
         End If
      End If
   End If
   
   Aux = Left(Aux + "000000", Dec + 1)

End If
Red = Aux
If Acr <> "F" Then Num = Num + 1
Aux = IIf(Dec > 0, Trim(Str(Fix(Num))), Trim(Str(Int(Num))))
Aux = Right(String(Tam, "0") + Aux, Tam) + Red
Red = Trim(Str(Int(Num)))

If Acr <> "F" Then Num = Num - 1

F_ZerEsq = IIf(Len(Red) > Len(Aux), String(Tam, "*") _
+ IIf(Dec > 0, "." + String(Dec, "*"), ""), Aux)

End Function
Function F_AbreBanco() As Boolean
    On Error Resume Next
    Dim conexao As String
    conn.Close
    
    P_BuscaIni
    On Error GoTo Erro
    Dim strCon As String
    DBEngine.DefaultType = dbUseODBC
    strCon = "Driver={MySQL ODBC 3.51 Driver};" & _
    "Server=" & sServer & ";" & _
    "Port=3306;" & _
    "Database=" & sDataBase & ";" & _
    "Uid=" & Trim(sUid) & ";" & _
    "Pwd=" & Trim(sPwd) & ";" & _
    "Option=1 2 8 32 2048 16384;" '131072
    
    
    
    
    
    
    
    conn.Open strCon
    conn_client.Open strCon
    conn_client.CursorLocation = adUseClient
    
    strConn_Old = "requinte"
'    conexao = "ODBC;DATABASE=" & sDataBase & ";port=3306;uid=root;server=10.1.1.10"
    Set Banco = OpenDatabase(sDataBase)
 '   Set banco = OpenDatabase(sDataBase, , , conexao)
   
    F_AbreBanco = True
    Exit Function
Erro:
    


MsgBox Err.Description & Chr(13) & strCon, vbCritical, App.EXEName
    P_Fim
    F_AbreBanco = False
End Function


Sub FechaObjeto(obj As Object)
    On Error Resume Next
    obj.Close
End Sub
Sub abreObjetoAdo(obj As Object, strSQL As String, Optional dataVinculado As Object = Null)
    On Error GoTo trata_erro
    FechaObjeto obj
    obj.Open strSQL, conn_client
    
    
    
        
        Set dataVinculado.DataSource = obj
trata_erro:
End Sub

Function F_GravaDec(sValor) As String
    On Error GoTo trata_erro
    
    F_GravaDec = Str(sValor)
    Exit Function
trata_erro:
    F_GravaDec = "0"
End Function
Function F_GravaHora(strHora, Optional formato As String = "HH:MM:SS") As String
Dim strDia, strMes, StrAno As String
  strHora = Format(strHora, formato)
  
  F_GravaHora = "'" & strHora & "'"
  
  

End Function

Function F_GravaData(strData) As String
Dim strDia, strMes, StrAno As String
  strData = Format(strData, "DD/MM/YYYY")
  strDia = Mid(strData, 1, 2)
  strMes = Mid(strData, 4, 2)
  StrAno = Right(strData, 4)
  F_GravaData = "'" & StrAno & "-" & strMes & "-" & strDia & "'"
  If F_GravaData = "'--'" Then
    F_GravaData = "'0000-00-00'"
  End If
  

End Function
Function F_GravaDataTempo(strData) As String
Dim strDia, strMes, StrAno As String
  On Error GoTo trata_erro
  F_GravaDataTempo = "'" & Format(strData, "yyyy-mm-dd hh:mm:ss") & "' "
  Exit Function
trata_erro:
  F_GravaDataTempo = "'0000-00-00 00:00:00'"
End Function

Sub P_SelecTexto(objText As Object)
    'Este procedimento seleciona o texto
    'do objeto em questão
    objText.SelStart = 0
    objText.SelLength = Len(RTrim(objText.Text))
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

