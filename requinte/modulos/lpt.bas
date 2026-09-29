Attribute VB_Name = "lpt"
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

Function GravaArquivoExterno1(caminho As String, conteudo) As Boolean
    
    On Error Resume Next
    Close #1
    
    Open caminho For Output As #1
    Print #1, conteudo
    
    Close #1
    
End Function
Function ImprimeEmArquivoExterno1(conteudo, buf As Integer) As Boolean
    On Error GoTo trata_Erro
    Print #buf, conteudo
    ImprimeEmArquivoExterno1 = True
    
    Exit Function
trata_Erro:
    ImprimeEmArquivoExterno1 = False
End Function
Function AbreArquivoExterno2(caminho As String, buf As Integer) As Boolean
    On Error GoTo trata_Erro
    If FechaArquivoExterno1(buf) Then
    End If
    Open caminho For Output As #buf
    AbreArquivoExterno2 = True
    Exit Function
trata_Erro:
    AbreArquivoExterno2 = False
End Function
Function FechaArquivoExterno1(buf As Integer) As Boolean
    On Error GoTo trata_Erro
    Close #buf
    FechaArquivoExterno1 = True
    Exit Function
trata_Erro:
    FechaArquivoExterno1 = False
End Function
