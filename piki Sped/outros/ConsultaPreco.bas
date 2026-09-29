Attribute VB_Name = "ConsultaPreco"
Global strNomeBanco As String
Global banco As database
Sub main()
    If App.PrevInstance = True Then
        MsgBox "Este Programa já Está sendo executado!", vbExclamation
        End
    End If
    P_BuscaNomeBanco
    If Not F_AbreBanco() Then
        End
    End If
    frmConsPreco.Show 1
End Sub
Function F_AbreBanco(Optional bBase_Teste As Boolean = False) As Boolean
    
'    On Error GoTo Erro
    DBEngine.DefaultType = dbUseODBC
    'Set bTopa = OpenDatabase("topatudo", , False, "ODBC;DATABASE=topatudo;")
    
    If bBase_Teste = True Then
        Set banco = OpenDatabase("market_teste", , False, "ODBC;DATABASE=market_teste;port=3306;")
    Else
        Set banco = OpenDatabase(strNomeBanco, , False, "ODBC;DATABASE=" & strNomeBanco & ";port=3306;")
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

Sub P_Fim()
    banco.Close
    End
End Sub
Sub P_SelecTexto(objText As Object)
    'Este procedimento seleciona o texto
    'do objeto em questão
    On Error Resume Next
    objText.SelStart = 0
    objText.SelLength = Len(RTrim(objText.Text))
End Sub
Sub P_AjustaTexto(ByRef objText As Object)
    Dim strAux As String
    Dim strAux2 As String
    Dim intCont As Integer
    On Error Resume Next
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


