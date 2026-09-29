Attribute VB_Name = "VariaveisGlobais"
Global strCNPJ_Emp As String
Global strFantasia_Emp As String
Global strCEP_Emp As String
Global strEndereco_Emp As String
Global strNumero_Emp As String
Global strComplemento_Emp As String
Global strBairro_Emp As String

Global strFone_Emp  As String
Global strFax_Emp As String
Global strEmail_Emp As String
Global sServer As String
Global sDriver As String
Global sUid As String
Global sPwd As String
Global sOptions As String
Global sDataBase As String

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
           strErro = strErro & "Código do erro:" & Str(Err) & Chr(13) & Chr(13)
           strErro = strErro & Err.Description & Chr(13) & Chr(13) & _
            Format(Now, "dd \de mmmm \de yyyy    hh:mm:ss")
  End Select

  MsgBox strErro, vbCritical + vbSystemModal, strProc
  P_Status ""
  
End Sub

Sub P_Status(Optional sMsg As String = "", Optional iPanel As Integer = 3, Optional bEspere As Boolean = True)
    On Error Resume Next
    Dim iCount As Integer
    mdiFrm.lblMsg.Caption = sMsg
    
    If Trim(sMsg) <> "" And bEspere Then
        For iCount = 0 To Forms.Count - 1
            Forms(iCount).MousePointer = 11
        Next
        frmMsg.lblMsg.Caption = sMsg
        frmMsg.Show 0, mdiFrm
        frmMsg.Refresh
    Else
        For iCount = 0 To Forms.Count - 1
            Forms(iCount).MousePointer = 0
        Next
        Unload frmMsg
    End If
    Exit Sub
trata_erro:
On Error Resume Next
mdiFrm.lblMsg.Caption = ""
End Sub

Sub P_PreencheImpostos(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_impostos "
    rPesq.Open strSQL, conn_client
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!cod_imp & "-" & rPesq!desc_imp
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
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

