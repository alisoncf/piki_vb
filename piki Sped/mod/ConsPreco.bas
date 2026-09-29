Attribute VB_Name = "Module1"
Public banco As Database
Public conn As New ADODB.Connection
Global sDriver As String
Global sServer As String
Global sUid As String
Global sDataBase As String
Global sPwd As String
Global sOption As String

Sub P_BuscaIni()
    Dim sNomeArq As String
    Dim sPath As String
    Dim s As String
    sNomeArq = "piki.ini"
    
    If Dir("c:\windows\system32\" & sNomeArq) <> "" Then
        sPath = "c:\windows\system32\" & sNomeArq
    ElseIf Dir(App.Path & "\" & sNomeArq) <> "" Then
        sPath = App.Path & "\" & sNomeArq
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
    
    strConnect = strConnect & "Server=" & sServer & ";Database=" & sDataBase & ";Uid=" & sUid & ";por=3306;pwd=" & sPwd & ";option=" & sOption
    
    'strConnect = "market"
   conn.Open strConnect
   conn.CursorLocation = adUseClient
   
        Set banco = OpenDatabase(sDataBase)

   
    F_AbreBanco = True
    Exit Function
erro:
    


MsgBox Err.Description & Chr(13) & strConnect, vbCritical, conSigla_Sistema
    F_AbreBanco = False
End Function
Sub main()
    P_BuscaIni
    If F_AbreBanco() Then
        frmConsulta.Show 1
    End If
End Sub

Sub P_fim()
On Error Resume Next
    conn.Close
    banco.Close
    End
End Sub
Sub GanhaFoco(objText As Object)

    objText.SelStart = 0
    objText.SelLength = Len(RTrim(objText.Text))
End Sub

