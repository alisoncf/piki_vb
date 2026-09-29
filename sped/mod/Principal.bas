Attribute VB_Name = "Principal"
Public conn As New ADODB.Connection
Public conn_client As New ADODB.Connection
Public strConnect As String

Sub main()
    If F_AbreBanco() Then
        frmPrincipal.Show 0
    End If
End Sub
Sub P_Fim()
    On Error Resume Next
    conn.Close
    conn_client.Close
    
    End
End Sub
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

