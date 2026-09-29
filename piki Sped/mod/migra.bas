Attribute VB_Name = "migra"
Sub MigraTabelaXML(tabela As String, arquivo As String)
    Dim adoR As ADODB.Recordset
    Dim strSQL As String
    Set adoR = New ADODB.Recordset
    Dim xml As String
    Dim i As Integer
    strSQL = "select * from " & tabela & " order by cod_pro "
    
    xml = "<?xml version='1.0' encoding='utf-8'?>" + Chr(13) + Chr(10)
    xml = xml + "<data>" & Now & "</data>"
    xml = xml + "<Empresa nome = '" & strNome_Emp & "' "
    xml = xml + "cnpj = '" & strCNPJ_Emp & "' </empresa> " + Chr(13) + Chr(10)
    xml = xml + "<tabela name = '" & tabela & "' </tabela>" + Chr(13) + Chr(10)
    Open arquivo For Output As #1
    Print #1, xml
    
    adoR.Open strSQL, Conn_Client
    
    With adoR
        Do While Not .EOF
            strSQL = ""
            For i = 0 To .Fields.count - 1
                
                strSQL = strSQL & "<field='" & .Fields(i).Name & "' "
                strSQL = strSQL & "type='" & TypeName(.Fields(i).Type) & "' "
                
                strSQL = strSQL & "value='" & .Fields(i).Value & "' />" + Chr(13) + Chr(10)
            Next
            
            Print #1, strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    
    
    
    Close #1
    
End Sub
Sub AtualizaDados(arquivo As String)

    Dim Linha As String
    Dim iLinha As Integer
    Dim chave As String
    Dim cont As Integer
    Dim marcador1 As Integer
    Dim marcador2 As Integer
    Dim tabela As String
    Dim valor As String
    Dim tipo As String
    Dim campos As String
    Dim iExec As Integer
    Open arquivo For Input As #1
    Dim cCampo As New Collection
    Dim cTipo As New Collection
    Dim cItem As New Collection
    Dim item As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strSQL As String
    Dim sWhere As String
    Dim sExec As String
    Dim logerro As String
    Dim tabelas As String
    On Error Resume Next
    Dim rKey As New ADODB.Recordset
    Do While Not EOF(1)
        cont = cont + 1
        If cont = 44 Then
            'MsgBox ""
        End If
        Line Input #1, Linha
                    
        If InStr(1, Linha, "<tabela>") > 0 Then
            marcador1 = InStr(1, Linha, "<tabela>")
            marcador2 = InStr(marcador1, Linha, "</tabela>")
            tabelas = Mid(Linha, marcador1 + 8, marcador2 - 9)
            
        
            
        End If
                    
        If InStr(1, Linha, "<campo>") > 0 Then
            marcador1 = InStr(1, Linha, "<campo>")
            marcador2 = InStr(marcador1, Linha, "</campo>")
            campos = Mid(Linha, marcador1 + 7, marcador2 - 8)
            Set cCampo = F_SeparaTexto(Mid(campos, 2, Len(campos) - 2), "|", False)
            
        
            
        End If
        If InStr(1, Linha, "<tipo>") > 0 Then
            marcador1 = InStr(1, Linha, "<tipo>")
            marcador2 = InStr(marcador1, Linha, "</tipo>")
            tipos = Mid(Linha, marcador1 + 6, marcador2 - 8)
            Set cTipo = F_SeparaTexto(Mid(tipos, 2, Len(tipos) - 1), "|", False)
            
        
            
        End If
        
        
        If InStr(1, Linha, "<item>") > 0 Then
            marcador1 = InStr(1, Linha, "<item>")
            marcador2 = InStr(marcador1, Linha, "</item>")
            item = Mid(Linha, marcador1 + 6, marcador2 - 7)
            Set cItem = F_SeparaTexto(Mid(item, 2, Len(item) - 2), "|", False)
            'update
        
            'update
            FechaObjeto rKey
            strSQL = "show keys from " & tabelas
            rKey.Open strSQL, Conn_Client
            
            chave = rKey!column_name
            rKey.Close
            
            sWhere = ""
            strSQL = " update " & tabelas & " set "
            For i = 1 To cItem.count
                If cCampo(i) = "TELEFONE_FOR" Then
'                    MsgBox ""
                End If
                If cCampo(i) = chave Then
                    If cItem(i) = "NULL" Then
                        sWhere = " where " & cCampo(i) & "=''"
                    Else
                        sWhere = " where " & cCampo(i) & "='" & cItem(i) & "'"
                    End If
                    
                Else
                    strSQL = strSQL & cCampo(i) & "="
                
                    If cItem(i) = "NULL" Then
                        strSQL = strSQL & "''"
                    Else
                        If cTipo(i) = "String" Then
                            
                            strSQL = strSQL & "'" & cItem(i) & "'"
                        ElseIf cTipo(i) = "Data" Then
                            strSQL = strSQL & "'" & cItem(i) & "'"
                        
                        ElseIf cTipo(i) = "Decimal" Or cTipo(i) = "Double" Or cTipo(i) = "Currency" Or cTipo(i) = "Integer" Then
                            
                            If cCampo(i) = "er_pro" Or cCampo(i) = "ef_pro" Or cCampo(i) = "r_pro" Then
                                strSQL = strSQL & cCampo(i)
                            Else
                                strSQL = strSQL & cItem(i)
                            End If
                            
                            
                        Else
                            strSQL = strSQL & "'" & cItem(i) & "'"
                        End If
                        
                    End If
                    If i < cItem.count Then
                        strSQL = strSQL & ","
                    End If
                    
                End If
                
                
                
                
            Next
            
            strSQL = strSQL & sWhere
            On Error Resume Next
            If sWhere <> "" Then
                Conn.Execute strSQL, iExec
            End If
            If iExec > 0 Then
                sExec = sExec & iExec & " records afetados em " & strSQL & Chr(13) + Chr(10)
            End If
        
            strSQL = "select count(*) as total from " & tabelas & sWhere
            
            rKey.Open strSQL, Conn_Client
            
            If rKey!total = 0 Then
                
                'insert significa que não achou
                strSQL = "insert into " & tabelas & campos & " values ("
                For i = 1 To cItem.count
                    If cCampo(i) = "ncm_pro" Then
    '                    MsgBox ""
                    End If
                    If cItem(i) = "NULL" Then
                        strSQL = strSQL & "''"
                    Else
                        If cTipo(i) = "String" Then
                            
                            strSQL = strSQL & "'" & cItem(i) & "'"
                        ElseIf cTipo(i) = "Data" Then
                            strSQL = strSQL & "'" & cItem(i) & "'"
                        Else
                            If cCampo(i) = "er_pro" Or cCampo(i) = "ef_pro" Or cCampo(i) = "r_pro" Then
                                strSQL = strSQL & "0"
                            Else
                                strSQL = strSQL & cItem(i)
                            End If
                        End If
                        
                    End If
                    If i < cItem.count Then
                        strSQL = strSQL & ","
                    End If
                    
                Next
                strSQL = strSQL & ")"
                strSQL = Replace(strSQL, "|", ",")
                
                Conn.Execute strSQL, iExec
                If iExec > 0 Then
                    sExec = sExec & iExec & " records afetados em " & strSQL & Chr(13) + Chr(10)
                End If
                rKey.Close
            End If
            
        End If
        
        
        
        P_Status "atualizando registro " & cont
        
        mdiFrm.pcStatus.Refresh
        
        
    Loop
    
    
    
    
    Open "c:\logpiki.txt" For Output As #2
        
        Print #2, "Log de atualização em " & Now
        Print #2, sExec
        Close #2
    
    
    Close #1
    
    Shell "notepad c:\logpiki.txt", vbMaximizedFocus
    
    P_Status ""

End Sub

Sub MigraTabela(tabela As String, arquivo As String)
    Dim adoR As ADODB.Recordset
    Dim strSQL As String
    Set adoR = New ADODB.Recordset
    Dim xml As String
    Dim i As Integer
    Dim tipo As String
    Dim campo As String
    Dim valor As String
    Dim cont As Integer
    Dim cTab As New Collection
    
    Set cTab = F_SeparaTexto(tabela, "|", False)
    
    
    
    xml = xml + "<data>" & Now & "</data>"
    xml = xml + "<empresa>" & strNome_Emp & "</empresa>"
    xml = xml + "<cnpj>" & strCNPJ_Emp & "</cnpj> "
    xml = xml + "<tabelas>" & tabela & "</tabelas>"
    
    Open arquivo For Output As #1
    Print #1, xml
    
    
    
    For cont = 1 To cTab.count
        Print #1, "<tabela>" & cTab(cont) & "</tabela>"
        strSQL = "select * from " & cTab(cont)
        adoR.Open strSQL, Conn_Client
        
        With adoR
            tipo = "<tipo>("
            campo = "<campo>("
            valor = "<valor>("
            For i = 0 To .Fields.count - 1
                    If .Fields(i).Type = 19 Then
                        tipo = tipo + TypeName(CInt(.Fields(i).Value))
                    Else
                        tipo = tipo + TypeName(.Fields(i).Value)
                    End If
                    campo = campo & .Fields(i).Name
                    
                        
                    
                  
                    If i <> .Fields.count - 1 Then
                        campo = campo + "|"
                        tipo = tipo + "|"
                    End If
                    
                Next
                campo = campo + ")</campo>"
                tipo = tipo + ")</tipo>"
                
                Print #1, campo
                Print #1, tipo
                
            Print #1, "<valor>" + Chr(13)
            
            Do While Not .EOF
                Print #1, "<item>";
                valor = "("
                For i = 0 To .Fields.count - 1
                    If .Fields(i).Type = 19 Then
                        If IsNull(CInt(.Fields(i).Value)) Or CInt(.Fields(i).Value) = 0 Then
                            valor = valor & "NULL"
                        Else
                            If TypeName(CInt(.Fields(i).Value)) = "Decimal" Or TypeName(CInt(.Fields(i).Value)) = "Double" Then
                                valor = valor & F_GravaDec(CInt(.Fields(i).Value))
                            Else
                                valor = valor & .Fields(i).Value
                            End If
                            
                        End If

                    Else
                
                        If IsNull(.Fields(i).Value) Or .Fields(i).Value = "" Then
                            valor = valor & "NULL"
                        Else
                            If TypeName(.Fields(i).Value) = "Decimal" Or TypeName(.Fields(i).Value) = "Double" Then
                                valor = valor & F_GravaDec(.Fields(i).Value)
                            Else
                                valor = valor & .Fields(i).Value
                            End If
                            
                        End If
                   End If
                    
                    
                    
                  
                    If i <> .Fields.count - 1 Then
                        valor = valor + "|"
                        
                    End If
                Next
                valor = valor & ")"
                Print #1, valor;
                Print #1, "</item>"
                .MoveNext
            Loop
            .Close
        End With
        
        Print #1, "</valor>"
        Print #1, "</tabela>"
        
    Next
    
    Close #1
    
End Sub
