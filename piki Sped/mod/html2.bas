Attribute VB_Name = "html2"
Sub P_HtmlAbreTable(arq As Integer, Optional border As Integer = 0, _
    Optional wid As String = "-1")
    Dim strSQL As String
    If wid <> "-1" Then
        strSQL = "width=" & wid
    End If
    Print #1, "<TABLE border=" & border; " cellpadding=0 cellspacing=0 style=border-collapse: collapse; BorderColor = '#111111' > "
End Sub
Sub P_HtmlFechaTable(arq As Integer)
    Print #1, "</TABLE >"
End Sub
Sub P_HtmlAbreLinha(arq As Integer)
    Print #1, "<TR>"
End Sub
Sub P_HtmlFechaLinha(arq As Integer)
    Print #1, "</TR>"
End Sub
Sub P_HtmlAbreColuna(arq As Integer, Texto As String, Optional align As String = "center", _
    Optional bFecha As Boolean = True, Optional wid As String = "-1")
    Dim strSQL As String
    
    Print #1, "<TD " & IIf(wid <> "1", wid, "") & "align='" & align & "'>" & Texto & IIf(bFecha, "</TD>", "")
End Sub
Sub P_HtmlAbreH(arq As Integer, Tam As Byte, Texto As String, Optional bFecha As Boolean = False)
     Print #arq, "<H" & Tam & " align='center'>" & Texto & IIf(bFecha, "</H>", "")
        
End Sub
Sub P_HtmlFechaH(arq As Integer)
    Print #arq, "</H>"
End Sub
Sub P_HtmlEspacoEmBranco(arq As Integer)
    Print #arq, "<P>&nbsp</P>"
End Sub

Sub P_MsflxToHtml(titulo As String, msflxgrd As MSFlexGrid)
    Dim sPath As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim jCount As Integer
    sPath = App.Path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        FileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
    
    Open sPath For Output As #1
    
    Print #1, "<HTML><HEAD><TITLE>" & titulo & "</TITLE></HEAD>"
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlEspacoEmBranco 1
    P_HtmlAbreTable 1, 0, 700
    With msflxgrd
        For jCount = 0 To .rows - 1
            .Row = jCount
            P_HtmlAbreLinha 1
            For iCount = 0 To .Cols - 1
                .Col = iCount
                P_HtmlAbreColuna 1, .Text
            Next
            P_HtmlFechaLinha 1
        Next
    End With
    
    
    P_HtmlFechaTable 1
    
    
    
    
    
    
    
    
    
    Print #1, "</HTML>"
    Close #1
    Shell "explorer " & sPath
End Sub


Sub P_DbgrdToHtml(titulo As String, dbgrd As DBGrid, dtVinculado As data, Optional border As Integer = 1)
    Dim sPath As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim jCount As Integer
    On Error Resume Next
    sPath = App.Path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        FileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
    
    Open sPath For Output As #1
    
    Print #1, "<HTML><HEAD><TITLE>" & titulo & "</TITLE></HEAD>"
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlEspacoEmBranco 1
    P_HtmlAbreTable 1, border, "700"
    With dbgrd
        P_HtmlAbreLinha 1
        For iCount = 0 To .Columns.count - 1
            
            P_HtmlAbreColuna 1, dbgrd.Columns(iCount).caption
            
            
        Next
        P_HtmlFechaLinha 1
        
        
        
    End With
    dtVinculado.Refresh
    With dtVinculado.Recordset
        
        Do While Not .EOF
            P_HtmlAbreLinha 1
            For iCount = 0 To dbgrd.Columns.count - 1
                
                dbgrd.Col = iCount
                P_HtmlAbreColuna 1, dbgrd.Text & "&nbsp", "left", True
                
            Next
            P_HtmlAbreLinha 1
            
            .MoveNext
        Loop
    End With
    
    P_HtmlFechaTable 1
    
    
    dtVinculado.Refresh
    
    
    
    
    
    
    Print #1, "</HTML>"
    Close #1
    Shell "explorer " & sPath
End Sub

Sub P_rPesqToHtml(titulo As String, rpesq As Recordset, Optional border As Integer = 1)
    Dim sPath As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim jCount As Integer
    On Error Resume Next
    sPath = App.Path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        FileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
    
    Open sPath For Output As #1
    
    Print #1, "<HTML><HEAD><TITLE>" & titulo & "</TITLE></HEAD>"
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlEspacoEmBranco 1
    P_HtmlAbreTable 1, border, "700"
    With rpesq
        P_HtmlAbreLinha 1
        For iCount = 0 To .Fields.count - 1
            
            P_HtmlAbreColuna 1, .Fields(iCount).Name
            
            
        Next
        P_HtmlFechaLinha 1
        
        
        
    End With
    
    
    With rpesq
        
        Do While Not .EOF
            P_HtmlAbreLinha 1
            For iCount = 0 To .Fields.count - 1
                
                dbgrd.Col = iCount
                P_HtmlAbreColuna 1, .Fields(iCount).Value & "&nbsp", "left", True
                
            Next
            P_HtmlAbreLinha 1
            
            .MoveNext
        Loop
    End With
    
    P_HtmlFechaTable 1
    
    

    
    
    
    
    
    
    Print #1, "</HTML>"
    Close #1
    Shell "explorer " & sPath
End Sub

Sub P_rPesqToTable(NumArq As Integer, rpesq As Recordset, Optional border As Integer = 1, Optional bHead As Boolean = True)
    Dim sPath As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim jCount As Integer
    On Error Resume Next
    
    'o arquivo já deve estar aberto
    
    P_HtmlAbreTable 1, border, "700"
    If bHead Then
        With rpesq
            
            P_HtmlAbreLinha 1
            For iCount = 0 To .Fields.count - 1
                
                P_HtmlAbreColuna 1, .Fields(iCount).Name
                
                
            Next
            P_HtmlFechaLinha 1
        End With
    End If
    With rpesq
        Do While Not .EOF
            P_HtmlAbreLinha 1
            For iCount = 0 To .Fields.count - 1
                
                dbgrd.Col = iCount
                P_HtmlAbreColuna 1, .Fields(iCount).Value & "&nbsp", "left", True
                
            Next
            P_HtmlAbreLinha 1
            
            .MoveNext
        Loop
    End With
    P_HtmlFechaTable 1
    
End Sub



