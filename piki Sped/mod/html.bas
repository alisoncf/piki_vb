Attribute VB_Name = "html"
Sub AbreTable(arq As Integer, Optional script As String = "", Optional largura As String = "100%")
    Print #1, "<table " & script & " width=" & largura & "> "
End Sub
Sub add_figura(arq As Integer, caminho As String, Optional wid As String = "", Optional height As String = "", _
    Optional script As String = "")
    Dim s As String
    s = "<img src=" & aspas & caminho & aspas
    If wid <> "" Then
        s = s & "width=" & aspas & wid & aspas
    End If
    If height <> "" Then
        s = s & "height=" & aspas & height & aspas
    End If
    Print #1, s; ">"
    
End Sub

Sub P_HtmlAbreTable(arq As Integer, Optional border As Integer = 1, _
    Optional wid As String = "-1")
    Dim strSQL As String
    If wid <> "-1" Then
        strSQL = "width=" & wid
    Else
        wid = ""
    End If
    Print #1, "<table border=" & border & " width=" & wid; " cellpadding=0 cellspacing=0 style=border-collapse: collapse; BorderColor = '#111111' > "
End Sub
Sub P_HtmlFechaTable(Optional arq As Integer = 1)
    Print #1, "</table >"
End Sub
Sub FechaTable(Optional arq As Integer = 1)
    P_HtmlFechaTable arq
End Sub

Sub P_HtmlAbreLinha(Optional arq As Integer = 1, Optional bkcolor = "")
    
    Print #1, "<tr>"
End Sub
Sub AbreLinha(Optional arq As Integer = 1, Optional script = "")
    
    Print #1, "<tr " & script & " >"
End Sub
Sub FechaLinha(Optional arq As Integer = 1)
    
    Print #arq, "</tr>"
End Sub

Sub AbreColuna(Optional arq As Integer = 1, Optional script = "")
    Print #arq, "<td " & script & ">"
    
End Sub
Sub AbreP(Optional arq As Integer = 1, Optional conteudo As String, Optional script = "", Optional fecha As Boolean = True)
    
    Print #arq, "<p " & script & ">" & conteudo; IIf(fecha, " </p>", "")
End Sub
Sub FechaP(Optional arq As Integer = 1)
    Print #arq, "</p>"
End Sub

Sub AbreBody(Optional arq As Integer = 1, Optional script = "")
    
    Print #1, "<body " & script & ">"
End Sub
Sub FechaBody(Optional arq As Integer = 1, Optional script = "")
    
    Print #1, "</body>"
End Sub

Sub AbreDiv(Optional arq As Integer = 1, Optional class = "")
    
    Print #1, "<div class=" & aspas & class & aspas & ">"
End Sub

Sub FechaDiv(Optional arq As Integer = 1, Optional class = "")
    
    Print #1, "</div>"
End Sub

Sub FechaColuna(Optional arq As Integer = 1)
    
    Print #1, "</td>"
End Sub
Sub AddColuna(Arquivo As Integer, Texto As String, Optional script As String = "")
    AbreColuna Arquivo, script
    Print #Arquivo, Texto
    FechaColuna
End Sub

Sub P_HtmlFechaLinha(Optional arq As Integer = 1)
    Print #1, "</tr>"
End Sub
Sub P_HtmlAbreColuna(arq As Integer, Texto As String, Optional align As String = "center", _
    Optional bFecha As Boolean = True, Optional wid As String = "-1", Optional colsPan As String = "0")
    Dim strSQL As String
    
    Print #1, "<td " & IIf(colsPan = "-1", "", "colspan=" & colsPan & " ") & IIf(wid > "0", "Width=" & wid, "") & " align='" & align & "'>" & Texto & "&nbsp" & IIf(bFecha, "</td>", "")
End Sub
Sub P_HtmlFechaColuna(Optional arq As Integer = 1)
    Print #arq, "</td>"
End Sub
Sub AbreH(tam As Byte, Texto As String, Optional bFecha As Boolean = False, Optional arq As Integer = 1)
     Print #arq, "<h" & tam & " align='center'>" & Texto & IIf(bFecha, "</h>", "")
        
End Sub
Sub P_HtmlAbreH(arq As Integer, tam As Byte, Texto As String, Optional bFecha As Boolean = False)
     Print #arq, "<H" & tam & " align='center'>" & Texto & IIf(bFecha, "</H>", "")
        
End Sub
Sub P_HtmlFechaH(Optional arq As Integer = 1)
    Print #arq, "</H>"
End Sub
Sub P_HtmlFechaTitle(Optional arq As Integer = 1)
    Print #arq, "</title>"
End Sub
Sub P_HtmlFechaBody(Optional arq As Integer = 1)
    Print #arq, "</body>"
End Sub
Sub P_HtmlAbreBody(Optional arq As Integer = 1)
    Print #arq, "<body>"
End Sub
Sub P_HtmlBR(Optional arq As Integer = 1)
    Print #arq, "<BR>"
End Sub
Sub P_HtmlAbreBlockQuote(Optional arq As Integer = 1)
    Print #arq, "<blockquote>"
End Sub
Sub P_HtmlAbreFechaBlockQuote(Optional arq As Integer = 1)
    Print #arq, "</blockquote>"
End Sub

Sub P_HtmlAbreP(Texto As String, fontsize As Integer, Optional arq As Integer = 1, Optional align As String = "left", Optional bgcolor As String = "black", _
    Optional fecha As Boolean = True)
    Print #arq, "<P font size=" & fontsize & " " & IIf(align <> "", "align='" & align & "'", "") & _
        IIf(bgcolor <> "", "bgcolor='" & bgcolor & "'", "") & ">" & Texto & _
        IIf(fecha = True, "</P>", "")
End Sub

Sub P_HtmlAbreTitlePlus(arq As Integer, title As String, Optional bFecha As Boolean = True)
    P_HtmlAbreHead arq
    Print #arq, "<title>" & title
    If bFecha Then
        Print #arq, "</title>"
    End If
    P_HtmlFechaHead arq
End Sub

Sub AbreTitle(arq As Integer, title As String, Optional bFecha As Boolean = True)
    P_HtmlAbreTitle arq, title, bFecha
End Sub

Sub P_HtmlAbreTitle(arq As Integer, title As String, Optional bFecha As Boolean = True)
    Print #arq, "<title>" & title
    If bFecha Then
        Print #arq, "</title>"
    End If
End Sub
Sub AbreHead(Optional arq As Integer = 1)
    P_HtmlAbreHead arq
End Sub
Sub P_HtmlAbreHead(Optional arq As Integer = 1)
    Print #arq, "<head>"
End Sub
Sub P_HtmlFechaHead(Optional arq As Integer = 1)
    Print #arq, "</head>"
End Sub
Sub FechaHead(Optional arq As Integer = 1)
    P_HtmlFechaHead arq
End Sub
Sub EspacoEmBrancoHtml(Optional arq As Integer = 1)
    P_HtmlEspacoEmBranco arq
End Sub

Sub P_HtmlEspacoEmBranco(Optional arq As Integer = 1)
    Print #arq, "<p>&nbsp</p>"
End Sub
Function F_SelCamposMsflxgrd(msflxgrd As MSFlexGrid) As String
    Dim iCount As Integer
    
    With msflxgrd
        .row = 0
        For iCount = 0 To .Cols - 1
            .col = iCount
            frmEscolherCamposImprHmtl.lst.AddItem .Text
            
        Next
    End With
    frmEscolherCamposImprHmtl.Show 1
End Function
Function F_SelCamposDbgrd(dbgrd As Object) As String
    Dim iCount As Integer
    On Error GoTo trata_Erro
    With dbgrd
        
        For iCount = 0 To .Columns.Count - 1
            
            frmEscolherCamposImprHmtl.lst.AddItem .Columns(iCount).caption
            
        Next
    End With
    frmEscolherCamposImprHmtl.Show 1
trata_Erro:
End Function

Sub P_MsflxToHtml(titulo As String, msflxgrd As Object, Optional H1 As String = "", _
    Optional extra_final As String = "", Optional ListColSoma As String = "")
    Dim sPath As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
    On Error Resume Next
    sPath = App.path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        fileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    strSQL = F_SelCamposMsflxgrd(msflxgrd)
        
        
    Dim c As Collection
    Set c = New Collection
    
    Set c = F_SeparaTexto(ListColSoma, ",", False)
    ReDim cTotal(c.Count) As Currency
        
        
    
    Open sPath For Output As #1
    
    Print #1, "<HTML><HEAD><TITLE>" & titulo & "</TITLE></HEAD>"
    Print #1, "<BODY>"
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlAbreH 1, 4, H1, True
    
    P_HtmlEspacoEmBranco 1
    P_HtmlAbreTable 1, 1, 700
    With msflxgrd
        For jCount = 0 To .rows - 1
            .row = jCount
            P_HtmlAbreLinha 1
            For iCount = 0 To .Cols - 1
                frmEscolherCamposImprHmtl.lst.ListIndex = iCount
                
                .col = iCount
                If frmEscolherCamposImprHmtl.lst.Selected(iCount) = True Then
                    P_HtmlAbreColuna 1, .Text
                End If
                
                
                If c.Count > 0 Then
                    i = IsInCollection(c, CStr(iCount))
                
                    If i > 0 Then
                        cTotal(i - 1) = cTotal(i - 1) + CCur(dbgrd.Columns(iCount).Text)
                    End If
                End If
                
                
            Next
            P_HtmlFechaLinha 1
        Next
    End With
        
    If c.Count > 0 Then
        P_HtmlAbreLinha 1
            For iCount = 0 To dbgrd.Columns.Count - 1
                i = IsInCollection(c, CStr(iCount))
                If i = 0 Then
                    P_HtmlAbreColuna 1, "&nbsp"
                Else
                    P_HtmlAbreColuna 1, FormatVal(cTotal(i - 1))
                End If
            Next
            
        P_HtmlFechaLinha 1
        
    End If
        
    
    P_HtmlFechaTable 1
    
    P_HtmlAbreH 1, 4, extra_final, True
    
    
    
    
    Unload frmEscolherCamposImprHmtl
    
    Print #1, "</BODY>"
    Print #1, "</HTML>"
    Close #1
    Shell "explorer " & sPath
End Sub
Function F_AbreHtml(nomeArq As String) As String
    Dim sPath As String
    sPath = App.path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        fileSystem.MkDir sPath
    End If
    nomeArq = sPath & "\" & nomeArq & ".html"
        
    
    Open nomeArq For Output As #1
    Print #1, "<html>"
    F_AbreHtml = nomeArq
End Function
Sub LinhaHorizontal(Optional arq As Integer = 1)
    Print #arq, "<hr>"
End Sub
Sub P_FechaHtml(Optional arq As Integer = 1, Optional FechaArq As Boolean = True, Optional sPath As String = "")


    Print #arq, "</Html>"
    If FechaArq Then
        Close #1
        Shell "explorer " & sPath
    End If
End Sub
Sub FechaHtml(Optional arq As Integer = 1, Optional FechaArq As Boolean = True, Optional sPath As String = "")


    Print #arq, "</html>"
    If FechaArq Then
        Close #1
        Shell "explorer " & sPath
    End If
End Sub
Sub P_DbgrdToHtml(titulo As String, dbgrd As Object, dtVinculado As Object, Optional extra As String = "", _
    Optional ListColSoma As String = "", Optional H1 As String = "")
    On Error GoTo trata_Erro
    Dim c As Collection
    Set c = New Collection
    
    Set c = F_SeparaTexto(ListColSoma, ",", False)
    ReDim cTotal(c.Count) As Currency
    
    
    
    Dim sPath As String
    Dim strSQL As String
    Dim iQtdColsSelec As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
'    On Error Resume Next
'    On Error GoTo trata_ERro
        sPath = App.path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        fileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
        
        
    strSQL = F_SelCamposDbgrd(dbgrd)
        
    P_Status "Gerando arquivo HTML..."
    Open sPath For Output As #1
    
    Print #1, "<HTML><HEAD><TITLE>" & titulo & "</TITLE></HEAD>"
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlEspacoEmBranco 1
    
    P_HtmlAbreH 1, 5, H1, True
    P_HtmlAbreH 1, 5, extra, True
    
    
    P_HtmlAbreTable 1, 1, 700
    With dbgrd
        P_HtmlAbreLinha 1
        For iCount = 0 To .Columns.Count - 1
            frmEscolherCamposImprHmtl.lst.ListIndex = iCount
            If frmEscolherCamposImprHmtl.lst.Selected(iCount) = True Then
                P_HtmlAbreColuna 1, dbgrd.Columns(iCount).caption
            End If
            
        Next
        P_HtmlFechaLinha 1
        
        
        
    End With
    'dtVinculado.Refresh
    With dtVinculado.Recordset
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            P_HtmlAbreLinha 1
            iQtdColsSelec = 0
            For iCount = 0 To dbgrd.Columns.Count - 1
                        
                dbgrd.col = iCount
                frmEscolherCamposImprHmtl.lst.ListIndex = iCount
                If frmEscolherCamposImprHmtl.lst.Selected(iCount) = True Then
                    P_HtmlAbreColuna 1, dbgrd.Text & "&nbsp", "left", True
                    iQtdColsSelec = iQtdColsSelec + 1
                End If
                If c.Count > 0 Then
                    i = IsInCollection(c, CStr(iCount))
                
                    If i > 0 Then
                        On Error Resume Next
                        cTotal(i - 1) = cTotal(i - 1) + CCur(dbgrd.Columns(iCount).Text)
                    End If
                End If
            Next
            P_HtmlAbreLinha 1
            P_Status "Escrevendo Item " & .AbsolutePosition + 1 & " de " & .RecordCount
            .MoveNext
        Loop
    End With
    
    If c.Count > 0 Then
        P_HtmlAbreLinha 1
            For iCount = 0 To dbgrd.Columns.Count - 1
                frmEscolherCamposImprHmtl.lst.ListIndex = iCount
                If frmEscolherCamposImprHmtl.lst.Selected(iCount) = True Then
                    i = IsInCollection(c, CStr(iCount))
                    If i = 0 Then
                        P_HtmlAbreColuna 1, "&nbsp"
                    Else
                        P_HtmlAbreColuna 1, FormatVal(cTotal(i - 1))
                    End If
                End If
            Next
            
        P_HtmlFechaLinha 1
        
    End If
    
    P_HtmlFechaTable 1
    
    If dtVinculado.Recordset.RecordCount > 0 Then
        dtVinculado.Recordset.MoveFirst
    End If
'    dtVinculado.Refresh
    Unload frmEscolherCamposImprHmtl
    
    
    

    
    Print #1, "</HTML>"
    Close #1
    P_Status ""
    
    Shell "explorer " & sPath
    Exit Sub
trata_Erro:
    MostraErro
    
    On Error Resume Next
    Close #1
    Unload frmEscolherCamposImprHmtl
End Sub
Sub P_RpesqToHTML1(titulo As String, rPesq As ADODB.Recordset, Optional extra1 As String = "", Optional extra2 As String = "")
    Dim sPath As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim jCount As Integer
    Dim sAux As String
    'On Error Resume Next
    sPath = App.path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        fileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
    
    Open sPath For Output As #1
    
    Print #1, "<HTML><HEAD><TITLE>" & titulo & "</TITLE></HEAD>"
    Print #1, "<BODY>"
    
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlEspacoEmBranco 1
    P_HtmlAbreH 1, 5, extra1, True
    P_HtmlAbreH 1, 5, extra2, True
    P_HtmlAbreTable 1, 0, 700
    With rPesq
        P_HtmlAbreLinha 1
        For iCount = 0 To .Fields.Count - 1
            
            P_HtmlAbreColuna 1, .Fields(iCount).Name
            
            
        Next
        P_HtmlFechaLinha 1
        
        
        
    End With
    
    With rPesq
        
        Do While Not .EOF
            P_HtmlAbreLinha 1
            For iCount = 0 To .Fields.Count - 1
                If .Fields(iCount).type = 7 Then
                    P_HtmlAbreColuna 1, FormatVal(.Fields(iCount).Value) & "&nbsp", "right", True
                    
                Else
                    P_HtmlAbreColuna 1, .Fields(iCount).Value & "&nbsp", "left", True
                End If
                
                
                
                
            Next
            P_HtmlAbreLinha 1
            
            .MoveNext
        Loop
    End With
    
    P_HtmlFechaTable 1
    
    
    
    
    
    
    
    Print #1, "<P>Impresso em " & Now & " por " & strNome_Usu & " </P> "
    Print #1, "</BODY>"
    Print #1, "</HTML>"
    
    Close #1
    Shell "explorer " & sPath
End Sub

Sub P_rPesqToHtml(titulo As String, rPesq As ADODB.Recordset, Optional extra1 As String = "", _
    Optional extra2 As String = "", Optional ListColSoma As String = "")
    Dim sPath As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim jCount As Integer
    Dim sAux As String
    'On Error Resume Next
    Dim c As Collection
    Set c = New Collection
    Dim cTotal() As Currency
    Set c = F_SeparaTexto(ListColSoma, ",", False)
    ReDim cTotal(c.Count) As Currency
    
    
    
    sPath = App.path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        fileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
    
    Open sPath For Output As #1
    
    Print #1, "<HTML><HEAD><TITLE>" & titulo & "</TITLE></HEAD>"
    Print #1, "<BODY>"
    
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlEspacoEmBranco 1
    P_HtmlAbreH 1, 5, extra1, True
    P_HtmlAbreH 1, 5, extra2, True
    P_HtmlAbreTable 1, 0, 700
    With rPesq
        P_HtmlAbreLinha 1
        For iCount = 0 To .Fields.Count - 1
            
            P_HtmlAbreColuna 1, .Fields(iCount).Name
            
            
        Next
        P_HtmlFechaLinha 1
        
        
        
    End With
    
    With rPesq
        
        Do While Not .EOF
            P_HtmlAbreLinha 1
            For iCount = 0 To .Fields.Count - 1
                If .Fields(iCount).type = 7 Then
                    P_HtmlAbreColuna 1, FormatVal(.Fields(iCount).Value) & "&nbsp", "right", True
                    
                Else
                    P_HtmlAbreColuna 1, .Fields(iCount).Value & "&nbsp", "left", True
                End If
                If c.Count > 0 Then
                    i = IsInCollection(c, CStr(iCount))
                
                    If i > 0 Then
                        cTotal(i - 1) = cTotal(i - 1) + CCur(rPesq.Fields(iCount).Value)
                    End If
                End If
                
                
                
            Next
            P_HtmlAbreLinha 1
            
            .MoveNext
        Loop
    End With
    
    P_HtmlFechaTable 1
    
    If c.Count > 0 Then
        P_HtmlAbreLinha 1
            For iCount = 0 To rPesq.RecordCount - 1
                
                    i = IsInCollection(c, CStr(iCount))
                    If i = 0 Then
                        P_HtmlAbreColuna 1, "&nbsp"
                    Else
                        P_HtmlAbreColuna 1, FormatVal(cTotal(i - 1))
                    End If
                
            Next
            
        P_HtmlFechaLinha 1
        
    End If
    
    
    
    
    
    Print #1, "<P>Impresso em " & Now & " por " & strNome_Usu & " </P> "
    Print #1, "</BODY>"
    Print #1, "</HTML>"
    
    Close #1
    Shell "explorer " & sPath
End Sub



Sub P_rPesqToTable(NumArq As Integer, rPesq As ADODB.Recordset, Optional border As Integer = 1, Optional bHead As Boolean = True)
    Dim sPath As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim jCount As Integer
    On Error Resume Next
    
    'o arquivo já deve estar aberto
    
    P_HtmlAbreTable 1, border, "700"
    If bHead Then
        With rPesq
            
            P_HtmlAbreLinha 1
            For iCount = 0 To .Fields.Count - 1
                
                P_HtmlAbreColuna 1, .Fields(iCount).Name
                
                
            Next
            P_HtmlFechaLinha 1
        End With
    End If
    With rPesq
        Do While Not .EOF
            P_HtmlAbreLinha 1
            For iCount = 0 To .Fields.Count - 1
                
                dbgrd.col = iCount
                P_HtmlAbreColuna 1, .Fields(iCount).Value & "&nbsp", "left", True
                
            Next
            P_HtmlAbreLinha 1
            
            .MoveNext
        Loop
    End With
    P_HtmlFechaTable 1
    
End Sub
Sub P_LinhaEmBranco(Optional arq As Integer = 1, Optional qtd As Integer = 1)
    Dim i As Integer
    For i = 1 To qtd
        Print #arq, "<br>"
    Next
End Sub

Sub P_DtgrdToHtml(titulo As String, dbgrd As Object, dtVinculado As Object, Optional extra As String = "", _
    Optional ListColSoma As String = "", Optional H1 As String = "")
    
    Dim c As Collection
    Set c = New Collection
    
    Set c = F_SeparaTexto(ListColSoma, ",", False)
    ReDim cTotal(c.Count) As Currency
    
    
    dbgrd.Visible = False
    Dim sPath As String
    Dim strSQL As String
    Dim iQtdColsSelec As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
'    On Error Resume Next
'    On Error GoTo trata_ERro
        sPath = App.path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        fileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
        
        
    strSQL = F_SelCamposDbgrd(dbgrd)
        
    P_Status "Gerando arquivo HTML..."
    Open sPath For Output As #1
    
    Print #1, "<HTML><HEAD><TITLE>" & titulo & "</TITLE></HEAD>"
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlEspacoEmBranco 1
    
    P_HtmlAbreH 1, 5, H1, True
    P_HtmlAbreH 1, 5, extra, True
    
    
    P_HtmlAbreTable 1, 1, 700
    With dbgrd
        P_HtmlAbreLinha 1
        If strSQL <> "" Then
            For iCount = 0 To .Columns.Count - 1
                frmEscolherCamposImprHmtl.lst.ListIndex = iCount
                If frmEscolherCamposImprHmtl.lst.Selected(iCount) = True Then
                    
                End If
                
            Next
        Else
            P_HtmlAbreColuna 1, dbgrd.Columns(iCount).caption
        End If
        P_HtmlFechaLinha 1
        
        
        
    End With
    'dtVinculado.Refresh
    With dtVinculado
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            P_HtmlAbreLinha 1
            iQtdColsSelec = 0
            For iCount = 0 To dbgrd.Columns.Count - 1
                        
                dbgrd.col = iCount
                frmEscolherCamposImprHmtl.lst.ListIndex = iCount
                If frmEscolherCamposImprHmtl.lst.Selected(iCount) = True Then
                    P_HtmlAbreColuna 1, dbgrd.Text & "&nbsp", "left", True
                    iQtdColsSelec = iQtdColsSelec + 1
                End If
                If c.Count > 0 Then
                    i = IsInCollection(c, CStr(iCount))
                
                    If i > 0 Then
                        On Error Resume Next
                        cTotal(i - 1) = cTotal(i - 1) + CCur(dbgrd.Columns(iCount).Text)
                    End If
                End If
            Next
            P_HtmlAbreLinha 1
            P_Status "Escrevendo Item " & .AbsolutePosition + 1 & " de " & .RecordCount
            .MoveNext
        Loop
    End With
    
    If c.Count > 0 Then
        P_HtmlAbreLinha 1
            For iCount = 0 To dbgrd.Columns.Count - 1
                frmEscolherCamposImprHmtl.lst.ListIndex = iCount
                If frmEscolherCamposImprHmtl.lst.Selected(iCount) = True Then
                    i = IsInCollection(c, CStr(iCount))
                    If i = 0 Then
                        P_HtmlAbreColuna 1, "&nbsp"
                    Else
                        P_HtmlAbreColuna 1, FormatVal(cTotal(i - 1))
                    End If
                End If
            Next
            
        P_HtmlFechaLinha 1
        
    End If
    
    P_HtmlFechaTable 1
    
    If dtVinculado.RecordCount > 0 Then
        dtVinculado.MoveFirst
    End If
'    dtVinculado.Refresh
    Unload frmEscolherCamposImprHmtl
    
    dbgrd.Visible = True
    
    
    
    Print #1, "</HTML>"
    Close #1
    P_Status ""
    
    Shell "explorer " & sPath
    Exit Sub
trata_Erro:
    P_MostraErro
    dbgrd.Visible = True
    On Error Resume Next
    Close #1
    Unload frmEscolherCamposImprHmtl
End Sub


Sub P_AdoToHtml(titulo As String, listaCab As String, listaCampos As String, _
    dtVinculado As ADODB.Recordset, Optional extra As String = "", _
    Optional ListColSoma As String = "", Optional H1 As String = "")
    
    Dim c As Collection
    Set c = New Collection
    
    Set c = F_SeparaTexto(ListColSoma, ",", False)
    ReDim cTotal(c.Count) As Currency
    
    Dim cab As New Collection
    Dim campos As New Collection
    
    
    Set cab = F_SeparaTexto(listaCab, ",", False)
    Set campos = F_SeparaTexto(listaCampos, ",", False)
    
    
    Dim sNomeArq As String
    Dim strSQL As String
    Dim sNomeRep As String
    Dim rPesq As New ADODB.Recordset
    Dim iDifD As Integer
    Dim bOk As Boolean
    Dim i As Integer
    Dim j As Integer
    
    
    sNomeArq = "rel" & F_SoNumero(Now) & ".html"
    sNomeArq = F_AbreHtml(sNomeArq)
    P_HtmlAbreTitlePlus 1, titulo
    
        AbreBody
    
        AbreH 3, titulo, True
    
        AbreTable 1
        AbreLinha
        For i = 1 To cab.Count
            AddColuna 1, cab(i)
        Next
        FechaLinha
    With ado
        Do While Not .EOF
            AbreLinha
            For i = 1 To campos.Count
                For j = 0 To ado.Columns - 1
                    
                Next
            Next
            .MoveNext
            FechaLinha
        Loop
    End With
    
    
    
    FechaBody
    
    P_FechaHtml 1, True, sNomeArq
    
End Sub


