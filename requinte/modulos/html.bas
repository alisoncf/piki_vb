Attribute VB_Name = "html"


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
Function F_HtmlAbreTable(arq As Integer, Optional border As Integer = 1, _
    Optional wid As String = "-1") As String
    Dim strSQL As String
    If wid <> "-1" Then
        strSQL = "width=" & wid
    Else
        wid = ""
    End If
    F_HtmlAbreTable = "<table border=" & border & " width=" & wid & "  cellpadding=0 cellspacing=0 style=border-collapse: collapse; BorderColor = '#111111' > "
End Function

Sub P_HtmlFechaTable(Optional arq As Integer = 1)
    Print #1, "</table >"
End Sub
Function F_HtmlFechaTable(Optional arq As Integer = 1) As String
    F_HtmlFechaTable = "</table >"
    
End Function

Sub P_HtmlAbreLinha(Optional arq As Integer = 1, Optional bkcolor = "")
    
    Print #1, "<tr  >"
End Sub
Function F_HtmlAbreLinha(Optional arq As Integer = 1, Optional bkcolor = "")
    
    F_HtmlAbreLinha = "<tr>"
End Function

Sub P_HtmlFechaLinha(Optional arq As Integer = 1)
    Print #1, "</tr>"
End Sub
Function F_HtmlFechaLinha(Optional arq As Integer = 1) As String
    F_HtmlFechaLinha = "</tr>"
End Function

Sub P_HtmlAbreColuna(arq As Integer, Texto As String, Optional align As String = "center", _
    Optional bFecha As Boolean = True, Optional wid As String = "-1", Optional colsPan As String = "0", Optional height As String = "-1")
    Dim strSQL As String
    
    Print #1, "<td " & IIf(colsPan = "-1", "", "colspan=" & colsPan & " ") & IIf(wid > "0", "Width=" & wid, "") & IIf(height > "-1", " height=" & height, "") & " align='" & align & "'>" & Texto & "&nbsp" & IIf(bFecha, "</td>", "")
End Sub
Function F_HtmlAbreColuna(arq As Integer, Texto As String, Optional align As String = "center", _
    Optional bFecha As Boolean = True, Optional wid As String = "-1", Optional colsPan As String = "0", Optional height As String = "-1") As String
    Dim strSQL As String
    
    F_HtmlAbreColuna = "<td " & IIf(colsPan = "-1", "", "colspan=" & colsPan & " ") & IIf(wid > "0", "Width=" & wid, "") & IIf(height > "-1", " height=" & height, "") & " align='" & align & "'>" & Texto & "&nbsp" & IIf(bFecha, "</td>", "")
End Function
Sub P_HtmlFechaColuna(Optional arq As Integer = 1)
    Print #arq, "</td>"
End Sub
Function F_HtmlFechaColuna(Optional arq As Integer = 1) As String
    F_HtmlFechaColuna = "</td>"
End Function

Sub P_HtmlAbreH(arq As Integer, Tam As Byte, Texto As String, Optional bFecha As Boolean = False)
     Print #arq, "<H" & Tam & " align='center'>" & Texto & IIf(bFecha, "</H>", "")
        
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
Sub P_HtmlAbreBody(Optional arq As Integer = 1, Optional Parametros As String = "")
    Print #arq, "<body " & Parametros & ">"
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
    Dim style As String
    style = aspas & "style" & aspas
    Print #arq, "<P font size=" & fontsize & " " & IIf(align <> "", "align='" & align & "'", "") & _
        IIf(bgcolor <> "", "bgcolor='" & bgcolor & "'", "") & ">" & Texto & _
        IIf(fecha = True, "</P>", "")
End Sub

Sub Paragrafo(arq As Integer, Texto As String, Optional fecha As Boolean = True, _
    Optional align As String = "justify", Optional size As String = "12", _
    Optional font As String = "Times", Optional color As String = "black", _
    Optional fontStyle As String = "bold", Optional textDecoration = "normal", Optional style As String = "")
    
    If style = "" Then
        style = "style=" & aspas & "font: " & size & "px " & font & ", sans-serif;text-align: " & _
        align & ";color: " & color & ";font-weight: " & fontStyle & ";text-decoration: " & textDecoration & ";" & aspas
            
    End If
    Print #arq, "<p " & style & ">" & Texto
        If fecha = True Then
            Print #arq, "</p>"
        End If
End Sub
Sub P_HtmlAbreTitlePlus(arq As Integer, title As String, Optional bFecha As Boolean = True)
    P_HtmlAbreHead arq
    Print #arq, "<title>" & title
    If bFecha Then
        Print #arq, "</title>"
    End If
    P_HtmlFechaHead arq
End Sub

Sub IMG(arq As Integer, scr As String, Optional script As String)
    ' style="width:500px;height:600px;
    Print #arq, "<img src=" & aspas & scr & aspas & " " & script & "/>"
    
    
    
End Sub


Sub P_HtmlAbreTitle(arq As Integer, title As String, Optional bFecha As Boolean = True)
    Print #arq, "<title>" & title
    If bFecha Then
        Print #arq, "</title>"
    End If
End Sub
Sub P_HtmlAbreHead(Optional arq As Integer = 1)
    Print #arq, "<head>"
End Sub
Sub P_HtmlAbreCenter(Optional arq As Integer = 1, Optional Texto As String = "")
    Print #arq, "<center>" & Texto
End Sub
Sub P_HtmlFechaCenter(Optional arq As Integer = 1)
    Print #arq, "</center>"
End Sub
Sub P_HtmlFechaHead(Optional arq As Integer = 1)
    Print #arq, "</head>"
End Sub


Sub P_HtmlEspacoEmBranco(Optional arq As Integer = 1)
    Print #arq, "<P>&nbsp</P>"
End Sub
Function F_SelCamposMsflxgrd(msflxgrd As MSFlexGrid) As String
    Dim iCount As Integer
    
    With msflxgrd
        .row = 0
        For iCount = 0 To .Cols - 1
            .col = iCount
            frmEscolherCamposImprHmtl.lst.AddItem .text
            
        Next
    End With
    frmEscolherCamposImprHmtl.Show 1
End Function
Function F_SelCamposDbgrd(dbgrd As Object, Optional listaPreSelecionados As String = "") As String
    Dim iCount As Integer
    
    With dbgrd
        
        For iCount = 0 To .Columns.Count - 1
            
            frmEscolherCamposImprHmtl.lst.AddItem .Columns(iCount).caption
            If listaPreSelecionados <> "" Then
                If InStr(listaPreSelecionados, CStr(frmEscolherCamposImprHmtl.lst.ListCount - 1)) > 0 Then
                    frmEscolherCamposImprHmtl.lst.Selected(frmEscolherCamposImprHmtl.lst.ListCount - 1) = True
                Else
                    frmEscolherCamposImprHmtl.lst.Selected(frmEscolherCamposImprHmtl.lst.ListCount - 1) = False
                End If
            Else
                frmEscolherCamposImprHmtl.lst.Selected(frmEscolherCamposImprHmtl.lst.ListCount - 1) = True
            End If
        Next
    End With
    frmEscolherCamposImprHmtl.Show 1
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
        FileSystem.MkDir sPath
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
                    P_HtmlAbreColuna 1, .text
                End If
                
                
                If c.Count > 0 Then
                    i = IsInCollection(c, CStr(iCount))
                
                    If i > 0 Then
                        cTotal(i - 1) = cTotal(i - 1) + CCur(dbgrd.Columns(iCount).text)
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
Function F_AbreHtml(nomeArq As String, Optional titulo = "", Optional meta = "") As String
    Dim sPath As String
    sPath = App.path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        FileSystem.MkDir sPath
    End If
    nomeArq = sPath & "\" & nomeArq & ".html"
        
    FechaArquivoExterno 1
    
    Open nomeArq For Output As #1
    Print #1, "<!DOCTYPE html>"
    Print #1, "<html>"
    If titulo <> "" Then
        Print #1, "<head>"
        Print #1, "<title>" & titulo & "</title>"
        
        Print #1, meta
        
        Print #1, "</head>"
    End If
    
    
    F_AbreHtml = nomeArq
End Function
Sub GeraHTMLAdo()

End Sub
Sub P_FechaHtml(Optional arq As Integer = 1, Optional FechaArq As Boolean = True, Optional sPath As String = "")


    Print #arq, "</Html>"
    If FechaArq Then
        Close #1
        Shell "explorer " & sPath
    End If
End Sub
Sub P_DbgrdToHtml(titulo As String, dbgrd As Object, dtVinculado As Object, Optional extra As String = "", _
    Optional ListColSoma As String = "", Optional H1 As String = "", Optional widthTable As String = "-1")
    
    Dim c As Collection
    Set c = New Collection
    
    Set c = F_SeparaTexto(ListColSoma, ",", False)
    ReDim cTotal(c.Count) As Currency
    
    
    Dim obj As Object
    If TypeName(dtVinculado) = "Data" Then
        Set obj = dtVinculado.Recordset
    ElseIf TypeName(dtVinculado) = "Adodc" Then
        Set obj = dtVinculado.Recordset
    Else
        Set obj = dtVinculado
    End If
    
    
    Dim sPath As String
    Dim strSQL As String
    Dim iQtdColsSelec As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
'    On Error Resume Next
    On Error GoTo trata_Erro
        sPath = App.path & "\html"
    If Dir(sPath, vbDirectory) = "" Then
        FileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
        
        
    strSQL = F_SelCamposDbgrd(dbgrd, "0,1,2")
        
    P_Status "Gerando arquivo HTML..."
    Open sPath For Output As #1
    
    Print #1, "<html><head><title>" & titulo & "</title></head>"
    P_HtmlAbreBody
    
    P_HtmlAbreH 1, 4, titulo, True
    P_HtmlEspacoEmBranco 1
    
    P_HtmlAbreH 1, 5, H1, True
    P_HtmlAbreH 1, 5, extra, True
    
    
    P_HtmlAbreTable 1, 1, widthTable
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
    With obj
        
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
                    If TypeName(obj) = "Recordset" Then
                        P_HtmlAbreColuna 1, dbgrd.Columns(iCount).text & "&nbsp", "left", True
                    Else
                        P_HtmlAbreColuna 1, dtVinculado(dbgrd.Columns(iCount).DataField) & "&nbsp", "left", True
                    End If
                    iQtdColsSelec = iQtdColsSelec + 1
                End If
                If c.Count > 0 Then
                    i = IsInCollection(c, CStr(iCount))
                
                    If i > 0 Then
                        On Error Resume Next
                        cTotal(i - 1) = cTotal(i - 1) + CCur(dbgrd.Columns(iCount).text)
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
    
    If TypeName(dtVinculado) = "Data" Or TypeName(dtVinculado) = "Adodc" Then
        
        If dtVinculado.Recordset.RecordCount > 0 Then
            dtVinculado.Recordset.MoveFirst
        End If
    Else
        If dtVinculado.RecordCount > 0 Then
            dtVinculado.MoveFirst
        End If
    End If
'    dtVinculado.Refresh
    Unload frmEscolherCamposImprHmtl
    
    
    
    
    P_HtmlFechaBody
    Print #1, "</html>"
    Close #1
    P_Status ""
    
    Shell "explorer " & sPath
    Exit Sub
trata_Erro:
    P_MostraErro
    
    On Error Resume Next
    Close #1
    Unload frmEscolherCamposImprHmtl
End Sub
Function F_DbgrdToTable(titulo As String, dbgrd As Object, dtVinculado As Object, _
    Optional ListColSoma As String = "", Optional tamTab As String = "700")
    
    Dim c As Collection
    Set c = New Collection
    
    Set c = F_SeparaTexto(ListColSoma, ",", False)
    ReDim cTotal(c.Count) As Currency
    
    
    Dim obj As Object
    If TypeName(dtVinculado) = "Data" Then
        Set obj = dtVinculado.Recordset
    ElseIf TypeName(dtVinculado) = "Adodc" Then
        Set obj = dtVinculado.Recordset
    Else
        Set obj = dtVinculado
    End If
    
    
    Dim sPath As String
    Dim strSQL As String
    Dim iQtdColsSelec As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
'    On Error Resume Next
'    On Error GoTo trata_ERro
    
        
    
    
    strSQL = F_HtmlAbreTable(1, 1, tamTab)
    With dbgrd
        
        strSQL = strSQL & Chr(13) & F_HtmlAbreLinha(1)
        
        For iCount = 0 To .Columns.Count - 1
            
            strSQL = strSQL & Chr(13) & F_HtmlAbreColuna(1, dbgrd.Columns(iCount).caption)
            

        Next
        strSQL = stsrql & F_HtmlFechaLinha
        
        
        
    End With
    'dtVinculado.Refresh
    With obj
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            strSQL = strSQL & Chr(13) & F_HtmlAbreLinha
            iQtdColsSelec = 0
            For iCount = 0 To dbgrd.Columns.Count - 1
                        
                dbgrd.col = iCount
                
                    If TypeName(obj) = "Recordset" Then
                        strSQL = strSQL & Chr(13) & F_HtmlAbreColuna(1, dbgrd.Columns(iCount).text & "&nbsp", "left", True)
                    Else
                        strSQL = strSQL & Chr(13) & F_HtmlAbreColuna(1, dtVinculado(dbgrd.Columns(iCount).DataField) & "&nbsp", "left", True)
                    End If
                    iQtdColsSelec = iQtdColsSelec + 1
                
                If c.Count > 0 Then
                    i = IsInCollection(c, CStr(iCount))
                
                    If i > 0 Then
                        On Error Resume Next
                        cTotal(i - 1) = cTotal(i - 1) + CCur(dbgrd.Columns(iCount).text)
                    End If
                End If
            Next
            strSQL = strSQL & Chr(13) & F_HtmlAbreLinha
            
            .MoveNext
        Loop
    End With
    
    If c.Count > 0 Then
        strSQL = strSQL & Chr(13) & F_HtmlAbreLinha
            For iCount = 0 To dbgrd.Columns.Count - 1
                
                    i = IsInCollection(c, CStr(iCount))
                    If i = 0 Then
                        strSQL = strSQL & Chr(13) & F_HtmlAbreColuna(1, "&nbsp")
                    Else
                        strSQL = strSQL & Chr(13) & F_HtmlAbreColuna(1, FormatVal(cTotal(i - 1)))
                    End If
                
            Next
            
        strSQL = strSQL & Chr(13) & F_HtmlFechaLinha
        
    End If
    
    strSQL = strSQL & Chr(13) & F_HtmlFechaTable
    
    If TypeName(dtVinculado) = "Data" Or TypeName(dtVinculado) = "Adodc" Then
        
        If dtVinculado.Recordset.RecordCount > 0 Then
            dtVinculado.Recordset.MoveFirst
        End If
    Else
        If dtVinculado.RecordCount > 0 Then
            dtVinculado.MoveFirst
        End If
    End If
'    dtVinculado.Refresh
    Unload frmEscolherCamposImprHmtl
    
    
    
    F_DbgrdToTable = strSQL
    
    
    
    P_Status ""
    
    
    Exit Function
trata_Erro:
    P_MostraErro
    
    On Error Resume Next
    
    Unload frmEscolherCamposImprHmtl
End Function

Sub P_DtgrdToHtml(titulo As String, dbgrd As Object, dtVinculado As Object, Optional extra As String = "", _
    Optional ListColSoma As String = "", Optional H1 As String = "", Optional FimdeTexto As String = "")
    
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
        FileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel.htm"
    
        
        
    strSQL = F_SelCamposDbgrd(dbgrd)
    'If strSQL = "" Then
    '    Exit Sub
    'End If
        
        
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
                    P_HtmlAbreColuna 1, dbgrd.text & "&nbsp", "left", True
                    iQtdColsSelec = iQtdColsSelec + 1
                End If
                If c.Count > 0 Then
                    i = IsInCollection(c, CStr(iCount))
                
                    If i > 0 Then
                        On Error Resume Next
                        cTotal(i - 1) = cTotal(i - 1) + CCur(dbgrd.Columns(iCount).text)
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
    
    
    
    If FimdeTexto <> "" Then
        Print #1, "<P>"
        Print #1, FimdeTexto
        Print #1, "</P>"
    End If
    
    
    Print #1, "</BODY>"
    Print #1, "</HTML>"
    Close #1
    P_Status ""
    
    Shell "explorer " & sPath

Exit Sub
trata_Erro:
    P_MostraErro
    
    On Error Resume Next
    Close #1
    Unload frmEscolherCamposImprHmtl
End Sub


Sub P_AdoToXML(ado As Object, operacao As String)
    
    Dim cabecalho As String
    
    
    Dim sPath As String
    Dim strSQL As String
    Dim iQtdColsSelec As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
'    On Error Resume Next
'    On Error GoTo trata_ERro
        sPath = App.path & "\xml"
    If Dir(sPath, vbDirectory) = "" Then
        FileSystem.MkDir sPath
    End If
    sPath = sPath & "\rel" & F_SoNumero(Now) & ".xml"
    
    cabecalho = "?xml version=" & """1.0"""
        
    
        
    P_Status "Gerando arquivo XML..."
    Open sPath For Output As #1
    
    Print #1, cabecalho
    
    
    'dtVinculado.Refresh
    With dtVinculado
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            P_HtmlAbreLinha 1
            iQtdColsSelec = 0
         Loop
    Close #1
    P_Status ""
    End With
    Shell "explorer " & sPath
    Exit Sub
trata_Erro:
    P_MostraErro
    
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
        FileSystem.MkDir sPath
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

Sub P_rPesqToHtml(titulo As String, rPesq As Object, Optional extra1 As String = "", _
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
        FileSystem.MkDir sPath
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
