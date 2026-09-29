Attribute VB_Name = "Grid"
Sub P_ajustaGridWid(Msflxgrd As Object, Optional iPrimLinha As Boolean = False)
    Dim iCount As Integer
    Dim iRow As Integer
    With Msflxgrd
        Msflxgrd.Visible = False
        For iRow = 0 To .rows - 1
            .row = iRow
            For iCount = 0 To .Cols - 1
                
                .col = iCount
                If .CellWidth < Printer.TextWidth(.Text) Then
                    .ColWidth(.col) = Printer.TextWidth(.Text)
                End If
            Next
        Next
        Msflxgrd.Visible = True
    End With
End Sub

Sub ZebraGrid(msf As Object, linhaI As Integer, _
    Optional cor1 = &H80000005, Optional cor2 = &HB8D6D8)
    Dim iCol As Integer
    Dim iRow As Integer
    msf.Visible = False
    With msf
        For iRow = linhaI To .rows - 1
            .row = iRow
            For iCol = 0 To .Cols - 1
                .col = iCol
                If iRow Mod 2 = 0 Then
                    .CellBackColor = cor2
                Else
                    .CellBackColor = cor1
                End If
            Next
        Next
    
        msf.Visible = True
    End With
    
End Sub
Sub ZebraGridPedaco(msf As MSFlexGrid, linhaI As Integer, linhaF As Integer, _
    Optional cor1 = &H80000005, Optional cor2 = &HB8D6D8)
    Dim iCol As Integer
    Dim iRow As Integer
    On Error GoTo trata_Erro
'    On Error Resume Next
    msf.Visible = False
    With msf
        For iRow = linhaI To linhaF
            .row = iRow
            For iCol = 0 To .Cols - 1
                .col = iCol
                If iRow Mod 2 = 0 Then
                    .CellBackColor = cor2
                Else
                    .CellBackColor = cor1
                End If
            Next
        Next
    
        msf.Visible = True
    End With
    Exit Sub
trata_Erro:
msf.Visible = True
End Sub
Sub ColoreLinhaGridKey(grid As Object, keysearch As String, colsearch As Integer, fonte As fonte)
    Dim i As Integer
    
    For i = 1 To grid.rows - 1
        grid.row = i
        grid.col = colsearch
        If UCase(grid.Text) = UCase(keysearch) Then
            ColoreLinhaGridFonte grid, i, fonte
        End If
    Next
    

End Sub
Sub ColoreLinhaGridFonte(grid As Object, Linha As Integer, font As fonte)
    Dim i As Integer
    With grid
        .row = Linha
        For i = 0 To .Cols - 1
            .col = i
            If font.cor_fundo > 0 Then
                .CellBackColor = font.cor_fundo
            End If
            If font.cor_face > 0 Then
                .CellForeColor = font.cor_face
            End If
            If font.nome <> "" Then
                .CellFontName = font.nome
            End If
            If font.tamanho > 0 Then
                .CellFontSize = font.tamanho
            End If
            .CellFontItalic = font.Italico
            .CellFontBold = font.negrito
            .CellFontUnderline = font.Sublinhado
        Next
    End With
End Sub
Sub ColoreLinhaGrid(grid As Object, Linha As Integer, Optional cor = &HD3E7EB, _
    Optional negrito As Boolean = False, Optional foreCor = "")



    On Error Resume Next
    Dim iCount As Integer
    Dim iRow As Integer
    With grid
        .row = Linha
        For iCount = 0 To .Cols - 1
            .col = iCount
            .CellBackColor = cor
            .CellFontBold = negrito
            If foreCor <> "" Then
                .CellForeColor = foreCor
            End If
        Next
        
    End With
End Sub


Sub ZebraLinha(msf As Object, Linha As Integer, _
    Optional cor1 = "&HFFFF", Optional corFonte = vbBlack, Optional negrito As Boolean = False)
    Dim iCol As Integer
    Dim iRow As Integer
    
    With msf
        
            .row = Linha
            For iCol = 0 To .Cols - 1
                .col = iCol
                
                    .CellBackColor = cor1
                    .CellForeColor = corFonte
                    .CellFontBold = negrito
            Next
        
    End With
    
End Sub

Sub TotalizaGrid(grid As MSFlexGrid, rowIni As Integer, rowFin As Integer, _
    ColIni As Integer, colFin As Integer, _
    Optional bTotalizaColuna As Boolean = True)
    
    Dim i As Integer
    Dim j As Integer
    Dim totalC As Double
    Dim totalR As Double
    With grid
        .Cols = .Cols + 1
        .Cols = .Cols - 1
        .row = 0
        .Text = "Total"
        For i = rowIni To rowFin - 1
            .row = i
            totalC = 0
            For j = ColIni To colFin - 1
                If .Text = "" Then .Text = "0"
                totalC = totalC + CCur(.Text)
                
                
            Next
            .col = .Cols - 1
            .Text = totalC
            
        Next
        
        
    End With
End Sub

Function buscaValorColuna(obj As Object, coluna As Integer) As String
    On Error Resume Next
    With obj
        .col = coluna
        buscaValorColuna = .Text
    End With


End Function

Function buscaValorCelula(obj As Object, coluna As Integer, Linha As Integer) As String
    On Error Resume Next
    With obj
        .col = coluna
        .row = Linha
        buscaValorCelula = .Text
    End With


End Function

Function buscaColunaHead(obj As Object, head As String, Optional Linha As Integer = 0) As Integer
    
    Dim i As Integer
    Dim r As Integer
    r = -1
    With obj
        .row = Linha
        For i = 0 To .Cols - 1
            .col = i
            If UCase(.Text) = UCase(head) Then
                r = i
                Exit For
            End If
        Next
    End With
    buscaColunaHead = r
End Function
