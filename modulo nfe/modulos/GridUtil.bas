Attribute VB_Name = "GridUtil"
Type ParValor
    titulo As String
    valor As Variant
End Type


Function ColunaGrid(grid As Object, col As Integer) As Collection
    Dim c As New Collection
    Dim i As Integer
    
    With grid
        For i = 1 To .rows - 1
            .row = i
            .col = col
            If .text <> "" Then
            
                c.Add .text
            End If
            
        Next
    End With
    
    Set ColunaGrid = c
End Function
Function NotIn(c As Collection, aspas As String) As String
    
    Dim s As String
    Dim i As Integer
    For i = 1 To c.Count
        
        s = s & aspas & c(i) & aspas
        If i < c.Count Then
            s = s & ", "
        End If
        
        
    Next
    NotIn = "(" & s & ") "
    
End Function
Sub OrdenaGrid(grid As Object, ado As ADODB.Recordset)
     On Error GoTo trata_Erro
    
    
    Dim s As String
    s = grid.Columns(ColIndex).DataField
          ado.Sort = s
        
 
    
    
trata_Erro:
End Sub
Sub AjustaGridData(grid As Object, nColunas As Integer, nLinhas As Integer, nColunasFixas As Integer, _
    tituloColunasGrid As String, camposData As String, tamanhoColunas As String, _
    Optional Source As Object = Nothing, Optional ColunasTotalizar As String, Optional bZebra As Boolean, Optional CorFix = &H95FAFF, _
    Optional zebraultimaLInha As Boolean = True)
    Dim titulo As New Collection
    Dim campos As New Collection
    Dim tamanho As New Collection
    
   
   
   
    On Error GoTo trata_Erro
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    Dim bItem As Boolean
    
    If CorFix = 0 Then CorFix = &H95FAFF
    
    
    
    
    
    Set titulo = F_SeparaTexto(tituloColunasGrid, ",", False)
    Set campos = F_SeparaTexto(camposData, ",", False)
    Set tamanho = F_SeparaTexto(tamanhoColunas, ", ", False)
    
    grid.Clear
    grid.Cols = titulo.Count
    grid.rows = 2
    grid.FixedRows = 1
    grid.rows = 2
    On Error Resume Next
    grid.FixedCols = nColunasFixas
    grid.Visible = False
   
    
    grid.AllowUserResizing = flexResizeBoth
    grid.row = 0
    For i = 1 To titulo.Count
        grid.col = i - 1
        grid.text = titulo(i)
    Next
    
    For i = 1 To tamanho.Count
         grid.ColWidth(i - 1) = CInt(tamanho(i))
       
    Next
    If (Source Is Nothing) = False Then
        s = LCase(campos(1))
        If s = "$item" Or s = "@Item" Or s = "@item" Or s = "$id" Or s = "@id" Then
            bItem = True
        End If
        With Source
            
            If .RecordCount > 0 Then
               
                .MoveFirst
            End If
            
            Do While Not .EOF
                
                If grid.row > 0 Then
                    grid.AddItem ""
                End If
                If bItem = True Then
                    grid.row = grid.rows - 1
                    grid.col = 0
                    grid.text = grid.row
                End If
                For j = 1 To campos.Count
                    grid.row = 0
                    
                    grid.col = j - 1
                    s = Trim(LCase(campos(j)))
                    For i = 0 To .Fields.Count - 1
                        
                        strSQL = .Fields(i).Name
                        If LCase(Trim(.Fields(i).Name)) = s Then
                            grid.row = grid.rows - 1
                            grid.text = TiraNull(.Fields(i).Value)
                            If .Fields(i).type = 5 Then
                                grid.text = FormatVal(.Fields(i).Value)
                            End If
                        End If
                        
                    Next
                Next
                .MoveNext
            Loop
        
        End With
    End If
    
    
    If grid.rows > 1 Then
        grid.FixedRows = 1
    End If
    If ColunasTotalizar <> "" Then
        TotalizarGrid grid, ColunasTotalizar
    End If
    If bZebra Then
        ZebraGrid grid, 1
    End If
    If zebraultimaLInha Then
        ZebraLinha grid, grid.rows - 1, CorFix, True
    End If
    
    
    grid.Visible = True
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub TotalizarGrid(grid As Object, colunas As String)

    Dim c As New Collection
    Set c = F_SeparaTexto(colunas, ",", False)
    ReDim t(c.Count) As Single
    Dim i As Integer
    Dim j As Integer
    
    With grid
        
        For i = 1 To .rows - 1
            .row = i
            For j = 1 To c.Count
                .col = CInt(c(j))
                If .text <> "" Then
                    t(j) = t(j) + CCur(.text)
                End If
            Next
        Next
        
    End With
    
    grid.AddItem "Total"
    grid.row = grid.rows - 1
    For i = 1 To c.Count
        grid.col = c(i)
        grid.text = FormatVal(t(i))
    Next
    

End Sub
Function getText(grid As Object, coluna As Integer, Optional Linha As Integer = -1) As String
    If Linha > -1 Then
        grid.row = Linha
    End If
    grid.col = coluna
    getText = grid.text

End Function
Function getTextByColumnHead(grid As Object, coluna As String, Optional Linha As Integer = -1) As String
    Dim i As Integer
    If Linha > -1 Then
        grid.row = Linha
    End If
    
    For i = 0 To grid.Cols - 1
        grid.col = i
        If LCase(grid.text) = LCase(coluna) Then
            getTextByColumnHead = grid.text
            Exit For
        End If
    Next
    
    
    
    

End Function

Function getColByText(grid As Object, text As String, row As Integer, Optional ilike As Boolean = False) As Integer
    Dim i As Integer
    Dim ret As Integer
    ret = -1
    With grid
        .row = row
        For i = 0 To .Cols - 1
            .col = i
            If ilike Then
                If InStr(LCase(.text), LCase(text)) > 0 Then
                    ret = .col
                    Exit For
                End If
            Else
                If LCase(.text) = LCase(text) Then
                    ret = .col
                    Exit For
                End If
            End If
        Next
    End With
    getColByText = ret
End Function

Function getRowByText(grid As Object, text As String, _
    col As Integer, Optional linhaInicio As Integer = 1) As Integer
    Dim i As Integer
    Dim ret As Integer
    ret = -1
    With grid
        .col = col
        For i = linhaInicio To .rows - 1
            .row = i
            If LCase(.text) = LCase(text) Then
                ret = .row
                getRowByText = ret
                Exit For
            End If
        Next
    End With
    getRowByText = ret
End Function

Sub ajustaGridTitulos(grade As Object, iColunas As Integer, tituloColunasGrid As String, Optional tamanhoColunas As String = "")
    Dim titulo As New Collection
    Dim tamanhoCol As New Collection
    
    grade.rows = 2
    grade.Cols = iColunas
    grade.rows = 1
    
    Set titulo = F_SeparaTexto(tituloColunasGrid, ",", False)
    Set tamanhoCol = F_SeparaTexto(tamanhoColunas, ",", False)
    For i = 1 To titulo.Count
        grade.col = i - 1
        grade.text = titulo(i)
    Next
    
    For i = 1 To tamanhoCol.Count
        grade.ColWidth(i - 1) = tamanhoCol(i)
        
    Next
    
End Sub
