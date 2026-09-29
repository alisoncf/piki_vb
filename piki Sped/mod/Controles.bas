Attribute VB_Name = "Controles"
Sub resizeForm(obj As Form)
    obj.height = mdiFrm.ScaleHeight - 200
    obj.Width = mdiFrm.ScaleWidth - 200
    If obj.MDIChild = True Then
        obj.Top = 10
        obj.Left = 10
    Else
        obj.Top = 300
        obj.Left = 10
    End If
    
End Sub
Sub posicionaADireita(obj As Object, container As Object)
    obj.Left = container.Width - obj.Width - 10
End Sub
Sub PosicionaGradeWidth(grade As Object, frm As Object, Optional recuo As Integer = 300)
    grade.Width = frm.Width - grade.Left - recuo
End Sub
Sub PosicionaGradeHeigth(grade As Object, frm As Object, Optional recuo As Integer = 200)
    grade.height = frm.Top - grade.Top - recuo
End Sub
Sub PosicionaGrade(grade As Object, objW As Object, objH As Object)
    PosicionaGradeWidth grade, objW
    PosicionaGradeHeigth grade, objH
End Sub

Sub P_AjustaTelefone(obj As Object)
    On Error GoTo trata_erro
        obj.Text = F_AjustaTelefone(obj.Text)
    Exit Sub
trata_erro:
    
End Sub
Sub P_ajustaGridWid(msflxgrd As Object, Optional iPrimLinha As Boolean = False)
    Dim iCount As Integer
    Dim iRow As Integer
    With msflxgrd
        msflxgrd.Visible = False
        For iRow = 0 To .rows - 1
            .row = iRow
            For iCount = 0 To .Cols - 1
                
                .Col = iCount
                If .CellWidth < Printer.TextWidth(.Text) Then
                    .ColWidth(.Col) = Printer.TextWidth(.Text)
                End If
            Next
        Next
        msflxgrd.Visible = True
    End With
End Sub

Sub ZebraGrid(msf As MSFlexGrid, linhaI As Integer, _
    Optional cor1 = &H80000005, Optional cor2 = &HB8D6D8)
    Dim iCol As Integer
    Dim iRow As Integer
    msf.Visible = False
    With msf
        For iRow = linhaI To .rows - 1
            .row = iRow
            For iCol = 0 To .Cols - 1
                .Col = iCol
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
    On Error GoTo trata_erro
'    On Error Resume Next
    msf.Visible = False
    With msf
        For iRow = linhaI To linhaF
            .row = iRow
            For iCol = 0 To .Cols - 1
                .Col = iCol
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
trata_erro:
msf.Visible = True
End Sub
Sub ColoreLinhaGrid(grid As MSFlexGrid, linha As Integer, Optional cor = &HD3E7EB, _
    Optional negrito As Boolean = False, Optional foreCor = "")



    On Error Resume Next
    Dim iCount As Integer
    Dim iRow As Integer
    With grid
        .row = linha
        For iCount = 0 To .Cols - 1
            .Col = iCount
            .CellBackColor = cor
            .CellFontBold = negrito
            If foreCor <> "" Then
                .CellForeColor = foreCor
            End If
        Next
        
    End With
End Sub

Function F_InFlxgrd(msf As Object, Text As String, colum As Integer) As Integer
    Dim iCount As Integer
    Dim bOk As Boolean
    With msf
        bOk = False
        For iCount = 0 To .rows - 1
            .row = iCount
            .Col = colum
            If UCase(.Text) = UCase(Text) Then
                bOk = True
                F_InFlxgrd = iCount
                Exit For
            End If
        Next
        If bOk = False Then
            F_InFlxgrd = -1
        End If
        
    End With
End Function
Function F_InFlxgrdCol(msf As Object, Text As String, row As Integer, Optional colI As Integer = -1, Optional colF As Integer = -1) As Integer
    Dim iCount As Integer
    Dim bOk As Boolean
    With msf
        bOk = False
        
        If colI = -1 Then
            colI = 0
        End If
        If colF = -1 Then
            colF = .Cols - 1
        End If
        For iCount = colI To colF
            .Col = iCount
            .row = row
            If UCase(.Text) = UCase(Text) Then
                bOk = True
                F_InFlxgrdCol = iCount
                Exit For
            End If
        Next
        If bOk = False Then
            F_InFlxgrdCol = -1
        End If
        
    End With
End Function
Function F_InList(msf As ListBox, Text As String) As Integer
    Dim iCount As Integer
    Dim bOk As Boolean
    With msf
        bOk = False
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            
            If UCase(.Text) = UCase(Text) Then
                bOk = True
                F_InList = iCount
                Exit For
            End If
        Next
        If bOk = False Then
            F_InList = -1
        End If
        
    End With
End Function
Function FormatCodigo(sCod As String) As String
    On Error Resume Next
    FormatCodigo = Format(sCod, "00#\.####\-#")
End Function

Sub P_AjustaNota(ByRef objNota As TextBox, _
Optional intTam As Integer = 8)
    'Formata um objeto com zeros na frente do número, que deve ser maior que 0
    On Error GoTo TrataErro:
    If objNota.Text = "" Then Exit Sub
    objNota = F_SoNumero(objNota.Text)
    
    objNota.Text = Format(objNota.Text, "0" & String(intTam - 1, "#"))
    Exit Sub

TrataErro:
    Exit Sub
    
End Sub


Sub P_PreencheTaxasICMS(cmbxOBJ As ComboBox)
    cmbxOBJ.Clear
    cmbxOBJ.AddItem "00 %"
    cmbxOBJ.AddItem "05 %"
    cmbxOBJ.AddItem "07 %"
    cmbxOBJ.AddItem "12 %"
    cmbxOBJ.AddItem "17 %"
    cmbxOBJ.AddItem "18 %"
End Sub



Sub P_PreencheNotaSaida(cmbxOBJ As Object)
    'Preenche o tipo de operação da Nota Fiscal
    cmbxOBJ.Clear
    cmbxOBJ.AddItem "5.12-Vendas Merc.Estadual"
    cmbxOBJ.AddItem "6.12-Vendas Merc.InterEstadual"
    cmbxOBJ.AddItem "5.32-Devolução Estadual"
    cmbxOBJ.AddItem "6.32-Devolução InterEstadual"
    cmbxOBJ.AddItem "5.22-Transf. Merc. Estadual"
    cmbxOBJ.AddItem "6.22-Transf. Merc InterEstadual"
    cmbxOBJ.AddItem "5.73-Venda Merc.Retida Estadual"
    cmbxOBJ.AddItem "5.99-Outras Saídas Estadual"
    cmbxOBJ.AddItem "6.99-Outras Saídas InterEstadual"
    cmbxOBJ.AddItem "5.93-Trans Mat.Consumo Estadual"
    cmbxOBJ.AddItem "6.93-Trans Mat.Consumo InterEstadual"
End Sub

Sub P_PreencheSituacaoTributaria(cmbxOBJ As ComboBox)
    'Carrega um combobox com as situações da tributação do ICMS
    cmbxOBJ.Clear
    cmbxOBJ.AddItem "00-Tributada Integralmente"
    cmbxOBJ.AddItem "01-Tributada e C/Cobr do ICMS por Subst"
    cmbxOBJ.AddItem "02-Com Redução de Base de Cálculo"
    cmbxOBJ.AddItem "03-Isenta ou não tributada C/Cobr ICMS Subst"
    cmbxOBJ.AddItem "04-Isenta ou não tributada"
    cmbxOBJ.AddItem "05-Com Suspensão ou diferimento"
    cmbxOBJ.AddItem "06-ICMS Cobrado Ant. por Subst Trib"
    cmbxOBJ.AddItem "07-C/Redução De Base de Calc C/Cobr ICMS Subst"
    cmbxOBJ.AddItem "09-Outras"
End Sub


Function IsLoad(Nomefrm As String) As Boolean
    Dim i As Integer
    Dim sText As String
    For i = 0 To Forms.count - 1
        If UCase(Forms(i).Name) = UCase(Nomefrm) Then
            IsLoad = True
            Exit Function
        End If
    Next
    IsLoad = False
End Function

Sub RefreshDataPosition(dt As Data)
    Dim iPos As Integer
    On Error GoTo trata_erro
    
    
    
    If dt.Recordset.RecordCount = 0 Then Exit Sub
    On Error Resume Next
    iPos = dt.Recordset.AbsolutePosition
    dt.Refresh
    dt.Recordset.Move iPos
    Exit Sub
trata_erro:
    Err.Clear
    On Error Resume Next
    
    dt.Refresh
    'MsgBox Err.Description & " erro em refreshdataposition", vbExclamation
End Sub

Sub RefreshAdoPosition(dt As ADODB.Recordset)
    Dim iPos As Integer
    On Error GoTo trata_erro
    
    
    
    If dt.RecordCount = 0 Then Exit Sub
    On Error Resume Next
    iPos = dt.AbsolutePosition
    dt.Requery
    dt.Move iPos
    Exit Sub
trata_erro:
    Err.Clear
    On Error Resume Next
    
    dt.Requery
    'MsgBox Err.Description & " erro em refreshdataposition", vbExclamation
End Sub


Sub LockForm(opcao As Boolean, frm As Form)
    'opcao true loca, false desloca
    Dim iCount As Integer
    For iCount = 0 To Forms.count - 1
        If opcao = True Then
            Forms(iCount).Enabled = False
        Else
            Forms(iCount).Enabled = True
        End If
        If Forms(iCount).Name = frm.Name Then
            Forms(iCount).Enabled = True
        End If
    Next
End Sub

Sub P_ValorLostFocus(ByRef objValor As Object, _
Optional strFormat As String = "###,##0.00")
    'Procedimento de perda de foco sobre valores
    Dim strExpr As String
    On Error GoTo trata_erro
    strExpr = F_SoNumero(objValor.Text, True)
    If F_Vazio(objValor.Text) Then
        strExpr = "0"
    End If
    If Right(strExpr, 1) = "." Then
       strExpr = "0" & strExpr
       Mid(strExpr, Len(strExpr)) = ","
    ElseIf Left(Right(strExpr, 2), 1) = "." Then
       strExpr = "0" & strExpr
       Mid(strExpr, Len(strExpr) - 1) = ","
    ElseIf Left(Right(strExpr, 3), 1) = "." Then
       strExpr = "0" & strExpr
       Mid(strExpr, Len(strExpr) - 2) = ","
    End If
    objValor.Text = strExpr
    'objValor.Text = (CCur(F_AjustaValor(objValor.Text)))
    objValor.Text = Format(objValor.Text, strFormat)
    
    If CCur(F_SoNumero(objValor.Text)) < 0 Then
       'Se o valor for menor que zero, vira zero
       objValor.Text = "0,00"
    End If
Exit Sub
trata_erro:
objValor.Text = "0,00"
End Sub

Sub P_SelecTexto(objText As Object)
    'Este procedimento seleciona o texto
    'do objeto em questão
    objText.SelStart = 0
    objText.SelLength = Len(RTrim(objText.Text))
End Sub


Sub ajustaTamGrid(msflxgrd As MSFlexGrid, container As Object)
    Dim iCount As Integer
    Dim iCol As Integer
    Dim tMax As String
    With msflxgrd
        For iCount = 0 To .rows - 1
            .row = iCount
            For iCol = 0 To .Cols - 1
                .Col = iCol
                If container.TextWidth(.Text) > .ColWidth(iCol) Then
                     .ColWidth(iCol) = container.TextWidth(.Text) + 100
                End If
                
            Next
        Next
    End With
End Sub

Sub ZebraLinha(msf As Object, linha As Integer, _
    Optional cor1 = "&HFFFF")
    Dim iCol As Integer
    Dim iRow As Integer
    
    With msf
        
            .row = linha
            For iCol = 0 To .Cols - 1
                .Col = iCol
                
                    .CellBackColor = cor1
                
            Next
        
    End With
    
End Sub


Function F_confMsflxCH(msflx As MSFlexGrid, colven As Integer, colch As Integer, _
    colage As Integer, colcco As Integer, colbco As Integer, colcpf As Integer, _
    colnome As Integer, Optional coltipo_doc As Integer = -1) As Boolean
     Dim iCount As Integer
     Dim sTipo As String
     Dim Retorno As Boolean
     Dim strSQL As String
     Dim iPassa As Boolean
     iPassa = True
     Retorno = True
     strSQL = "os seguintes dados estão faltando: " & Chr(13) & Chr(13)
     With msflx
        For iCount = 1 To .rows - 1
            iPassa = True
            .row = iCount
                If coltipo_doc > -1 Then
                    .Col = coltipo_doc
                    If Left(.Text, 1) <> "2" And Left(.Text, 1) <> "C" Then
                        iPassa = False
                    End If
                End If
                
                If iPassa = True Then
                    
                    .Col = colch
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        Retorno = False
                    End If
                    .Col = colbco
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        Retorno = False
                    End If
                    .Col = colage
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " agência" & Chr(13)
                        Retorno = False
                    End If
                    .Col = colcco
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " conta" & Chr(13)
                        Retorno = False
                    End If
                    .Col = colnome
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " nome do titular" & Chr(13)
                        Retorno = False
                    End If
                    .Col = colcpf
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " CPF do titular" & Chr(13)
                        Retorno = False
                    End If
                End If
                
                
            
        Next
     End With
     F_confMsflxCH = Retorno
     If Retorno = False Then
        MsgBox strSQL, vbExclamation
        
     End If
End Function

Function IsInCollection(c As Collection, valor As String) As Integer
    
     Dim i As Integer
     For iCount = 1 To c.count
        If Trim(c(iCount)) = Trim(valor) Then
            IsInCollection = iCount
            Exit For
        End If
     Next

End Function

Function F_NotIn(campo As String, tabela As String, where As String, Optional sep As String = "'") As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select  " & campo & " as campo from " & tabela & " "
    If where <> "" Then
        strSQL = strSQL & " where " & where
    End If
    rPesq.Open strSQL, Conn_Client
    With rPesq
        strSQL = "("
        Do While Not .EOF
        
            strSQL = strSQL & sep & !campo & sep
            
            .MoveNext
            If rPesq.EOF Then
                strSQL = strSQL & ")"
            Else
                strSQL = strSQL & ","
            End If
        Loop
    End With
    F_NotIn = strSQL
End Function
Function F_DigitoVerificador(numero As String, Optional qtd_dig_ver As Integer = 1) As String
    Dim strSQL As String
    Dim iCount As Integer
    Dim iVal As Single
    Dim sFor As String
    Dim sDig As String
    For iCount = 1 To Len(numero)
        iVal = iVal + Asc(Mid(numero, iCount, 1))
    Next
    iVal = iVal
    sFor = String(qtd_dig_ver - 1, "0") & "#"
    sDig = Mid(iVal, 2, qtd_dig_ver)
    
    F_DigitoVerificador = numero & Format(sDig, sFor)
End Function
Sub AlteraCabGrid(grid As MSFlexGrid, colunas As String)

End Sub


