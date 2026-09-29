Attribute VB_Name = "Troca"
Sub ImprimeTroca(seq_tro As String)
    
    imprTroca seq_tro, frmVisualiza.Pic
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
        imprTroca seq_tro, Printer
        Printer.EndDoc
        
    End If
    Unload frmVisualiza
End Sub
Sub imprTroca(seq_tro As String, obj As Object)
    On Error Resume Next
    Dim strSQL As String
    Dim sSEQ As String
    Dim iCount As Integer
    Dim cQtd_Ant As Currency
    Dim cQtd_Atual As Currency
    Dim sPro As String
    Dim sCod_Pro As String
    Dim cFinal As Currency
    Dim sSeq_ite As String
    Dim sUnit_Ite As String
    Dim abater As Currency
    Dim acrescer As Currency
    Dim desconto As Currency
    
    
    On Error GoTo trata_erro
    
    Dim rPesq As New ADODB.Recordset
    
    
    strSQL = "select tv.*,tt.*,tit.*,tp.nome_pro from tab_vendas tv "
    strSQL = strSQL & " inner join tab_trocas as tt on tv.num_ven=tt.num_ven "
    strSQL = strSQL & " inner join tab_ittrocas as tit on tit.seq_tro=tt.seq_tro "
    strSQL = strSQL & " inner join tab_produtos as tp on tp.cod_pro=tit.cod_pro "
    strSQL = strSQL & " where tt.seq_tro='" & seq_tro & "' "
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    P_CabRelTimbr "TROCA REFERENTE À VENDA " & rPesq!num_ven & " EM " & Date, obj
    P_ImprX obj, "Venda: " & rPesq!num_ven, 0
    P_ImprX obj, "" & rPesq!nomecli_ven, 30, False
    P_ImprX obj, "" & rPesq!data_ven, 0
    P_ImprX obj, "" & rPesq!fechavenda_usu, 45
    P_ImprX obj, "" & FormatVal(rPesq!valor_ven), 100
    P_ImprX obj, "" & Format(rPesq!desconto_ven), 160, False
    P_ImprLinha obj
    'P_ImprX obj, "Inicial: " & lblin, 0
    'P_ImprX obj, "À Vista: " & lblAvIsta.caption, 30
    'P_ImprX obj, "A Prazo: " & lblAPrazo.caption, 60
    'P_ImprX obj, "Financeira: " & lblFinan.caption, 90
    
    
    'P_ImprX obj, "Desconto: R$ " & lblDesconto.caption, 130
    'P_ImprX obj, "Final: " & lblFinal.caption, 170, False
    P_ImprLinha obj, "="
    With rPesq
        desconto = !desconto_tro
        obj.FontBold = True
        abater = !abat_tro
        acrescer = !acres_tro
        P_ImprX obj, "...", 0
        P_ImprX obj, "Cód", 5
        P_ImprX obj, "Descrição", 24
        P_ImprX obj, "Qtd", 120, True, True
        P_ImprX obj, "Unit", 135, True, True
        
        P_ImprX obj, "Valor", 175, True, direita
        P_ImprX obj, "Situação", 180, False, esquerda
        obj.FontBold = False
        P_ImprLinha obj
        Do While Not .EOF
            '.row = iCount
            
            
            sPro = !cod_pro
            
            cQtd_Atual = !qtdatual_itt - !qtdant_itt
            
            P_ImprX obj, .AbsolutePosition, 0
            P_ImprX obj, !cod_pro, 5
            P_ImprX obj, StrConv(Left(!nome_pro, 50), vbProperCase), 24
            P_ImprX obj, !qtdatual_itt - !qtdant_itt, 120, True, True
            P_ImprX obj, cQtd_Atual & "X " & FormatVal((!unit_itt)), 135, True, esquerda
            P_ImprX obj, FormatVal(!unit_itt / (!qtdatual_itt - !qtdant_itt) * (!qtdatual_itt - !qtdant_itt)), 175, True, direita
            If cQtd_Atual < 0 Then
                P_ImprX obj, "Devolvendo", 180, False, esquerda
            Else
                P_ImprX obj, "Levando", 180, False, esquerda
            End If
            
            cFinal = cFinal + cQtd_Atual * (!unit_itt / (!qtdatual_itt - !qtdant_itt))
            
            obj.FontBold = False
            obj.ForeColor = vbBlack
            .MoveNext
        Loop
        obj.FontBold = True
        P_ImprLinha obj
        P_ImprX obj, "A Abater: " & FormatVal(abater), 10
        P_ImprX obj, "A Acrescer: " & FormatVal(acrescer), 60
        If desconto > 0 Then
            P_ImprX obj, "Desconto: " & FormatVal(desconto), 120
        End If
        P_ImprX obj, "Final: " & FormatVal(acrescer - abater - desconto), 190, False, direita
        'P_ImprX obj, "", False
        P_ImprLinha obj, "_"
    End With
Exit Sub
trata_erro:
    MsgBox "Erro na impressão de trocas"
End Sub

