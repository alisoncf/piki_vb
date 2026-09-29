VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Contas a Receber/Pagar"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "frmRelMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10155
   Begin VB.Frame frControle 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2400
         ScaleHeight     =   345
         ScaleWidth      =   7545
         TabIndex        =   18
         Top             =   5760
         Width           =   7575
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sugerido:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5160
            TabIndex        =   26
            Top             =   45
            Width           =   855
         End
         Begin VB.Label lblSug 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   360
            Left            =   6120
            TabIndex        =   25
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Juros:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2760
            TabIndex        =   24
            Top             =   45
            Width           =   525
         End
         Begin VB.Label lblJuros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   360
            Left            =   3480
            TabIndex        =   23
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   360
            Left            =   720
            TabIndex        =   20
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   38
            Width           =   510
         End
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmRelMov.frx":030A
         Height          =   5415
         Left            =   2400
         OleObjectBlob   =   "frmRelMov.frx":031E
         TabIndex        =   21
         Top             =   240
         Width           =   7575
      End
      Begin VB.CheckBox chkAdiantamento 
         BackColor       =   &H80000000&
         Caption         =   "Considerar Adiantamento"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxdataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1200
            TabIndex        =   8
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Tipo:"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2175
         Begin VB.ComboBox cmbxTipo 
            Height          =   315
            ItemData        =   "frmRelMov.frx":22F7
            Left            =   120
            List            =   "frmRelMov.frx":2304
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelMov.frx":2350
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelMov.frx":265A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelMov.frx":2964
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Frame Frame6 
         Caption         =   "E&xpressão:"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2175
         Begin VB.TextBox txtbxExpr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnTerc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            Picture         =   "frmRelMov.frx":322E
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "&Buscar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2175
         Begin VB.ComboBox cmbxBuscar 
            Height          =   315
            ItemData        =   "frmRelMov.frx":35B8
            Left            =   120
            List            =   "frmRelMov.frx":35C5
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Data dtNot 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Ordenar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   2175
         Begin VB.ComboBox cmbxOrder 
            Height          =   315
            ItemData        =   "frmRelMov.frx":35E5
            Left            =   120
            List            =   "frmRelMov.frx":35F5
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmRelMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub soma()
    Dim cTotal As Currency
    Dim cJuros As Currency
    Dim cSugerido As Currency
    With dtNot.Recordset
       If cmbxTipo.ListIndex = 0 Or cmbxTipo.ListIndex = 1 Then
            Do While Not .EOF
            cTotal = cTotal + !reman_itd
            If !tipo_doc = "R" Then
                If (!valor_juro < 0 And chkAdiantamento.Value = 1) Or (!valor_juro > 0) Then
                    cJuros = cJuros + !valor_juro
                    cSugerido = cSugerido + !sugerido
                End If
            End If
            .MoveNext
            Loop
       Else
            Do While Not .EOF
                cTotal = cTotal + !valor_itd
                cJuros = cJuros + !juros_itd
                .MoveNext
            Loop
        End If
    End With
    lblTotal.caption = FormatVal(cTotal)
    lblJuros.caption = FormatVal(cJuros)
    lblSug.caption = FormatVal(cSugerido)
End Sub

Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub cmndbtnBuscar_Click()
    
    Dim strSQL As String
    Dim sOrderBy As String
    Dim sCampo As String
    On Error GoTo trata_Erro
    
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "venc_Itd"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "Nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "num_doc,venc_itd"
    Else
        sOrderBy = "reman_itd"
    End If
        
    
    If cmbxTipo.ListIndex = 0 Then
        strSQL = "select tipo_doc,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, tc.cod_cli as codigo, tc.nome_cli as nome,td.historico_doc,td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc,"
        strSQL = strSQL & " ((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, ((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, "
        strSQL = strSQL & "ti.* from tab_clientes as tc, tab_doc as td,tab_itdoc as ti,tab_parametros as tp  "
        strSQL = strSQL & " where tc.cod_cli=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='R' "
        strSQL = strSQL & " and ti.reman_itd > 0 "
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
        
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and tc.cod_Cli =  '" & txtbxExpr.tag & "' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            sCampo = "ti.valor_itd = " & str(txtbxExpr.Text)
        
        End If
        strSQL = strSQL & " order by " & sOrderBy
            
        
    ElseIf cmbxTipo.ListIndex = 1 Then
        strSQL = "select tipo_doc, tc.cod_for as codigo,tc.nome_for as nome,td.historico_doc,td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc,"
        strSQL = strSQL & "ti.* from tab_fornecedores as tc, tab_doc as td,tab_itdoc as ti  "
        strSQL = strSQL & " where tc.cod_for=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='P' "
        
        strSQL = strSQL & " and ti.reman_itd > 0 "
        
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd < " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd > " & F_GravaData(txtbxDataIni.Text)
        End If
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and tc.cod_for =  '" & txtbxExpr.tag & "' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            sCampo = "ti.valor_itd = " & str(txtbxExpr.Text)
        
        End If
        strSQL = strSQL & " order by " & sOrderBy
    Else
        strSQL = "select tipo_doc,tc.cod_cli as codigo, tc.nome_cli as nome,td.historico_doc,td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc,"
        strSQL = strSQL & "ti.* from tab_clientes as tc, tab_doc as td,tab_itdoc as ti  "
        strSQL = strSQL & " where tc.cod_cli=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='R' "
        strSQL = strSQL & " and ti.reman_itd = 0 "
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.baixa_itd <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.baixa_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and tc.cod_Cli =  '" & txtbxExpr.tag & "' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            strSQL = strSQL & " and td.historico_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 3 Then
            strSQL = strSQL & " and ti.reman_itd = " & str(txtbxExpr.Text)
        End If
        strSQL = strSQL & " order by " & sOrderBy
    End If
    
    
    
    
    
    
    dtNot.Connect = Banco.Connect
    dtNot.RecordSource = strSQL
    dtNot.Refresh
    If dtNot.Recordset.RecordCount = 0 Then
        MsgBox "Não foram encontrados registros para serem listados!" & Chr(13) & _
                "Verifique sua Pesquisa!", vbExclamation, constNome_sis
        
    End If
    soma
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnLocaliza_Click()
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 1
End Sub

Sub imprime(obj As Object)
    Dim iItens As Integer
    Dim iPag As Integer
    Dim strSQL As String
    Dim cTotal As Currency
    Dim sTipo As String
    Dim cPerc As Currency
    Dim sCod_Ter As String
    Dim cParcial As Currency
    Dim cJurosParc As Currency
    Dim cSugParc As Currency
    Dim cTotalJuros As Currency
    Dim cTotalSug As Currency
    
    If cmbxTipo.ListIndex = 0 Then
        sTipo = "Relatório de Contas a Receber no Período"
        
    
    
    Else
        sTipo = "Relatório de Contas a Pagar no Período"
    
    End If
    P_CabRelTimbr sTipo & " de " & txtbxDataIni.Text & _
        " Até " & txtbxdataFin.Text, obj
    obj.FontName = "arial"
    obj.fontsize = 12
    
    P_ImprLinha obj
    obj.fontsize = 9
    
    P_ImprX obj, "Documento", 0
    P_ImprX obj, "Terceiro", 20
    P_ImprX obj, "Histórico", 72
    P_ImprX obj, "Valor", 148, True, True
    P_ImprX obj, "Acrésc.", 160, True, True
    P_ImprX obj, "Sugerido.", 180, True, True
    P_ImprX obj, "Vencimento", 200, False, True
    
    
    
    P_ImprLinha obj
    
    With dtNot.Recordset
        
        .MoveFirst
        strSQL = CStr(.RecordCount / 50)
        If InStr(1, strSQL, ",") <> 0 Then
            iPag = Int(strSQL) + 1
        End If
        sCod_Ter = !codigo
        Do While Not .EOF
            If sCod_Ter <> !codigo And cmbxOrder.ListIndex = 1 Then
                P_ImprLinha obj, "_"
                P_ImprX obj, "Total Parcial.....: " & FormatVal(cParcial), 200, True, True
                P_ImprX obj, "Sugerido Parcial.....: " & FormatVal(cSugParc), 150, True, True
                P_ImprX obj, "Juros Parcial.....: " & FormatVal(cJurosParc), 100, False, True
                P_ImprLinha obj, "-"
                cParcial = 0
            End If
            iItens = iItens + 1
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If TypeName(obj) = "Printer" Then
                    P_Rodape "Continua...", True, "Pág:" & IIf(TypeName(obj) = "Printer", obj.Page, "%")
                Else
                    P_RodapeVis obj, "Continua...", "Pág:" & IIf(TypeName(obj) = "Printer", "N", "%")
                End If
                iItens = 0
                obj.FontName = "arial"
                obj.fontsize = 9
                P_ImprX obj, "Documento", 0
                P_ImprX obj, "Terceiro", 20
                P_ImprX obj, "Histórico", 72
                P_ImprX obj, "Valor", 148, True, True
                P_ImprX obj, "Acrésc.", 160, True, True
                P_ImprX obj, "Sugerido.", 180, True, True
                P_ImprX obj, "Vencimento", 200, False, True
                P_ImprLinha obj
            End If
            obj.FontName = "arial"
            obj.fontsize = 9
            P_ImprX obj, !num_doc, 0
            P_ImprX obj, Left(!nome, 25), 20
            P_ImprX obj, Left(!historico_doc, 30), 72
            
            P_ImprX obj, FormatVal(!reman_itd), 148, True, True
            If !tipo_doc = "R" Then
                If !valor_juro > 0 Or (!valor_juro < 0 And chkAdiantamento.Value = 1) Then
                    P_ImprX obj, FormatVal(!valor_juro), 160, True, True
                    P_ImprX obj, FormatVal(!sugerido), 180, True, True
                Else
                    P_ImprX obj, FormatVal("0,00"), 160, True, True
                    P_ImprX obj, FormatVal("0,00"), 180, True, True
                End If
            End If
            P_ImprX obj, !venc_itd, 200, False, True
            
            
            cTotal = cTotal + !reman_itd
            If !tipo_doc = "R" Then
                cTotalJuros = cTotalJuros + !valor_juro
                cTotalSug = cTotalSug + !sugerido
            End If
            sCod_Ter = !codigo
            
            
            cParcial = cParcial + !reman_itd
            If !tipo_doc = "R" Then
                If !valor_juro > 0 Or (!valor_juro < 0 And chkAdiantamento.Value = 1) Then
                    cSugParc = cSugParc + !sugerido
                    cJurosParc = cJurosParc + !valor_juro
                End If
            End If
            
            .MoveNext
        Loop
        P_ImprLinha obj
        
        'P_ImprX obj, "" & lblTotalGerado.caption, 160, True, True
        
'        P_ImprX obj, "" & lblLiq.caption, 180, True, True
        P_ImprX obj, "Total Total.....: " & FormatVal(cTotal), 200, True, True
                P_ImprX obj, "Sugerido Total.....: " & lblSug.caption, 150, True, True
                P_ImprX obj, "Juros Total.....: " & lblJuros.caption, 100, False, True
        On Error Resume Next
        If TypeName(obj) = "Printer" Then
            P_RodapeVis obj, strNome_Emp, "Pág: " & obj.Page & " de " & obj.Page
        End If
    End With
    
    If TypeName(obj) = "Printer" Then
        obj.EndDoc
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
    End If
    Unload frmVisualiza
End Sub

Private Sub cmndbtnSair_Click()
        Unload Me
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.BorderStyle = 0
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dbgrdChe_GotFocus()
    dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChe_LostFocus()
    dbgrdChe.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    Dim strSQL As String
    
    
    cmbxTipo.ListIndex = 0
    cmbxOrder.ListIndex = 0
    dtNot.Connect = Banco.Connect

    P_CentraForm Me
    
'    strSQL = "select tn.*,tt.* from tab_compras as tn, tab_fornecedores as tt "
'    strSQL = strSQL & " where tt.cod_for=tn.cod_for "
'
'    strSQL = strSQL & " order by num_not "
'    dtNot.RecordSource = strSQL
'    dtNot.Refresh
    cmbxBuscar.ListIndex = 0
    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    cmndbtnBuscar_Click
End Sub

Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
 
    
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
            
        'cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    If cmbxBuscar.ListIndex = 0 Then
        txtbxExpr.Text = F_SoNumero(txtbxExpr.Text)
    Else
        P_AjustaTexto txtbxExpr
    End If
End Sub


