VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Cheques A Receber/ Pagar"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   Icon            =   "frmRelCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10965
   Begin VB.Frame frControle 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.Frame Frame5 
         Caption         =   "&Ordenar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   2175
         Begin VB.ComboBox cmbxOrder 
            Height          =   315
            ItemData        =   "frmRelCheques.frx":030A
            Left            =   120
            List            =   "frmRelCheques.frx":031D
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   1935
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
            ItemData        =   "frmRelCheques.frx":0358
            Left            =   120
            List            =   "frmRelCheques.frx":0365
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "E&xpressão:"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2175
         Begin VB.CommandButton cmndbtnTerc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            Picture         =   "frmRelCheques.frx":0385
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtbxExpr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelCheques.frx":070F
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4920
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelCheques.frx":0FD9
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelCheques.frx":12E3
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3240
         Width           =   2175
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
            ItemData        =   "frmRelCheques.frx":15ED
            Left            =   120
            List            =   "frmRelCheques.frx":15FA
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
         Begin VB.TextBox txtbxdataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   7
            Top             =   240
            Width           =   975
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   2145
         TabIndex        =   18
         Top             =   5760
         Width           =   2175
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
            TabIndex        =   20
            Top             =   120
            Width           =   510
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
            TabIndex        =   19
            Top             =   82
            Width           =   570
         End
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmRelCheques.frx":1648
         Height          =   6135
         Left            =   2400
         OleObjectBlob   =   "frmRelCheques.frx":165C
         TabIndex        =   21
         Top             =   240
         Width           =   8415
      End
   End
End
Attribute VB_Name = "frmRelCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub soma()
    Dim cTotal As Currency
    With adc.Recordset
        Do While Not .EOF
                        
            cTotal = cTotal + !Valor_che
            
            .MoveNext
        Loop
    End With
    lblTotal.caption = FormatVal(cTotal)
    
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
        sOrderBy = "venc_che"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "Nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "num_doc"
    ElseIf cmbxOrder.ListIndex = 3 Then
        sOrderBy = "valor_che"
    Else
        sOrderBy = "dep_che"
    End If
    If cmbxTipo.ListIndex = 0 Then
        dbgrd.Columns(4).caption = "Banco"
        strSQL = "select tc.cod_cli as codigo,tb.nome_bco as banco, tc.nome_cli as nome,td.historico_doc,td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc,"
        strSQL = strSQL & "ti.* from tab_clientes as tc, tab_doc as td,tab_cheques as ti,  "
        strSQL = strSQL & "tab_bancos as tb where tb.cod_bco=ti.cod_bco "
        strSQL = strSQL & " and tc.cod_cli=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='R' "
        strSQL = strSQL & " and ti.dep_che is null "
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_che <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_che >= " & F_GravaData(txtbxDataIni.Text)
        End If
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and tc.cod_Cli =  '" & txtbxExpr.tag & "' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            sCampo = "ti.valor_che = " & str(txtbxExpr.Text)
        
        End If
        strSQL = strSQL & " order by " & sOrderBy
    ElseIf cmbxTipo.ListIndex = 1 Then
        dbgrd.Columns(4).caption = "Conta"
        strSQL = "select  tc.cod_for as codigo,tb.nome_con as banco, tc.nome_for as nome,td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc,"
        strSQL = strSQL & "ti.* from tab_fornecedores as tc, tab_doc as td,tab_cheques as ti,  "
        strSQL = strSQL & "tab_contas as tb where tb.cod_con=ti.cod_con"
        strSQL = strSQL & " and tc.cod_for=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='P' "
        strSQL = strSQL & " and ti.dep_che is null "
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_che < " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_che > " & F_GravaData(txtbxDataIni.Text)
        End If
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and tc.cod_for =  '" & txtbxExpr.tag & "' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            sCampo = "ti.valor_che = " & str(txtbxExpr.Text)
        End If
        strSQL = strSQL & " order by " & sOrderBy
    Else
        dbgrd.Columns(4).caption = "Banco"
        strSQL = "select tc.cod_cli as codigo,tb.nome_bco as banco, tc.nome_cli as nome,td.historico_doc,td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc,"
        strSQL = strSQL & "ti.* from tab_clientes as tc, tab_doc as td,tab_cheques as ti,  "
        strSQL = strSQL & "tab_bancos as tb where tb.cod_bco=ti.cod_bco "
        strSQL = strSQL & " and tc.cod_cli=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='R' "
        strSQL = strSQL & " and ti.dep_che <> '0000-00-00' "
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_che <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_che >= " & F_GravaData(txtbxDataIni.Text)
        End If
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and tc.cod_Cli =  '" & txtbxExpr.tag & "' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            sCampo = "ti.valor_che = " & str(txtbxExpr.Text)
        End If
        strSQL = strSQL & " order by " & sOrderBy
    End If
    adc.ConnectionString = Conn.ConnectionString
    adc.RecordSource = strSQL
    adc.Refresh
    If adc.Recordset.RecordCount = 0 Then
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
    
    
    If cmbxTipo.ListIndex = 0 Then
        sTipo = "Relatório de Cheques a Receber no Período"
        
    
    
    Else
        sTipo = "Relatório de Cheques a Pagar no Período"
    
    End If
    P_CabRelTimbr sTipo & " de " & txtbxDataIni.Text & _
        " Até " & txtbxdataFin.Text, obj
    obj.FontName = "arial"
    obj.fontsize = 12
    
    P_ImprLinha obj
    obj.fontsize = 9
    
    P_ImprX obj, "Seq.:", 0
    P_ImprX obj, "Terceiro", 18
    P_ImprX obj, "Nº Cheque", 80
    P_ImprX obj, "Conta", 110
    P_ImprX obj, "Agência", 130
    P_ImprX obj, "Banco", 150
    P_ImprX obj, "Valor(R$)", 180, True, True
    P_ImprX obj, "Vencimento", 200, False, True
    P_ImprLinha obj
    
    With adc.Recordset
        
        .MoveFirst
        strSQL = CStr(.RecordCount / 60)
        If InStr(1, strSQL, ",") <> 0 Then
            iPag = Int(strSQL) + 1
        End If
        
        Do While Not .EOF
            iItens = iItens + 1
            If iItens = 60 Then
                P_RodapeVis obj, "Continua...", "Pág:" & IIf(TypeName(obj) = "Printer", "N", "%") & " de " & iPag
                iItens = 0
                obj.FontName = "arial"
                obj.fontsize = 9
                
                P_ImprX obj, "Seq.:", 0
                P_ImprX obj, "Terceiro", 18
                P_ImprX obj, "Nº Cheque", 80
                P_ImprX obj, "Conta", 110
                P_ImprX obj, "Agência", 130
                P_ImprX obj, "Banco", 150
                P_ImprX obj, "Valor(R$)", 180, True, True
                P_ImprX obj, "Vencimento", 200, False, True

                P_ImprLinha obj
            End If
            obj.FontName = "arial"
            obj.fontsize = 9
            P_ImprX obj, !seq_che, 0
            P_ImprX obj, Left(!nome, 32), 18
            P_ImprX obj, !num_che, 80
            P_ImprX obj, !cco_che, 110
            P_ImprX obj, !age_che, 130
            P_ImprX obj, !cod_bco, 150
            P_ImprX obj, FormatVal(!Valor_che), 180, True, True
            P_ImprX obj, !Venc_che, 200, False, True
            cTotal = cTotal + !Valor_che
           
            .MoveNext
        Loop
        P_ImprLinha obj
        
        'P_ImprX obj, "" & lblTotalGerado.caption, 160, True, True
        
'        P_ImprX obj, "" & lblLiq.caption, 180, True, True
        P_ImprX obj, "" & Round(cTotal, 2), 200, False, True
        
        
        P_ImprX obj, lblTotal.caption, 130, False
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

Private Sub cmndbtnTerc_Click()
    
    If cmbxBuscar.ListIndex <> 1 Then
        cmndbtnBuscar_Click
        Exit Sub
    End If
    If cmbxTipo.ListIndex = 0 Then
        frmSelectCli.txtbxNOme.tag = "C"
    Else
        frmSelectCli.txtbxNOme.tag = "F"
    End If
    frmSelectCli.txtbxNOme.Text = txtbxExpr.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxExpr.Text = frmSelectCli.adoR!nome
    txtbxExpr.tag = frmSelectCli.adoR!cod
    Unload frmSelectCli
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
    adc.ConnectionString = Conn.ConnectionString

    P_CentraForm Me
    
'    strSQL = "select tn.*,tt.* from tab_compras as tn, tab_fornecedores as tt "
'    strSQL = strSQL & " where tt.cod_for=tn.cod_for "
'
'    strSQL = strSQL & " order by num_not "
'    adc.RecordSource = strSQL
'    adc.Refresh
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

Private Sub txtbxExpr_Change()
    'cmndbtnTerc_Click
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    If cmbxBuscar.ListIndex = 0 Then
        txtbxExpr.Text = F_SoNumero(txtbxExpr.Text)
    Else
        P_AjustaTexto txtbxExpr
    End If
End Sub




