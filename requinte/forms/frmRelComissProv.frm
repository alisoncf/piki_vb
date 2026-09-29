VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelCarnesRecebidos 
   Caption         =   "Relatório de Carnês Recebidos"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   Icon            =   "frmRelComissProv.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   11580
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   7
      Top             =   0
      Width           =   11580
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   4200
         Picture         =   "frmRelComissProv.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   6000
         TabIndex        =   13
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   4800
         TabIndex        =   12
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11550
      TabIndex        =   1
      Top             =   7935
      Width           =   11580
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   0
         Picture         =   "frmRelComissProv.frx":07CC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico de Vendas"
         Height          =   735
         Left            =   4560
         Picture         =   "frmRelComissProv.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   9120
         Picture         =   "frmRelComissProv.frx":0CA0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   2280
         Picture         =   "frmRelComissProv.frx":0D95
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Detalhes"
         Height          =   735
         Left            =   6840
         Picture         =   "frmRelComissProv.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelComissProv.frx":128C
      Height          =   6855
      Left            =   120
      OleObjectBlob   =   "frmRelComissProv.frx":129D
      TabIndex        =   0
      Top             =   960
      Width           =   11295
   End
End
Attribute VB_Name = "frmRelCarnesRecebidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxFunc_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = "select ts.nome_sub,sum(final_ite) as total,(sum(final_ite)*ts.comiss_sub/100) as comiss,"
    strSQL = strSQL & " tf.guerra_fun from tab_vendas as tv, tab_itvendas as ti,tab_produtos as tp, tab_subsec as ts,"
    strSQL = strSQL & " tab_funcionarios as tf where tv.abrevenda_usu=tf.cod_fun "
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_PesqData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_PesqData(txtbxdataFin.Text)
    If cmbxFunc.Text <> "" Then
        strSQL = strSQL & " and guerra_fun='" & cmbxFunc.Text & "' "
    End If
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tp.cod_sub=ts.cod_sub "
    strSQL = strSQL & " and situacao_ven=4 "
    strSQL = strSQL & " group by ts.cod_sub, tf.cod_fun order by tf.guerra_fun "
    dt.RecordSource = strSQL
    dt.Refresh
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    With dt.Recordset
        Do While Not .EOF
            cTotal = cTotal + !total
            .MoveNext
        Loop
        lblVendas.caption = FormatVal(cTotal)
        dt.Refresh
    End With
    
End Sub

Private Sub cmndbtnCliente_Click()
    buscaRep
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = representante.F_BuscaRep(txtbxRepr.Text)
    If r.cod_rep <> "" Then
        txtbxRepr.tag = r.cod_rep
        txtbxRepr.Text = r.nome_rep
        P_LocalizaOperacao r.tipocob, cmbxTipoCob
        
        
        
    End If
    
    
End Sub
Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Relatório de Comissões de " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, obj
    P_ImprLinha obj, "-"
    P_ImprX obj, "Vendedor", 10
    P_ImprX obj, "Vendido", 70, True, True
    P_ImprX obj, "Comissão", 100, True, True
    P_ImprX obj, "Perc.", 115, False, 3
    P_ImprLinha obj, "-"
    With dt.Recordset
        .MoveFirst
        Do While Not .EOF
            P_ImprX obj, !guerra_fun, 10
            P_ImprX obj, FormatVal(!total), 70, True, True
            P_ImprX obj, FormatVal(!comiss), 100, True, True
            P_ImprX obj, FormatVal(!total / CCur(lblVendas.caption) * 100), 115, False, 3

            .MoveNext
        Loop
        P_RodapeVis obj, constNome_sis
    End With
    
    
End Sub
Private Sub Form_Load()
    dt.Connect = Banco.Connect
    P_CentraForm Me
    P_PreencheTipoCob cmbxTipoCob
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
