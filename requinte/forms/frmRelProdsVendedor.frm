VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelProdsVendedor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vendas por Marca/Vendedor"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtbxSub 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   40
      TabIndex        =   33
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtbxSecao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      MaxLength       =   40
      TabIndex        =   32
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtbxExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9600
      MaxLength       =   40
      TabIndex        =   26
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtbxFab 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      MaxLength       =   40
      TabIndex        =   25
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico de Vendas"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdsVendedor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3120
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   825
         ScaleWidth      =   2145
         TabIndex        =   10
         Top             =   5640
         Width           =   2175
         Begin VB.Label lblDev 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   16
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Devol:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   540
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   45
            Width           =   585
         End
         Begin VB.Label lblVendas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   840
            TabIndex        =   13
            Top             =   -15
            Width           =   480
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comiss:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lblComissao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   11
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdsVendedor.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdsVendedor.frx":023F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdsVendedor.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbxFunc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Detalhes"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelProdsVendedor.frx":09D6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   20
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   870
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
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optnVendas 
      Caption         =   "Vendas"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   8640
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Diferença de Trocas"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   8640
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelProdsVendedor.frx":0AC0
      Height          =   7575
      Left            =   2520
      OleObjectBlob   =   "frmRelProdsVendedor.frx":0AD1
      TabIndex        =   21
      Top             =   840
      Width           =   10695
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   9720
      TabIndex        =   22
      Top             =   8400
      Width           =   3615
      Begin VB.OptionButton optnTipoPag 
         Caption         =   "Por Tipo Pag"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optnGeral 
         Caption         =   "Geral"
         Height          =   195
         Left            =   1680
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmndbtnBuscaProd 
      Height          =   285
      Left            =   12960
      Picture         =   "frmRelProdsVendedor.frx":19BA
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblSubCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Subseção:"
      Height          =   195
      Left            =   5040
      TabIndex        =   30
      Top             =   120
      Width           =   765
   End
   Begin VB.Label lblSecCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Seção:"
      Height          =   195
      Left            =   2640
      TabIndex        =   29
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E&xpressão/Produto:"
      Height          =   195
      Left            =   9600
      TabIndex        =   28
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricante/Marca:"
      Height          =   195
      Left            =   7440
      TabIndex        =   27
      Top             =   120
      Width           =   1320
   End
End
Attribute VB_Name = "frmRelProdsVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbxSec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxSec.ListIndex = -1
    End If
End Sub

Private Sub cmbxSub_Click()
    cmbxSub.tag = F_BuscaUmCampo("cod_sub", "tab_subsec", "nome_sub='" & cmbxSub.Text & "'")
End Sub

Private Sub cmbxSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxSub.ListIndex = -1
    End If
End Sub

Private Sub cmbxFunc_Click()
    cmbxFunc.tag = F_BuscaUmCampo("cod_fun", "tab_funcionarios", "guerra_fun='" & cmbxFunc.Text & "'")
End Sub

Private Sub cmndbtnBuscaProd_Click()
    Dim p As ProdutoSaida
    p = F_BuscaProdSaida(txtbxExpr.Text)
    If p.cod_pro <> "" Then
        txtbxExpr.Text = p.nome_pro
    
        txtbxExpr.tag = p.cod_pro
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    
    
    If cmbxFunc.Text <> strNome_Usu Then
        If strVisrelest_usu <> "S" Then
            MsgBox "Você não pode visualizar uma comissão de outra pessoa!", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
        strSQL = "select distinct tp.cod_pro, tp.nome_pro,sum(final_ite) as total,sum(ti.qtd_ite) as qtd,"
        strSQL = strSQL & " tf.cod_fun,tf.guerra_fun,tf.nome_fun from tab_vendas as tv,"
        strSQL = strSQL & " tab_funcionarios as tf,tab_itvendas as ti,tab_produtos as tp, "
        strSQL = strSQL & " tab_fornecedores as tfo, tab_secoes as ts, tab_subsec as tss "
        strSQL = strSQL & " where tv.abrevenda_usu=tf.cod_fun "
        strSQL = strSQL & " and tv.num_ven=ti.NUm_ven"
        strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
        strSQL = strSQL & " and tp.cod_fab=tfo.cod_for "
        strSQL = strSQL & " and tp.cod_sub=tss.cod_sub "
        strSQL = strSQL & " and tss.cod_sec=ts.cod_sec "
        
        
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and tv.dataefet_ven>=" & F_PesqData(txtbxDataIni.Text)
        End If
        If txtbxDataFin.Text <> "" Then
            strSQL = strSQL & " and tv.dataefet_ven<=" & F_PesqData(txtbxDataFin.Text)
        End If
        
        If txtbxExpr.tag = "" Then
            If txtbxExpr.Text <> "" Then
                strSQL = strSQL & " and nome_pro like '%" & txtbxExpr.Text & "%' "
            End If
        Else
            strSQL = strSQL & " and tp.cod_pro='" & txtbxExpr.tag & "' "
        End If
        If txtbxFab.tag <> "" Then
            strSQL = strSQL & " and tp.cod_Fab='" & txtbxFab.tag & "' "
        End If
        If txtbxSecao.tag <> "" Then
            strSQL = strSQL & " and ts.cod_sec='" & txtbxSecao.tag & "'"
        End If
        If txtbxSub.Text <> "" Then
            strSQL = strSQL & " and tp.cod_sub='" & txtbxSub.tag & "'"
        End If
        If cmbxFunc.Text <> "" Then
            strSQL = strSQL & " and abrevenda_usu='" & cmbxFunc.tag & "' "
        End If
        
        
        strSQL = strSQL & " and situacao_ven=4 "
        
        strSQL = strSQL & " group by tf.cod_fun,tp.cod_pro order by total desc"
        dt.RecordSource = strSQL
        dt.Refresh
        
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DbgrdToHtml "Relatório de Produtos Vendidos/Vendedores", dbgrd, dt, _
    "Sec: " & txtbxSecao.Text & " Sub: " & txtbxSub.Text & "Fab: " & txtbxFab.Text & _
    "  Prod: " & txtbxExpr.Text, 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dt.Connect = Banco.Connect
    P_CentraForm Me
    P_PreencheVendedor cmbxFunc, 1
    
    
    If strVisrelfin_usu <> "S" Then
        P_LocalizaOperacao strNome_Usu, cmbxFunc
        
    End If
End Sub

Private Sub txtbxdatafin_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdatafin_LostFocus()
    P_AjustaData txtbxDataFin
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

Private Sub txtbxExpr_Change()
    txtbxExpr.tag = ""
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturnt Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFor
    End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxFab.Text)
    If f.cod_for <> "" Then
        txtbxFab.Text = f.nome_for
        txtbxFab.tag = f.cod_for
        'busca
    End If
    
End Sub

Private Sub txtbxSecao_Change()
    txtbxSecao.tag = ""
End Sub

Private Sub txtbxSecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaSec
    End If
End Sub
Sub buscaSec()
    Dim s As secao.secao
    s = secao.F_BuscaSec(txtbxSecao.Text)
    txtbxSecao.Text = s.Nome_Sec
    txtbxSecao.tag = s.cod_sec
End Sub

Private Sub txtbxSub_Change()
    txtbxSub.tag = ""
End Sub

Private Sub txtbxSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaSub
    End If
End Sub

Sub buscaSub()
    Dim s As subsec
    s = F_BuscaSubSec(txtbxSub.Text, txtbxSecao.tag)
    If s.cod_sub <> "" Then
        txtbxSub.Text = s.nome_sub
        txtbxSub.tag = s.cod_sub
    End If
End Sub
