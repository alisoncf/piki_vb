VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmReemissao 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vendas na Data:"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9375
   Icon            =   "frmReemissao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pc 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   9195
      TabIndex        =   12
      Top             =   6240
      Width           =   9255
   End
   Begin VB.Frame frControle 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CheckBox chkUsu 
         Caption         =   "Vendas do Usuário"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4920
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   2025
         TabIndex        =   13
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox txtbxData 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmReemissao.frx":0442
         Height          =   5895
         Left            =   2280
         OleObjectBlob   =   "frmReemissao.frx":0459
         TabIndex        =   11
         Top             =   240
         Width           =   6855
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Buscar"
            Height          =   660
            Left            =   120
            Picture         =   "frmReemissao.frx":14FA
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "&Emitir"
            Height          =   660
            Left            =   120
            Picture         =   "frmReemissao.frx":15F3
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   660
            Left            =   120
            Picture         =   "frmReemissao.frx":197D
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnVis 
            Caption         =   "&Visualizar"
            Height          =   660
            Left            =   120
            Picture         =   "frmReemissao.frx":1A72
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1680
            Width           =   1815
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         Caption         =   "Caixa:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmReemissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTodas_Click()
    P_Busca
End Sub

Private Sub chkUsu_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmndbtnBuscar_Click()
    P_Busca
End Sub

Private Sub cmndbtnImpr_Click()
    Dim strSQL As String
    Dim sCPF As String
    sCPF = F_BuscaUmCampo("cpfcli_ven", "tab_vendas", "num_ven='" & dtVendas.Recordset!num_ven & "'")
    If MsgBox("Confirma reemissão da venda? CPF:" & sCPF, vbYesNo + vbQuestion) = vbYes Then
    
        
    
        If F_EmiteCupomVenda(dtVendas.Recordset!num_ven, 12, dtVendas.Recordset!nomecli_ven, Format(dtVendas.Recordset!cpfcli_ven), , IIf(pc.BackColor = vbBlue, True, False)) Then
            P_Status ""
            MsgBox "Cupom Impresso!", vbInformation
            P_BaixaSaida dtVendas.Recordset!num_ven
            P_GravaSaida dtVendas.Recordset!num_ven, strNum_Cpm, IIf(pc.BackColor = vbBlue, True, False)
        End If
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVis_Click()
On Error GoTo trata_erro:
    ImprVenda dtVendas.Recordset!num_ven
    Exit Sub
trata_erro:
P_MostraErro
    Exit Sub
End Sub
Function F_JustificaTexto(sText As String, tamanho As Integer) As String
    If Len(sText) >= tamanho Then
        sText = Left(sText, tamanho)
    Else
        sText = sText + Space(tamanho - Len(sText))
    End If
    F_JustificaTexto = sText
End Function

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    dbgrd.MarqueeStyle = 0
    P_Busca
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        pc.BackColor = vbGreen
        P_Busca
        cmndbtnImpr.SetFocus
        
    ElseIf KeyCode = vbKeyF9 Then
        pc.BackColor = vbYellow
        P_Busca
        cmndbtnImpr.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtbxData.Text = Date
    dtVendas.Connect = Banco.Connect
    pc.BackColor = vbGreen
    P_PreencheCaixas cmbxCaixa
    
        P_LocalizaOperacao strCod_Cai, cmbxCaixa
        
    'If strExclcai_usu <> "S" Then
    '    cmbxCaixa.Enabled = False
        'End If
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    On Error GoTo trata_erro
  
    
    strSQL = "select tc.nome_cli,tv.*,tf.guerra_fun,tf2.guerra_fun as nome "
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " tab_vendas as tv,tab_funcionarios as tf, tab_funcionarios as tf2 "
    If pc.BackColor = vbGreen Then
        strSQL = strSQL & ",tab_saidas as ts where ts.num_ven=tv.num_ven  and  "
    Else
        strSQL = strSQL & " where "
    End If
    
    strSQL = strSQL & " Dataefet_ven=" & F_GravaData(txtbxData.Text) & " "
    strSQL = strSQL & " and tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " and tf2.cod_fun=tv.fechavenda_usu  "
    strSQL = strSQL & " and tv.situacao_ven > 2 "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    If chkUsu.Value = 1 Then
        strSQL = strSQL & " and tf.cod_fun='" & strCod_Usu & "' "
    End If
    strSQL = strSQL & " and tv.cod_cai='" & Left(cmbxCaixa.Text, 2) & "'"
        
    strSQL = strSQL & " order by num_ven "
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    Exit Sub
trata_erro:
P_MostraErro
    
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        If dtVendas.Recordset.RecordCount > 0 Then
            dbgrd.SetFocus
        End If
        
        P_AjustaData txtbxData
    End If
    If KeyCode = vbKeyDown Then
        dbgrd.SetFocus
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
    P_Busca
End Sub

