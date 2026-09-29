VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiFrm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00755C33&
   Caption         =   "Estoque"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11400
   Icon            =   "mdiFrm_Vendas.frx":0000
   LinkTopic       =   "Geral"
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11370
      TabIndex        =   7
      Top             =   0
      Width           =   11400
      Begin MSComctlLib.TabStrip tbst 
         Height          =   390
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   688
         Style           =   2
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   840
   End
   Begin VB.PictureBox pcStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11400
      TabIndex        =   0
      Top             =   9000
      Width           =   11400
      Begin VB.TextBox txtbxShell 
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   11415
      End
      Begin VB.PictureBox pcRig 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   10200
         ScaleHeight     =   735
         ScaleWidth      =   1215
         TabIndex        =   4
         Top             =   0
         Width           =   1215
         Begin VB.Label lblHora 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   75
            TabIndex        =   5
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.Label lblCodUsu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   315
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   45
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Usuário"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5610
         TabIndex        =   3
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lblUsu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Usuário"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   315
         Width           =   1995
      End
      Begin VB.Label lblUs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   585
      End
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   4320
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnPrinc 
      Caption         =   "&Principal"
      Begin VB.Menu mnRecados 
         Caption         =   "Recados"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnEditCli 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mnProd 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu div4 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair do Piki Box"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnOrc 
      Caption         =   "&Vendas"
      Visible         =   0   'False
      Begin VB.Menu mnVendaAtac 
         Caption         =   "&Nova Venda"
      End
      Begin VB.Menu mnEdiVenAtac 
         Caption         =   "&Edição de Vendas"
      End
      Begin VB.Menu divOrc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnTroca 
         Caption         =   "&Troca"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnRel 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnRelCom 
         Caption         =   "&Acompanhamento de Comissoes"
      End
      Begin VB.Menu mnRelAcomp 
         Caption         =   "A&companhamento por Cliente"
      End
      Begin VB.Menu mnRelCliAniv 
         Caption         =   "Relatório de Clientes Aniversariantes"
      End
   End
   Begin VB.Menu mnParam 
      Caption         =   "Parâ&metros"
      Visible         =   0   'False
      Begin VB.Menu mnAdm 
         Caption         =   "&Aministração do Sistema"
         Enabled         =   0   'False
         Begin VB.Menu mnSQLPlu 
            Caption         =   "&SQL Plus - Console SQL"
         End
         Begin VB.Menu mnBAckUp 
            Caption         =   "&BackUp"
         End
         Begin VB.Menu mnTransf 
            Caption         =   "&Utilitário de Transferência de Dados"
         End
         Begin VB.Menu mnMigra 
            Caption         =   "&Utilitário de Migração"
         End
      End
      Begin VB.Menu mnUtilECF 
         Caption         =   "&Utilitário de ECF"
      End
      Begin VB.Menu mnImp 
         Caption         =   "&Impostos"
      End
      Begin VB.Menu mnCont 
         Caption         =   "C&ontagens"
      End
      Begin VB.Menu mnCaixas 
         Caption         =   "&Caixas"
      End
      Begin VB.Menu mnDePara 
         Caption         =   "&De Para"
      End
      Begin VB.Menu mnCalMed 
         Caption         =   "&Cálculo de Média Diária"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnConfig 
         Caption         =   "&Configurações do Sistema"
      End
      Begin VB.Menu mnUsu 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnProm 
         Caption         =   "&Promoções"
      End
      Begin VB.Menu mnMescComp 
         Caption         =   "&Mesclar Componentes"
      End
      Begin VB.Menu mnMudSen 
         Caption         =   "&Mudar Senha"
      End
      Begin VB.Menu mnGeraEti 
         Caption         =   "&Gerar Etiquetas"
      End
      Begin VB.Menu div5 
         Caption         =   "-"
      End
      Begin VB.Menu mnFormasPag 
         Caption         =   "&Formas de Pagamento"
      End
      Begin VB.Menu mnApar 
         Caption         =   "&Aparência"
      End
      Begin VB.Menu mnCalc 
         Caption         =   "Ca&lculadora"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnJan 
      Caption         =   "&Janelas"
      WindowList      =   -1  'True
      Begin VB.Menu mnCasc 
         Caption         =   "&Cascata"
      End
      Begin VB.Menu mnLado 
         Caption         =   "&Lado a Lado"
      End
      Begin VB.Menu mnTodasJan 
         Caption         =   "&Todas"
      End
      Begin VB.Menu mnDescanso 
         Caption         =   "Descanso de Tela"
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "A&juda"
      Begin VB.Menu mnAjuda 
         Caption         =   "&Indices da Ajuda"
      End
      Begin VB.Menu mnCred 
         Caption         =   "&Créditos"
      End
   End
End
Attribute VB_Name = "mdiFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub mnCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub cmndbtnCaixa_Click()
    frmSaidas.Show 0
    frmSaidas.SetFocus
End Sub

Private Sub cmndbtnCompras_Click()
    frmEdiEntr.Show 0
End Sub

Private Sub cmndbtnECF_Click()
    On Error Resume Next
    frmImpressoraECF.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmndbtnFechar_Click()

End Sub

Private Sub cmndbtnDocs_Click()
    frmEditaMov.Show 0
End Sub

Private Sub cmndbtnJan_Click()
    frmEdiFin.Show 0
End Sub

Private Sub cmndbtnNovaVenda_Click()
    mnEdiVen_Click
End Sub

Private Sub cmndbtnSair_Click()
    mnSair_Click
End Sub

Private Sub cmndbtnUsu_Click()
    frmOptUsu.Show 0, mdiFrm
End Sub

Private Sub mnAjEstRes_Click()
    
End Sub

Private Sub mnArr_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiArr.Show 0
    frmEdiArr.SetFocus
End Sub


Private Sub MDIForm_Resize()
    pcRig.Left = Me.Width - pcRig.Width - 100
    lblMsg.Left = (Me.Width / 2) - (lblMsg.Width / 2)
    txtbxShell.Width = Me.Width - 100
    tbst.Width = Me.Width - 100
    
    
End Sub

Private Sub mnAbreCx_Click()
    frmAberturaCaixa.Show 0, mdiFrm
End Sub

Private Sub mnBol_Click()
    frmBoletos.Show 0
End Sub

Private Sub mnBuscaOrc_Click()
    frmEdiOrc.Show 0
    frmEdiOrc.SetFocus
End Sub

Private Sub mnCaixas_Click()
    If strFechames_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbCritical
        Exit Sub
    
    End If
    frmCaixas.Show 0
End Sub

Private Sub mnCalc_Click()
    On Error Resume Next
    Shell "C:\windows\calc.exe", vbNormalFocus
    Shell "C:\windows\system32\calc.exe", vbNormalFocus
End Sub

Private Sub mnCalMed_Click()
    If strViscusto_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para esta operação!", vbExclamation
        Exit Sub
        
    End If
    frmMediaConsumo_Mer.Show 1
End Sub

Private Sub mnCOmport_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmComport.Show 0
End Sub

Private Sub mnCompProCli_Click()
    frmComportProCli.Show 0
    frmComportProCli.SetFocus
End Sub

Private Sub mnCons_Click()
    If strFechames_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmFechaFinan.Show 1
    End If
End Sub

Private Sub mnConsPre_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmConsPreco.Show 0
    frmConsPreco.SetFocus
End Sub

Private Sub mnCont_Click()
    frmContagens.Show 0
    frmContagens.SetFocus
End Sub

Sub mnCOntEcf_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmImpressoraECF.Show 1
End Sub

Private Sub cmndbtnCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub cmndbtnDev_Click()
    'frmDevolucoes.Show 0
End Sub

Private Sub cmndbtnNewVEnda_Click()
    On Error Resume Next
    frmVenda.Show 1
End Sub

Private Sub cmndbtnPEd_Click()
    frmEdiPed.Show 0
End Sub

Private Sub cmndbtnProd_Click()
    frmEdiPro.Show 0
    frmEdiPro.SetFocus
End Sub

Private Sub cmndbtnVendas_Click()
    mnNovaVenda_Click
End Sub

Private Sub MDIForm_Activate()
     On Error Resume Next
    VerificaAtualizacao
    frmMenu.Show 0
    Me.caption = constNome_sis & " " & constVersao_Sis & " - " & strNome_Emp & " - " & strCidade_Emp & "/" & strUF_Emp & " .:. CNPJ.: " & strCNPJ_Emp
    'TaskBar False
   ' If F_BuscaUmCampo("cod_loj", "tab_parametros", "") = "" Then
    '    MsgBox "Os Dados desta Loja não estão definidos corretamente!" & Chr(13) & _
     '       "Informe esta mensagem ao Responsável!", vbExclamation
      '  P_Fim
        'End If
   StatusBar1.Panels(6).Text = Winsock1.LocalIP & "-" & Winsock1.LocalHostName
   strIP_Usu = Winsock1.LocalIP
   strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixas", "ip_cai='" & strIP_Usu & "'")
   
   
   strImprFiscal = F_BuscaUmCampo("impressora_cai", "tab_caixas", "ip_cai='" & strIP_Usu & "'")
   If strImprFiscal = "" Then
        strImprFiscal = "B"
   End If
   If strCod_Cai = "" Then
        strCod_Cai = "00"
   End If
   strCod_Tur = F_BuscaUmCampo("Cod_tur", "tab_caixas", "cod_Cai='" & strCod_Cai & "'")
    If bActive <> True Then
'        frmMenu.Show 0, Me
'        frmMenu.Left = 0
'        frmMenu.Top = 600
'        frmMenu.Height = StatusBar1.Top - 50
        bActive = True
        
    End If
    'pcCalend.Top = pcStatus.Top - pcCalend.Height
    lblDia.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase)
    lblNumDia.caption = Day(Date)
    lblMes.caption = StrConv(MonthName(Month(Date)), vbProperCase)
    lblAno.caption = Year(Date)
    
    
    
End Sub

Private Sub MDIForm_DblClick()
    'frmVenda.Show 0, Me
End Sub

Private Sub MDIForm_Load()
'    pcCalend.Top = pcStatus.Top - pcCalend.Height
    
    'frmLogo.Show 0
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnPrinc
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'TaskBar True
    P_Fim
End Sub

Private Sub mnAjuda_Click()
    MsgBox "em breve"
End Sub

Private Sub mnApar_Click()
    frmConfig.Show 0
End Sub

Private Sub mnBan_Click()
    frmEdiBan.Show 0
End Sub

Private Sub mnCadDiv_Click()
    frmEdiDiv.Show 0
End Sub

Private Sub mnCadEsp_Click()
    frmEdiMar.Show 0
End Sub

Private Sub mnCadGen_Click()
    frmEdiCat.Show 0
End Sub

Private Sub mnCadCargos_Click()
    frmEdiCargos.Show 0
End Sub

Private Sub mnCadCat_Click()
    frmEdiCat.Show 0
End Sub

Private Sub mnCadMar_Click()
    frmEdiMar.Show 0
End Sub

Private Sub mnCadPla_Click()
    frmEdiPlanoContas.Show 0
    frmEdiPlanoContas.SetFocus
End Sub

Private Sub mnCalGuias_Click()
    frmLanDoc.Show 0
End Sub

Private Sub mnCasc_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnCheq_Click()
    frmControleCheques.Show 0
End Sub

Private Sub mnComDes_Click()
    frmEdiDes.Show 0
End Sub

Private Sub mnCOn_Click()
    frmEdiCon.Show 0
End Sub

Private Sub mnConfig_Click()
    frmParametros.Show 0
End Sub

Private Sub mnContasPagar_Click()
    frmEditaMov.Show 0
End Sub

Private Sub mnConv_Click()
    frmEdiConv.Show 0
End Sub

Private Sub mnConveniados_Click()
    frmEdiConveniados.Show 0
End Sub

Private Sub mnCred_Click()
    On Error Resume Next
    frmCred.Show 1
End Sub

Private Sub mnDemoInv_Click()
    frmDemonstrativoInvent.Show 0
    
End Sub

Private Sub mnDePara_Click()
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    frmDePara.Show 1
End Sub

Private Sub mnDescanso_Click()
    frmCaixaLivre.Show 1
End Sub

Private Sub mnDEv_Click()
    Dim strSQL As String
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    
    
    
    
    'frmDevolucoes.caption = "STYLO PRÓPRIO -Devolução a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
  
    On Error Resume Next
    frmDevolucoes.Show 1
    
End Sub

Private Sub mnEcf_Click()
    frmECF.Show 0, mdiFrm
End Sub

Private Sub mnEdiComp_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiEntr.Show 0
End Sub


Private Sub mnEdiMovFin_Click()
    
    frmEdiMovFin.Show 0
End Sub

Private Sub mnEdiMovInt_Click()
    frmEdiMovInt.Show 0
    frmEdiMovInt.SetFocus
End Sub

Private Sub mnEdiNotaSai_Click()
    frmEdiEspelhoNota.Show 0
    frmEdiEspelhoNota.SetFocus
End Sub

Private Sub mnEdiPer_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiPerdas.Show 0
    frmEdiPerdas.SetFocus
End Sub

Private Sub mnEdiProcomp_Click()
    frmEdiProcomp.Show 0
    frmEdiProcomp.SetFocus
End Sub

Private Sub mnEditCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub mnEditLan_Click()
     
    frmEditaMov.Show 0
End Sub

Private Sub mnEdiTransf_Click()
    frmEdiTransf.Show 0
    frmEdiTransf.SetFocus
End Sub

Private Sub mnEdiTro_Click()
    frmEdiTrocas.Show 0
    frmEdiTrocas.SetFocus
End Sub

Private Sub mnEdiUsu_Click()
    frmEdiUsu.Show 0
End Sub

Private Sub mnEdiVen_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiVendas.Show 0
End Sub

Private Sub mnEfetivVen_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    If IsLoad("frmEfetivaVen") Then
        frmEfetivaVen.SetFocus
    Else
        frmEfetivaVen.Show 0
    End If
End Sub

Private Sub mnEdiVenAtac_Click()
    frmEdiVendas.Show 0
End Sub

Private Sub mnEfetVen_Click()
    frmEfetivaVen.Show 0
End Sub

Private Sub mnEntrada_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmNotaEntrada.Show 1
End Sub

Private Sub mnEntrPed_Click()
    frmEfetivaPed.Show 0
End Sub

Private Sub mnEntrVen_Click()
    frmEntrega.Show 0
End Sub

Private Sub mnEstIni_Click()
    On Error Resume Next
    frmAjEstIni.Show 1
End Sub

Private Sub mnEti_Click()
    P_ImprxEstiquetas frmVisualiza.Pic
    frmVisualiza.Show 1
End Sub

Private Sub mnFechaCaixa_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim fr As New frmLanCaixa
    fr.Show 0
End Sub

Private Sub mnEspelho_Click()
    frmEspelhoNota.Show 0
    frmEspelhoNota.SetFocus
End Sub

Private Sub mnEtiquetas_Click()
    frmGeraEtiquetas.Show 0
End Sub

Private Sub mnFabr_Click()
    frmEdiFab.Show 0
End Sub

Private Sub mnFinanceiras_Click()
    frmEdiFinan.Show 0
End Sub

Private Sub mnFormasPag_Click()
    frmEdiFormasPag.Show 0
End Sub

Private Sub mnForn_Click()
    frmEdiFor.Show 0
End Sub

Private Sub mnIncCli_Click()
    frmIncClientes.Show 0
End Sub

Private Sub mnIncTer_Click()
    frmIncTer.Show 0
End Sub



Private Sub mnLivIcms_Click()
    frmLivIcms.Show 0
End Sub

Private Sub mnLivIpi_Click()
    frmLivIpi.Show 0
End Sub

Private Sub mnMudaCli_Click()
    On Error Resume Next
    frmSelectCliente.Show 1
End Sub

Private Sub mnPer_Click()
    frmPerfis.Show 0
End Sub

Private Sub mnFreteiros_Click()
    frmEdiFreteiros.Show 0
    frmEdiFreteiros.SetFocus
End Sub

Private Sub mnFun_Click()
    frmEdiFun.Show 0
End Sub

Private Sub mnGeraEti_Click()
    frmGeraEtiquetas.Show 0
    frmGeraEtiquetas.SetFocus
End Sub

Private Sub mnGeraNota_Click()
    frmGeraNota.Show 1
End Sub

Private Sub mnHistCheCPf_Click()
    frmHistoricoChe.Show 0, mdiFrm
End Sub

Private Sub mnImo_Click()
    frmEdiImo.Show 0
End Sub

Private Sub mnGerarEti_Click()
    frmGeraEtiquetas.Show 0
End Sub

Private Sub mnImp_Click()
    frmImpost.Show 1
End Sub

Private Sub mnIncMer_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim fr As New frmEntradas
    fr.Show 0, mdiFrm
End Sub

Private Sub mnInvent_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmInventario.Show 0, mdiFrm
    frmInventario.SetFocus
End Sub

Private Sub mnLado_Click()
    Me.Arrange vbVertical
End Sub

Private Sub mnLanCaixa_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiCai.Show 0
End Sub

Private Sub mnLojas_Click()
    frmEdiLoj.Show 0
End Sub

Private Sub mnLanCont_Click()
    frmLanCont.Show 1
End Sub

Private Sub mnLoj_Click()
    frmEdiLoj.Show 0
    frmEdiLoj.SetFocus
End Sub

Private Sub mnMescComp_Click()
    frmMesclaComp.Show 0, mdiFrm
End Sub

Private Sub mnMonit_Click()
    frmMonitoramentoCaixas.Show 0, mdiFrm
    frmMonitoramentoCaixas.SetFocus
End Sub

Private Sub mnMovFin_Click()
    frmEdiFin.Show 0
    frmEdiFin.SetFocus
End Sub

Private Sub mnMovUsu_Click()
    frmFechaCaixaUsu.Show 0, mdiFrm
End Sub

Private Sub mnMudSen_Click()
    On Error Resume Next
    frmMudaSenha.Show 1
End Sub

Private Sub mnNovaVenda_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = F_GeraNovaVenda("00000")
    
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    'On Error Resume Next
    frmVenda.Atacado = False
    frmVenda.Show 0, mdiFrm
    
End Sub

Private Sub mnPed_Click()
    frmEdiPed.Show 0
End Sub

Private Sub mnPisos_Click()
    frmEstoqueIni.Show 1
End Sub

Private Sub mnNFDev_Click()
    frmGeraNotaAvulsa.Show 1
End Sub

Private Sub mnNovoMovInt_Click()
   ' Dim fr As New frmMovInt
   ' fr.Show 0
End Sub

Private Sub mnNovoOrc_Click()
    'Dim fr As New frmOrcamento
    'fr.Show 0
End Sub

Private Sub mnPrincipio_Click()
    principio.P_EditaPrincipio
End Sub

Private Sub mnProd_Click()
    frmEdiPro.Show 0
End Sub

Private Sub mnRelGuias_Click()
    frmRelGuias.Show 0
End Sub

Private Sub mnRecCli_Click()
    frmRecebimentos.Show 0
    frmRecebimentos.SetFocus
End Sub

Private Sub mnRelAltPre_Click()
    frmAltPreco.Show 0, mdiFrm
End Sub

Private Sub mnRelComp_Click()
If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
End Sub

Private Sub mnProdProm_Click()
    
        frmProdProm.Show 0
        frmProdProm.SetFocus
    
    
End Sub

Private Sub mnProm_Click()
    
        
        frmProdProm.Show 0
    
End Sub

Private Sub mnRelAnaFat_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmRelFatMensal.Show 0
        frmRelFatMensal.SetFocus
    End If
End Sub

Private Sub mnRelAnaIndFluxo_Click()
    If strVisrelest_usu = "S" Then
        frmRelAnaliseIndFluxo.Show 0
        frmRelAnaliseIndFluxo.SetFocus
    Else
        MsgBox "Acesso Inválido!", vbCritical
    End If
End Sub

Private Sub mnRelAtvCli_Click()
    If strVisrelfin_usu = "S" Then
        frmRelAtvCli.Show 0
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mNRelAtvPro_Click()
    If strVisrelfin_usu = "S" Then
        frmRelAtvCliProd.Show 0
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mnRelChq_Click()
    frmRelCheques.Show 0
    frmRelCheques.SetFocus
End Sub

Private Sub mnRelCmpFor_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelProdFor.Show 0
    frmRelProdFor.SetFocus
End Sub

Private Sub mnRelCOmiss_Click()
        If strVisrelfin_usu <> "S" Then
        MsgBox "Você não tem permissão para este aesso!", vbExclamation
    Else
        frmRelComiss.Show 0
        frmRelComiss.SetFocus
    End If
End Sub

Private Sub mnRelCompEst_Click()
    If strVisrelest_usu = "S" Then
        frmRelEstFis.Show
    End If
End Sub

Private Sub mnRelDev_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelDevProd.Show 0
End Sub

Private Sub mnRelFatVenda_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmRelFat.Show 0
        frmRelFat.SetFocus
    End If
End Sub

Private Sub mnRelFluxo_Click()
    If strVisrelest_usu = "S" Then
        frmRelAnaliseFluxo.Show 0
        frmRelAnaliseFluxo.SetFocus
    Else
        MsgBox "Acesso Inválido!", vbCritical
    End If
End Sub

Private Sub mnRelGiro_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelGiro.Show 0
    frmRelGiro.SetFocus
End Sub

Private Sub mnRelLucrComp_Click()
    If strVisrelest_usu = "S" Then
        frmRelLucrComp.Show 0
        frmRelLucrComp.SetFocus
    End If
End Sub

Private Sub mnRelMMD_Click()
    frmRelMMD.Show 0
    frmRelMMD.SetFocus
End Sub

Private Sub mnRelPagReceb_Click()
    frmRelMov.Show 0
    frmRelMov.SetFocus
End Sub

Private Sub mnRelProdForn_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelFornProd.Show 0
    frmRelFornProd.SetFocus
End Sub

Private Sub mnRelProdLucr_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbCritical
        Exit Sub
    End If
    frmRelProdLucr.Show 0
End Sub

Private Sub mnRelReceitasDesp_Click()
    frmRelReceitasDespesas.Show 0
End Sub

Private Sub mnRelVend_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelSaiCaixa.Show 0
    frmRelSaiCaixa.SetFocus
End Sub

Private Sub mnTer_Click()
    If strCOD_CLI = "" Then
        If MsgBox("Não há cliente selecionado." & Chr(13) & "Deseja selecionar um agora?" _
            , vbYesNo + vbQuestion) = vbYes Then
            On Error Resume Next
            frmSelectCliente.Show 1
            If strCOD_CLI = "" Then
                MsgBox "Não foi selecionado nehum cliente", vbExclamation
            Else
                frmSelectTer.Show
            End If
            
        Else
            MsgBox "Não posso manipular terceiros sem que você escolha um cliente", vbCritical
            Exit Sub
        End If
    Else
        frmSelectTer.Show 0
    End If
    
End Sub

Private Sub mnSaiDia_Click()
    On Error Resume Next
    frmSaidaDiaria.Show 1
End Sub

Private Sub mnResumo_Click()
    frmResumoContasPagar.Show 0
End Sub

Private Sub mnRelVendas_Click()
    frmRelComiss.Show 0
End Sub

Private Sub mnRecados_Click()
    frmNovoRecado.Show 0, mdiFrm
    
End Sub

Private Sub mnRelAcomp_Click()
    frmRelVendasVendedor.Show 0
End Sub

Private Sub mnRelCliAniv_Click()
    frmCliAniv.Show 0
End Sub

Private Sub mnRelCom_Click()
    frmRelComiss.Show 0
    
End Sub

Private Sub mnSair_Click()
    If MsgBox("Deseja Realmente Sair?", vbYesNo + vbQuestion) = vbYes Then
        P_Fim
    Else
        Exit Sub
    End If
End Sub

Private Sub mnServicos_Click()
    frmEdiServ.Show 0
End Sub

Private Sub mnSec_Click()
    secao.P_EditaSecao
End Sub

Private Sub mnSubSec_Click()
    frmEdiSubSec.Show 0
    frmEdiSubSec.SetFocus
End Sub

Private Sub mnSujComp_Click()
    If strVisrelest_usu = "S" Then
        frmSujComp.Show 0
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
    End If
    
End Sub

Private Sub mnTodasJan_Click()
    'On Error Resume Next
   
   If Forms.count > 1 Then
        frmForms.Show , mdiFrm
    End If
   
End Sub

Private Sub mnTransf_Click()
    On Error Resume Next
    frmTransf.Show 1
End Sub

Private Sub mnTroca_Click()
    If strAbresaida_usu = "S" Then
        On Error Resume Next
        frmTroca.Show 1
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mnTransfComp_Click()
    frmTransfComp.Show 0, mdiFrm
    
End Sub

Private Sub mnUsu_Click()
    Dim fr As New frmIncUsu
    fr.Show 0
End Sub

Private Sub mnUtilECF_Click()
    P_PortaComunicacaoECF
    frmUtilitarioECF.Show 1
End Sub

Sub mnVendaAtac_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = Vendas.F_GeraNovaVenda("00000")
    
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
    
    frmVenda.Show 0, mdiFrm
    

End Sub

Private Sub mnVendasPorCli_Click()
    frmRelVendasVendedor.Show 0
    frmRelVendasVendedor.SetFocus
End Sub

Private Sub mnVinc_Click()
    frmVinculaDoc.Show 1
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
     If Panel.Index = 1 Then
        mnMudaCli_Click
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim iCount As Integer
    If Button.Index = 1 Then
        mnEditCli_Click
    ElseIf Button.Index = 2 Then
        mnForn_Click
    ElseIf Button.Index = 3 Then
        mnProd_Click
    ElseIf Button.Index = 4 Then
        mnEdiVen_Click
    ElseIf Button.Index = 5 Then
        mnNovaVenda_Click
    ElseIf Button.Index = 6 Then
        mnLanCaixa_Click
    ElseIf Button.Index = 7 Then
        mnEdiComp_Click
    ElseIf Button.Index = 8 Then
        mnServicos_Click
    ElseIf Button.Index = 9 Then
        If Forms.count = 1 Then
            mnSair_Click
        Else
            'For iCount = 1 To Forms.count - 1
                
               ' MsgBox Forms(iCount).Name
                 Unload Screen.ActiveForm
            'Next
        End If
    End If
    
End Sub



Private Sub tbst_Click()
    Dim iCount As Integer
    On Error GoTo trata_Erro
    For iCount = 1 To Forms.count - 1
        If Forms(iCount).Name = tbst.SelectedItem.tag Then
            Forms(iCount).SetFocus
        End If
    Next
    Exit Sub
trata_Erro:
End Sub


Private Sub Timer1_Timer()
    lblHora.caption = Time
    buscaFrm
End Sub

Private Sub txtbxShell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_shell
    End If
    
End Sub
Sub P_shell()
    Dim s As String
    s = LCase(txtbxShell.Text)
    If s = "vendas" Then
        
    End If
End Sub
Sub buscaFrm()
    On Error Resume Next
    tbst.Tabs.Clear
    
    Dim iCount As Integer
    For iCount = 1 To Forms.count - 1
        If Forms(iCount).Visible = True Then
            tbst.Tabs.Add , , Forms(iCount).caption
            tbst.Tabs(iCount).tag = Left(Forms(iCount).Name, 14)
        End If
    Next
End Sub
