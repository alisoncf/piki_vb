VERSION 5.00
Begin VB.Form frmTroco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troco"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmTroco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "frmTroco.frx":0CCA
      ScaleHeight     =   735
      ScaleWidth      =   9735
      TabIndex        =   30
      Top             =   0
      Width           =   9735
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Troco"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   720
         TabIndex        =   31
         Top             =   -120
         Width           =   1680
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   9615
      TabIndex        =   27
      Top             =   840
      Width           =   9615
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Left            =   6600
         Picture         =   "frmTroco.frx":16EC
         ScaleHeight     =   1800
         ScaleWidth      =   3000
         TabIndex        =   35
         Top             =   480
         Width           =   3000
         Begin VB.Label lblTroco 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   600
            Left            =   1080
            TabIndex        =   36
            Top             =   480
            Width           =   945
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1650
         Left            =   3480
         Picture         =   "frmTroco.frx":8BBF
         ScaleHeight     =   1650
         ScaleWidth      =   3000
         TabIndex        =   34
         Top             =   480
         Width           =   3000
         Begin VB.TextBox txtbxDinheiro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00A23202&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   720
            Left            =   600
            TabIndex        =   1
            Text            =   "0,00"
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Left            =   120
         Picture         =   "frmTroco.frx":10637
         ScaleHeight     =   1800
         ScaleWidth      =   3000
         TabIndex        =   32
         Top             =   480
         Width           =   3000
         Begin VB.Label lblValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   600
            Left            =   1080
            TabIndex        =   33
            Tag             =   "!"
            Top             =   480
            Width           =   945
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Dinheiro:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4320
         TabIndex        =   0
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Troco:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   7680
         TabIndex        =   29
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Final:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         TabIndex        =   28
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.CheckBox chkSEPD 
      Caption         =   "Emitir SEPD"
      Height          =   735
      Left            =   1320
      Picture         =   "frmTroco.frx":17A11
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Imprime Nota Fiscal"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.PictureBox pcCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6DADD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   5775
      TabIndex        =   6
      Top             =   3720
      Width           =   5775
      Begin VB.TextBox txtbxRestante 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4680
         TabIndex        =   25
         Text            =   "0,00"
         Top             =   1155
         Width           =   735
      End
      Begin VB.TextBox txtbxEmUso 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   1155
         Width           =   735
      End
      Begin VB.TextBox txtbxLImite 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   1155
         Width           =   735
      End
      Begin VB.CommandButton cmndbtnInfo 
         Caption         =   "Informações do Cliente"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Top             =   435
         Width           =   2055
      End
      Begin VB.TextBox txtbxNome_Cli 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   8
         Tag             =   "00000"
         Text            =   "Venda A Consumidor"
         Top             =   75
         Width           =   4215
      End
      Begin VB.TextBox txtbxCNPJ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   10
         Top             =   435
         Width           =   2055
      End
      Begin VB.TextBox txtbxEndereco 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   12
         Top             =   795
         Width           =   4215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restante:"
         Height          =   195
         Left            =   3840
         TabIndex        =   24
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em Uso:"
         Height          =   195
         Left            =   2160
         TabIndex        =   22
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - C&liente:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF/C&NPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CheckBox chkEmite 
      Caption         =   "Romaneio"
      Height          =   735
      Left            =   120
      Picture         =   "frmTroco.frx":17B0D
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txtbxDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6120
      TabIndex        =   15
      Text            =   "0,00"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnDesconto 
      Caption         =   "De&sconto:"
      Height          =   285
      Left            =   6120
      TabIndex        =   14
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncelar (F12)"
      Height          =   735
      Left            =   8160
      Picture         =   "frmTroco.frx":17D69
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "C&onfirmar (F7)"
      Height          =   735
      Left            =   6480
      Picture         =   "frmTroco.frx":182B7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   120
      ScaleHeight     =   15
      ScaleWidth      =   10335
      TabIndex        =   13
      Top             =   5520
      Width           =   10335
   End
   Begin VB.OptionButton optnCH 
      BackColor       =   &H00C6DADD&
      Caption         =   "Venda &Condicional (F6)"
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Width           =   7935
   End
   Begin VB.OptionButton optnAv 
      BackColor       =   &H00C6DADD&
      Caption         =   "&A Vista (F5)"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Label lblAvisoAtacado 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Atenção. Vendas em atacado apenas pagamento em dinheiro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   37
      Top             =   5520
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.Label lblNome_ven 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   17
      Top             =   3960
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atendente:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6120
      TabIndex        =   16
      Top             =   3720
      Width           =   780
   End
End
Attribute VB_Name = "frmTroco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public operacao As String
Dim sBloqueado_cai As String
Public bTEF As Boolean 'essa merda aqui é um gato, pra economizar muito trabalho por causa do tef
'B - buffer, C - concomitante
Sub buscaVendaAtacado()
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("seq_ite", "tab_itvendas", "num_ven='" & Me.tag & "' and atacado='S'")
    If strSQL <> "" Then
        optnCH.Enabled = False
        lblAvisoAtacado.Visible = True
    End If
End Sub
Private Sub Command1_Click()
    
End Sub

Private Sub cmndbtnAv_Click()

End Sub


Private Sub cmndbtnCancel_Click()
    frmVenda.Enabled = True
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnDesconto_Click()
    Dim sUsu As String
    sUsu = F_SenhaCodBar
    If sUsu <> "" Then
        txtbxDesconto.Enabled = True
        cmndbtnDesconto.tag = sUsu
        txtbxDesconto.SetFocus
    End If
    
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    bTEF = False
    On Error GoTo trata_Erro
    If CCur(lblTroco.caption) < 0 Then
        MsgBox "Informe o Valor Pago!", vbExclamation
        txtbxDinheiro.SetFocus
        Exit Sub
       
    End If
    
    If F_Vazio(txtbxNome_Cli.Text) Then
        MsgBox "Informe o nome do cliente!", vbExclamation
        txtbxNome_Cli.SetFocus
        Exit Sub
    End If
    If txtbxNome_Cli.tag = "" Then
        MsgBox "Escolha o Cliente", vbExclamation
        txtbxNome_Cli.SetFocus
        Exit Sub
        
    End If
    If optnAv.Value Then
            strSQL = "delete from tab_caixa where num_ven='" & Me.tag & "'"
            conn.Execute strSQL
            strSQL = F_GravaCai("C", Date, CCur(lblValor.caption), "VENDA: " & Me.tag, _
                Me.tag, "DIN", "01", CCur(txtbxDinheiro.Text), CCur(lblTroco.caption), _
                strCod_Cai, strCod_Tur)
        If frmVenda.pcBaixa.Visible = True Then
            
        strSQL = F_GravaCai("D", Date, CCur(lblValor.caption), "F10-VENDA: " & Me.tag, _
                Me.tag, "DIN", "01", CCur(txtbxDinheiro.Text), CCur(lblTroco.caption), _
                strCod_Cai, strCod_Tur)
            
        End If
        strSQL = "update tab_vendas set nomecli_ven = '" & txtbxNome_Cli.Text & "', "
        strSQL = strSQL & " cod_cli='" & txtbxNome_Cli.tag & "',"
        strSQL = strSQL & "valor_ven=" & F_GravaDec(CCur(lblValor.caption))
        strSQL = strSQL & ",situacao_ven='4' " 'efetivada
        strSQL = strSQL & ", DATaefet_ven=" & F_GravaData(Date)
        strSQL = strSQL & ",desconto_ven=" & F_GravaDec(CCur(txtbxDesconto.Text))
        strSQL = strSQL & ",cod_Ven='" & lblNome_ven.tag & "' "
        strSQL = strSQL & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        
        strSQL = "delete from tab_venfec where num_ven='" & Me.tag & "'"
        conn.Execute strSQL
        strSQL = "insert into tab_venfec(num_ven,avista_ven,COD_CAI,cod_tur,valorfinal_ven)values("
        strSQL = strSQL & "'" & Me.tag & "',"
        If frmVenda.pcBaixa.Visible = False Then
            strSQL = strSQL & F_GravaDec(lblValor.caption) & ","
        Else
            strSQL = strSQL & "0,"
        End If
        strSQL = strSQL & "'" & strCod_Cai & "', "
        strSQL = strSQL & "'" & strCod_Tur & "', "
        If frmVenda.pcBaixa.Visible = False Then
            strSQL = strSQL & F_GravaDec(lblValor.caption) & ")"
        Else
            strSQL = strSQL & "0)"
        End If
        conn.Execute strSQL
        
        
        P_BaixaVenda Me.tag
        If chkSEPD.Value = 1 Then
            If operacao = "C" Or operacao = "B" Then
                P_notaSEPD Me.tag, strNum_Cpm
            Else
                P_notaSEPD Me.tag, ""
            End If
        End If
        If chkEmite.Value Then
            'ImprVenda Me.tag
            Vendas.ImprimeComprovanteNaoFiscal Me.tag
        End If
        
        frmVenda.Enabled = True
    Else
        
        'If MsgBox("Deseja informar os Dados da Venda agora?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update tab_vendas set nomecli_ven = '" & txtbxNome_Cli.Text & "' "
            strSQL = strSQL & ",cod_cli='" & txtbxNome_Cli.tag & "' "
            'strSQL = strSQL & ",valor_ven=" & F_GravaDec(CCur(lblValor.caption))
            strSQL = strSQL & ",desconto_ven=" & F_GravaDec(CCur(txtbxDesconto.Text))
            strSQL = strSQL & ",situacao_ven='2' "
            strSQL = strSQL & ",cod_Ven='" & lblNome_ven.tag & "' "
            strSQL = strSQL & " where num_ven='" & Me.tag & "' "
            conn.Execute strSQL
            frmTipoVenda.operacao = "Venda"
            frmTipoVenda.lblNome_cli.caption = txtbxNome_Cli.Text
            frmTipoVenda.lblNome_cli.tag = txtbxNome_Cli.tag
            frmTipoVenda.tag = Me.tag
            frmTipoVenda.lblNum_Ven.caption = Me.tag
            frmTipoVenda.lblValor_Ven.caption = lblValor.caption
            frmTipoVenda.lblValorFinal.caption = lblValor.caption
            frmTipoVenda.lblVendedor.caption = strNome_Usu
            frmTipoVenda.txtbxDesconto.Text = FormatVal(txtbxDesconto.Text)
            frmTipoVenda.lblCpf.caption = txtbxCNPJ.Text
            frmTipoVenda.lblLimite.caption = txtbxLImite.Text
            frmTipoVenda.lblLimiteRest.caption = txtbxRestante.Text
            frmTipoVenda.lblLimiteEmUso.caption = txtbxEmUso.Text
            
            If CCur(txtbxRestante.Text) <= 0 Then
                frmTipoVenda.chkAprazo.Enabled = False
                
            End If
            
            frmTipoVenda.Show 0, mdiFrm
            frmTipoVenda.SetFocus
            Do While frmTipoVenda.Visible = True
                DoEvents
            Loop
            frmVenda.lblMensagem.caption = "Concluindo Venda..."
            If frmTipoVenda.tag = "" Then
                
                MsgBox "Os Valores não foram informados corretamente!", vbExclamation
                
                strSQL = "update tab_vendas set "
                strSQL = strSQL & "situacao_ven='1' "
                strSQL = strSQL & " where num_ven='" & Me.tag & "' "
                conn.Execute strSQL
                Me.tag = ""
                P_Status ""
                Unload frmTipoVenda
                frmVenda.lblMensagem.caption = ""
                Unload Me
                Exit Sub
            End If
            Unload frmTipoVenda
            If chkEmite.Value Then
                ImprimeComprovanteNaoFiscal Me.tag
            End If
        'End If
    End If
    If CCur(txtbxDesconto.Text) > 0 Then
        Dim cDesconto As Currency
        
        cDesconto = CCur(txtbxDesconto.Text) / CCur(lblValor.tag)
        strSQL = "update tab_itvendas set final_ite=(qtd_ite*unit_ite)-((qtd_ite*unit_ite)*" & F_GravaDec(cDesconto) & ")"
        strSQL = strSQL & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
    Else
        strSQL = "update tab_itvendas set final_ite=(qtd_ite*unit_ite) "
        strSQL = strSQL & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
    End If
    
    If pc.BackColor = vbGreen Or pc.BackColor = vbBlue And bTEF = False Then
        
        If operacao = "B" Then
            If F_EmiteCupomVenda(Me.tag, 17, txtbxNome_Cli.Text, _
                    txtbxCNPJ.Text, (CCur(txtbxDinheiro.Text)), IIf(pc.BackColor = vbBlue, True, False)) Then
                'MsgBox "Cupom impresso!", vbInformation
                
                'P_BaixaSaida Me.Tag
                'P_GravaSaida Me.Tag, strNum_Cpm, IIf(pc.BackColor = vbBlue, True, False)
                P_GravaCupom Me.tag
            Else
                
                'P_PortaComunicacaoECF
                If F_EmiteCupomVenda(Me.tag, 17, txtbxNome_Cli.Text, _
                    F_BuscaUmCampo("cpf_cli", "tab_Clientes", "cod_cli='" & _
                    txtbxNome_Cli.tag & "'"), CCur(txtbxDinheiro.Text), IIf(pc.BackColor = vbBlue, True, False)) Then
                    'MsgBox "Cupom impresso!", vbInformation
                End If
            End If
        ElseIf operacao = "C" Then
            If Me.tag = "" Then
                MsgBox "Não encontrada!", vbExclamation
                Unload Me
            End If
            P_Msg "Obrigado Pela Preferência...", "Finalizando Venda"
                        
            P_Msg "Finalizando Cupom Fiscal..." & strNum_Cpm
            
            frmVenda.lblMensagem.caption = "Finalizando Cupom Fiscal..." & strNum_Cpm
            frmVenda.Refresh
            'P_Msg "Finalizando Cupom Fiscal..." & strNum_Cpm, "Obrigado Pela Preferência"
            
            
            i = ecf_IniciaFechamentoCupom("D", "$", txtbxDesconto.Text)
            
            i = ecf_EfetuaFormaPagamento("Dinheiro", txtbxDinheiro.Text)
            
            strSQL = F_BuscaUmCampo("msgecf_par", "tab_parametros", "")
            strSQL = strSQL & Chr(13) & txtbxNome_Cli.Text & "-" & txtbxCNPJ.Text
            strSQL = strSQL & "-" & txtbxEndereco.Text
            strSQL = strSQL & Chr(13) + Chr(10) & "Caixa: " & strCod_Cai & " Atendente: " & strNome_Usu
            i = ecf_TerminaFechamentoCupom(strSQL)
            
            i = F_RetornoVariaveis
            
            'P_GravaSaida Me.Tag, strNum_Cpm, False
            P_GravaCupom Me.tag
            P_Msg ""
            frmVenda.lblMensagem.caption = ""
        End If
    Else
            
    End If
    
    
    
    intCUPOM = intCUPOM + 1
    If intCUPOM = 3 Then
        intCUPOM = 0
    End If
    P_Msg ""
    Me.tag = "X"
    Me.Hide
    frmVenda.Enabled = True
    Exit Sub
trata_Erro:
    MostraErro
End Sub

Private Sub Form_Activate()
    If operacao = "C" Then
        pc.BackColor = vbGreen
    End If
    txtbxDinheiro_GotFocus
    optnCH_Click
    sBloqueado_cai = F_BuscaUmCampo("bloqueado_cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
    If sBloqueado_cai = "S" Then
        If operacao = "B" Then
            pc.BackColor = vbGreen
        End If
    End If
    lblNome_ven.caption = strNome_Usu
    buscaVendaAtacado
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If operacao <> "C" Then
    If KeyCode = vbKeyF8 Then
        pc.BackColor = vbGreen
        cmndbtnOk.SetFocus
    ElseIf KeyCode = vbKeyF9 Then
        If sBloqueado_cai = "S" Then
            Exit Sub
        End If
        pc.BackColor = vbRed
        cmndbtnOk.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
        pc.BackColor = vbBlue
        cmndbtnOk.SetFocus
    ElseIf KeyCode = vbKeyF10 Then
        
    End If
    
        
    End If
    If KeyCode = vbKeyF5 Then
        optnAv.SetFocus
    End If
    If KeyCode = vbKeyF6 Then
        If optnCH.Enabled Then
            optnCH.SetFocus
        Else
            MsgBox "Esta opção não está acessível no momento!", vbExclamation
            Exit Sub
        End If
        
    End If
    If KeyCode = vbKeyF4 Then
        txtbxNome_Cli.SetFocus
    End If
    If KeyCode = vbKeyF7 Then
        cmndbtnOk_Click
    End If
    If KeyCode = vbKeyF12 Then
        cmndbtnCancel_Click
    End If
    
End Sub

Private Sub Form_Load()

    'pc.BackColor = vbGreen
    pc.BackColor = vbRed
    txtbxNome_Cli.tag = "00000"
    txtbxNome_Cli.Text = "VENDA A CONSUMIDOR"
End Sub

Private Sub optnAv_Click()
'    frNome_cli.Visible = False
    
End Sub

Private Sub optnAv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnCH_Click()
'    frNome_cli.Visible = True
    txtbxDinheiro.Text = lblValor.caption
    txtbxDinheiro_LostFocus
End Sub

Private Sub optnCH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnCred_Click()
    frNome_cli.Visible = True
End Sub

Private Sub optnCred_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnFin_Click()
    frNome_cli.Visible = True
End Sub

Private Sub optnFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       ' trocatabkey keycode
    End If
End Sub

Private Sub pcCli_GotFocus()
    txtbxNome_Cli.SetFocus
End Sub


Private Sub txtbxCNPJ_GotFocus()
    ganhaFoco txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    If Len(txtbxCNPJ.Text) > 14 Then
        txtbxCNPJ.Text = Format(txtbxCNPJ.Text, "0#\.###\.###\/####\-##")
    Else
        txtbxCNPJ.Text = Format(txtbxCNPJ.Text, "00#\.###\.###\-##")
    End If
End Sub

Private Sub txtbxDesconto_GotFocus()
    ganhaFoco txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    Dim cDesc As Currency
    On Error GoTo trata_Erro
    PerdeFoco txtbxDesconto
    cDesc = CCur(F_BuscaUmCampo("desconto_usu", "tab_defusu", _
        "cod_usu='" & cmndbtnDesconto.tag & "'"))
    
    If CCur(txtbxDesconto.Text) > CCur(lblValor.caption) * cDesc / 100 Then
        MsgBox "Desconto muito alto!", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    
    End If
    lblValor.caption = FormatVal(lblValor.tag) - CCur(txtbxDesconto.Text)
    txtbxDesconto.Enabled = False
    Exit Sub
trata_Erro:
    MostraErro
    txtbxDesconto.Text = "0"
End Sub

Private Sub txtbxDinheiro_GotFocus()
    ganhaFoco txtbxDinheiro
End Sub

Private Sub txtbxDinheiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDown Then
        txtbxNome_Cli.SetFocus
    End If
End Sub

Private Sub txtbxDinheiro_LostFocus()
    PerdeFoco txtbxDinheiro
    lblTroco.caption = FormatVal(CCur(txtbxDinheiro.Text) - CCur(lblValor.caption))
End Sub

Private Sub txtbxEndereco_GotFocus()
    ganhaFoco txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxEndereco_LostFocus()
    AjustaTexto txtbxEndereco
End Sub

Private Sub txtbxNome_Cli_GotFocus()
    ganhaFoco txtbxNome_Cli
End Sub

Private Sub txtbxNome_Cli_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim c As tCliente
    If KeyCode = vbKeyReturn Then
        c = F_BuscaCliente("C", txtbxNome_Cli.Text)
        txtbxNome_Cli.Text = c.nome_cli
        txtbxNome_Cli.tag = c.cod_cli
        txtbxCNPJ.Text = c.cnpj_Cli
        txtbxEndereco.Text = c.endereco_cli
        txtbxLImite.Text = FormatVal(c.limite_cli)
        txtbxEmUso.Text = FormatVal(F_RetornaLimiteUsoCli(c.cod_cli))
        txtbxRestante.Text = FormatVal(c.limite_cli - CCur(txtbxEmUso.Text))
        If optnCH.Value Then
            If CCur(txtbxEmUso.Text) + CCur(lblValor.caption) > c.limite_cli Then
                MsgBox "O limite do Cliente não é suficiente para a compra!", vbExclamation
                
            End If
        End If
        
    End If
End Sub

Private Sub txtbxNome_Cli_LostFocus()
    On Error Resume Next
    AjustaTexto txtbxNome_Cli
End Sub
