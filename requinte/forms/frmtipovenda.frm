VERSION 5.00
Begin VB.Form frmTipoVenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione o Tipo de Pagamento"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   9360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFinanceira 
      Caption         =   "Cartão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Picture         =   "frmtipovenda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox pcFinan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1095
      ScaleWidth      =   7095
      TabIndex        =   21
      Top             =   2400
      Width           =   7095
      Begin VB.TextBox txtbxValorFin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.PictureBox pcAp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1095
      ScaleWidth      =   7095
      TabIndex        =   25
      Top             =   3600
      Width           =   7095
      Begin VB.TextBox txtDataBase 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   53
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtbxFormaPag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   50
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtbxValorAP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Text            =   "0,00"
         Top             =   495
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   285
         Left            =   3480
         Picture         =   "frmtipovenda.frx":0DAD
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base:"
         Height          =   195
         Left            =   3840
         TabIndex        =   52
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblParcelas 
         AutoSize        =   -1  'True
         Caption         =   "Parcelas:"
         Height          =   195
         Left            =   6120
         TabIndex        =   54
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblPlaCap 
         BackStyle       =   0  'Transparent
         Caption         =   "&Plano Pgto"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblAcresc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acréscimo por Parcela"
         ForeColor       =   &H8000000C&
         Height          =   195
         Left            =   5280
         TabIndex        =   49
         Top             =   0
         Width           =   1590
      End
      Begin VB.Label lblFinAcresc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Final:"
         ForeColor       =   &H8000000C&
         Height          =   195
         Left            =   5970
         TabIndex        =   48
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.PictureBox pcCHV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1095
      ScaleWidth      =   7095
      TabIndex        =   6
      Top             =   4920
      Width           =   7095
      Begin VB.TextBox txtbxCPF 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Top             =   740
         Width           =   1935
      End
      Begin VB.TextBox txtbxTitular 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   740
         Width           =   2775
      End
      Begin VB.TextBox txtbxCco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxAgeChe 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtbxBanco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtbxValorChe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxNumChe 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCPFCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CPF/CNPJ:"
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         Top             =   520
         Width           =   825
      End
      Begin VB.Label lblTit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Titular:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   520
         Width           =   480
      End
      Begin VB.Label lblCco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Conta:"
         Height          =   195
         Left            =   3840
         TabIndex        =   15
         Top             =   0
         Width           =   465
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Agência:"
         Height          =   195
         Left            =   3120
         TabIndex        =   13
         Top             =   0
         Width           =   630
      End
      Begin VB.Label lblBan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Banco:"
         Height          =   195
         Left            =   2400
         TabIndex        =   11
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblNumChe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cheque:"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox pcDin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1095
      ScaleWidth      =   7095
      TabIndex        =   1
      Top             =   1200
      Width           =   7095
      Begin VB.TextBox txtPIX 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   58
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox chkPIX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PIX"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   57
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtbxValorAV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbPix 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta/Chave:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2400
         TabIndex        =   59
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.CheckBox chkAv 
      Caption         =   "&Dinheiro/PIX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Picture         =   "frmtipovenda.frx":0EA6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox chkCheque 
      Caption         =   "Cheque"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Picture         =   "frmtipovenda.frx":13A6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CheckBox chkAprazo 
      Caption         =   "A Prazo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Picture         =   "frmtipovenda.frx":153C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3600
      Width           =   1815
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   9360
      TabIndex        =   31
      Top             =   0
      Width           =   9360
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Tag             =   "!"
         Text            =   "0,00"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblLimite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5640
         TabIndex        =   56
         Top             =   0
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite:"
         Height          =   195
         Left            =   5040
         TabIndex        =   55
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF/CNPJ:"
         Height          =   195
         Left            =   2280
         TabIndex        =   46
         Top             =   0
         Width           =   825
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000.000.000-00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3360
         TabIndex        =   47
         Top             =   0
         Width           =   1140
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   4440
         TabIndex        =   45
         Tag             =   "!"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblRestante 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   44
         Tag             =   "!"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restante:"
         Height          =   195
         Left            =   3360
         TabIndex        =   43
         Tag             =   "!"
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblValor_Ven 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Tag             =   "!"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venda:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Tag             =   "!"
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Desconto($):"
         Height          =   195
         Left            =   1200
         TabIndex        =   40
         Tag             =   "!"
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lblValorFinal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   39
         Tag             =   "!"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Final:"
         Height          =   195
         Left            =   2280
         TabIndex        =   38
         Tag             =   "!"
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   0
         Width           =   510
      End
      Begin VB.Label lblNum_Ven 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   35
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblNome_cli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   33
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblVendedor 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   32
         Top             =   720
         Width           =   2460
      End
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Caption         =   "F7 - &Confirmar"
      Height          =   735
      Left            =   5400
      Picture         =   "frmtipovenda.frx":17EE
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnCancelar 
      Cancel          =   -1  'True
      Caption         =   "F12 - C&ancelar"
      Height          =   735
      Left            =   7320
      Picture         =   "frmtipovenda.frx":1938
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6240
      Width           =   1815
   End
End
Attribute VB_Name = "frmTipoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public operacao As String
Public ok As Boolean
Dim bActive As Boolean
Dim cli As cliente
Private Sub chkAprazo_Click()
    If chkAprazo.Value = 1 Then
        If F_BuscaUmCampo("status_cli", TAB_CLIENTES, "cod_cli='" & _
            lblNome_Cli.tag & "'") <> "1" Then
            If MsgBox("O Status do cliente não permite esta operação! Deseja Continuar?", vbYesNo + vbExclamation) = vbNo Then
                Exit Sub
            End If
        End If
        
    
        
        
    End If
    If chkAprazo.Value = 0 Then
        pcAp.Enabled = False
        txtbxValorAP.Enabled = False
        'cmbxPlanoPag.Enabled = False
        txtbxValorAP.Text = "0,00"
        
    Else
        pcAp.Enabled = True
        txtbxValorAP.Text = FormatVal(F_Restante)
        txtbxValorAP.Enabled = True
        'cmbxPlanoPag.Enabled = True
        txtbxValorAP.SetFocus
        
        
    End If
    
End Sub
Sub P_Soma()
    lblRestante.caption = FormatVal(F_Restante)
    
End Sub
Private Sub chkAv_Click()
    If chkAv.Value = 1 Then
        txtbxValorAV.Enabled = True
        txtbxValorAV.SetFocus
        txtbxValorAV.Text = FormatVal(F_Restante)
    Else
        txtbxValorAV.Text = "0,00"
        txtbxValorAV.Enabled = False
        
    End If
End Sub

Private Sub chkAv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        chkCheque.SetFocus
    End If
End Sub

Private Sub chkCheque_Click()
    
    If chkCheque.Value = 1 Then
        If F_BuscaUmCampo("status_cli", "tab_clientes", "cod_cli='" & _
            frmTroco.txtbxNomeCli.tag & "'") <> "1" Then
            If MsgBox("O Status do cliente não permite esta operação! Deseja Continuar?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
         
        
        End If
    End If
    If chkCheque.Value = 1 Then
        txtbxValorChe.Enabled = True
        txtbxValorChe.SetFocus
        txtbxValorChe.Text = FormatVal(F_Restante)
        habChe True
        txtbxTitular.Text = lblNome_Cli.caption
        txtbxCPF.Text = lblcpf.caption
        
    Else
        txtbxValorChe.Text = "0,00"
        txtbxValorChe.Enabled = False
        habChe False
    End If
End Sub
Sub habChe(opcao As Boolean)
    txtbxNumChe.Enabled = opcao
    txtbxBanco.Enabled = opcao
    txtbxAgeChe.Enabled = opcao
    txtbxTitular.Enabled = opcao
    txtbxCPF.Enabled = opcao
    lblCPFCap.Enabled = opcao
    lblNumChe.Enabled = opcao
    lblBan.Enabled = opcao
    lblAge.Enabled = opcao
    lblTit.Enabled = opcao
    lblCco.Enabled = opcao
    txtbxCco.Enabled = opcao
End Sub

Private Sub chkCheque_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        chkFinanceira.SetFocus
    End If
End Sub

Private Sub chkFinanceira_Click()
    If chkFinanceira.Value = 0 Then
        txtbxValorFin.Enabled = False
        
        
        txtbxValorFin.Text = "0,00"
    Else
        txtbxValorFin.Text = FormatVal(F_Restante)
        txtbxValorFin.Enabled = True
                txtbxValorFin.SetFocus
    End If
    
    
End Sub

Private Sub chkFinanceira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        chkAprazo.SetFocus
    End If
End Sub


Private Sub cmbxFinan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxPlanoPag_Click()
    Dim strSQL As String
    lblAcresc.caption = "Acréscimo por Parcela: R$ "
    strSQL = F_BuscaUmCampo("juros_tip", "tab_tipoven", "cod_tip='" & _
        Left("", 2) & "'")
    If strSQL = "" Then strSQL = "0"
    lblAcresc.caption = lblAcresc.caption & _
        FormatVal(CCur(strSQL) * CCur(txtbxValorAP.Text) / 100)
End Sub

Private Sub cmbxPlanoPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub chkPIX_Click()
    If chkPIX.Value = 1 Then
        txtPIX.Enabled = True
        lbPix.Enabled = True
        txtPIX.SetFocus
    Else
        txtPIX.Enabled = False
        lbPix.Enabled = False
        
    End If
End Sub

Private Sub cmndbtnCancelar_Click()
    Me.tag = ""
    Me.Hide
End Sub
Sub informa()
    Dim strSQL As String
    Dim sSitucao As String
    Dim sNum_Cpm As String
    Dim i As Integer
    Dim sSeq As String
    Dim sSeqDoc As String
    Dim sSeqItd As String
    
    If CCur(txtbxValorAP.Text) > 0 Then
        If CCur(txtbxValorAP.Text) > CCur(lblLimite.caption) Then
            If MsgBox("Limite insuficiente para compra a prazo! Deseja continuar mesmo assim?", vbExclamation + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        
        If lblNome_Cli.tag = "" Or lblNome_Cli.tag = "0" Then
            MsgBox "Não é possível gerar parcelas para cliente não cadastrado.", vbExclamation
            Exit Sub
        End If
        
        sSeqDoc = doc.F_GravaDocCompleto("TROCA VENDA " & Me.tag & "EM " & Now, Me.tag, txtbxValorAP.Text, txtDataBase.Text, lblNome_Cli.tag, _
            "R", Me.tag, lblParcelas.caption, "V", "C", "000")
        If sSeqDoc = "" Then
            MsgBox "erro na gravação do documento!", vbExclamation
            Exit Sub
        End If
        
        
        APIInserirDoc sSeqDoc, True
        
        
        
        Me.tag = sSeqDoc
        
        
        
            
    End If
    If CCur(txtbxValorAV.Text) > 0 Then
        
        If chkPIX.Value = 0 Then
        
            strSQL = F_GravaCai("C", Date, txtbxValorAV.Text, "Referente a  " & operacao & _
            " " & Me.tag, Me.tag, "DIN", "01", 0, 0, strCod_Cai, strCod_Tur, "001", "S")
        Else
            strSQL = F_GravaCai("C", Date, txtbxValorAV.Text, "PIX " & txtPIX.Text & " Ref. " & operacao & _
            " " & Me.tag, Me.tag, "PIX", "01", txtbxValorAV.Text, txtbxValorAV.Text, strCod_Cai, strCod_Tur, "001", "S")
        End If
        If strSQL = "" Then
            MsgBox "Erro!", vbExclamation
            
        End If
    End If
     
    If CCur(txtbxValorFin.Text) > 0 Then
        
        
        sSeq = Cartoes.F_PagamentoCartaoAvulso("Troca " & Me.tag & "-" & lblNome_Cli.caption, txtbxValorFin.Text)
                'já efetua gravação no caixa
                If sSeq = "" Then
                    Me.tag = ""
                    MsgBox "pagamento não efetuado!", vbExclamation
                    Exit Sub
                End If

        Me.tag = sSeq
        Me.Hide
                
        
        
    End If
    If CCur(txtbxValorChe.Text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtbxValorChe.Text, "Referente a  " & operacao & _
        " " & Me.tag & " CHV: " & txtbxNumChe.Text, Me.tag, "CHV", "01", 0, 0, strCod_Cai, strCod_Tur, "001", "S")
        If strSQL = "" Then
            MsgBox "Erro!", vbExclamation
            Exit Sub
        End If
        If Not F_GravaCHV(strSQL, txtbxNumChe.Text, txtbxBanco.Text, _
            txtbxAgeChe.Text, txtbxCco.Text, txtbxCPF.Text, txtbxTitular.Text, txtbxValorChe.Text) Then
            P_MostraErro "Erro na gravação do cheque A Vista"
        End If
    End If
        
            
End Sub

Private Sub cmndbtnCliente_Click()
    Dim strSQL As String
    Dim c As cliente
    c = F_BuscaCliente("")
    If c.cod_cli <> "" Then
        strSQL = "Confirma mudança do cliente: " & lblNome_Cli.caption
        strSQL = strSQL & Chr(13) & "Para o cliente: " & c.nome_cli & "?"
        If MsgBox(strSQL, vbYesNo + vbQuestion) = vbYes Then
            lblNome_Cli.caption = c.nome_cli
            lblNome_Cli.tag = c.cod_cli
            lblcpf.caption = c.cnpj_Cli
        End If
    End If
End Sub

Private Sub cmndbtnConfirmar_Click()
    If CCur(lblRestante.caption) <> 0 Then
        MsgBox "A " & operacao & " não está completa!", vbExclamation
        Exit Sub
    End If
    Dim sNum_Cpm As String
    sNum_Cpm = Space(6)
    If MsgBox("Confirma Fechamento da " & operacao & "? ", vbYesNo + vbQuestion) = vbYes Then
      '  grava
        
        If LCase(operacao) = "venda" Then
            
            
                frmFecharVenda.Command4.Enabled = True
                
                informa
                P_Status
                
            Me.tag = "X"
            Me.Hide
        ElseIf LCase(operacao) = "troca" Then
            frmFecharVenda.Command4.Enabled = False
            informa
            
            If Me.tag <> "" Then
                Unload frmFecharVenda
                On Error Resume Next
                Me.Hide
            End If
        
        End If
        'Unload frmFecharVenda
        
    Else
        'Unload frmFecharVenda
        'Me.Tag = ""
        'Me.Hide
    End If
    
    
    
    
    
 End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "delete from tab_venfec where num_ven='" & lblNum_Ven.caption & "' "
    Conn.Execute strSQL
    
    
    
    strSQL = "insert into tab_venfec(num_ven,avista_ven,aprazo_ven,"
    strSQL = strSQL & "finan_ven,desconto_ven,valorfinal_ven,cod_fin,cod_tip,cod_usu,tipo_ven,cod_Cai,cod_Tur)values( "
    strSQL = strSQL & "'" & lblNum_Ven.caption & "',"
    strSQL = strSQL & F_GravaDec(txtbxValorAV.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxValorAP.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxValorFin.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxDesconto.Text) & ","
    strSQL = strSQL & F_GravaDec(lblValorFinal.caption) & ","
    strSQL = strSQL & "'" & "', "
    strSQL = strSQL & "'" & txtbxFormaPag.tag & "', "
    strSQL = strSQL & "'" & strCod_Usu & "', "
    If CCur(txtbxValorAP.Text) > 0 Then
        strSQL = strSQL & "'D'"
    Else
        strSQL = strSQL & "'1', "
    End If
        
        
        
        strSQL = strSQL & "'" & strCod_Cai & "',"
        strSQL = strSQL & "'" & strCod_Tur & "') "
    
    'Banco.Execute strSQL
    
    If LCase(operacao) = "venda" Then
        strSQL = "update tab_vendas set situacao_ven='2', "
        strSQL = strSQL & "nomecli_ven='" & lblNome_Cli.caption & "', "
        strSQL = strSQL & " cod_cli='" & lblNome_Cli.tag & "'  where num_ven='" & lblNum_Ven.caption & "' "
        
     '   Banco.Execute strSQL
    End If
End Sub
Function soma() As Currency
    soma = CCur(txtbxValorAP.Text) + CCur(txtbxValorAV.Text) + CCur(txtbxValorFin.Text) + CCur(txtbxValorChe.Text)
    
End Function
Function F_Restante() As Currency
    F_Restante = CCur(lblValorFinal.caption) - (CCur(txtbxValorAP.Text) + CCur(txtbxValorAV.Text) + CCur(txtbxValorFin.Text) + CCur(txtbxValorChe.Text))
    
End Function

Private Sub cmndbtnFormasPag_Click()
    buscaFormas
End Sub
Sub buscaFormas()
    Dim f As FormasPag.FormasPag
    
    f = FormasPag.F_BuscaFormasPag
    
    If f.cod_For <> "" Then
        txtbxFormaPag.Text = f.nome_for
        txtbxFormaPag.tag = f.cod_For
        lblParcelas.caption = f.parcelas_for
    End If
    
End Sub
Private Sub Form_Activate()
    On Error GoTo trata_Erro
    If bActive Then Exit Sub
    P_Soma
    bActive = True
    
    buscaCli
trata_Erro:
    
    
End Sub
Sub buscaCli()
    cli = F_BuscaClienteCodigo(lblNome_Cli.tag)
    lblLimite.caption = cli.limite_cli - F_BuscaLimiteEmUsoCli(cli.cod_cli)
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        If chkAv.Value = 1 Then
            chkAv.Value = 0
        Else
            chkAv.Value = 1
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If chkCheque.Value = 1 Then
            chkCheque.Value = 0
        Else
            chkCheque.Value = 1
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If chkFinanceira.Value = 1 Then
            chkFinanceira.Value = 0
        Else
            chkFinanceira.Value = 1
        End If

    ElseIf KeyCode = vbKeyF4 Then
        If chkAprazo.Value = 1 Then
            chkAprazo.Value = 0
        Else
            chkAprazo.Value = 1
        End If
    ElseIf KeyCode = vbKeyF7 Then
        cmndbtnConfirmar_Click
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnCancelar_Click
    End If
End Sub


Private Sub Form_Load()
    'P_PreencheFinanceiras cmbxFinan
    'P_PreencheFormasPag cmbxPlanoPag, IIf(optnCheque.Value, "C", "D")
    lblValorFinal.caption = FormatVal(CCur(lblValor_Ven.caption) - CCur(txtbxDesconto.Text))
    
    If F_BuscaUmCampo("status_cli", "tab_clientes", "cod_cli='" & _
        frmTroco.txtbxNomeCli.tag & "'") <> "1" Then
        pcAp.Enabled = False
        'optnCheque.Enabled = False
        'optnDup.Enabled = False
        MsgBox "A venda só poderá ser fechada a vista ou a cartão!", vbInformation
    End If
    txtDataBase.Text = Date
End Sub

Private Sub optnCheque_Click()
    'P_PreencheFormasPag cmbxPlanoPag, "C"
End Sub

Private Sub optnDup_Click()
    'P_PreencheFormasPag cmbxPlanoPag, "DUP"
End Sub

Private Sub txtbxAgeChe_GotFocus()
    P_SelecTexto txtbxAgeChe
End Sub

Private Sub txtbxAgeChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxAgeChe_LostFocus()
    P_AjustaTexto txtbxAgeChe
End Sub

Private Sub txtbxAge_Che_Change()

End Sub

Private Sub txtbxBanco_GotFocus()
    P_SelecTexto txtbxBanco
End Sub

Private Sub txtbxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxBanco_LostFocus()
    P_AjustaTexto txtbxBanco
End Sub

Private Sub txtbxConta_Change()

End Sub

Private Sub txtbxccoche_Change()

End Sub

Private Sub txtbxCco_GotFocus()
    P_SelecTexto txtbxCco
End Sub

Private Sub txtbxCco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCco_LostFocus()
    P_AjustaTexto txtbxCco
End Sub

Private Sub txtbxComprovante_GotFocus()
    'P_SelecTexto txtbxComprovante
End Sub

Private Sub txtbxComprovante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCpf_GotFocus()
    P_SelecTexto txtbxCPF
End Sub

Private Sub txtbxCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCPF_LostFocus()
    txtbxCPF.Text = F_SoNumero(txtbxCPF.Text)
    If Len(txtbxCPF.Text) = 0 Then
        
    ElseIf Len(txtbxCPF.Text) >= 14 Then
        lblCPFCap.caption = "&CNPJ:"
        P_FormataCGC txtbxCPF
        If Not F_CGC(txtbxCPF.Text) Then
            MsgBox "CNPJ Inválido!", vbExclamation
            txtbxCPF.SetFocus
            Exit Sub
        End If
    Else
        lblCPFCap.caption = "&CPF:"
        P_FormataCPF txtbxCPF
        If Not F_CPF(txtbxCPF.Text) Then
            MsgBox "CPF Inválido!", vbExclamation
            txtbxCPF.SetFocus
            Exit Sub
        End If
    End If
End Sub


Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    P_ValorLostFocus txtbxDesconto
    If CCur(txtbxDesconto.Text) > 50 Then
        MsgBox "Desconto muito alto!", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    End If
    lblValorFinal.caption = FormatVal(CCur(lblValor_Ven.caption) - (CCur(lblValor_Ven.caption) * CCur(txtbxDesconto.Text) / 100))
    
End Sub

Private Sub txtbxNumChe_GotFocus()
    P_SelecTexto txtbxNumChe
End Sub

Private Sub txtbxNumChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxNumChe_LostFocus()
    P_AjustaTexto txtbxNumChe
End Sub

Private Sub txtbxTitular_GotFocus()
    P_SelecTexto txtbxTitular
End Sub

Private Sub txtbxTitular_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxTitular_LostFocus()
    P_AjustaTexto txtbxTitular
End Sub

Private Sub txtbxValorAV_GotFocus()
    P_SelecTexto txtbxValorAV
End Sub

Private Sub txtbxValorAV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorAV_LostFocus()
    Dim cSoma As Currency
    On Error Resume Next
    P_ValorLostFocus txtbxValorAV
    cSoma = soma - CCur(txtbxValorAV)
    If CCur(txtbxValorAV.Text) > CCur(lblValorFinal.caption) - cSoma Then
        MsgBox "Este valor excede o restante!", vbExclamation
        txtbxValorAV.SetFocus
        Exit Sub
    End If
    P_Soma
End Sub
Private Sub txtbxValorAP_GotFocus()
    P_SelecTexto txtbxValorAP
End Sub

Private Sub txtbxValorAP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorAP_LostFocus()
    P_ValorLostFocus txtbxValorAP
    If F_Restante < 0 Then
        MsgBox "Este valor excede o restante!", vbExclamation
        txtbxValorAP.SetFocus
        Exit Sub
    End If
    P_Soma
    soma
End Sub

Private Sub txtbxValorChe_GotFocus()
    P_SelecTexto txtbxValorChe
End Sub

Private Sub txtbxValorChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorChe_LostFocus()
    Dim cSoma As Currency
    On Error Resume Next
    P_ValorLostFocus txtbxValorChe
    cSoma = soma - CCur(txtbxValorChe.Text)
    If CCur(txtbxValorChe.Text) > CCur(lblValorFinal.caption) - cSoma Then
        MsgBox "Este valor excede o restante!", vbExclamation
        txtbxValorChe.SetFocus
        Exit Sub
    End If
    P_Soma

End Sub

Private Sub txtbxValorFin_GotFocus()
    P_SelecTexto txtbxValorFin
End Sub

Private Sub txtbxValorFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorFin_LostFocus()
    Dim cSoma As Currency
    On Error Resume Next
    P_ValorLostFocus txtbxValorFin
    cSoma = soma - CCur(txtbxValorFin.Text)
    If F_Restante < 0 Then
        MsgBox "Este valor excede o restante!", vbExclamation
        txtbxValorFin.Text = "0"
        txtbxValorFin.SetFocus
        Exit Sub
    End If
    P_Soma
End Sub

Private Sub txtDataBase_GotFocus()
    P_SelecTexto txtDataBase
End Sub

Private Sub txtDataBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDataBase_LostFocus()
    P_AjustaData txtDataBase
End Sub
