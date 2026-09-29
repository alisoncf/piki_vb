VERSION 5.00
Begin VB.Form frmTipoVenda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecione o Tipo de Pagamento"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmndbtnCancelar 
      Cancel          =   -1  'True
      Caption         =   "F12 - C&ancelar"
      Height          =   735
      Left            =   5880
      Picture         =   "frmTipoVenda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CheckBox chkFinanceira 
      Caption         =   "F3 - Financeira"
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
      Picture         =   "frmTipoVenda.frx":054E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.PictureBox pcFinan 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1095
      ScaleWidth      =   5655
      TabIndex        =   21
      Top             =   3600
      Width           =   5655
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
      Begin VB.ComboBox cmbxFinan 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmTipoVenda.frx":2256
         Left            =   1320
         List            =   "frmTipoVenda.frx":2258
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtbxComprovante 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   27
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblTef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   62
         Top             =   600
         Width           =   225
      End
      Begin VB.Label lblTipoFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   195
         Left            =   360
         TabIndex        =   61
         Top             =   240
         Width           =   90
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
      Begin VB.Label lblCapFin 
         BackStyle       =   0  'Transparent
         Caption         =   "&Financeira:"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblComprov 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº do Comprovante:"
         Height          =   195
         Left            =   1560
         TabIndex        =   26
         Top             =   645
         Width           =   1440
      End
   End
   Begin VB.PictureBox pcAp 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2040
      ScaleHeight     =   1335
      ScaleWidth      =   5655
      TabIndex        =   29
      Top             =   4800
      Width           =   5655
      Begin VB.TextBox txtbxValorAP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Text            =   "0,00"
         Top             =   495
         Width           =   975
      End
      Begin VB.ComboBox cmbxPlanoPag 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   480
         Width           =   3735
      End
      Begin VB.OptionButton optnDup 
         BackColor       =   &H00D3E7EB&
         Caption         =   "F5 - &Duplicata"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Tag             =   "!"
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton optnCheque 
         BackColor       =   &H00D3E7EB&
         Caption         =   "F6 - C&heque:"
         Height          =   195
         Left            =   1680
         TabIndex        =   31
         Tag             =   "!"
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Final:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3120
         TabIndex        =   64
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acrésc. P/ Parc:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   285
         TabIndex        =   63
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblPlaCap 
         BackStyle       =   0  'Transparent
         Caption         =   "&Plano Pgto"
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblAcresc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1590
         TabIndex        =   56
         Top             =   885
         Width           =   570
      End
      Begin VB.Label lblFinAcresc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   330
         Left            =   3960
         TabIndex        =   55
         Top             =   870
         Width           =   570
      End
   End
   Begin VB.PictureBox pcCHV 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1095
      ScaleWidth      =   5655
      TabIndex        =   6
      Top             =   2400
      Width           =   5655
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
      BackColor       =   &H00D3E7EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1095
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   1200
      Width           =   5655
      Begin VB.Frame frNaoFiscal 
         Caption         =   "&Recebimento Não Fiscal:"
         Height          =   975
         Left            =   1800
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
         Begin VB.TextBox txtbxNaoFiscal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            TabIndex        =   59
            Text            =   "0,00"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblValor 
            AutoSize        =   -1  'True
            Caption         =   "&Valor do Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   480
            Width           =   1500
         End
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
      Caption         =   "F1 - &Dinheiro"
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
      Picture         =   "frmTipoVenda.frx":225A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox chkCheque 
      Caption         =   "F2 - Cheque À Vista"
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
      Picture         =   "frmTipoVenda.frx":37F7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CheckBox chkAprazo 
      Caption         =   "F4 - A Prazo"
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
      Picture         =   "frmTipoVenda.frx":537B
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H007DBFDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   8850
      TabIndex        =   38
      Top             =   0
      Width           =   8850
      Begin VB.CommandButton cmndbtnCliente 
         Caption         =   "Mudar Cliente"
         Height          =   255
         Left            =   5205
         TabIndex        =   57
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
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
         TabIndex        =   44
         Tag             =   "!"
         Text            =   "0,00"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restante:"
         Height          =   195
         Left            =   7320
         TabIndex        =   70
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em Uso:"
         Height          =   195
         Left            =   7320
         TabIndex        =   69
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite:"
         Height          =   195
         Left            =   7320
         TabIndex        =   68
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lblLimite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8040
         TabIndex        =   67
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblLimiteEmUso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8040
         TabIndex        =   66
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblLimiteRest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8040
         TabIndex        =   65
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF/CNPJ:"
         Height          =   195
         Left            =   2280
         TabIndex        =   53
         Top             =   0
         Width           =   825
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000.000.000-00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   54
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   4440
         TabIndex        =   52
         Tag             =   "!"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblRestante 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3360
         TabIndex        =   51
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
         TabIndex        =   50
         Tag             =   "!"
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblValor_Ven 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   49
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
         TabIndex        =   48
         Tag             =   "!"
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Desconto(%):"
         Height          =   195
         Left            =   1200
         TabIndex        =   47
         Tag             =   "!"
         Top             =   480
         Width           =   945
      End
      Begin VB.Label lblValorFinal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2280
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   43
         Top             =   0
         Width           =   510
      End
      Begin VB.Label lblNum_Ven 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   42
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblNome_cli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label lblVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4440
         TabIndex        =   39
         Top             =   720
         Width           =   2460
      End
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Caption         =   "F7 - &Confirmar"
      Height          =   735
      Left            =   3960
      Picture         =   "frmTipoVenda.frx":6415
      Style           =   1  'Graphical
      TabIndex        =   36
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
Private Sub chkAprazo_Click()
    If chkAprazo.Value = 1 Then
    If F_BuscaUmCampo("status_cli", "tab_clientes", "cod_cli='" & _
        frmTroco.txtbxNome_Cli.tag & "'") <> "1" Then
        MsgBox "O Status do cliente não permite esta operação!", vbExclamation
        Exit Sub
        
    End If
    End If
    If chkAprazo.Value = 0 Then
        pcAp.Enabled = False
        txtbxValorAP.Enabled = False
        cmbxPlanoPag.Enabled = False
        txtbxValorAP.Text = "0,00"
        
    Else
        pcAp.Enabled = True
        txtbxValorAP.Text = FormatVal(F_Restante)
        txtbxValorAP.Enabled = True
        cmbxPlanoPag.Enabled = True
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
            frmTroco.txtbxNome_Cli.tag & "'") <> "1" Then
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
        txtbxTitular.Text = lblNome_cli.caption
        txtbxCPF.Text = lblCPF.caption
        
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
        cmbxFinan.Enabled = False
        txtbxComprovante.Enabled = False
        txtbxValorFin.Text = "0,00"
    Else
        txtbxValorFin.Text = FormatVal(F_Restante)
        txtbxValorFin.Enabled = True
        cmbxFinan.Enabled = True
        txtbxComprovante.Enabled = True
        txtbxValorFin.SetFocus
    End If
    
    
End Sub

Private Sub chkFinanceira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        chkAprazo.SetFocus
    End If
End Sub

Private Sub cmbxFinan_Click()
    cmbxFinan.tag = F_BuscaUmCampo("cod_fin", "tab_financeiras", "nome_fin='" & cmbxFinan.Text & "'")
       lblTipoFin.caption = F_BuscaUmCampo("Tipo_fin", "tab_financeiras", "cod_fin='" & cmbxFinan.tag & "'")
    If lblTipoFin.caption = "3" Then
        lblTef.caption = "TEF"
        pcFinan.BackColor = vbGreen
    Else
        lblTef.caption = "POS"
        pcFinan.BackColor = vbYellow
    End If
    
End Sub

Private Sub cmbxFinan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxPlanoPag_Click()
    Dim strSQL As String
    
    strSQL = F_BuscaUmCampo("juros_tip", "tab_tipoven", "cod_tip='" & _
        Left(cmbxPlanoPag.Text, 2) & "'")
    If strSQL = "" Then strSQL = "0"
    lblAcresc.caption = "" & _
        FormatVal(CCur(strSQL) * CCur(txtbxValorAP.Text) / 100)
        
    lblFinAcresc.caption = FormatVal(CCur(lblValor_Ven.caption) + CCur(lblAcresc.caption))
End Sub

Private Sub cmbxPlanoPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
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
    
'    If bTefDedicado Then
'        If Not F_AbreBancoODBC(bDedicado, "dedicado") Then
'        End If
'    End If
    ok = True
    If CCur(txtbxValorAP.Text) > 0 Then
    
        frmFecharVenda.operacao = operacao
        frmFecharVenda.lblvlrVenda.caption = FormatVal(lblValorFinal.caption)
        frmFecharVenda.lblDesconto.caption = FormatVal(txtbxDesconto.Text)
        
            frmFecharVenda.frAp.Enabled = True
            frmFecharVenda.frForm.Enabled = True
            frmFecharVenda.frVal.Enabled = True
            If cmbxPlanoPag.Text <> "" Then
                frmFecharVenda.cmbxFormas_Pag.tag = Left(cmbxPlanoPag.Text, 2)
            End If
        
        frmFecharVenda.lblAprazo.caption = txtbxValorAP.Text
        frmFecharVenda.lblAprazo.tag = txtbxValorAP.Text
        frmFecharVenda.lblAvista.caption = txtbxValorAV.Text
        frmFecharVenda.lblFinan.caption = txtbxValorFin.Text
        
        frmFecharVenda.lblvlrVenda.caption = lblValor_Ven.caption
        frmFecharVenda.lblTotal.caption = lblValorFinal.caption
        frmFecharVenda.lblNome_cli.caption = lblNome_cli.caption
        frmFecharVenda.lblNome_cli.tag = lblNome_cli.tag
            
        If optnCheque.Value Then
            frmFecharVenda.optnCheque.Value = True
        Else
            frmFecharVenda.otpnCart.Value = True
            
        End If
        frmFecharVenda.optnCheque.Enabled = False
        frmFecharVenda.otpnCart.Enabled = False
        
        frmFecharVenda.tag = lblnum_ven.caption
        frmFecharVenda.caption = "Fechar " & operacao & " " & Me.tag & " de " & lblNome_cli.caption
        frmFecharVenda.Show 1
    End If
    If CCur(txtbxValorFin.Text) > 0 Then
        
        If lblTipoFin.caption = "3" Then
            bUsaTEF = True
        Else
            bUsaTEF = False
        End If
        If bUsaTEF Then
        
            If bTefDedicado = True Then
                If efetua_tef_dedicado = False Then
                    ok = False
                    Exit Sub
                Else
                    ok = True
                End If
            Else
                strNum_Cpm = "      "
                i = ecf_NumeroCupom(strNum_Cpm)
                frmVenda.lblMensagem.caption = "Iniciando transação TEF"
                P_Msg "Iniciando transação TEF", "Transação TEF"
                
                'P_Status "Processo TEF"
                frmTroco.bTEF = True 'isso é um gato pra poder não imprimir denovo o cupom
                
                If Not F_VendaCupomTEF(operacao, Me.tag, lblNome_cli.caption, lblCPF.caption, CCur(txtbxValorFin.Text), _
                    True, CCur(txtbxValorAV.Text), CCur(txtbxValorAP.Text)) Then
                    killIntPosResp
                    'P_BaixaSaida Me.Tag
                    'P_GravaSaida Me.Tag, sNum_Cpm
                    ok = False
                    MsgBox "Não foi possível efetuar a transação TEF!" & Chr(13) & "Tente novamente ou escolha outra forma de pagamento!", vbCritical
                    killIntPosResp
                    Exit Sub
                    P_Status
                Else
                    ok = True
                    P_Status
                    'Exit Sub
                End If
                P_Msg ""
            End If
        End If
        'me.tag está passando em branco, por isso dá erro de run-time e o formulário de fecharvenda não fecha.
        If Not F_GravaVenFin(Me.tag, cmbxFinan.tag, CCur(txtbxValorFin.Text), _
            CStr(Date), strCod_Usu, lblNome_cli.tag, txtbxComprovante.Text) Then
            MsgBox "Erro!", vbExclamation
            ok = True
        End If
        strSQL = F_GravaCai("C", Date, txtbxValorFin.Text, "Referente a  " & operacao & _
            " " & Me.tag & " FINAN: " & cmbxFinan.Text, Me.tag, "FIN", "01", 0, 0, strCod_Cai, strCod_Tur)
        If strSQL = "" Then
             MsgBox "Erro!", vbExclamation
            
        End If
    End If
    
    If CCur(txtbxValorAV.Text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtbxValorAV.Text, "Referente a  " & operacao & _
            " " & Me.tag, Me.tag, "DIN", "01", 0, 0, strCod_Cai, strCod_Tur)
        If strSQL = "" Then
            MsgBox "Erro!", vbExclamation
            
        End If
    End If
    
    If CCur(txtbxValorChe.Text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtbxValorChe.Text, "Referente a  " & operacao & _
        " " & Me.tag & " CHV: " & txtbxNumChe.Text, Me.tag, "CHV", "01", 0, 0, strCod_Cai, strCod_Tur)
        If strSQL = "" Then
            MsgBox "Erro!", vbExclamation
            Exit Sub
        End If
        If Not F_GravaCHV(strSQL, txtbxNumChe.Text, txtbxBanco.Text, _
            txtbxAgeChe.Text, txtbxCco.Text, txtbxCPF.Text, txtbxTitular.Text, txtbxValorChe.Text) Then
            MostraErro "Erro na gravação do cheque A Vista"
        End If
    End If
        
    If CCur(txtbxValorAP.Text) > 0 Then 'se houver a prazo baixa no frmfecharVenda
        If frmFecharVenda.tag = "" Then
            sSitucao = "2" ' e vai como venda fechada caso não tenha finalizado
        Else
            sSitucao = "4"
            Unload frmFecharVenda
        End If
    Else
        sSitucao = "4"  'se não houver a prazo, baixa tudo aqqui
        ok = True
        P_BaixaVenda Me.tag
    End If
    
        strSQL = "update " & TAB_VENDAS & " set situacao_ven=" & sSitucao & ", "
        strSQL = strSQL & "dataefet_ven=" & F_GravaData(Date) & ","
        strSQL = strSQL & "desconto_ven=" & F_GravaDec(txtbxDesconto.Text) & ","
        strSQL = strSQL & "cod_tip='" & Left(cmbxPlanoPag.Text, 2) & "' "
        'strSQL = strSQL & "valor_ven=" & F_GravaDec(lblTotal.caption)
        strSQL = strSQL & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        '  MsgBox "Venda Concluída", vbInformation
End Sub

Private Sub cmndbtnCliente_Click()
    Dim strSQL As String
    Dim c As tCliente
    c = F_BuscaCliente("C", "")
    If c.cod_cli <> "" Then
        strSQL = "Confirma mudança do cliente: " & lblNome_cli.caption
        strSQL = strSQL & Chr(13) & "Para o cliente: " & c.nome_Cli & "?"
        If MsgBox(strSQL, vbYesNo + vbQuestion) = vbYes Then
            lblNome_cli.caption = c.nome_Cli
            lblNome_cli.tag = c.cod_cli
            lblCPF.caption = c.cnpj_Cli
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
    
        grava
        
        If LCase(operacao) = "venda" Then
                frmFecharVenda.Command4.Enabled = True
                
                informa
                
                P_Status
                If Not ok And bUsaTEF = True Then
                    P_Msg "", ""
                    Me.tag = ""
                    Me.Hide
                    Exit Sub
                End If
                If ok = True Then
                    Me.Hide
                End If
        ElseIf LCase(operacao) = "troca" Then
            frmFecharVenda.Command4.Enabled = False
            informa
            Me.tag = "X"
        End If
        Unload frmFecharVenda
        
        Me.Hide
 End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "delete from  " & TAB_VENFEC & "  where num_ven='" & lblnum_ven.caption & "' "
    conn.Execute strSQL
    strSQL = "insert into  " & TAB_VENFEC & " (num_ven,avista_ven,aprazo_ven,"
    strSQL = strSQL & "finan_ven,desconto_ven,valorfinal_ven,cod_fin,cod_tip,cod_usu,tipo_ven,cod_Cai,cod_Tur)values( "
    strSQL = strSQL & "'" & lblnum_ven.caption & "',"
    strSQL = strSQL & F_GravaDec(txtbxValorAV.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxValorAP.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxValorFin.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxDesconto.Text) & ","
    strSQL = strSQL & F_GravaDec(lblValorFinal.caption) & ","
    strSQL = strSQL & "'" & F_BuscaUmCampo("cod_fin", "tab_financeiras", "nome_fin='" & (cmbxFinan.Text) & "'") & "', "
    strSQL = strSQL & "'" & Left(cmbxPlanoPag.Text, 2) & "', "
    strSQL = strSQL & "'" & strCod_Usu & "', "
    If CCur(txtbxValorAP.Text) > 0 Then
        strSQL = strSQL & IIf(optnCheque.Value = True, "'C',", "'D',")
    Else
        strSQL = strSQL & "'1', "
    End If
        
        
        
        strSQL = strSQL & "'" & strCod_Cai & "',"
        strSQL = strSQL & "'" & strCod_Tur & "') "
    
    conn.Execute strSQL
    
    If LCase(operacao) = "venda" Then
        strSQL = "update " & TAB_VENDAS & " set situacao_ven='2', "
        If CCur(lblFinAcresc.caption) > 0 Then
            strSQL = strSQL & "valor_ven=" & F_GravaDec(lblFinAcresc.caption) & ","
        End If
        strSQL = strSQL & "nomecli_ven='" & lblNome_cli.caption & "', "
        strSQL = strSQL & " cod_cli='" & lblNome_cli.tag & "'  where num_ven='" & lblnum_ven.caption & "' "
        
        conn.Execute strSQL
    End If
End Sub
Function soma() As Currency
    soma = CCur(txtbxValorAP.Text) + CCur(txtbxValorAV.Text) + CCur(txtbxValorFin.Text) + CCur(txtbxValorChe.Text)
    
End Function
Function F_Restante() As Currency
    F_Restante = CCur(lblValorFinal.caption) - (CCur(txtbxValorAP.Text) + CCur(txtbxValorAV.Text) + CCur(txtbxValorFin.Text) + CCur(txtbxValorChe.Text))
    
End Function

Private Sub Form_Activate()
       
    P_Soma
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        If chkAv.Value = 1 Then
            chkAv.Value = 0
        Else
            chkAv.Value = 1
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If chkCheque.Enabled = True Then
            If chkCheque.Value = 1 Then
                chkCheque.Value = 0
            Else
                chkCheque.Value = 1
            End If
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If chkFinanceira.Enabled = True Then
            If chkFinanceira.Value = 1 Then
                chkFinanceira.Value = 0
            Else
                chkFinanceira.Value = 1
            End If
        End If

    ElseIf KeyCode = vbKeyF4 Then
        If chkAprazo.Enabled = True Then
            If chkAprazo.Value = 1 Then
                chkAprazo.Value = 0
            Else
                chkAprazo.Value = 1
            End If
        Else
            
            
        End If
    ElseIf KeyCode = vbKeyF5 Then
        optnDup.Value = True
    ElseIf KeyCode = vbKeyF6 Then
        optnCheque.Value = True
    ElseIf KeyCode = vbKeyF7 Then
        cmndbtnConfirmar_Click
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnCancelar_Click
    End If
End Sub

Private Sub Form_Load()
    P_PreencheFinanceiras cmbxFinan
    P_PreencheFormasPag cmbxPlanoPag, IIf(optnCheque.Value, "C", "D")
    lblValorFinal.caption = FormatVal(CCur(lblValor_Ven.caption) - CCur(txtbxDesconto.Text))
    
    If F_BuscaUmCampo("status_cli", "tab_clientes", "cod_cli='" & _
        frmTroco.txtbxNome_Cli.tag & "'") <> "1" Then
        pcAp.Enabled = False
        optnCheque.Enabled = False
        optnDup.Enabled = False
'        MsgBox "A venda só poderá ser fechada a vista ou a cartão!", vbInformation
    End If

End Sub

Private Sub optnCheque_Click()
    P_PreencheFormasPag cmbxPlanoPag, "C"
End Sub

Private Sub optnDup_Click()
    P_PreencheFormasPag cmbxPlanoPag, "D"
End Sub

Private Sub txtbxAgeChe_GotFocus()
    ganhaFoco txtbxAgeChe
End Sub

Private Sub txtbxAgeChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtbxAgeChe_LostFocus()
    AjustaTexto txtbxAgeChe
End Sub

Private Sub txtbxAge_Che_Change()

End Sub

Private Sub txtbxBanco_GotFocus()
    ganhaFoco txtbxBanco
End Sub

Private Sub txtbxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaBCO
        TrocaTabKey KeyCode
    End If
End Sub
Sub buscaBCO()
    Dim b As TBanco
    b = F_BuscaBanco(txtbxBanco.Text)
    If b.cod_bco <> "" Then
        txtbxBanco.Text = b.cod_bco
    End If
End Sub
Private Sub txtbxBanco_LostFocus()
    AjustaTexto txtbxBanco
End Sub

Private Sub txtbxConta_Change()

End Sub

Private Sub txtbxccoche_Change()

End Sub

Private Sub txtbxCco_GotFocus()
    ganhaFoco txtbxCco
End Sub

Private Sub txtbxCco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCco_LostFocus()
    AjustaTexto txtbxCco
End Sub

Private Sub txtbxComprovante_GotFocus()
    ganhaFoco txtbxComprovante
End Sub

Private Sub txtbxComprovante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCpf_GotFocus()
    ganhaFoco txtbxCPF
End Sub

Private Sub txtbxCpf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
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
    ganhaFoco txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    PerdeFoco txtbxDesconto
    If CCur(txtbxDesconto.Text) > 50 Then
        MsgBox "Desconto muito alto!", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    End If
    lblValorFinal.caption = FormatVal(CCur(lblValor_Ven.caption) - (CCur(lblValor_Ven.caption) * CCur(txtbxDesconto.Text) / 100))
    
End Sub

Private Sub txtbxNumChe_GotFocus()
    ganhaFoco txtbxNumChe
End Sub

Private Sub txtbxNumChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtbxNumChe_LostFocus()
    AjustaTexto txtbxNumChe
End Sub

Private Sub txtbxTitular_GotFocus()
    ganhaFoco txtbxTitular
End Sub

Private Sub txtbxTitular_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxTitular_LostFocus()
    AjustaTexto txtbxTitular
End Sub

Private Sub txtbxValorAV_GotFocus()
    ganhaFoco txtbxValorAV
End Sub

Private Sub txtbxValorAV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValorAV_LostFocus()
    Dim cSoma As Currency
    On Error Resume Next
    PerdeFoco txtbxValorAV
    cSoma = soma - CCur(txtbxValorAV)
    If CCur(txtbxValorAV.Text) > CCur(lblValorFinal.caption) - cSoma Then
        MsgBox "Este valor excede o restante!", vbExclamation
        txtbxValorAV.SetFocus
        Exit Sub
    End If
    P_Soma
End Sub
Private Sub txtbxValorAganhafoco()
    ganhaFoco txtbxValorAP
End Sub

Private Sub txtbxValorAP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValorAP_LostFocus()
    PerdeFoco txtbxValorAP
    If F_Restante < 0 Then
        MsgBox "Este valor excede o restante!", vbExclamation
        txtbxValorAP.SetFocus
        Exit Sub
    End If
    P_Soma
    soma
End Sub

Private Sub txtbxValorChe_GotFocus()
    ganhaFoco txtbxValorChe
End Sub

Private Sub txtbxValorChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValorChe_LostFocus()
    Dim cSoma As Currency
    On Error Resume Next
    PerdeFoco txtbxValorChe
    cSoma = soma - CCur(txtbxValorChe.Text)
    If CCur(txtbxValorChe.Text) > CCur(lblValorFinal.caption) - cSoma Then
        MsgBox "Este valor excede o restante!", vbExclamation
        txtbxValorChe.SetFocus
        Exit Sub
    End If
    P_Soma

End Sub

Private Sub txtbxValorFin_GotFocus()
    ganhaFoco txtbxValorFin
End Sub

Private Sub txtbxValorFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxValorFin_LostFocus()
    Dim cSoma As Currency
    On Error Resume Next
    PerdeFoco txtbxValorFin
    cSoma = soma - CCur(txtbxValorFin.Text)
    If F_Restante < 0 Then
        MsgBox "Este valor excede o restante!", vbExclamation
        txtbxValorFin.SetFocus
        Exit Sub
    End If
    P_Soma
End Sub
Function efetua_tef_dedicado() As Boolean
    
    Dim i As Integer
    Dim strSQL As String
    Dim s As String
    
    
    
    efetua_tef_dedicado = Dedicado.efetua_venda(txtbxValorFin.Text, txtbxValorAV.Text, txtbxValorAP.Text, , lblNome_cli.caption, lblCPF.caption, "")
    
    
'Exit Function
'
'
'    i = ecf_NumeroCupom(strNum_Cpm)
'
'    If Not F_GravaTransacao(Me.Tag, 1, strNum_Cpm) Then
'        MsgBox "Erro na gravação da transação!", vbExclamation
'        efetua_tef_dedicado = False
'        Exit Function
'    End If
'    If Not gravaVendaTefDedicado Then
'        efetua_tef_dedicado = False
'        MsgBox "erro"
'        Exit Function
'    End If
'
'    'Shell strCaminhoPkXp, vbNormalFocus
'    Sleep 10000
'    s = "1"
'
'    Do While (s = "1" Or s = "2")
'        Sleep 1000
'        s = F_BuscaUmCampoOutroBanco(bDedicado, "status_tra", "tab_transacoes", "num_ven='" & Me.Tag & "'")
'
'
'    Loop
'
'    If s = "3" Then
'        'se o status é 2 significa que deu tudo certo
'        efetua_tef_dedicado = True
'    ElseIf s = "5" Then
'        'não deu certo, escolha outra forma de pagamento
'        efetua_tef_dedicado = False
'
'    ElseIf s = "51" Then
'
'    ElseIf s = "2" Then
'        efetua_tef_dedicado = True
'
'    End If
'
End Function
'Function gravaVendaTefDedicado() As Boolean
'    Dim strSQL As String
'    Dim rPesq as new adodb.recordset
'    strSQL = "select tv.*,ti.* from " & TAB_VENDAS & " as tv, " & TAB_itVENDAS & " as ti"
'    strSQL = strSQL & " where tv.num_ven=ti.num_ven"
'    strSQL = strSQL & " and tv.num_ven='" & Me.Tag & "'"
'
'    rpesq.open  strsql, conn
'    With rPesq
'        If Not F_GravaVenda(Me.Tag, !cod_cli, !nomecli_ven, !valor_ven, txtbxValorAV.Text, _
'            txtbxValorAP.Text, txtbxValorFin.Text, !desconto_ven, lblCPF.caption, "", _
'            strNome_Usu) Then
'            gravaVendaTefDedicado = False
'        End If
'        Do While Not .EOF
'            If Not F_GravaItvendas(Me.Tag, !cod_pro, !descricao_ite, !qtd_ite, !unit_ite, 0, "I", "$", F_IcmsProd(!cod_pro, 3), _
'                "2") Then
'                gravaVendaTefDedicado = False
'            End If
'            .MoveNext
'        Loop
'        .Close
'    End With
'
'    gravaVendaTefDedicado = True
'End Function
