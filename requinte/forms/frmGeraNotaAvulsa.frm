VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGeraNotaAvulsa 
   Caption         =   "Nota Fiscal de Saída Avulsa"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   Icon            =   "frmGeraNotaAvulsa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxST 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   35
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtbxUnidade 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      MaxLength       =   50
      TabIndex        =   33
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtbxUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   39
      Text            =   "0,0000"
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtbxQtd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   37
      Text            =   "0"
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtbxTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   41
      Text            =   "0,00"
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtbxIpi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7560
      TabIndex        =   43
      Text            =   "0,00"
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Inserir"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtbxCodBarra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   31
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox txtbxIcms 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8280
      TabIndex        =   45
      Text            =   "0,0"
      Top             =   4800
      Width           =   615
   End
   Begin VB.CheckBox chkLpt 
      Caption         =   "Imprimir LPT"
      Height          =   255
      Left            =   120
      TabIndex        =   79
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txtbxGrid 
      Height          =   405
      Left            =   6120
      TabIndex        =   76
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Emissão:"
      Height          =   735
      Left            =   8160
      TabIndex        =   8
      Top             =   0
      Width           =   1695
      Begin VB.TextBox txtbxEmiss 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Sair"
      Height          =   615
      Left            =   8280
      Picture         =   "frmGeraNotaAvulsa.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Gravar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      Picture         =   "frmGeraNotaAvulsa.frx":0F8C
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnImprNota 
      Caption         =   "&Imprimir Nota"
      Height          =   615
      Left            =   4920
      Picture         =   "frmGeraNotaAvulsa.frx":10D6
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnFrete 
      Caption         =   "&Frete"
      Height          =   615
      Left            =   1560
      Picture         =   "frmGeraNotaAvulsa.frx":1220
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnEspelho 
      Caption         =   "&Espelho de Nota"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      Picture         =   "frmGeraNotaAvulsa.frx":136A
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   47
      Top             =   5160
      Width           =   9735
      Begin VB.CommandButton cmndbtnDesconto 
         Caption         =   "De&sconto:"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   82
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Enabled         =   0   'False
         ForeColor       =   &H0000C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   81
         Text            =   "0,00"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtbxNota_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   450
         Left            =   8280
         MaxLength       =   8
         TabIndex        =   70
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtbxDesp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   5520
         TabIndex        =   66
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtbxTotalNot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   6960
         TabIndex        =   68
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtbxTotalProd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   6960
         TabIndex        =   58
         Text            =   "0,00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxVIcms 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   2640
         TabIndex        =   53
         Text            =   "0,00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxPicms 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   120
         TabIndex        =   49
         Text            =   "12"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1200
         TabIndex        =   51
         Text            =   "0,00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxBCSubs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   4080
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxVlrSubs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   5520
         TabIndex        =   56
         Text            =   "0,00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxValorFre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1200
         TabIndex        =   60
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtbxValorIpi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   4080
         TabIndex        =   64
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtbxValorSeg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   2640
         TabIndex        =   62
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblCUPOM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Não"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9000
         TabIndex        =   78
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CUPOM:"
         Height          =   195
         Left            =   8280
         TabIndex        =   77
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "&Número Nota:"
         Height          =   195
         Left            =   8280
         TabIndex        =   69
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Despesas Acess:"
         Height          =   195
         Left            =   5520
         TabIndex        =   65
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total da Nota:"
         Height          =   195
         Left            =   6960
         TabIndex        =   67
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Produtos"
         Height          =   195
         Left            =   6960
         TabIndex        =   57
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS(%)"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Base Cálc Icms.:"
         Height          =   195
         Left            =   1200
         TabIndex        =   50
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Icms:"
         Height          =   195
         Left            =   2640
         TabIndex        =   52
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&B.Cálc Icms Sub:"
         Height          =   195
         Left            =   4080
         TabIndex        =   54
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Icms Subs:"
         Height          =   195
         Left            =   5520
         TabIndex        =   55
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do &Seguro:"
         Height          =   195
         Left            =   2640
         TabIndex        =   61
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do &Frete:"
         Height          =   195
         Left            =   1200
         TabIndex        =   59
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do &IPI:"
         Height          =   195
         Left            =   4080
         TabIndex        =   63
         Top             =   720
         Width           =   870
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtbxNatOp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "VENDAS MERC. ESTADUAL"
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtbxCFOP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   4
         Text            =   "5102"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxIESubTrib 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Natureza da Operação:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&CFOP:"
         Height          =   195
         Left            =   4440
         TabIndex        =   3
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "I.E. do Subst. Trib."
         Height          =   195
         Left            =   5280
         TabIndex        =   5
         Top             =   120
         Width           =   1320
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Destinatário/Remetente:"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   9735
      Begin VB.TextBox txtbxCNPJDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtbxEndDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtbxBairroDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   22
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtbxCepDesp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   24
         Text            =   "75.000-000"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtbxMuniDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         TabIndex        =   26
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtbxFoneDet 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtbxIeDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtbxUFDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9240
         TabIndex        =   28
         Text            =   "GO"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtbxNomeDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Tag             =   "00000"
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&CNPJ/CPF:"
         Height          =   195
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/Razão Social"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Bairro/Distrito:"
         Height          =   195
         Left            =   3240
         TabIndex        =   21
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Left            =   5760
         TabIndex        =   23
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Município"
         Height          =   195
         Left            =   6840
         TabIndex        =   25
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Fone/Fax:"
         Height          =   195
         Left            =   8280
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Insc. Est:"
         Height          =   195
         Left            =   6840
         TabIndex        =   15
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   9240
         TabIndex        =   27
         Top             =   840
         Width           =   255
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxGrd 
      Height          =   2175
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3836
      _Version        =   393216
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&ST:"
      Height          =   195
      Left            =   3960
      TabIndex        =   34
      Tag             =   "B"
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&UN:"
      Height          =   195
      Left            =   3360
      TabIndex        =   32
      Tag             =   "B"
      Top             =   4560
      Width           =   285
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   6600
      TabIndex        =   40
      Top             =   4560
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Qtd:"
      Height          =   195
      Left            =   4920
      TabIndex        =   36
      Top             =   4560
      Width           =   300
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unitário"
      Height          =   195
      Left            =   5640
      TabIndex        =   38
      Top             =   4560
      Width           =   540
   End
   Begin VB.Label lblExpr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Produto"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Tag             =   "B"
      Top             =   4560
      Width           =   555
   End
   Begin VB.Label lblIPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Ipi(R$):"
      Height          =   195
      Left            =   7560
      TabIndex        =   42
      Top             =   4560
      Width           =   510
   End
   Begin VB.Label lblIcmsCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I&CMS(%)"
      Height          =   195
      Left            =   8280
      TabIndex        =   44
      Top             =   4560
      Width           =   600
   End
   Begin VB.Label lblItens 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Itens:"
      Height          =   195
      Left            =   8520
      TabIndex        =   80
      Top             =   6480
      Width           =   390
   End
   Begin VB.Menu mnMenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnTrib 
         Caption         =   "&Tributar Todas"
      End
      Begin VB.Menu mnNaoTrib 
         Caption         =   "&Não Tributar Todas"
      End
   End
End
Attribute VB_Name = "frmGeraNotaAvulsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bActive As Boolean

Private Sub chkClick_Click()

End Sub

Private Sub cmndbtnArred_Click()
    P_ValorLostFocus txtbxBC
    P_ValorLostFocus txtbxBCSubs
    P_ValorLostFocus txtbxVIcms
    P_ValorLostFocus txtbxVlrSubs
    P_ValorLostFocus txtbxTotalNot
    P_ValorLostFocus txtbxTotalProd
End Sub

Private Sub cmndbtnClose_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnConf_Click()
    If MsgBox("Confirma conclusão da NF?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    sNum = MaxKeySQL("tab_saidas", "nota_sai", "", "0000000#")
    txtbxNota_Sai.Text = sNum
    p_imprNF Printer, False
    If MsgBox("A Impressão saiu OK?", vbYesNo + vbQuestion) = vbNo Then
        If MsgBox("A Impressora está pronta para imprimir novamente?", vbOKCancel) = vbOK Then
            cmndbtnConf_Click
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    If cmndbtnConf.Tag = "" Then
        grava
    Else
        altera
    End If
    Me.Tag = "x"
    Me.Hide
    
    
    'Exit Sub
   
    
End Sub
Sub p_imprNF(obj As Object, Optional Entrada As Boolean = False)
    Dim iCount As Integer, sProd As String
    Dim sAltura As Single
    Dim sLargura As Single
    obj.FontSize = "10"
    obj.FontName = "Arial"
    Dim sYAj As Single
    sYAj = -15
    'On Error Resume Next
    sAltura = 210
    sLargura = 160 '17 cm
    obj.ScaleMode = vbMillimeters
'    Printer.PaperSize = 256
'    sAltura = obj.ScaleY(sAltura, vbMillimeters, vbTwips)
'    sLargura = obj.ScaleX(sLargura, vbMillimeters, vbTwips)
    
    Dim rPesq As Recordset
        strSQL = "select * from tab_confignf "
        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
        obj.FontSize = 12
        'Printer.BackStyle = 0
       ' Printer.PaperSize = 11 'a 5
        P_PrintXY obj, txtbxNota_Sai.Text, !numx_con, !numy_con + sYAj
        P_PrintXY obj, "X", IIf(Not Entrada, !marcax_con, !marcax_con + 15), !marcay_con + sYAj
        obj.FontName = "arial"
        obj.FontSize = 9
        P_PrintXY obj, txtbxNatOp.Text, !natopx_con, !natopy_con + sYAj
        P_PrintXY obj, txtbxCFOP.Text, !cfopx_con, !cfopy_con + sYAj
        P_PrintXY obj, txtbxIESubTrib.Text, !iex_con, !iey_con + sYAj
        P_PrintXY obj, txtbxNomeDest.Text, !nomedestx_con, !nomedesty_con + sYAj
        P_PrintXY obj, txtbxCNPJDest.Text, !cnpjx_con, !cnpjy_con + sYAj
        P_PrintXY obj, txtbxEmiss.Text, !emissx_con, !emissy_con + sYAj
        P_PrintXY obj, txtbxEndDest.Text, !enddestx_con, !enddesty_con + sYAj
        P_PrintXY obj, txtbxBairroDest.Text, !baidestx_con, !baidesty_con + sYAj
        P_PrintXY obj, txtbxCepDesp.Text, !cepdestx_con, !cepdesty_con + sYAj
        'P_PrintXY obj, txtbxDataSai.Text, 185, 50
        P_PrintXY obj, txtbxMuniDest.Text, !munidestx_con, !munidesty_con + sYAj
        P_PrintXY obj, txtbxFoneDet.Text, !fonedestx_con, !fonedesty_con + sYAj
        P_PrintXY obj, txtbxUFDest.Text, !ufdestx_con, !ufdesty_con + sYAj
        P_PrintXY obj, txtbxIeDest.Text, !iedestx_con, !iedesty_con + sYAj
    '    P_PrintXY obj, txtbxHoraSai.Text, 185, 57
        With msflxGrd
            obj.CurrentY = rPesq!iniprody_con + sYAj
            For iCount = 1 To .Rows - 1
                .Row = iCount
                If iItem = 20 Then
                    iItem = 0
                    obj.NewPage
                    
                    obj.CurrentY = rPesq!iniprody_con + sYAj
                     MsgBox "Clique em OK quando a impressora estiver pronta!", vbInformation
                    
                End If
                sProd = ""
                .Col = 0
                P_ImprX obj, CStr(iCount), 0 '
                P_ImprX obj, .Text, rPesq!codprod_con - 2
                .Col = 1
                P_ImprX obj, Left(.Text, 35), rPesq!descprod_con + 1
                .Col = 2
                P_ImprX obj, .Text, rPesq!stprod_con
                .Col = 3
                P_ImprX obj, .Text, rPesq!unid_con
                .Col = 4
                P_ImprX obj, .Text, rPesq!qtdprod_con, True, True
                .Col = 5
                P_ImprX obj, .Text, rPesq!unitprod_con, True, True
                .Col = 6
                P_ImprX obj, .Text, rPesq!totalprod_con, True, True
                .Col = 7
                P_ImprX obj, .Text, rPesq!aliqprod_con, False
                'obj.CurrentY = obj.CurrentY - 0.7
                iItem = iItem + 1
            Next
        End With
        obj.CurrentY = 240
        
        P_PrintXY obj, txtbxBC.Text, !bcx_con, !bcy_con + sYAj
        P_PrintXY obj, txtbxVIcms.Text, !vlricmsx_con, !vlricmsy_con + sYAj
        P_PrintXY obj, txtbxBCSubs.Text, !bcsubsx_con, !bcsubsy_con + sYAj
        P_PrintXY obj, txtbxVlrSubs.Text, !vlrsubsx_con, !vlrsubsy_con + sYAj
        P_PrintXY obj, txtbxTotalProd.Text, !vlrprodx_con, !vlrprody_con + sYAj
        P_PrintXY obj, txtbxValorFre.Text, !vlrfretex_con, !vlrfretey_con + sYAj
        P_PrintXY obj, txtbxValorSeg.Text, !vlrsegx_con, !vlrsegy_Con + sYAj
        P_PrintXY obj, txtbxDesp.Text, !outrasdespx_con, !outrasdespy_con + sYAj
        P_PrintXY obj, txtbxValorIpi.Text, !vlripix_con, !vlripiy_con + sYAj
        P_PrintXY obj, txtbxTotalNot.Text, !vlrtotalx_con, !vlrtotaly_con + sYAj
        P_PrintXY obj, txtbxNota_Sai.Text, !numrodapex_con, !numrodapey_con + sYAj
        
        If IsLoad("frmFrete") Then
            P_PrintXY obj, frmFrete.txtbxNomeFre.Text, !nomefrex_con, !nomefrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxContaFre.Text, !tipofrex_con, !tipofrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxPlacaFre.Text, !placafrex_con, !placafrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxUfVeiFre.Text, !ufveix_con, !ufveiy_con + sYAj
            P_PrintXY obj, frmFrete.txtbxCnpjFre.Text, !cnpjfrex_con, !cnpjfrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxEnderecoFre.Text, !endfrex_con, !endfrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxMuniFre.Text, !munifrex_con, !munifrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxUfFre.Text, !uffrex_con, !uffrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxIEFre.Text, !iefrex_con, !iefrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxQtdFre.Text, !qtdfrex_con, !qtdfrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxEspFre.Text, !espfrex_con, !espfrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxMarcaFre.Text, !marcafrex_con, !marcafrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxNumFre.Text, !numfrex_con, !numfrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxPBrutoFre.Text, !pbrutofrex_con, !pbrutofrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxPLiqFre.Text, !pliqfrex_con, !pliqfrey_con + sYAj
            P_PrintXY obj, frmFrete.txtbxDadosAdd.Text, !dadosaddx_con, !dadosaddy_con + sYAj
            If TypeName(obj) = "Printer" Then
                Printer.EndDoc
            End If
        End If
        .Close
    End With
End Sub
Sub p_imprNFLPT(Optional Entrada As Boolean = False)
    Dim iCount As Integer, sProd As String
    Dim sAltura As Single
    Dim sLargura As Single
    Dim obj As Object
    Dim sYAj As Single
    Dim Margem As Integer
    Dim sDesc As String
    Dim iItem As Integer
    Margem = 0
    Dim rPesq As Recordset
        strSQL = "select * from tab_confignf "
        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
        If Not F_OpenLpt Then
            Exit Sub
        End If
        P_FeedLpt1
        P_FontLpt "M", "4"
        P_PrintLpt "X", IIf(Not Entrada, 43, 48), True
        P_PrintLpt txtbxNota_Sai.Text, 18, False
        P_FeedLpt1 2
        
        
        P_PrintLpt F_EsticaTexto(txtbxNatOp.Text, 20) & " " & F_EsticaTexto(txtbxCFOP, 5) & F_EsticaTexto(txtbxIESubTrib.Text, 13), 0, False
        P_FeedLpt1 2
        P_PrintLpt F_EsticaTexto(txtbxNomeDest.Text, 43) & " " & F_EsticaTexto(txtbxCNPJDest.Text, 16) & F_EsticaTexto(txtbxEmiss.Text, 13), 0, False
        
        P_PrintLpt F_EsticaTexto(txtbxEndDest.Text, 31) & " " & F_EsticaTexto(txtbxBairroDest.Text, 18) & " " & F_EsticaTexto(txtbxCepDesp.Text, 6), 0, False
        
        P_FeedLpt1 1
        sDesc = F_EsticaTexto(txtbxMuniDest.Text, 26) & " "
        sDesc = sDesc & F_EsticaTexto(txtbxFoneDet.Text, 12) & " "
        sDesc = sDesc & F_EsticaTexto(txtbxUFDest.Text, 2) & " "
        sDesc = sDesc & F_EsticaTexto(txtbxIeDest.Text, 12)
        P_PrintLpt sDesc, 0, False
        P_FeedLpt1 1
        With msflxGrd
            
            For iCount = 1 To .Rows - 1
                .Row = iCount
                If iItem = 15 Then
                    
                    iItem = 0
                    P_FeedLpt1 26
                    P_PrintLpt "X", IIf(Not Entrada, 43, 48), True
                    P_PrintLpt txtbxNota_Sai.Text, 18, False
                    P_FeedLpt1 8
                    P_PrintLpt "Continuação NF: " & txtbxNota_Sai.Text, 0, False
                    
                End If
                sProd = ""
                .Col = 0
                
                sDesc = F_EsticaTexto(.Text, 8) & " "
                .Col = 1
                sDesc = sDesc & F_EsticaTexto(.Text, 26) & " "
                .Col = 2
                sDesc = sDesc & F_EsticaTexto(.Text, 2) & " "
                .Col = 3
                sDesc = sDesc & F_EsticaTexto(.Text, 3) & " "
                .Col = 4
                sDesc = sDesc & F_EsticaTexto(.Text, 5, 2) & " "
                .Col = 5
                sDesc = sDesc & F_EsticaTexto(.Text, 7, 2) & " "
                .Col = 6
                sDesc = sDesc & F_EsticaTexto(.Text, 10, 2) & " "
                .Col = 7
                sDesc = sDesc & F_EsticaTexto(.Text, 4)
                P_PrintLpt sDesc, 0, False
                
                iItem = iItem + 1
            Next
        End With
         
        P_FeedLpt1 16 - iItem
        sDesc = F_EsticaTexto(txtbxBC.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxVIcms.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxBCSubs.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxVlrSubs.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxTotalProd, 18, 2)
        P_PrintLpt sDesc, 0, False
        
        sDesc = F_EsticaTexto(txtbxValorFre.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxValorSeg.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxDesp.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxValorIpi.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxTotalNot.Text, 18, 2)
        P_PrintLpt sDesc, 0, False
        
        
        
        
        
        If IsLoad("frmFrete") Then
            sDesc = F_EsticaTexto(frmFrete.txtbxNomeFre.Text, 20) & Space(3)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxContaFre.Text, 2)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxPlacaFre.Text, 8)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxUfVeiFre.Text, 2)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxCnpjFre.Text, 16)
            P_PrintLpt sDesc, 0, False
            
            sDesc = F_EsticaTexto(frmFrete.txtbxEnderecoFre.Text, 20)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxMuniFre.Text, 14)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxUfFre.Text, 3)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxIEFre.Text, 12)
            P_PrintLpt sDesc, 0, False
            
            sDesc = F_EsticaTexto(frmFrete.txtbxQtdFre.Text, 20)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxEspFre.Text, 14)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxMarcaFre.Text, 3)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxNumFre.Text, 12)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxPBrutoFre.Text, 3)
            sDesc = sDesc & F_EsticaTexto(frmFrete.txtbxPLiqFre.Text, 12)
            P_PrintLpt sDesc, 0, False
            
            P_FeedLpt1
            P_PrintLpt F_Quebra(frmFrete.txtbxDadosAdd.Text, 14), 0, False
            P_FeedLpt1 4
        Else
            P_FeedLpt1 16
        End If
        .Close
        P_PrintLpt F_EsticaTexto(txtbxNota_Sai.Text, 65, 2)
        P_FeedLpt1 5
    End With
    
    If Not F_CloseLpt Then
            Exit Sub
        End If
End Sub

Sub imprimeEspelho(obj As Object)
    Dim iCount As Integer
    Dim iItem As Integer
    Dim sProd As String
    P_CabRelTimbr "Espelho de Nota Fiscal: " & txtbxNota_Sai & "-" & txtbxNomeDest.Text, obj
    P_ImprLinha obj
    P_ImprX obj, "Número: " & txtbxNota_Sai.Text, 0
    P_ImprX obj, "Base Cálc: " & txtbxBC.Text, 32
    P_ImprX obj, "ICMS(%): " & txtbxPicms.Text & "%", 62
    P_ImprX obj, "Vlr ICMS: " & txtbxVIcms.Text, 87
    P_ImprX obj, "Desp: " & txtbxDesp.Text, 115
    P_ImprX obj, "Produtos: " & txtbxTotalProd, 140
    P_ImprX obj, "Total: " & txtbxTotalProd, 175, False
    P_ImprLinha obj
    P_ImprX obj, "Cód", 0
    P_ImprX obj, "Descrição dos Produtos", 12
    P_ImprX obj, "Un.", 140
    P_ImprX obj, "Qtd.", 160, True, True
    P_ImprX obj, "Unit.", 175, True, True
    P_ImprX obj, "Total", 200, False, True
    P_ImprLinha obj
    With msflxGrd
        
        For iCount = 1 To .Rows - 1
            .Row = iCount
            If iItem = 45 Then
                iItem = 0
                If TypeName(obj) = "Printer" Then
                    P_Rodape strNome_Emp, True, "Continua..."
                End If
                P_ImprX obj, "Cód", 0
                P_ImprX obj, "Descrição dos Produtos", 12
                P_ImprX obj, "Un.", 140
                P_ImprX obj, "Qtd.", 160, True, True
                P_ImprX obj, "Unit.", 175, True, True
                P_ImprX obj, "Total", 200, False, True
                P_ImprLinha obj
                
            End If
            sProd = ""
            .Col = 0
            P_ImprX obj, .Text, 0
            .Col = 2
            sProd = "Art. Emb " & .Text
            .Col = 1
            sProd = sProd & " Ref. Ord: " & .Text
            .Col = 4
            sProd = sProd & "(" & .Text & ")"
            P_ImprX obj, sProd, 12
            .Col = 3
            P_ImprX obj, .Text, 140
            .Col = 5
            P_ImprX obj, .Text, 160, True, True
            .Col = 6
            P_ImprX obj, .Text, 175, True, True
            .Col = 7
            P_ImprX obj, .Text, 200, False, True
            
            
            iItem = iItem + 1
        Next
        P_ImprLinha obj
        'P_ImprX obj, lblTotalArt.caption, 160, True, False
            
    End With
    
    
End Sub
Sub buscaTerceiro()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error Resume Next
    strSQL = "select * from tab_fornecedores where cod_for='" & txtbxNomeDest.Tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then Exit Sub
        txtbxBairroDest.Text = !bairro_FOR
        txtbxCepDesp.Text = !cep_for
        txtbxCNPJDest.Text = !cnpj_for
        txtbxEndDest.Text = !endereco_for
        txtbxFoneDet.Text = !telefone_for
        txtbxIeDest.Text = !ie_for
        txtbxMuniDest.Text = !cidade_for
        txtbxNomeDest.Text = !nome_for
        txtbxUFDest.Text = !uf_for
        
    End With
End Sub

Private Sub cmndbtnConfirma_Click()
    Dim strSQL As String
    strSQL = "" & Chr(9)
    strSQL = strSQL & txtbxCodBarra.Text & Chr(9)
    strSQL = strSQL & txtbxST.Text & Chr(9)
    strSQL = strSQL & txtbxUnidade.Text & Chr(9)
    strSQL = strSQL & txtbxQtd.Text & Chr(9)
    strSQL = strSQL & txtbxUnit.Text & Chr(9)
    strSQL = strSQL & txtbxTotal.Text & Chr(9)
    strSQL = strSQL & txtbxIcms.Text & Chr(9)
    strSQL = strSQL & txtbxIpi.Text
    msflxGrd.AddItem strSQL
    
    txtbxCodBarra.SetFocus
    
    
End Sub

Private Sub cmndbtnDesconto_Click()
    Dim sUsu As String
    sUsu = F_SenhaCodBar
    If txtbxTotalNot.Text = "0,00" Then Exit Sub
    If sUsu <> "" Then
        txtbxDesconto.Enabled = True
        cmndbtnDesconto.Tag = sUsu
        txtbxDesconto.SetFocus
    End If
    
End Sub

Private Sub cmndbtnEspelho_Click()
    
        imprimeEspelho frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.Tag <> "" Then
            imprimeEspelho Printer
            P_Rodape strNome_Emp, True, "(F)"
            Printer.EndDoc
        End If
        Unload frmVisualiza
        
        
        
    
End Sub

Private Sub cmndbtnFrete_Click()
    frmFrete.Show 1
End Sub

Private Sub cmndbtnImprNota_Click()
    
    
    
'    If Me.Tag = "" Then
'        MsgBox "Você deve primeiramente gravar a NF!", vbInformation
'        Exit Sub
'    End If
    Dim sAltura As Single
    Dim sLargura As Single
    If chkLpt.Value = 1 Then
        p_imprNFLPT
        Exit Sub
    End If
    
    sAltura = 210
    sLargura = 157 '17 cm
    Printer.ScaleMode = vbMillimeters
    'Printer.PaperSize = 256
    sAltura = Printer.ScaleY(sAltura, vbMillimeters, vbTwips)
    sLargura = Printer.ScaleX(sLargura, vbMillimeters, vbTwips)
    Printer.Height = sAltura
    Printer.Width = sLargura
    
    'p_imprNF frmVisualiza.Pic
    'frmVisualiza.Show 1
    'Unload frmVisualiza
    'Exit Sub
    p_imprNF Printer
    Printer.EndDoc
    If MsgBox("A Impressão foi bem sucedida?", vbYesNo + vbQuestion) = vbYes Then
        cmndbtnSalvar_Click
        bActive = False
        msflxGrd.Rows = 1
        Form_Activate
        txtbxVenda.Text = ""
        soma
        txtbxVenda.SetFocus
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Sub grava()
    Dim strSQL As String
    Dim sSeq As String
    Dim iCount As Integer
    sSeq = MaxKeySQL("tab_saidasnf", "seq_sai", "", "0000#")
    strSQL = "INSERT INTO tab_saidasNF (SEQ_SAI, COD_CLI, NOTA_SAI, SERIE_SAI, "
    strSQL = strSQL & "NATOP_SAI, CFOP_SAI, VALOR_SAI, BC_SAI, VICMS_SAI, PICMS_SAI, "
    strSQL = strSQL & "VALP_SAI, COD_USU, EMISSAO_SAI, situacao_sai, despesas_Sai,num_ven) VALUES ("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & txtbxNomeDest.Tag & "',"
    strSQL = strSQL & "'" & txtbxNota_Sai.Text & "',"
    strSQL = strSQL & "'" & "1" & "',"
    strSQL = strSQL & "'" & txtbxNatOp.Text & "',"
    strSQL = strSQL & "'" & txtbxCFOP.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxTotalNot.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxBC.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxVIcms.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxPicms.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxTotalProd.Text) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & F_GravaData(txtbxEmiss.Text) & ","
    strSQL = strSQL & "'1',"
    strSQL = strSQL & F_GravaDec(txtbxDesp.Text) & ", "
    strSQL = strSQL & "'" & txtbxVenda.Text & "') "
    Banco.Execute strSQL
    With msflxGrd
        For iCount = 1 To .Rows - 1
            strSQL = "insert into tab_itsainf(seq_sai,seq_its,cod_pro,qtd_its,unit_its)values("
            strSQL = strSQL & "'" & sSeq & "',"
            strSQL = strSQL & iCount & ","
            .Col = 0
            strSQL = strSQL & "'" & .Text & "', "
            .Col = 4
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .Col = 5
            strSQL = strSQL & F_GravaDec(.Text)
            strSQL = strSQL & ")"
            Banco.Execute strSQL
        Next
    End With
    Me.Tag = sSeq
    P_GravaFrete
    bActive = False
    Form_Activate
End Sub
Sub P_GravaFrete()
    If IsLoad("frmFrete") Then
        With frmFrete
            strSQL = "INSERT INTO tab_fretes (seq_sai, nome_fre, plveic_fre, ufveic_fre, cnpj_fre, "
            strSQL = strSQL & "endereco_fre, municipio_fre, uf_fre, ie_fre, qdt_fre, esp_fre, "
            strSQL = strSQL & "marca_fre, num_fre, pb_fre, pl_fre, dadosadd_fre) VALUES ("
            strSQL = strSQL & "'" & Me.Tag & "',"
            strSQL = strSQL & "'" & .txtbxNomeFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxPlacaFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxUfVeiFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxCnpjFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxEnderecoFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxMuniFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxUfFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxIEFre.Text & "',"
            strSQL = strSQL & F_GravaDec(.txtbxQtdFre.Text) & ","
            strSQL = strSQL & "'" & .txtbxEspFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxMarcaFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxNumFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxPBrutoFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxPLiqFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxDadosAdd.Text & "')"
            Banco.Execute strSQL
        End With
        
    End If
    
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If msflxGrd.Rows = 1 Then
        MsgBox "Não há itens..."
        Exit Sub
    End If
    If txtbxNota_Sai.Text = "" Then
        MsgBox "Informe o número da Nota!", vbExclamation
        txtbxNota_Sai.SetFocus
        Exit Sub
    End If
'    strSQL = F_BuscaUmCampo("nota_sai", "tab_saidasnf", "num_ven='" & txtbxVenda.Text & "'")
'    If strSQL <> "" Then
'        If MsgBox("Já existe uma nota para esta venda: " & strSQL, vbYesNo + vbCritical) = vbNo Then
'            Exit Sub
'        End If
'    End If
'    If MsgBox("Confirma Gravação da NF?", vbYesNo + vbQuestion) = vbYes Then
 '       grava
'    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    'P_Status "Aguarde..."
    If bActive Then Exit Sub
    busca
    sNum = MaxKeySQL("tab_saidasnf", "nota_sai", "", "0000000#")
    txtbxNota_Sai.Text = sNum
    bActive = True
End Sub
Sub somaItem()
    Dim cQtd As Currency
    Dim cUnit As Currency
    With msflxGrd
        .Col = 4: cQtd = .Text
        .Col = 5: cUnit = .Text
        .Col = 6: .Text = FormatVal(cQtd * cUnit)
        
    End With
End Sub
Sub soma()
    Dim iCount As Integer
    Dim cTotal As Currency
    Dim cICMS As Currency
    Dim cBC As Currency
    Dim cTotProd As Currency
    Dim iLinhas As Integer
    Dim iColula As Integer
    With msflxGrd
        iLinhas = .Row
        iColula = .Col
        For iCount = 1 To .Rows - 1
            .Row = iCount
            .Col = 6
            cTotProd = .Text
            cTotal = cTotal + CCur(.Text)
            .Col = 7
            cICMS = cICMS + (cTotProd * (CCur(.Text) / 100))
            If CCur(.Text) > 0 Then
                .Col = 6
                cBC = cBC + CCur(.Text)
            End If
            
            
        Next
        txtbxBC.Text = FormatVal(cBC)
        txtbxTotalNot.Text = FormatVal(cTotal)
        txtbxTotalProd.Text = FormatVal(cTotal)
        txtbxVIcms.Text = FormatVal(cICMS)
        .Row = iLinhas
        .Col = iColula
    End With
End Sub
Sub gravaNota()
    On Error GoTo trata_erro
    Dim sNum As String
    sNum = MaxKeySQL("tab_saidas", "nota_sai", "", "0000000#")
    txtbxNota_Sai.Text = sNum
    strSQL = "delete from tab_notasai where seq_sai='" & Me.Tag & "' "
    Banco.Execute strSQL
    strSQL = "insert into tab_notasai(seq_sai,nomedest_sai, cnpjdest_sai, enddest_sai,"
    strSQL = strSQL & "bairrodest_sai, cepdest_sai, munidest_sai, fonedest_sai,"
    strSQL = strSQL & "ufdest_sai,iedest_sai, emissao_sai, datahora_sai, BC_SAI,"
    strSQL = strSQL & "VICMS_SAI,PICMS_SAI,bcsubs_SAI,vsubs_SAI,vipi_sai,vfrete_sai,"
    strSQL = strSQL & "vseg_sai,totalprod_sai,totalnota_sai,nomefre_sai,contafre_sai,"
    strSQL = strSQL & "plveicfre_sai,ufveicfre_sai,cnpjfre_sai,endfre_sai,munifre_sai,"
    strSQL = strSQL & "uffre_sai,iefre_sai,qtdfre_sai,espfre_sai,numfre_sai,"
    strSQL = strSQL & "marcafre_sai,pbrutofre_sai, pliqfre_sai,dadosad_sai)values("
    strSQL = strSQL & "'" & Me.Tag & "',"
    strSQL = strSQL & "'" & txtbxNomeDest.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJDest.Text & "',"
    strSQL = strSQL & "'" & txtbxEndDest.Text & "',"
    strSQL = strSQL & "'" & txtbxBairroDest.Text & "',"
    strSQL = strSQL & "'" & txtbxCepDesp.Text & "',"
    strSQL = strSQL & "'" & txtbxMuniDest.Text & "',"
    strSQL = strSQL & "'" & txtbxFoneDet.Text & "',"
    strSQL = strSQL & "'" & txtbxUFDest.Text & "',"
    strSQL = strSQL & "'" & txtbxIeDest.Text & "',"
    strSQL = strSQL & F_GravaData(txtbxEmiss.Text) & ","
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & F_GravaDec(txtbxBC.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxVIcms.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxPicms.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxBCSubs.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxVlrSubs.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxValorIpi.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxValorFre.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxValorSeg.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxTotalProd.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxTotalNot.Text) & ", "
    strSQL = strSQL & "'" & txtbxNomeFre.Text & "', "
    strSQL = strSQL & "'" & txtbxContaFre.Text & "', "
    strSQL = strSQL & "'" & txtbxPlacaFre.Text & "',"
    strSQL = strSQL & "'" & txtbxUfVeiFre.Text & "',"
    strSQL = strSQL & "'" & txtbxCnpjFre.Text & "',"
    strSQL = strSQL & "'" & txtbxEnderecoFre.Text & "',"
    strSQL = strSQL & "'" & txtbxMuniFre.Text & "',"
    strSQL = strSQL & "'" & txtbxUfFre.Text & "',"
    strSQL = strSQL & "'" & txtbxIEFre.Text & "',"
    strSQL = strSQL & "'" & txtbxQtdFre.Text & "',"
    strSQL = strSQL & "'" & txtbxEspFre.Text & "',"
    strSQL = strSQL & "'" & txtbxNumFre.Text & "',"
    strSQL = strSQL & "'" & txtbxMarcaFre.Text & "',"
    strSQL = strSQL & "'" & txtbxPBrutoFre.Text & "',"
    strSQL = strSQL & "'" & txtbxPLiqFre.Text & "',"
    strSQL = strSQL & "'" & txtbxDadosAd.Text & "')"
    Banco.Execute strSQL
    strSQL = "update tab_saidas set nota_sai='" & sNum & "' "
    strSQL = strSQL & " where seq_sai='" & Me.Tag & "' "
    Banco.Execute strSQL
Exit Sub
trata_erro:
MsgBox Err.Description, vbCritical



End Sub
Sub altera()
    On Error GoTo trata_erro
    strSQL = "update tab_notasai set "
    
    
    
    strSQL = strSQL & "nomedest_sai='" & txtbxNomeDest.Text & "',"
    strSQL = strSQL & "cnpjdest_sai='" & txtbxCNPJDest.Text & "',"
    strSQL = strSQL & "enddest_sai='" & txtbxEndDest.Text & "',"
    strSQL = strSQL & "bairrodest_sai='" & txtbxBairroDest.Text & "',"
    strSQL = strSQL & "cepdest_sai='" & txtbxCepDesp.Text & "',"
    strSQL = strSQL & "munidest_sai='" & txtbxMuniDest.Text & "',"
    strSQL = strSQL & "fonedest_sai='" & txtbxFoneDet.Text & "',"
    strSQL = strSQL & "ufdest_sai='" & txtbxUFDest.Text & "',"
    strSQL = strSQL & "iedest_sai='" & txtbxIeDest.Text & "',"
    strSQL = strSQL & "emissao_sai=" & F_GravaData(txtbxEmiss.Text) & ","
    strSQL = strSQL & "datahora_sai=" & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "BC_SAI=" & F_GravaDec(txtbxBC.Text) & ", "
    strSQL = strSQL & "VICMS_SAI=" & F_GravaDec(txtbxVIcms.Text) & ", "
    strSQL = strSQL & "PICMS_SAI=" & F_GravaDec(txtbxPicms.Text) & ", "
    strSQL = strSQL & "bcsubs_SAI=" & F_GravaDec(txtbxBCSubs.Text) & ", "
    strSQL = strSQL & "vsubs_SAI=" & F_GravaDec(txtbxVlrSubs.Text) & ", "
    strSQL = strSQL & "vipi_sai=" & F_GravaDec(txtbxValorIpi.Text) & ", "
    strSQL = strSQL & "vfrete_sai=" & F_GravaDec(txtbxValorFre.Text) & ", "
    strSQL = strSQL & "vseg_sai=" & F_GravaDec(txtbxValorSeg.Text) & ", "
    strSQL = strSQL & "totalprod_sai=" & F_GravaDec(txtbxTotalProd.Text) & ", "
    strSQL = strSQL & "totalnota_sai=" & F_GravaDec(txtbxTotalNot.Text) & ", "
    strSQL = strSQL & "nomefre_sai='" & txtbxNomeFre.Text & "', "
    strSQL = strSQL & "contafre_sai='" & txtbxContaFre.Text & "', "
    strSQL = strSQL & "plveicfre_sai='" & txtbxPlacaFre.Text & "',"
    strSQL = strSQL & "ufveicfre_sai='" & txtbxUfVeiFre.Text & "',"
    strSQL = strSQL & "cnpjfre_sai='" & txtbxCnpjFre.Text & "',"
    strSQL = strSQL & "endfre_sai='" & txtbxEnderecoFre.Text & "',"
    strSQL = strSQL & "munifre_sai='" & txtbxMuniFre.Text & "',"
    strSQL = strSQL & "uffre_sai='" & txtbxUfFre.Text & "',"
    strSQL = strSQL & "iefre_sai='" & txtbxIEFre.Text & "',"
    strSQL = strSQL & "qtdfre_sai='" & txtbxQtdFre.Text & "',"
    strSQL = strSQL & "espfre_sai='" & txtbxEspFre.Text & "',"
    strSQL = strSQL & "numfre_sai='" & txtbxNumFre.Text & "',"
    strSQL = strSQL & "marcafre_sai='" & txtbxMarcaFre.Text & "',"
    strSQL = strSQL & "pbrutofre_sai='" & txtbxPBrutoFre.Text & "',"
    strSQL = strSQL & "pliqfre_sai='" & txtbxPLiqFre.Text & "', "
    strSQL = strSQL & "Dadosad_sai='" & txtbxDadosAd.Text & "' "
    strSQL = strSQL & " where seq_sai='" & Me.Tag & "' "
    
    Banco.Execute strSQL
Exit Sub
trata_erro:
MsgBox Err.Description, vbCritical
End Sub

Sub busca()
    Dim rPesq As Recordset
    Dim rOrd As Recordset
    Dim sAux As String
    Dim sArt As String
    Dim cPreco As Currency
    Dim cQtd As Currency
    Dim bIgual As Boolean
    msflxGrd.Rows = 1
    
    strSQL = "select tv.* ,tc.* from tab_vendas as tv,tab_clientes as tc where num_Ven='" & Me.Tag & "' "
    strSQL = strSQL & " and tv.cod_cli=tc.cod_cli "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        sArt = ""
        Do While Not .EOF
            
            txtbxNomeDest.Text = !nomecli_ven
            buscaCLi
            txtbxCNPJDest.Text = !cpf_cli
            txtbxIeDest.Text = Format(!rg_cli)
            txtbxEndDest.Text = Format(!endereco_cli)
            txtbxBairroDest.Text = Format(!bairro_cli)
            txtbxFoneDet.Text = Format(!telefone_cli)
            txtbxEndDest.Text = Format(!endereco_cli)
            txtbxBairroDest.Text = Format(!bairro_cli)
            txtbxCepDesp.Text = Format(!cep_cli)
            txtbxMuniDest.Text = Format(!cidade_cli)
            txtbxUFDest.Text = Format(!uf_cli)
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select tp.*,ti.* from tab_itvendas as ti,tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.num_ven='" & Me.Tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            sAux = ""
            sAux = sAux & !cod_pro & Chr(9)
            sAux = sAux & !nome_pro & Chr(9)
            sAux = sAux & !st_pro & Chr(9)
            sAux = sAux & !unidade_pro & Chr(9)
            sAux = sAux & Format(!qtd_ite) & Chr(9)
            sAux = sAux & FormatVal(!final_ite / !qtd_ite, "###,##0.00") & Chr(9)
            sAux = sAux & FormatVal(!final_ite, "###,##0.000") & Chr(9)
            If !st_pro = "T" Then
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & "15"
                Else
                    sAux = sAux & "12"
                End If
            Else
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & "0"
                Else
                    sAux = sAux & "12"
                End If
            End If
            If bIgual = False Then
                msflxGrd.AddItem sAux
            End If
            .MoveNext
        Loop
    End With
    msflxGrd.Visible = False
    soma
    zebraGrid msflxGrd, 2
    lblItens.Caption = msflxGrd.Rows - 1
    msflxGrd.Visible = True
    If IsLoad("frmFrete") Then
        Unload frmFrete
    End If
End Sub

Private Sub Form_Load()
    bActive = False
    txtbxEmiss.Text = Date
    With msflxGrd
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = 9
        .Col = 0: .Text = "Cód"
        .Col = 1: .Text = "Descrição"
        .Col = 2: .Text = "St"
        .Col = 3: .Text = "Un."
        .Col = 4: .Text = "Qtd"
        .Col = 5: .Text = "Unit."
        .Col = 6: .Text = "Total"
        .Col = 7: .Text = "ICMS(%)"
        .Col = 8: .Text = "IPI(%)"
        .ColWidth(0) = 1000
        .ColWidth(1) = 3500
        .ColWidth(2) = 600
        .ColWidth(3) = 800
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 700
        .ColWidth(8) = 700
        .ColAlignment(0) = 0
        .ColAlignment(1) = 0
        .ColAlignment(2) = 0
    End With
End Sub

Private Sub mnNaoTrib_Click()
'
End Sub

Private Sub mnTrib_Click()
    '
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And msflxGrd.Rows > 1 Then
        With msflxGrd
            If .Col = 4 Or .Col = 5 Or .Col = 7 Then
                txtbxGrid.Text = .Text
                txtbxGrid.Left = .CellLeft + .Left
                txtbxGrid.Top = .CellTop + .Top
                txtbxGrid.Width = .CellWidth
                txtbxGrid.Height = .CellHeight
                txtbxGrid.Visible = True
                txtbxGrid.SetFocus
            End If
            
        End With
    End If
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With msflxGrd
        If .Col = 7 And Button = 2 Then
           Me.PopupMenu mnMenu, , .CellLeft + .Left, .CellTop + .Top
        End If
    End With
End Sub

Private Sub txtbxBairroDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCepDesp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCnpjFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    Dim cDesc As Currency
    On Error GoTo trata_erro
    P_ValorLostFocus txtbxDesconto
    cDesc = CCur(F_BuscaUmCampo("desconto_usu", "tab_defusu", _
        "cod_usu='" & cmndbtnDesconto.Tag & "'"))
    
    If CCur(txtbxDesconto.Text) > CCur(txtbxTotalNot.Text) * cDesc / 100 Then
        MsgBox "Desconto muito alto!", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    
    End If
    lblValor.Caption = FormatVal(lblValor.Tag) - CCur(txtbxDesconto.Text)
    txtbxDesconto.Enabled = False
    Exit Sub
trata_erro:
    P_MostraErro
    txtbxDesconto.Text = "0"
End Sub

Private Sub txtbxEmiss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEndDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEnderecoFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEspFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFoneDet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxHoraSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxGrid_GotFocus()
    P_SelecTexto txtbxGrid
End Sub

Private Sub txtbxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If msflxGrd.Col = 4 Then
            P_ValorLostFocus txtbxGrid, ""
        Else
            P_ValorLostFocus txtbxGrid
        End If
        msflxGrd.Text = txtbxGrid.Text
        txtbxGrid_LostFocus
        somaItem
        
        soma
        
    End If
End Sub

Private Sub txtbxGrid_LostFocus()
    txtbxGrid.Visible = False
    msflxGrd.SetFocus
End Sub

Private Sub txtbxIcms_GotFocus()
    P_SelecTexto txtbxIcms
End Sub

Private Sub txtbxIcms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIcms_LostFocus()
    P_ValorLostFocus txtbxIcms, "##"
End Sub

Private Sub txtbxIeDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIEFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIESubTrib_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMarcaFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIpi_GotFocus()
    P_SelecTexto txtbxIpi
End Sub

Private Sub txtbxIpi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIpi_LostFocus()
    P_ValorLostFocus txtbxIpi
End Sub

Private Sub txtbxMuniDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMuniFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNatOp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCFOP
    End If
End Sub
Sub buscaCFOP()
    Dim strSQL As String
    strSQL = F_CFOP
    If strSQL <> "" Then
        txtbxNatOp.Text = Mid(strSQL, 6, Len(strSQL))
        txtbxCFOP.Text = Left(strSQL, 4)
        
    End If
End Sub
Private Sub txtbxNomeDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCLi
    End If
End Sub
Sub buscaCLi()
    Dim t As Cliente
    t = F_BuscaCliente(txtbxNomeDest.Text)
    If t.cod_cli <> "" Then
        txtbxNomeDest.Tag = t.cod_cli
        txtbxNomeDest.Text = t.nome_Cli
        txtbxCNPJDest.Text = t.cnpj_Cli
        txtbxIeDest.Text = t.ie_Cli
        txtbxEndDest.Text = t.endereco_cli
        txtbxMuniDest.Text = t.cidade_cli
        txtbxBairroDest.Text = t.bairro_cli
        txtbxCepDesp.Text = t.cep_cli
        txtbxUFDest.Text = t.uf_cli
        txtbxFoneDet.Text = t.telefone_cli
        buscaUF
    End If
End Sub
Private Sub txtbxNomeFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNumFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPBrutoFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPlacaFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPLiqFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtdFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###"
    p_totalItem
End Sub

Private Sub txtbxST_GotFocus()
    P_SelecTexto txtbxST
End Sub

Private Sub txtbxST_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxST_LostFocus()
    P_AjustaTexto txtbxST
End Sub

Private Sub txtbxUFDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUfFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUfVeiFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub


Private Sub txtbxUFDest_LostFocus()
    buscaUF
End Sub
Sub buscaUF()
    If txtbxUFDest.Text = "GO" Then
        With msflxGrd
            For iCount = 1 To .Rows - 1
                .Col = 2
                If .Text = "T" Then
                    .Col = 7
                    .Text = "15"
                End If
            Next
        End With
    Else
        With msflxGrd
            For iCount = 1 To .Rows - 1
                
                    .Col = 7
                    .Text = "12"
                
            Next
        End With
    End If
End Sub

Sub p_totalItem()
    txtbxTotal.Text = FormatVal((txtbxUnit.Text) * CCur(txtbxQtd.Text))
End Sub
Private Sub txtbxVenda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxVenda_LostFocus
        Me.Tag = txtbxVenda.Text
        busca
        
    End If
End Sub

Private Sub txtbxVenda_LostFocus()
    txtbxVenda.Text = Format(txtbxVenda.Text, "000000#")
End Sub

Private Sub txtbxUnidade_GotFocus()
    P_SelecTexto txtbxUnidade
End Sub

Private Sub txtbxUnidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUnidade_LostFocus()
    P_AjustaTexto txtbxUnidade
End Sub

Private Sub txtbxUnit_GotFocus()
    P_SelecTexto txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit
    p_totalItem
End Sub
