VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmChecarInutilizacoes 
   Caption         =   "Checagem de Notas Emitidas"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   15615
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9128
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   15585
      TabIndex        =   0
      Top             =   5775
      Width           =   15615
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   14520
         Picture         =   "frmChecarInutilizacoes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdInutilizar 
         Caption         =   "Inutilizar"
         Height          =   735
         Left            =   13320
         Picture         =   "frmChecarInutilizacoes.frx":024A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   735
         Left            =   12120
         Picture         =   "frmChecarInutilizacoes.frx":1269
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdRetransm 
         Caption         =   "Contingência"
         Height          =   735
         Left            =   10800
         Picture         =   "frmChecarInutilizacoes.frx":1698
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Verificar"
         Height          =   735
         Left            =   8880
         Picture         =   "frmChecarInutilizacoes.frx":1F62
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frmChecarInutilizacoes.frx":2404
         Left            =   4560
         List            =   "frmChecarInutilizacoes.frx":2414
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbxModelo 
         Height          =   315
         ItemData        =   "frmChecarInutilizacoes.frx":244F
         Left            =   6480
         List            =   "frmChecarInutilizacoes.frx":245C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbxSerie 
         Height          =   315
         ItemData        =   "frmChecarInutilizacoes.frx":2480
         Left            =   7920
         List            =   "frmChecarInutilizacoes.frx":248D
         TabIndex        =   2
         Text            =   "cmbxSerie"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdEmpresa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4200
         Picture         =   "frmChecarInutilizacoes.frx":24B1
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   4560
         TabIndex        =   13
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         Height          =   195
         Left            =   1320
         TabIndex        =   12
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo:"
         Height          =   195
         Left            =   6480
         TabIndex        =   10
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         Height          =   195
         Left            =   7920
         TabIndex        =   9
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   120
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmChecarInutilizacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInutilizar_Click()
    frmInutilizarNFe.Show 1
End Sub
