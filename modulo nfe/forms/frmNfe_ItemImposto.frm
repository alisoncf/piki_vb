VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNfe_ItemImposto 
   Caption         =   "Imposto do Item da NFE"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12750
   Icon            =   "frmNfe_ItemImposto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   12750
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab sstb 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "ICMS"
      TabPicture(0)   =   "frmNfe_ItemImposto.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTipoIcms"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label24"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblT"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label22"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "gridValores"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAlterar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAdd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExcluir"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbOrigem"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbxTipoICMS"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "PIS"
      TabPicture(1)   =   "frmNfe_ItemImposto.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "COFINS"
      TabPicture(2)   =   "frmNfe_ItemImposto.frx":1D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "IPI"
      TabPicture(3)   =   "frmNfe_ItemImposto.frx":1D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.ComboBox cmbxTipoICMS 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   8535
      End
      Begin VB.ComboBox cmbOrigem 
         Height          =   315
         ItemData        =   "frmNfe_ItemImposto.frx":1D3A
         Left            =   120
         List            =   "frmNfe_ItemImposto.frx":1D59
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   8535
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   11280
         TabIndex        =   3
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   10080
         TabIndex        =   2
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   8880
         TabIndex        =   1
         Top             =   2640
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid gridValores 
         Height          =   1935
         Left            =   8880
         TabIndex        =   4
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3413
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblT 
         Caption         =   "Tipo de ICMS"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Valores:"
         Height          =   195
         Left            =   8880
         TabIndex        =   8
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblTipoIcms 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo_Icms"
         Height          =   195
         Left            =   7920
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmNfe_ItemImposto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
