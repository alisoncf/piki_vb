VERSION 5.00
Begin VB.Form frmNCM 
   Caption         =   "Operações Fiscais sobre NCM"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcCST 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6225
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton Command1 
         Caption         =   "Mudar Todos"
         Height          =   495
         Left            =   4200
         TabIndex        =   28
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtCst_Icms 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Tag             =   "00"
         Text            =   "00 - Tributada Integralmente"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.PictureBox pcCstIcms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         Picture         =   "frmNCM.frx":0000
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   15
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtCst_Ipi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   3735
      End
      Begin VB.PictureBox pcCstIpi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         Picture         =   "frmNCM.frx":03AB
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   13
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox txtCst_Pis 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Tag             =   "49"
         Text            =   "49 - Outras Operações de Saída"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.PictureBox pcCstPis 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         Picture         =   "frmNCM.frx":0756
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   11
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtCst_Cofins 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Tag             =   "49"
         Text            =   "49 - Outras Operações de Saída"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.PictureBox pcCstCofins 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         Picture         =   "frmNCM.frx":0B01
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   9
         Top             =   2880
         Width           =   255
      End
      Begin VB.TextBox txtbxICMS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Tag             =   "17"
         Text            =   "ICMS 17%"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.PictureBox pcIcms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         Picture         =   "frmNCM.frx":0EAC
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   7
         Top             =   1680
         Width           =   255
      End
      Begin VB.PictureBox pcGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         Picture         =   "frmNCM.frx":1257
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   6
         Top             =   4080
         Width           =   255
      End
      Begin VB.TextBox txtGenero 
         Appearance      =   0  'Flat
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
         TabIndex        =   5
         Top             =   4080
         Width           =   3735
      End
      Begin VB.TextBox txtOrigem 
         Appearance      =   0  'Flat
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
         TabIndex        =   4
         Top             =   4680
         Width           =   3735
      End
      Begin VB.PictureBox pcOrigem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         Picture         =   "frmNCM.frx":1602
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   3
         Top             =   4680
         Width           =   255
      End
      Begin VB.TextBox txtNCMCopia 
         Appearance      =   0  'Flat
         BackColor       =   &H0073E1F4&
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
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtEanCopia 
         Appearance      =   0  'Flat
         BackColor       =   &H0073E1F4&
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cst ICMS:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cst IPI:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   510
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cst PIS:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cst COFINS:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Códigos de Situação tributária em operações de saída (vendas):"
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
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   5265
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "&ICMS:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gênero: (Apenas para produtos primários ex.: insumos)"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Width           =   3840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origem:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   4440
         Width           =   540
      End
      Begin VB.Label lblIcms 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Situação Trib."
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
         Height          =   195
         Left            =   2760
         TabIndex        =   19
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&NCM:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN/GTIN:"
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmNCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
