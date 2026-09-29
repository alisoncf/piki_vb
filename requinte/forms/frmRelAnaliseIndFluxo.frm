VERSION 5.00
Begin VB.Form frmRelAnaliseIndFluxo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Análise Direcionada de Fluxo"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame7 
         Caption         =   "Impostos:"
         Height          =   1095
         Left            =   5160
         TabIndex        =   55
         Top             =   4800
         Width           =   3375
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICMS:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   675
            Width           =   435
         End
         Begin VB.Label lblICMS 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1200
            TabIndex        =   58
            Top             =   600
            Width           =   570
         End
         Begin VB.Label lblImpFed 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1200
            TabIndex        =   57
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Federativo:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   315
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Demonstrativo Indireto: (Somatória dos Itens * Custo de Saída)"
         Height          =   2295
         Left            =   2640
         TabIndex        =   11
         Top             =   120
         Width           =   5895
         Begin VB.Line Line3 
            X1              =   255
            X2              =   5775
            Y1              =   1635
            Y2              =   1635
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Percentual F:"
            Height          =   195
            Left            =   3480
            TabIndex        =   27
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferença:"
            Height          =   195
            Left            =   3480
            TabIndex        =   26
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Percentual F:"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   1875
            Width           =   945
         End
         Begin VB.Line Line2 
            X1              =   240
            X2              =   5760
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferença:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entradas:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   315
            Width           =   675
         End
         Begin VB.Label lblEntradaE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   2340
            TabIndex        =   22
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compras:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   795
            Width           =   660
         End
         Begin VB.Label lblEntradaF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   2340
            TabIndex        =   20
            Top             =   675
            Width           =   570
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Giro:"
            Height          =   195
            Left            =   3480
            TabIndex        =   19
            Top             =   315
            Width           =   330
         End
         Begin VB.Label lblGiroE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   5100
            TabIndex        =   18
            Top             =   240
            Width           =   570
         End
         Begin VB.Label lblGiroF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   5130
            TabIndex        =   17
            Top             =   675
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas:"
            Height          =   195
            Left            =   3480
            TabIndex        =   16
            Top             =   720
            Width           =   585
         End
         Begin VB.Label lblPercEnt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0 %"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   2400
            TabIndex        =   15
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label lblDifEnt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   2340
            TabIndex        =   14
            Top             =   1200
            Width           =   570
         End
         Begin VB.Label lblPerSai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0 %"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   5160
            TabIndex        =   13
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label lblDifSai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   5100
            TabIndex        =   12
            Top             =   1200
            Width           =   570
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2415
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Buscar"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelAnaliseIndFluxo.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "Im&primir"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelAnaliseIndFluxo.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnFechar 
            Cancel          =   -1  'True
            Caption         =   "&Sair"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelAnaliseIndFluxo.frx":0797
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnGraf 
            Caption         =   "&Gráfico Comparativo"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelAnaliseIndFluxo.frx":0902
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1680
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxDAtaFIn 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D&e:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1320
            TabIndex        =   4
            Top             =   120
            Width           =   285
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Demonstrativo Direto: (Somatória dos Valores das Notas)"
         Height          =   2295
         Left            =   2640
         TabIndex        =   28
         Top             =   2520
         Width           =   5895
         Begin VB.Label lblDifSai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   5100
            TabIndex        =   44
            Top             =   1200
            Width           =   570
         End
         Begin VB.Label lblPerSai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0 %"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   5160
            TabIndex        =   43
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label lblDifEnt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   2340
            TabIndex        =   42
            Top             =   1200
            Width           =   570
         End
         Begin VB.Label lblPercEnt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0 %"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   2400
            TabIndex        =   41
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas:"
            Height          =   195
            Left            =   3480
            TabIndex        =   40
            Top             =   720
            Width           =   585
         End
         Begin VB.Label lblGiroF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080FF80&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   5100
            TabIndex        =   39
            Top             =   675
            Width           =   570
         End
         Begin VB.Label lblGiroE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   5100
            TabIndex        =   38
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Giro:"
            Height          =   195
            Left            =   3480
            TabIndex        =   37
            Top             =   315
            Width           =   330
         End
         Begin VB.Label lblEntradaF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080FF80&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   2340
            TabIndex        =   36
            Top             =   675
            Width           =   570
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compras:"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   795
            Width           =   660
         End
         Begin VB.Label lblEntradaE 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   2340
            TabIndex        =   34
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada:"
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferença:"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   1275
            Width           =   735
         End
         Begin VB.Line Line4 
            X1              =   240
            X2              =   5760
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Percentual F:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   1875
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferença:"
            Height          =   195
            Left            =   3480
            TabIndex        =   30
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Percentual F:"
            Height          =   195
            Left            =   3480
            TabIndex        =   29
            Top             =   1920
            Width           =   945
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   5760
            Y1              =   1635
            Y2              =   1635
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Estoque:"
         Height          =   1095
         Left            =   120
         TabIndex        =   45
         Top             =   4800
         Width           =   2415
         Begin VB.Label lblInvent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1080
            TabIndex        =   49
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventário:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   750
         End
         Begin VB.Label lblEstoque 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1080
            TabIndex        =   47
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   315
            Width           =   405
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Receita:"
         Height          =   1095
         Left            =   2640
         TabIndex        =   50
         Top             =   4800
         Width           =   2415
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Real:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblReceitaR 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   720
            TabIndex        =   53
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lblReceitaF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   720
            TabIndex        =   51
            Top             =   600
            Width           =   570
         End
      End
   End
End
Attribute VB_Name = "frmRelAnaliseIndFluxo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub buscaDir()

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select sum(valor_not) as total from tab_Compras where entrada_not>=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and entrada_not<=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and status_not>=3 "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            lblEntradaF(1).caption = FormatVal(!total)
        End If
        .Close
    End With
    
    
    strSQL = "select sum(valor_not) as total from tab_Cmp where entrada_not>=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and entrada_not<=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and status_not>=3 "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblEntradaE(1).caption = FormatVal(!total)
        End If
        .Close
    End With
    
    lblDifEnt(1).caption = FormatVal(CCur(lblEntradaE(1).caption) - CCur(lblEntradaF(1).caption))
    If CCur(lblEntradaE(1).caption) * 100 > 0 Then
        lblPercEnt(1).caption = FormatVal(CCur(lblEntradaF(1).caption) / CCur(lblEntradaE(1).caption) * 100) & "%"
    End If
    strSQL = "select sum(valor_cup) as total from tab_cupons where emissao_sai>=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and emissao_sai<=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and serie_sai <>'CANC' "
    strSQL = strSQL & " and serie_sai <>'CNC' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblGiroF(1).caption = FormatVal(!total)
        End If
        .Close
    End With
    
    
    strSQL = "select sum(tf.valor_ven) as total from tab_vendas as tv, "
    strSQL = strSQL & " tab_venfec as tf where dataefet_ven>=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and tv.situacao_ven=4"
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount > 0 Then
            lblGiroE(1).caption = FormatVal(!total)
        End If
        .Close
    End With
    
    lblDifSai(1).caption = FormatVal(CCur(lblGiroE(1).caption) - CCur(lblGiroF(1).caption))
    If CCur(lblGiroE(1).caption) > 0 Then
    lblPerSai(1).caption = FormatVal(CCur(lblGiroF(1).caption) / CCur(lblGiroE(1).caption) * 100) & "%"
    End If

End Sub
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTime As String
    sTime = Now
    P_Status "Aguarde enquanto eu calculo..."
    strSQL = "select sum(custosai_ite*qtd_ite*qtd_emb) as total from tab_Compras as tc, tab_itcompras as ti "
    strSQL = strSQL & " where entrada_not>=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and entrada_not<=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and ti.seq_not=tc.seq_not "
    strSQL = strSQL & " and tc.status_not>=3"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblEntradaF(0).caption = FormatVal(!total)
        End If
        .Close
    End With
    
    
    strSQL = "select sum(custosai_ite*qtd_ite*qtd_emb) as total from tab_Cmp as tc, tab_itcmp as ti "
    strSQL = strSQL & " where entrada_not>=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and entrada_not<=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and ti.seq_not=tc.seq_not "
    strSQL = strSQL & " and tc.status_not>=3"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblEntradaE(0).caption = FormatVal(!total)
        End If
        .Close
    End With
    
    lblDifEnt(0).caption = FormatVal(CCur(lblEntradaE(0).caption) - CCur(lblEntradaF(0).caption))
    If CCur(lblEntradaE(0).caption) > 0 Then
        lblPercEnt(0).caption = FormatVal(CCur(lblEntradaF(0).caption) / CCur(lblEntradaE(0).caption) * 100) & "%"
    End If
    strSQL = "select if(cr_pro>0, sum(cf_pro*qtd_its),sum(cr_pro*qtd_its))  as total "
    strSQL = strSQL & " from tab_cupons as tc, tab_itcuup as ti, tab_produtos as tp "
    strSQL = strSQL & " where emissao_sai>=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and emissao_sai<=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and tc.cancelado_cup='N' "
    strSQL = strSQL & " and tc.seq_cup=ti.seq_cup "
    
    strSQL = strSQL & " and tc.cod_cai=ti.cod_Cai "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblGiroF(0).caption = FormatVal(!total)
        End If
        .Close
    End With
    
    
    strSQL = "select sum(cr_pro*qtd_ite) as total from tab_vendas as tv, "
    strSQL = strSQL & " tab_venfec as tf,tab_produtos as tp, "
    strSQL = strSQL & " tab_itvendas as ti "
    strSQL = strSQL & " where dataefet_ven>=" & F_GravaData(txtbxDataIni.text)
    strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxDataFin.text)
    strSQL = strSQL & " and tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and tv.situacao_ven=4"
    strSQL = strSQL & " and tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblGiroE(0).caption = FormatVal(!total)
        End If
        .Close
    End With
    
    lblDifSai(0).caption = FormatVal(CCur(lblGiroE(0).caption) - CCur(lblGiroF(0).caption))
    If CCur(lblGiroE(0).caption) > 0 Then
    lblPerSai(0).caption = FormatVal(CCur(lblGiroF(0).caption) / CCur(lblGiroE(0).caption) * 100) & "%"
    End If
    P_Status "Mais um momento por favor..."
    buscaDir
    
    strSQL = "select if (cf_pro>0,sum(cf_pro*ef_pro),sum(cr_pro*ef_pro)) as f,"
    strSQL = strSQL & "sum(cr_pro*er_pro) as r "
    strSQL = strSQL & " from tab_estoque as te, tab_produtos as tp where tp.cod_pro=te.cod_pro "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblEstoque.caption = FormatVal(!r)
            lblInvent.caption = FormatVal(!f)
        End If
        .Close
    End With
    lblReceitaF.caption = FormatVal(CCur(lblGiroF(1).caption) - CCur(lblGiroF(0).caption))
    lblReceitaR.caption = FormatVal(CCur(lblGiroE(1).caption) - CCur(lblGiroE(0).caption))
    lblImpFed.caption = FormatVal(CCur(lblGiroF(1).caption) * crcTxFederativa_Par / 100)
    buscaImp
    P_Status ""
    MsgBox "Busca concluída em " & DateDiff("s", sTime, Now) & " Segundos...", vbInformation
End Sub
Sub buscaImp()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select sum(unit_its*qtd_its*icms_its/100)  as total "
    
    strSQL = strSQL & " from tab_cupons as tc, tab_itcup as ti "
    strSQL = strSQL & " where data_cup>=" & F_GravaDataTempo(txtbxDataIni.text)
    strSQL = strSQL & " and data_cup<=" & F_GravaDataTempo(txtbxDataFin.text & "23:59:59")
    strSQL = strSQL & " and tc.cancelado_cup='N' "
    
    strSQL = strSQL & " and tc.seq_sai=ti.seq_sai "
    strSQL = strSQL & " and tc.cod_cai=ti.cod_Cai "
    
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        lblIcms.caption = FormatVal(rPesq!total)
    End If
    rPesq.Close
End Sub
Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error Resume Next
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    On Error Resume Next
    P_CabRelTimbr Me.caption & " de " & txtbxDataIni.text & " até " & txtbxDataFin.text, obj
    obj.Print ""
    obj.fontsize = 12
    obj.FontBold = True
    P_ImprX obj, "DEMONSTRATIVO INDIRETO", 10, False
    obj.FontBold = False
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    P_ImprX obj, "Entradas:", 15
    P_ImprX obj, lblEntradaE(0).caption, 65, True, True
    P_ImprX obj, "Giro:", 70
    P_ImprX obj, lblGiroE(0).caption, 120, False, True
    P_ImprX obj, "Compras:", 15
    P_ImprX obj, lblEntradaF(0).caption, 65, True, True
    P_ImprX obj, "Vendas:", 70
    P_ImprX obj, lblGiroF(0).caption, 120, False, True
    obj.Print ""
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    P_ImprX obj, "Diferença:", 15
    P_ImprX obj, lblDifEnt(0).caption, 65, True, True
    P_ImprX obj, "Diferença:", 70
    P_ImprX obj, lblDifSai(0).caption, 120, False, True
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    P_ImprX obj, "Percentual F:", 15
    P_ImprX obj, lblPercEnt(0).caption, 65, True, True
    P_ImprX obj, "Percentual F:", 70
    P_ImprX obj, lblPerSai(0).caption, 120, False, True
    obj.Print ""
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    
    obj.FontBold = True
    P_ImprX obj, "DEMONSTRATIVO DIRETO", 10, False
    obj.FontBold = False
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    P_ImprX obj, "Entradas:", 15
    P_ImprX obj, lblEntradaE(1).caption, 65, True, True
    P_ImprX obj, "Giro:", 70
    P_ImprX obj, lblGiroE(1).caption, 120, False, True
    P_ImprX obj, "Compras:", 15
    P_ImprX obj, lblEntradaF(1).caption, 65, True, True
    P_ImprX obj, "Vendas:", 70
    P_ImprX obj, lblGiroF(1).caption, 120, False, True
    obj.Print ""
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    P_ImprX obj, "Diferença:", 15
    P_ImprX obj, lblDifEnt(1).caption, 65, True, True
    P_ImprX obj, "Diferença:", 70
    P_ImprX obj, lblDifSai(1).caption, 120, False, True
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    P_ImprX obj, "Percentual F:", 15
    P_ImprX obj, lblPercEnt(1).caption, 65, True, True
    P_ImprX obj, "Percentual F:", 70
    P_ImprX obj, lblPerSai(1).caption, 120, False, True
    obj.Print ""
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
        
    '
    obj.FontBold = True
    P_ImprX obj, "ESTOQUE", 10, False
    obj.FontBold = False
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    P_ImprX obj, "Total:", 15
    P_ImprX obj, lblEstoque.caption, 65, True, True
    P_ImprX obj, "Inventário:", 70
    P_ImprX obj, lblInvent.caption, 120, False, True
    obj.Print ""
    obj.FontBold = True
    P_ImprX obj, "RECEITA", 10, False
    obj.FontBold = False
    P_Line obj, 5, obj.CurrentY, 180, obj.CurrentY
    P_ImprX obj, "Real:", 15
    P_ImprX obj, lblReceitaR.caption, 65, True, True
    P_ImprX obj, "Total:", 70
    P_ImprX obj, lblReceitaF.caption, 120, False, True
    obj.Print ""
    P_FimRel obj, strNome_Emp
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

