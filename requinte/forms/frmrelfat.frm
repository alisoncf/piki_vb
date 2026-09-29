VERSION 5.00
Begin VB.Form frmRelFat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumo de Faturamento Total (Saidas)"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12135
   Icon            =   "frmrelfat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   5160
      ScaleHeight     =   2745
      ScaleWidth      =   6825
      TabIndex        =   68
      Top             =   5280
      Width           =   6855
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   14415
      TabIndex        =   67
      Top             =   0
      Width           =   14445
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   570
         Left            =   2520
         Picture         =   "frmrelfat.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   570
         Left            =   4320
         Picture         =   "frmrelfat.frx":053B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   75
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   570
         Left            =   10320
         Picture         =   "frmrelfat.frx":0948
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   75
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico Comparativo"
         Height          =   570
         Left            =   6120
         Picture         =   "frmrelfat.frx":0AB3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   75
         Width           =   1695
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D&e:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox pcTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   8640
      ScaleHeight     =   2985
      ScaleWidth      =   3345
      TabIndex        =   51
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblDupTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   66
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   65
         Top             =   2520
         Width           =   1995
      End
      Begin VB.Label lblTotalTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   64
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dinheiro(DIN):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lblDinTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   62
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Duplicata(DUP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   61
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Crédito(CRT):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lblCrtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   59
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblChvTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   58
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheque Pré-Datado(CHP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   57
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheques a Vista (CHV):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label lblChpTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   55
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblCrdTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   54
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Débito(CRD):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   53
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   120
         TabIndex        =   52
         Top             =   0
         Width           =   810
      End
   End
   Begin VB.PictureBox pcAcerto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5160
      ScaleHeight     =   2985
      ScaleWidth      =   3345
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Débito(CRD):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lblAcertoCrd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   38
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acertos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   120
         TabIndex        =   35
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label lblAcertoCHP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   34
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheques a Vista (CHV):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheque Pré-Datado(CHP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label lblAcertoCHV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   31
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblAcertoCrt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   30
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Crédito(CRT):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lblAcertoDup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   28
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Duplicata(DUP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblAcertoDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   26
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dinheiro(DIN):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lblTotalAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   24
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1995
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7185
      ScaleWidth      =   4665
      TabIndex        =   8
      Top             =   840
      Width           =   4695
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2025
         ScaleWidth      =   4425
         TabIndex        =   40
         Top             =   4440
         Width           =   4455
         Begin VB.Line Line5 
            X1              =   0
            X2              =   3360
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label lblMenorDia 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Menor Dia"
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
            Left            =   120
            TabIndex        =   50
            Top             =   840
            Width           =   885
         End
         Begin VB.Label lblDiaMaior 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maior Dia"
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
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Menor Movimentação:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   600
            Width           =   1590
         End
         Begin VB.Label lblValorMenor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1830
            TabIndex        =   47
            Top             =   600
            Width           =   390
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maior movimentação:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label lblVlrMaior 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1830
            TabIndex        =   45
            Top             =   120
            Width           =   390
         End
         Begin VB.Label lblPrev 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1680
            TabIndex        =   44
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previsão Mensal:"
            Height          =   195
            Left            =   1680
            TabIndex        =   43
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Média Diária:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   1320
            Width           =   930
         End
         Begin VB.Label lblMedDia 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   120
            TabIndex        =   41
            Top             =   1560
            Width           =   525
         End
      End
      Begin VB.Label lblTotalQueima 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   85
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total de Vendas Queima"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   84
         Top             =   3240
         Width           =   1995
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desconto de Queima"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   83
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label lblDescQueima 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   82
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3480
         TabIndex        =   81
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblDescTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   80
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label lblDescCrd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   79
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label lblDescCrt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   78
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblDescDup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   77
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label lblDescChp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   76
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblDescCHV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   75
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblDescDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3360
         TabIndex        =   74
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recebido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2160
         TabIndex        =   73
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Trabalhados:"
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   6600
         Width           =   1290
      End
      Begin VB.Label lblTrab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1800
         TabIndex        =   71
         Top             =   6600
         Width           =   120
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Restantes:"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   6840
         Width           =   1125
      End
      Begin VB.Label lblRest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1800
         TabIndex        =   69
         Top             =   6840
         Width           =   120
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Débito(CRD):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lblCrd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   36
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   1995
      End
      Begin VB.Label lblTotalVendas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   20
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dinheiro(DIN):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lblDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   18
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Duplicata(DUP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblDup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   16
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Crédito(CRT):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lblCrt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   14
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblChv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   13
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheque Pré-Datado(CHP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheques a Vista (CHV):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label lblCHP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   10
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmRelFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub buscaMov()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim sOr As String
    
    strSQL = "select  sum(tv.valor_ven) as total ,"
    strSQL = strSQL & " dataefet_ven from tab_Vendas as tv"
    strSQL = strSQL & " where tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and tv.situacao_ven=4  "
    strSQL = strSQL & "group by dataefet_ven"
    
    sOrd = " order by total desc limit 1 "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        lblDiaMaior.caption = !dataefet_ven & " " & WeekdayName(Weekday(!dataefet_ven))
        lblVlrMaior.caption = FormatVal(!total)
        .Close
    End With
    sOrd = " order by total asc limit 1 "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        lblMenorDia.caption = !dataefet_ven & " " & WeekdayName(Weekday(!dataefet_ven))
        lblValorMenor.caption = FormatVal(!total)
        .Close
    End With
End Sub
Sub buscaDias()
    Dim sAux As String
    Dim sMes As String
    Dim strSQL As String
    Dim sDia As String
    strSQL = F_BuscaUmCampo("count(distinct dataefet_ven)", "tab_vendas", _
        "dataefet_ven>=" & F_GravaData(txtbxDataIni.Text) & " and " & _
        "dataefet_ven<=" & F_GravaData(txtbxdataFin.Text))
    lblTrab.caption = strSQL
    sAux = Month(txtbxdataFin.Text)
    sAux = "01/" & sAux & "/" & Year(txtbxdataFin.Text)
    sAux = DateAdd("M", 1, sAux)
    sAux = DateAdd("D", -1, sAux)
    sAux = DateDiff("D", txtbxdataFin.Text, sAux)
    lblRest.caption = sAux
    sDia = Weekday(txtbxdataFin.Text)
    If sDia = 0 Then
        sAux = Int(sAux) - 1
    End If
    sAux = Int(sAux - (Int(sAux) / 7))
    lblRest.caption = sAux
    
    
    
    
End Sub
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim cTotal As Currency
    Dim cTotalDesc As Currency
    On Error GoTo trata_erro
    buscaDias
    'lblVendaFin.caption = "0"
    strSQL = " select sum(tv.valor_ven) as valor,sum(tv.desconto_ven) as desconto,tp.tipo_cai "
    strSQL = strSQL & " from tab_vendas as tv, tab_formaspag as tp  "
    strSQL = strSQL & " where tv.cod_for=tp.cod_for "
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and tv.situacao_ven=4  "
    strSQL = strSQL & " group by tp.tipo_cai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If !tipo_cai = "DIN" Then
                lblDin.caption = FormatVal(!valor)
                lblDescDin.caption = FormatVal(!desconto)
            ElseIf !tipo_cai = "DUP" Then
                lblDup.caption = FormatVal(!valor)
                lblDescDup.caption = FormatVal(!desconto)
            ElseIf !tipo_cai = "CRT" Then
                lblCrt.caption = FormatVal(!valor)
                lblDescCrt.caption = FormatVal(!desconto)
            ElseIf !tipo_cai = "CRD" Then
                lblCrd.caption = FormatVal(!valor)
                lblDescCrd.caption = FormatVal(!desconto)
            ElseIf !tipo_cai = "CHP" Then
                lblCHP.caption = FormatVal(!valor)
                lblDescChp.caption = FormatVal(!desconto)
            ElseIf !tipo_cai = "CHV" Then
                lblChv.caption = FormatVal(!valor)
                lblDescCHV.caption = FormatVal(!desconto)
            
            Else
                MsgBox ""
            End If
            
            cTotal = cTotal + !valor
            cTotalDesc = cTotalDesc + !desconto
            .MoveNext
        Loop
        
        .Close
    End With
    lblTotalVendas.caption = FormatVal(cTotal)
    lblDescTotal.caption = FormatVal(cTotalDesc)
    
    strSQL = "select sum(desconto_ite) as desconto,sum(final_ite) as final from tab_itvendas as ti, tab_vendas as tv "
    strSQL = strSQL & " where tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and tv.situacao_ven=4  "
    strSQL = strSQL & " and pq_itd>0 "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            lblDescQueima.caption = FormatVal(!desconto)
            lblTotalQueima.caption = FormatVal(!final)
        Else
            lblDescQueima.caption = "0,00"
            lblTotalQueima.caption = "0,00"
        End If
        
    End With
    
    cTotal = 0
    strSQL = " select sum(ta.valor_ace) as valor,tp.tipo_cai "
    strSQL = strSQL & " from tab_acertos as ta, tab_formaspag as tp  "
    strSQL = strSQL & " where ta.cod_for=tp.cod_for "
    strSQL = strSQL & " and data_Ace>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and data_ace<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and status_ace=2"
    strSQL = strSQL & " group by tp.tipo_cai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If !tipo_cai = "DIN" Then
                lblAcertoDin.caption = FormatVal(!valor)
                
            ElseIf !tipo_cai = "DUP" Then
                lblAcertoDup.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "CRT" Then
                lblAcertoCrt.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "CRD" Then
                lblAcertoCrd.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "CHP" Then
                lblAcertoCHP.caption = FormatVal(!valor)
            ElseIf !tipo_cai = "CHV" Then
                lblAcertoCHV.caption = FormatVal(!valor)
            
            Else
                MsgBox ""
            End If
            
            cTotal = cTotal + !valor
            .MoveNext
        Loop
        .Close
    End With
    lblTotalAcerto.caption = FormatVal(cTotal)
    
    lblDinTotal.caption = FormatVal(CCur(lblDin.caption) + CCur(lblAcertoDin.caption))
    lblChpTotal.caption = FormatVal(CCur(lblCHP.caption) + CCur(lblAcertoCHP.caption))
    lblChvTotal.caption = FormatVal(CCur(lblChv.caption) + CCur(lblAcertoCHV.caption))
    lblChvTotal.caption = FormatVal(CCur(lblChv.caption) + CCur(lblAcertoCHV.caption))
    lblDupTotal.caption = FormatVal(CCur(lblDup.caption) + CCur(lblAcertoDup.caption))
    lblCrtTotal.caption = FormatVal(CCur(lblCrt.caption) + CCur(lblAcertoCrt.caption))
    lblCrdTotal.caption = FormatVal(CCur(lblCrd.caption) + CCur(lblAcertoCrd.caption))
    lblTotalTotal.caption = FormatVal(CCur(lblTotalVendas.caption) + CCur(lblTotalAcerto.caption))
    buscaMov
'    strSQL = "select sum(acres_tro-abat_tro) as total from tab_trocas "
'    strSQL = strSQL & " where left(data_tro,10)>=" & F_GravaData(txtbxDataIni.Text)
'    strSQL = strSQL & " and left(data_tro,10)<=" & F_GravaData(txtbxDAtaFIn.Text)
'    abreobjetoado rpesq,strsql
'    With rPesq
'            lblTroca.caption = FormatVal(!total)
'            If lblTroca.caption = "" Then lblTroca.caption = "0,00"
'            lblTotal.caption = FormatVal(CCur(lblTotal.caption) + CCur(lblTroca.caption))
'        .Close
'    End With
    strSQL = "select sum(gerado_dev) as total from tab_devolucoes "
    strSQL = strSQL & " where left(data_dev,10)>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and left(data_Dev,10)<=" & F_GravaData(txtbxdataFin.Text)
    abreObjetoAdo rPesq, strSQL
    With rPesq
            'lblDev.caption = FormatVal(!total)
'            If lblDev.caption = "" Then lblDev.caption = "0,00"
 '           lblTotal.caption = FormatVal(CCur(lblTotal.caption) - CCur(lblDev.caption))
        .Close
    End With
    lblMedDia.caption = FormatVal(CCur(lblTotalVendas.caption / CCur(lblTrab.caption)))
    lblPrev.caption = FormatVal(CCur(lblMedDia.caption) * (CCur(lblRest.caption) + CCur(lblTrab.caption)))
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnGraf_Click()
    Dim b(4) As valChart
    Exit Sub
    
    b(0).caption = "À Vista " & lblVendaAvista.caption
    b(0).Value = CCur(lblVendaAvista.caption)
    b(1).caption = "Cheques-Pré " & lblTotalCHP.caption
    b(1).Value = CCur(lblTotalCHP.caption)
    b(2).caption = "Duplicatas " & lblTotalDup.caption
    b(2).Value = CCur(lblTotalDup.caption)
    b(3).caption = "Cartões " & lblVendaFin.caption
    b(3).Value = CCur(lblVendaFin.caption)
    P_Graph b, 4, 1, True, "Gráfico de Análise de Vendas", "Período entre: " & txtbxDataIni.Text & " e: " & txtbxdataFin.Text
End Sub

Private Sub Command1_Click()

End Sub



Private Sub cmndbtnImprimir_Click()
    
    imprHTML
    
End Sub
Sub imprHTML()
    Dim i As Integer
    Dim sNomeArq As String
    
    
    sNomeArq = F_AbreHtml("rel_fat")
    
    P_HtmlAbreHead 1
    P_HtmlAbreTitle 1, "Relatório de Faturamento", True
    P_HtmlAbreTitle 1, "De: " & txtbxDataIni.Text & " Até " & txtbxdataFin.Text, True
    
    
    P_HtmlFechaHead 1
    P_HtmlAbreBody
    
    P_HtmlAbreH 1, 3, "Relatório de Faturamento por Vendas", True
    
    
    P_HtmlAbreTable 1, 1
        P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "Forma de Pagamento:"
            P_HtmlAbreColuna 1, "Vendas:"
            P_HtmlAbreColuna 1, "Acertos"
            P_HtmlAbreColuna 1, "Total"
        P_HtmlFechaLinha
        P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "Dinheiro(DIN):"
            P_HtmlAbreColuna 1, lblDin.caption, "right"
            P_HtmlAbreColuna 1, lblAcertoDin.caption, "right"
            P_HtmlAbreColuna 1, lblDinTotal.caption, "right"
            
        P_HtmlFechaLinha
        P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "Cheques À Vista(CHV):"
            P_HtmlAbreColuna 1, lblChv.caption, "right"
            P_HtmlAbreColuna 1, lblAcertoCHV.caption, "right"
            P_HtmlAbreColuna 1, lblChvTotal.caption, "right"
        P_HtmlFechaLinha
        P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "Cheque-Pré(CHP):"
            P_HtmlAbreColuna 1, lblCHP.caption, "right"
            P_HtmlAbreColuna 1, lblAcertoCHP.caption, "right"
            P_HtmlAbreColuna 1, lblChpTotal.caption, "right"
        P_HtmlFechaLinha
        P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "Duplicata(DUP):"
            P_HtmlAbreColuna 1, lblDup.caption, "right"
            P_HtmlAbreColuna 1, lblAcertoDup.caption, "right"
            P_HtmlAbreColuna 1, lblDupTotal.caption, "right"
        P_HtmlFechaLinha
        P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "Cartão de Crédito(CRT):"
            P_HtmlAbreColuna 1, lblCrt.caption, "right"
            P_HtmlAbreColuna 1, lblAcertoCrt.caption, "right"
            P_HtmlAbreColuna 1, lblCrtTotal.caption, "right"
        P_HtmlFechaLinha
        P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "Cartão de Débito(CRD):"
            P_HtmlAbreColuna 1, lblCrd.caption, "right"
            P_HtmlAbreColuna 1, lblAcertoCrd.caption, "right"
            P_HtmlAbreColuna 1, lblCrdTotal.caption, "right"
        P_HtmlFechaLinha
    
        P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "Total:"
            P_HtmlAbreColuna 1, lblTotalVendas.caption, "right"
            P_HtmlAbreColuna 1, lblTotalAcerto.caption, "right"
            P_HtmlAbreColuna 1, lblTotalTotal.caption, "right"
        P_HtmlFechaLinha
    
    P_HtmlFechaTable
    
    P_HtmlAbreP Now, 1
    P_HtmlFechaBody
    P_FechaHtml 1, True, sNomeArq
    
    
    
    
    
    
    
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Resumo de Faturamento de " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, obj
    P_ImprLinha obj
    obj.fontsize = 14
    
    Exit Sub
    P_ImprX obj, "Vendas...", 10
    P_ImprX obj, "Venda à Vista............................: ", 10
    P_ImprX obj, lblVendaAvista.caption, 160, False, 3
    P_ImprX obj, "Venda a Prazo...........................: ", 10
    P_ImprX obj, lblVendaAp.caption, 160, False, 3
    P_ImprX obj, "Venda por Financeiras..............: ", 10
    P_ImprX obj, lblVendaFin.caption, 160, False, 3
    P_ImprLinha obj, "_"
    P_ImprX obj, "Total............................................: ", 10
    P_ImprX obj, lblTotal.caption, 160, False, 3
    obj.fontsize = 10
    P_RodapeVis obj, constNome_sis & " " & constVersao_Sis
End Sub
Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        pcAcerto.Visible = pcAcerto.Visible = False
        pcTotal.Visible = pcTotal.Visible = False
    End If
End Sub

Private Sub Form_Load()
    txtbxdataFin.Text = F_UltDiaMes(Date)
    txtbxDataIni.Text = F_PrimDiaMes(Date)
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
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
