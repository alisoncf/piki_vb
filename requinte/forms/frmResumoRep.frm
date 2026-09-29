VERSION 5.00
Begin VB.Form frmResumoRep 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumo de Atividades de Representante"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxNomeCli 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   2880
      TabIndex        =   29
      Top             =   240
      Width           =   4575
   End
   Begin VB.CommandButton cmndbtnCliente 
      Height          =   285
      Left            =   7440
      Picture         =   "frmResumoRep.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtbxAte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   27
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   240
      ScaleHeight     =   5145
      ScaleWidth      =   4065
      TabIndex        =   6
      Top             =   840
      Width           =   4095
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comissões Geradas"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   33
         Top             =   4200
         Width           =   2475
      End
      Begin VB.Label Label12 
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
         Left            =   2580
         TabIndex        =   32
         Top             =   4200
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acertos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   960
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transferências:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   24
         Top             =   3240
         Width           =   2475
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Devoluções:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   2475
      End
      Begin VB.Label lblConsign 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   2580
         TabIndex        =   22
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label lblDevolucoes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   2580
         TabIndex        =   21
         Top             =   3720
         Width           =   1155
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cartão de Débito(CRD):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   2475
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
         Left            =   2580
         TabIndex        =   19
         Top             =   2160
         Width           =   1155
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
         Left            =   2580
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   720
         Width           =   2475
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheque Pré-Datado(CHP):"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   2475
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
         Left            =   2580
         TabIndex        =   15
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
         Left            =   2580
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   1800
         Width           =   2475
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
         Left            =   2580
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1440
         Width           =   2475
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
         Left            =   2580
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   360
         Width           =   2475
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
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2580
         TabIndex        =   8
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   2475
      End
   End
   Begin VB.TextBox txtbxDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnExpor 
      Caption         =   "&Detalhar"
      Height          =   685
      Left            =   2040
      Picture         =   "frmResumoRep.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnGraf 
      Caption         =   "&Gráfico"
      Height          =   685
      Left            =   5880
      Picture         =   "frmResumoRep.frx":04E6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   685
      Left            =   7800
      Picture         =   "frmResumoRep.frx":0630
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   685
      Left            =   3960
      Picture         =   "frmResumoRep.frx":079B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   685
      Left            =   120
      Picture         =   "frmResumoRep.frx":0BA8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Representante:"
      Height          =   195
      Left            =   2880
      TabIndex        =   31
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Até:"
      Height          =   195
      Left            =   1560
      TabIndex        =   28
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "De:"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmResumoRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnCliente_Click()
    buscaRep
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxNomeCli.Text)
    If r.cod_rep <> "" Then
        txtbxNomeCli.tag = r.cod_rep
        txtbxNomeCli.Text = r.nome_rep
    End If
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    
    
    
End Sub

Private Sub cmndbtnExpor_Click()
    Dim strSQL As String
    P_RelResumoAtv txtbxNomeCli.tag, txtbxDe.Text, txtbxAte.Text
    
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub
Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

