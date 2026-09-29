VERSION 5.00
Begin VB.Form frmCadOrgaos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Orgaos"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   1440
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   3480
      Width           =   5175
   End
   Begin VB.TextBox txtAtlz 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   14
      TabIndex        =   18
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtCobrador 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      ToolTipText     =   "Tecle enter para selecionar um bairro"
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Esc - &Fecha&r"
      Height          =   855
      Left            =   4680
      Picture         =   "frmCadOrgaos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "F7 - Sal&var"
      Height          =   855
      Left            =   2520
      Picture         =   "frmCadOrgaos.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox txtbxFone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      MaxLength       =   14
      TabIndex        =   16
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtbxNomeGru 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Tecle enter para selecionar um bairro"
      Top             =   240
      Width           =   4935
   End
   Begin VB.TextBox txtbxNome 
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
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   5175
   End
   Begin VB.TextBox txtbxEndereco 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox txtbxCidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   11
      Text            =   "Anápolis"
      Top             =   2280
      Width           =   4575
   End
   Begin VB.TextBox txtbxCEP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   14
      TabIndex        =   14
      Text            =   "75.000-000"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtbxBairro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Tecle enter para selecionar um bairro"
      Top             =   1320
      Width           =   4935
   End
   Begin VB.CommandButton cmndbtnBuscaGru 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      Picture         =   "frmCadOrgaos.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmndbtnCobrador 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      Picture         =   "frmCadOrgaos.frx":085E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdCob 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      Picture         =   "frmCadOrgaos.frx":0BE8
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Obs"
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   360
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Atualização"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cobrador"
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
      Left            =   360
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefone:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4200
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grupo:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Nome:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Endereço:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblBairro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Set&or:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ci&dade:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblUF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GO"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6000
      TabIndex        =   12
      Top             =   2280
      Width           =   600
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CE&P:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmCadOrgaos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCob_Click()
    buscaCobrador
    
    
End Sub
Sub buscaCobrador()
    Dim c As ContaCob
    c = F_BuscaContaCob(txtCobrador.Text)
    If c.cod_cob <> "" Then
        
        txtCobrador.Text = c.nome_cob
        txtCobrador.tag = c.cod_cob
    End If
    
End Sub
Sub buscaCobrador1()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectCobrador("")
    
    If f.cod_fun <> "" Then
        txtCobrador.tag = f.cod_fun
        txtCobrador.Text = f.nome_fun
        
    End If
    
    
End Sub

Private Sub cmndbtnBuscaGru_Click()
    Dim g As grupoCli
    g = F_BuscaGrupoCli
    
    If g.cod_gru <> "" Then
        txtbxNomeGru.Text = g.nome_gru
        txtbxNomeGru.tag = g.cod_gru
        txtbxNome.SetFocus
        
    End If
    
    
    
End Sub

Private Sub cmndbtnCobrador_Click()
    Dim s As setor
    
    
    s = F_BuscaSetor(txtbxBairro.Text)
    If s.cod <> "" Then
        txtbxBairro.Text = s.nome
        txtbxBairro.tag = s.cod
        txtbxCidade.Text = s.Cidade
        lblUF.caption = s.UF
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.tag = "" Then
        If F_GravaOrgao(txtbxNomeGru.tag, txtbxNome.Text, txtbxBairro.tag, txtbxEndereco.Text, txtbxFone.Text, txtbxCEP.Text, txtAtlz.Text, txtObs.Text) <> "" Then
            MsgBox "Cadastrado com sucesso!", vbInformation
            txtbxNome.Text = ""
            txtbxEndereco.Text = ""
            txtbxNome.SetFocus
        End If
    Else
        orgao.P_AlteraOrgao Me.tag, txtbxNomeGru.tag, txtbxNome.Text, txtbxBairro.tag, txtbxEndereco.Text, txtbxCEP.Text, txtbxFone.Text, txtCobrador.tag, txtObs.Text, txtAtlz.Text
        
        MsgBox "Altedado!", vbInformation
        
    End If
End Sub

Private Sub txtAtlz_GotFocus()
    P_SelecTexto txtAtlz
End Sub

Private Sub txtAtlz_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAtlz_LostFocus()
    P_AjustaData txtAtlz
End Sub

Private Sub txtbxFone_GotFocus()
    P_SelecTexto txtbxFone
End Sub

Private Sub txtbxFone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxFone_LostFocus()
    P_AjustaTexto txtbxFone
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub
Private Sub txtbxEndereco_GotFocus()
    P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEndereco_LostFocus()
    P_AjustaTexto txtbxEndereco
End Sub
Private Sub txtbxNomeGru_GotFocus()
    P_SelecTexto txtbxNomeGru
End Sub

Private Sub txtbxNomeGru_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaGru_Click
    End If
End Sub

Private Sub txtCobrador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCobrador
    End If
    
End Sub
