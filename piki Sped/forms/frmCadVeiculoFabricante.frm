VERSION 5.00
Begin VB.Form frmCadVeiculoFabricante 
   Caption         =   "Fabricante de Veículo"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6750
   Icon            =   "frmCadVeiculoFabricante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6750
      TabIndex        =   4
      Top             =   1635
      Width           =   6750
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   4920
         Picture         =   "frmCadVeiculoFabricante.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   3360
         Picture         =   "frmCadVeiculoFabricante.frx":10DD8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "frmCadVeiculoFabricante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNovo_Click()
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    If F_BuscaUmCampo("id", "veiculo_fabricante", "descricao='" & txtDescricao.Text & "'") <> "" Then
        MsgBox "Já existe este fabricante!", vbExclamation
        Exit Sub
    End If
    If Veiculos.GravaVeiculo_Fabricante(txtDescricao.Text) Then
        MsgBox "Registro gravado com sucesso!", vbInformation
        txtDescricao.Text = ""
    End If
End Sub
Sub altera()
    If F_BuscaUmCampo("id", "veiculo_fabricante", "descricao='" & txtDescricao.Text & "' and id<>'" & Me.tag & "'") <> "" Then
        MsgBox "Já existe este fabricante!", vbExclamation
        Exit Sub
    End If
    If Veiculos.AlteraVeiculo_Fabricante(Me.tag, txtDescricao.Text) Then
        MsgBox "Registro gravado com sucesso!", vbInformation
        
    End If
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtDescricao_GotFocus()
    P_SelecTexto txtDescricao
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDescricao_LostFocus()
    P_AjustaTexto txtDescricao
End Sub
