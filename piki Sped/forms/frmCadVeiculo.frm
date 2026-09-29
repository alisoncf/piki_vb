VERSION 5.00
Begin VB.Form frmCadVeiculo 
   Caption         =   "Veículos"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6900
   Icon            =   "frmCadVeiculo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFabricante 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   5655
   End
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
      ScaleWidth      =   6900
      TabIndex        =   7
      Top             =   2595
      Width           =   6900
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   5160
         Picture         =   "frmCadVeiculo.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   3600
         Picture         =   "frmCadVeiculo.frx":10DD8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   285
      Left            =   5880
      Picture         =   "frmCadVeiculo.frx":1125F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricante"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   750
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
Attribute VB_Name = "frmCadVeiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    Dim f As Veiculo_Fabricante
    f = Veiculos.BuscaVeiculo_Fabricante(txtFabricante.Text)
    If f.id > 0 Then
        txtFabricante.Text = f.descricao
        txtFabricante.tag = f.id
    End If
End Sub

Private Sub cmdNovo_Click()
    If txtDescricao.Text = "" Then
        MsgBox "Informe o nome do veículo!", vbExclamation
        txtDescricao.SetFocus
        Exit Sub
    End If
    If txtFabricante.tag = "" Then
        MsgBox "Escolha o fabricante!", vbExclamation
        txtFabricante.SetFocus
        Exit Sub
    End If
    
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    If F_BuscaUmCampo("id", "veiculo", "descricao='" & txtDescricao.Text & "' and id_fab='" & txtFabricante.tag & "'") <> "" Then
        MsgBox "Já existe este veículo!", vbExclamation
        Exit Sub
    End If
    If Veiculos.GravaVeiculo(txtFabricante.tag, txtDescricao.Text) Then
        MsgBox "Registro gravado com sucesso!", vbInformation
        txtDescricao.Text = ""
    End If
End Sub
Sub altera()
    If F_BuscaUmCampo("id", "veiculo", "descricao='" & txtDescricao.Text & "' and id<>'" & Me.tag & "' and id_fab='" & txtFabricante.tag & "'") <> "" Then
        MsgBox "Já existe este veiculo!", vbExclamation
        Exit Sub
    End If
    If Veiculos.AlteraVeiculo(Me.tag, txtFabricante.tag, txtDescricao.Text) Then
        MsgBox "Registro gravado com sucesso!", vbInformation
        
    End If
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    txtCodigo.Text = MaxKeySQL("veiculo", "id", "", "")
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

Private Sub txtFabricante_Change()
    txtFabricante.tag = ""
End Sub

Private Sub txtFabricante_GotFocus()
    P_SelecTexto txtFabricante
End Sub

Private Sub txtFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdBuscar_Click
    End If
End Sub
