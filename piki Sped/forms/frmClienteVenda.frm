VERSION 5.00
Begin VB.Form frmClienteVenda 
   Caption         =   "Informe os Dados do Cliente"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10080
   Icon            =   "frmClienteVenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3780
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   9735
   End
   Begin VB.TextBox txtCPF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00916024&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   10080
      TabIndex        =   4
      Top             =   2805
      Width           =   10080
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (Esc)"
         Height          =   735
         Left            =   8400
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar (F7)"
         Height          =   735
         Left            =   6720
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF/CNPJ:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "frmClienteVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bOk As Boolean
Private Sub cmdCancelar_Click()
    Me.tag = ""
    Me.Hide
    
End Sub

Private Sub cmdConfirmar_Click()
    If txtCPF.Text = "" And txtNome.Text = "" Then
        Me.Hide
        Exit Sub
    End If
    If txtCPF.Text <> "" And Not bOk Then
        MsgBox "CPF/CNPJ Inválido!", vbExclamation
        Exit Sub
    End If
    
    Me.tag = "x"
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmdConfirmar_Click
    End If
End Sub

Private Sub txtCPF_GotFocus()
    P_SelecTexto txtCPF
    
End Sub

Private Sub txtCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCli
        Sendkeys Chr(9)
    End If
End Sub
Sub buscaCli()
    Dim c As cliente
    c = F_BuscaClienteCPFCNPJ(txtCPF.Text)
    
    
End Sub
Private Sub txtCPF_LostFocus()
    
    If F_CPF(txtCPF.Text) Then
        lblNome.caption = "Nome"
        bOk = True
    ElseIf F_CGC(txtCPF.Text) Then
        lblNome.caption = "Razão Social"
        bOk = True
    ElseIf txtCPF.Text = "" Or txtCPF.Text = "0" Then
        bOk = True
    Else
        bOk = False
        MsgBox "CPF ou CNPJ inválido!", vbExclamation
        
    
    End If
    
    
End Sub

Private Sub txtNome_Change()
    txtNome.tag = ""
End Sub

Private Sub txtNome_GotFocus()
    P_SelecTexto txtNome
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        P_AjustaTexto txtNome, 3
        cmdConfirmar_Click
    ElseIf KeyCode = vbKeyF2 Then
        buscaCliente
    End If
    
    
End Sub
Sub buscaCliente()
    Dim c As cliente
    
    c = F_BuscaCliente(txtNome.Text)
    If c.cod_cli <> "" Then
        txtNome.Text = c.nome_Cli
        txtNome.tag = c.cod_cli
        If c.tipo_cli = "F" Then
            txtCPF.Text = c.cpf_cli
            
        Else
            txtCPF.Text = c.cnpj_Cli
        End If
        txtCPF.tag = c.cod_cli
        txtNome.BackColor = vbGreen
    Else
        txtCPF.tag = ""
        txtNome.BackColor = vbWhite
    End If
    
End Sub
Private Sub txtNome_LostFocus()
    P_AjustaTexto txtNome, caso.Maiusculo
End Sub
