VERSION 5.00
Begin VB.Form frmCadConv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Convênios"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnConveniados 
      Caption         =   "Conveniados"
      Height          =   735
      Left            =   120
      Picture         =   "frmCadConv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ComboBox cmbxStatus 
      Height          =   315
      ItemData        =   "frmCadConv.frx":08CA
      Left            =   1680
      List            =   "frmCadConv.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtbxDesc 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Sal&var"
      Height          =   735
      Left            =   3240
      Picture         =   "frmCadConv.frx":08F0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fecha&r"
      Height          =   735
      Left            =   5160
      Picture         =   "frmCadConv.frx":0A3A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Height          =   255
      Left            =   6600
      Picture         =   "frmCadConv.frx":0B84
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtbxCliente 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Left            =   870
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto:"
      Height          =   195
      Left            =   630
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   525
   End
End
Attribute VB_Name = "frmCadConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim c As Cliente
    c = F_BuscaCliente(txtbxCliente.Text, False)
    
    If c.cod_cli <> "" Then
        txtbxCliente.Text = c.nome_Cli
        txtbxCliente.tag = c.cod_cli
        
    End If
    
    
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    
    If txtbxCliente.tag = "" Then
        MsgBox "Escolha o cliente!", vbExclamation
        txtbxCliente.SetFocus
        Exit Sub
    End If
    
    strSQL = F_senhaReturnUsu(False)
    
    If strSQL = "" Then Exit Sub
    
    
    'ver permissões do usuário
    
    
    If Me.tag = "" Then
        strSQL = convenio.F_GravaConv(txtbxCliente.tag, txtbxDesc)
        If strSQL <> "" Then
            MsgBox "Convênio Gravado com Sucesso!", vbInformation
        End If
    Else
        If convenio.F_AlteraConv(Me.tag, txtbxDesc.Text, cmbxStatus.Text) Then
            MsgBox "Convênio alterado com sucesso!", vbInformation
            Unload Me
        End If
    
    End If
End Sub

Private Sub txtbxCliente_GotFocus()
    P_SelecTexto txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCliente_LostFocus()
    P_AjustaTexto txtbxCliente
End Sub

Private Sub txtbxDesc_GotFocus()
    P_SelecTexto txtbxDesc
End Sub

Private Sub txtbxDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDesc_LostFocus()
    P_ValorLostFocus txtbxDesc
End Sub
