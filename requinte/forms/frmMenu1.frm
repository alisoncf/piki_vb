VERSION 5.00
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7665
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1335
   ControlBox      =   0   'False
   DrawStyle       =   2  'Dot
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   1335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnProd 
      Caption         =   "&Produtos"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnCli 
      Caption         =   "&Clientes"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   770
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnECF 
      Caption         =   "Controle EC&F"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1540
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnVendas 
      Caption         =   "&Vendas"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3065
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnNovaVenda 
      Caption         =   "&Nova Venda"
      Default         =   -1  'True
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3815
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnCaixa 
      Caption         =   "Cai&xa"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4565
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnUsu 
      Caption         =   "Usuá&rio"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":15E4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5315
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnJan 
      Caption         =   "&Janelas"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":1EAE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6065
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Sair"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":21B8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6820
      Width           =   1260
   End
   Begin VB.CommandButton cmndbtnCompras 
      Caption         =   "&Compras"
      Height          =   735
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmMenu1.frx":24C2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2300
      Width           =   1260
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnCaixa_Click()
    frmSaidas.Show 0
    frmSaidas.SetFocus
End Sub

Private Sub cmndbtnCli_Click()
    
    frmEdiClientes.Show 0
    frmEdiClientes.SetFocus
End Sub

Private Sub cmndbtnCompras_Click()
    If strAbrenota_usu = "S" Then
        frmEdiEntr.Show 0
        frmEdiEntr.SetFocus
    Else
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
    End If
End Sub

Private Sub cmndbtnECF_Click()
    If strFechacai_usu = "S" Then
        On Error Resume Next
        frmImpressoraECF.Show 1
    Else
        MsgBox "Você não tem permissão para manipular a ECF!", vbExclamation
    End If
End Sub

Private Sub cmndbtnJan_Click()
    Me.PopupMenu mdiFrm.mnJan
End Sub

Private Sub cmndbtnNovaVenda_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = F_GeraNovaVenda("00000", strCod_Usu)
    
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.Tag = strSQL
    On Error Resume Next
    frmVenda.Show 1

End Sub

Private Sub cmndbtnProd_Click()
    frmEdiPro.Show 0
    frmEdiPro.SetFocus
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub cmndbtnSair_Click()
    If MsgBox("Deseja realmente sair?", vbYesNo + vbQuestion) = vbYes Then
        P_Fim
    Else
        Exit Sub
    End If
End Sub

Private Sub cmndbtnUsu_Click()
    frmOptUsu.Show 0, mdiFrm
    
End Sub

Private Sub cmndbtnVendas_Click()
     
    If strAbrevenda_usu = "S" Then
        frmEdiVendas.Show 0
        frmEdiVendas.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
        mdiFrm.SetFocus
        SendKeys vbAltMask
    End If
    If KeyCode = vbKeyF3 Then
        Me.Hide
        mdiFrm.SetFocus
        
    End If

End Sub

