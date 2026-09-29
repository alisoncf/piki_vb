VERSION 5.00
Begin VB.Form frmRecados 
   Caption         =   "Recados"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8730
   Icon            =   "frmRecados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdRecado 
      Caption         =   "Novo Recado"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdExcl 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdLido 
      Caption         =   "Marcar como Lido"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtbxMsg 
      Height          =   2415
      Left            =   120
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1680
      Width           =   8055
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Verificar Mensagem:"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8055
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4575
      End
      Begin VB.OptionButton optnRep 
         Caption         =   "Para um Representante:"
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   2175
      End
      Begin VB.OptionButton optnVoce 
         Caption         =   "Para Você"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnTerceiro 
         Height          =   315
         Left            =   4680
         Picture         =   "frmRecados.frx":992A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   780
   End
End
Attribute VB_Name = "frmRecados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim TipoTer As String

Private Sub cmdExcl_Click()
    On Error GoTo trata_erro
    If MsgBox("Deseja excluir o recado?", vbYesNo + vbQuestion) = vbYes Then
        mensagem.ExcluiMsg dt.Recordset!seq_rec
        dt.Refresh
    End If
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmdRecado_Click()
    Me.tag = "x"
    Me.caption = "Postar uma mensagem"
    cmdSalvar.Enabled = True
    bActive = False
    Form_Activate
     cmdSalvar.Enabled = True
    
End Sub

Private Sub cmdSalvar_Click()
    If Me.tag <> "" Then
        If txtbxTerceiro.tag = "" Then
            MsgBox "Escolha o destino", vbExclamation
            txtbxTerceiro.SetFocus
            Exit Sub
        End If
        
        mensagem.PostaMensagem TipoTer, txtbxTerceiro.tag, txtbxMsg.Text
        MsgBox "Mensagem enviada", vbInformation
        Me.tag = ""
        Form_Activate
        
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    If optnRep.Value Then
    
        strSQL = F_SelectMensagem("R", txtbxTerceiro.tag)
    Else
        strSQL = F_SelectMensagem("U", strCod_Usu)
        
        
        
    End If
    
    dt.RecordSource = strSQL
        dt.Refresh
End Sub
Private Sub cmndbtnTerceiro_Click()
    If optnRep.Value Then
        buscaRep
    Else
        buscausu
    End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxTerceiro.Text)
    If r.cod_Rep <> "" Then
        txtbxTerceiro.tag = r.cod_Rep
        txtbxTerceiro.Text = r.nome_rep
    End If
End Sub
Sub buscausu()
    Dim r As funcionario.funcionario
    r = funcionario.F_SelectFuncionario(txtbxTerceiro.Text)
    If r.cod_fun <> "" Then
        txtbxTerceiro.tag = r.cod_fun
        txtbxTerceiro.Text = r.guerra_fun
    End If
End Sub

Private Sub dt_Reposition()
    On Error Resume Next
    dt.caption = "Recado " & dt.Recordset.AbsolutePosition + 1 & " de " & dt.Recordset.RecordCount
    
    If dt.Recordset.RecordCount > 0 Then
        cmdExcl.Enabled = True
        cmdLido.Enabled = True
    Else
        cmdExcl.Enabled = False
        cmdLido.Enabled = False
    End If
    
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        optnVoce.caption = "Usuário"
        txtbxMsg.Enabled = True
        txtbxMsg.Locked = False
        cmdLido.Enabled = False
        cmdExcl.Enabled = False
        cmdRecado.Enabled = True
        cmdSalvar.Enabled = False
        cmndbtnTerceiro.Enabled = True
        txtbxTerceiro.Enabled = True
        txtbxMsg.Text = ""
        dt.Visible = True
        
    Else
        optnVoce.caption = "Para Você"
        txtbxMsg.Text = ""
        txtbxTerceiro.Enabled = False
        txtbxMsg.Locked = True
        cmdLido.Enabled = True
        cmdLido.Enabled = True
        cmdExcl.Enabled = True
        cmdRecado.Enabled = True
        cmdSalvar.Enabled = False
        dt.Visible = True
        TipoTer = "U"
        optnVoce.Value = True
        cmndbtnTerceiro.Enabled = False
        txtbxTerceiro.Enabled = False
    End If
    
    
    
    
    bActive = True
End Sub

Private Sub Form_Load()
    dt.Connect = Banco.Connect
    bActive = False
End Sub

Private Sub optnRep_Click()
    TipoTer = "R"
    If Me.tag = "" Then
        txtbxTerceiro.Enabled = True
        cmndbtnTerceiro.Enabled = True
    End If
End Sub

Private Sub optnVoce_Click()
    TipoTer = "U"
    If Me.tag = "" Then
        txtbxTerceiro.Enabled = False
        cmndbtnTerceiro.Enabled = False
    End If
End Sub

Private Sub txtbxTerceiro_GotFocus()
    P_SelecTexto txtbxTerceiro
End Sub

Private Sub txtbxTerceiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerceiro_Click
    End If
End Sub
