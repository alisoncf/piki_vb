VERSION 5.00
Begin VB.Form frmMesclarOrgao 
   Caption         =   "Mesclar órgãos"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   735
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdConf 
      Caption         =   "Confirmar"
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtFica 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4815
   End
   Begin VB.TextBox txtSai 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label lblCodNovo 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   5280
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCodAnt 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   5280
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Órgão que tomará seu lugar"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Órgão que irá desaparecer:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmMesclarOrgao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConf_Click()
    Dim strSQL As String
    Dim i
    If MsgBox("Deseja realmente excluir o órgão " & txtSai.Text & " alterar seus registros para o órgão " & txtFica.Text & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "update tab_clientes set cod_org='" & txtFica.tag & "' where "
    strSQL = strSQL & " cod_org='" & txtSai.tag & "'"
    i = Conn.Execute(strSQL)
    strSQL = "delete from tab_orgaos  "
    strSQL = strSQL & "where cod_org='" & txtSai.tag & "'"
    Conn.Execute strSQL
    
    
    MsgBox " registros alterados "
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtFica_Change()
    txtFica.tag = ""
End Sub

Private Sub txtFica_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFica
    End If
End Sub

Private Sub txtSai_Change()
    txtSai.tag = ""
End Sub

Private Sub txtSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaSai
    End If
End Sub
Sub buscaSai()
    Dim o As orgao.orgao
    o = F_BuscaOrgao(txtSai)
    If o.cod_org <> "" Then
        
        txtSai.Text = o.nome_org
        txtSai.tag = o.cod_org
        lblCodAnt.caption = o.cod_org
    End If
    
    
    
End Sub
Sub buscaFica()
    Dim o As orgao.orgao
    o = F_BuscaOrgao(txtFica)
    If o.cod_org <> "" Then
        txtFica.Text = o.nome_org
        txtFica.tag = o.cod_org
        
        lblCodNovo.caption = o.cod_org
        'cmdConf.SetFocus
    End If
    
    
    
End Sub

