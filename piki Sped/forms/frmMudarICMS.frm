VERSION 5.00
Begin VB.Form frmMudarICMS 
   Caption         =   "Alteração de ICMS"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355
   LinkTopic       =   "Form2"
   ScaleHeight     =   3075
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   6840
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   3960
      ScaleHeight     =   1515
      ScaleWidth      =   3315
      TabIndex        =   8
      Top             =   360
      Width           =   3375
      Begin VB.Label lblRetorno 
         Caption         =   "Produtos encontrados"
         Height          =   1095
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.TextBox txtEmpresa 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtICMS 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "17"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtCFOP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdEmpresa 
      Height          =   285
      Left            =   3000
      Picture         =   "frmMudarICMS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdCFOP 
      Height          =   285
      Left            =   3000
      Picture         =   "frmMudarICMS.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Novo ICMS (%)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   420
   End
End
Attribute VB_Name = "frmMudarICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
    alterar
End Sub

Private Sub cmdCFOP_Click()
    Dim e As Padrao1
    e = Sped_Lista_CFOP(txtCFOP.Text)
    txtCFOP.Text = e.codigo & "-" & e.nome
    txtCFOP.tag = e.codigo
    busca
End Sub

Private Sub cmdEmpresa_Click()
    Dim e As sped_empresa.sped_empresa
    e = sped_empresa.BuscaSpedEmpresa("")
    txtEmpresa.Text = e.apelido
    txtEmpresa.tag = e.codigo
    busca
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select count(*) as total from estoque where cfop_consumidor='" & txtCFOP.tag & "' "
    strSQL = strSQL & " and cod_emp='" & txtEmpresa.tag & "'"
    abreObjetoAdo rPesq, strSQL
    
    lblRetorno.caption = rPesq!total & " Produtos que serão afetados "
End Sub
Sub alterar()
    Dim strSQL As String
    Dim i As Integer
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão pra essa operação!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma alteração do ICMS de todos os produtos com essa condição?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    strSQL = "update estoque set picms=" & F_GravaDec(txtICMS.Text) & " where cod_emp='" & txtEmpresa.tag & "' "
    strSQL = strSQL & " and cfop_consumidor='" & txtCFOP.tag & "' "
    conn.Execute strSQL, i
    MsgBox i & " registros alterados", vbInformation
    
    busca
End Sub
