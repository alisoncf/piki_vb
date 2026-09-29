VERSION 5.00
Begin VB.Form frmAlteraImpostos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impostos"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Para o Imposto:"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
      Begin VB.CommandButton cmndbtnConf 
         Caption         =   "Confirmar"
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Top             =   300
         Width           =   1575
      End
      Begin VB.ComboBox cmbxNovoImp 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Imposto:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mudar Todos os Produtos:"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtbxDescricao 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   3615
      End
      Begin VB.ComboBox cmbxCategoria 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ComboBox cmbxImposto 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Produtos Afetados:"
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label lblProdutos 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Imposto:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmAlteraImpostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_erro
    strSQL = "select count(tp.cod_pro) as total from " & TAB_PRODUTOS & " as tp, tab_categorias as tc "
    strSQL = strSQL & " where tp.cod_cat=tc.cod_Cat "
    If cmbxCategoria.Text <> "" Then
        strSQL = strSQL & " and nome_Cat='" & cmbxCategoria.Text & "' "
    End If
    If cmbxImposto.Text <> "" Then
        strSQL = strSQL & " and cod_imp='" & Left(cmbxImposto.Text, 3) & "' "
    End If
    strSQL = strSQL & " and nome_pro like '%" & txtbxDescricao.Text & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then
            lblProdutos.caption = rPesq!total
        Else
            lblProdutos.caption = "0"
        End If
        .Close
    End With
    Exit Sub
trata_erro:
    MostraErro

End Sub

Private Sub cmbxCategoria_Click()
    cmbxCategoria.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCategoria.Text & "'")
    Busca
    
End Sub

Private Sub cmbxCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxImposto.ListIndex = -1
    End If
End Sub

Private Sub cmbxImposto_Click()
    Busca
End Sub

Private Sub cmbxImposto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxImposto.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnConf_Click()
    Dim strSQL As String
    If MsgBox("Confirma alteração dos produtos?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "update " & TAB_PRODUTOS & " set cod_imp='" & Left(cmbxNovoImp.Text, 3) & "' "
    strSQL = strSQL & " where nome_pro like '%" & txtbxDescricao.Text & "%' "
    If cmbxImposto.tag <> "" Then
        strSQL = strSQL & " and cod_imp='" & Left(cmbxImposto.Text, 3) & "'"
    End If
    If cmbxCategoria.tag <> "" Then
        strSQL = strSQL & " and cod_cat='" & cmbxCategoria.tag & "'"
    End If
    conn.Execute strSQL
    MsgBox "Produtos alterados! " & banco.RecordsAffected, vbInformation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    P_PreencheImpostos cmbxImposto
    P_PreencheImpostos cmbxNovoImp
    P_PreencheCateg cmbxCategoria
End Sub

Private Sub txtbxDescricao_Change()
    Busca
End Sub
