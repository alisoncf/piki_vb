VERSION 5.00
Begin VB.Form frmAltProPCP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Incluir/Alterar Produto do Processo de Compra:"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnHist 
      Caption         =   "Histórico de Compras"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame frAltMer 
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtbxQFA_Ipc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "0,00"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnSim 
         Caption         =   "Sim"
         Height          =   615
         Left            =   7200
         Picture         =   "frmAltProPCP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   390
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnNao 
         Caption         =   "Não"
         Height          =   615
         Left            =   8520
         Picture         =   "frmAltProPCP.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   390
         Width           =   1215
      End
      Begin VB.CommandButton cmdForn 
         Height          =   285
         Left            =   6600
         Picture         =   "frmAltProPCP.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Unitário:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total.:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornecedor:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblNome_Mer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Produto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   45
         Width           =   1710
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. Comprar:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº:"
      Height          =   195
      Left            =   2160
      TabIndex        =   25
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label lblNumNot 
      AutoSize        =   -1  'True
      Caption         =   "00/00/0000"
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
      Left            =   2400
      TabIndex        =   24
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label lblEmb 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
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
      Left            =   6480
      TabIndex        =   23
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Embalagem:"
      Height          =   195
      Left            =   5520
      TabIndex        =   22
      Top             =   1800
      Width           =   870
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
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
      Left            =   8520
      TabIndex        =   21
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço Unitário:"
      Height          =   195
      Left            =   7320
      TabIndex        =   20
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label lblQtd 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
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
      Left            =   4800
      TabIndex        =   19
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd Comprada:"
      Height          =   195
      Left            =   3600
      TabIndex        =   18
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label lblNomeFor 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Fornecedor"
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
      Left            =   5520
      TabIndex        =   17
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor:"
      Height          =   195
      Left            =   4560
      TabIndex        =   16
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblUltComp 
      AutoSize        =   -1  'True
      Caption         =   "00/00/0000"
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
      Left            =   3360
      TabIndex        =   15
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Última Compra:"
      Height          =   195
      Left            =   2160
      TabIndex        =   14
      Top             =   1560
      Width           =   1065
   End
End
Attribute VB_Name = "frmAltProPCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As String
Private Sub cmbxForn_Click()
    'cmbxForn.tag = F_BuscaUmCampo("cod_For", TAB_FORNECEDORES, "nome_for='" & cmbxForn.Text & "'")
End Sub

Private Sub cmbxForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmdForn_Click()
    buscaFor
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_BuscaFornecedor(txtForn.Text)
    If f.cod_For <> "" Then
        txtForn.Text = f.nome_For
        txtForn.tag = f.cod_For
    End If
End Sub
Private Sub cmndbtnHist_Click()
    P_InfoEntradasPro lblNome_Mer.tag
End Sub

Private Sub cmndbtnNao_Click()
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iRow As Integer

        
        strSQL = "update tab_itenspcp set sts_ipc='N',"
        strSQL = strSQL & "QFA_ipc=" & str(txtbxQFA_Ipc.Text)
        
        strSQL = strSQL & " where Cod_pro='" & lblNome_Mer.tag & "' "
        strSQL = strSQL & " and seq_pcp='" & Me.tag & "' "
        conn.Execute strSQL
        
        Me.Hide
    


End Sub

Private Sub cmndbtnSim_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iRow As Integer
        If txtForn.tag = "" Then
            MsgBox "Informe o fornecedor!", vbExclamation
            txtForn.SetFocus
            Exit Sub
        End If
        If Int(F_BuscaUmCampo("count(*)", "tab_itenspcp", "cod_pro='" & lblNome_Mer.tag & "'" & _
            " and seq_pcp='" & Me.tag & "'")) > 1 Then   'tem q ser mudado em caso de mais lugares
            MsgBox "Este item já está duplicado!", vbCritical
            Exit Sub
        End If
        
        
        strSQL = "update tab_itenspcp set sts_ipc='S',"
        strSQL = strSQL & "QFA_ipc=" & str(txtbxQFA_Ipc.Text) & ","
        strSQL = strSQL & "total_ipc=" & str(txtbxTotal.Text) & ","
        strSQL = strSQL & "cod_for='" & txtForn.tag & "' "
        strSQL = strSQL & " where Cod_pro='" & lblNome_Mer.tag & "' "
        strSQL = strSQL & " and seq_pcp='" & Me.tag & "' "
        
        
        conn.Execute strSQL
        
    
        Me.Hide

End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Activate()
    BuscaUltEntrada
End Sub
Sub VerificaProcesso()
    Dim strSQL As String
    
End Sub
Private Sub Form_Load()
    'MudaAparencia Me
    'P_PreencheFornecedores cmbxForn
End Sub
Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select num_doc,dt_e_s,tf.nome,ti.*,tfor.nome_For,tfor.cod_for  "
    strSQL = strSQL & " from sped_nf as nf inner join sped_nf_itens as ti on nf.id=ti.id_nf "
    strSQL = strSQL & " left join " & "sped_participante" & " as tf  on tf.codigo=nf.cod_part"
    strSQL = strSQL & " left join " & TAB_FORNECEDORES & " as tfor  on tf.cod_for=tfor.cod_for "
    
    strSQL = strSQL & " where ti.cod_pro='" & lblNome_Mer.tag & "' "
    strSQL = strSQL & " order by dt_e_s desc limit 1 "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            lblUltComp.caption = ""
            lblNomeFor.caption = ""
            lblNumNot.caption = ""
            lblQtd.caption = ""
            lblUnit.caption = ""
            lblEmb.caption = ""
        Else
            lblUltComp.caption = !DT_E_S
            lblNomeFor.caption = !nome
            lblNumNot.caption = !num_doc
            lblQtd.caption = !qtd
            lblUnit.caption = !vl_item
            lblEmb.caption = !qtd_convert
            txtForn.Text = !nome_For
            txtForn.tag = !cod_For
        End If
        .Close
    End With
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Private Sub Text1_GotFocus()
    
End Sub

Private Sub txtbxQFA_Ipc_GotFocus()
    ganhaFoco txtbxQFA_Ipc
End Sub

Private Sub txtbxQFA_Ipc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxQFA_Ipc_LostFocus()
    PerdeFoco txtbxQFA_Ipc
    TotalItem
End Sub

Sub TotalItem()
    If CCur(txtbxQFA_Ipc.Text) > 0 Then
        txtbxUnit.Text = FormatVal(CCur(txtbxTotal.Text) / CCur(txtbxQFA_Ipc.Text), "###,##0.0000")
    Else
        txtbxUnit.Text = "0"
    End If
End Sub
Private Sub txtbxTotal_GotFocus()
    ganhaFoco txtbxTotal
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxTotal_LostFocus()
    PerdeFoco txtbxTotal
    TotalItem
End Sub
