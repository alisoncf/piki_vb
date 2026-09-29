VERSION 5.00
Begin VB.Form frmCadFormasPag 
   Caption         =   "Cadastro de Formas de Pagamento"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbxTipoPagNfe 
      Height          =   315
      ItemData        =   "frmCadFormasPag.frx":0000
      Left            =   240
      List            =   "frmCadFormasPag.frx":0002
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   2400
      Width           =   5415
   End
   Begin VB.CheckBox chkGerencial 
      Caption         =   "Gerencial"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.CheckBox chkAvista 
      Caption         =   "À Vista"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtIntervalo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Text            =   "30"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtParcelas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "1"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtTipoPag 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   5175
   End
   Begin VB.TextBox txtDesconto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "0,00"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtTipoCai 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5415
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   4080
      Picture         =   "frmCadFormasPag.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Gravar"
      Height          =   735
      Left            =   2400
      Picture         =   "frmCadFormasPag.frx":016F
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdTipoPag 
      Height          =   285
      Left            =   5400
      Picture         =   "frmCadFormasPag.frx":0617
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmndbtnBuscaRep 
      Height          =   285
      Left            =   2280
      Picture         =   "frmCadFormasPag.frx":09A1
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Pagamento NFe:"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   2160
      Width           =   1560
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Dias"
      Height          =   195
      Left            =   3000
      TabIndex        =   12
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Intervalo"
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Pagamento"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto(%):"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Fluxo de Caixa:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frmCadFormasPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmdTipoNfe_Click()

End Sub

Private Sub cmdTipoPag_Click()
    Dim t As tipopag
    t = BuscaTipoPag("")
    If t.codigo <> "" Then
        txtTipoPag.Text = t.descricao
        txtTipoPag.tag = t.codigo
    End If
    
    
End Sub

Private Sub cmndbtnBuscaRep_Click()
    buscaTipoCai
End Sub
Sub buscaTipoCai()
    Dim e As esp_cai
    
    e = CAIXA.BuscaEspCai(txtTipoCai.Text)
    
    If e.sigla <> "" Then
        txtTipoCai.Text = e.sigla
    End If
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If strAltformaspag_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If Not F_senha Then Exit Sub
    If Me.tag = "" Then
        grava
    Else
        altera
    
    End If
    
End Sub
Sub grava()
    If FormasPag.GravaFormasPag(Me.tag, txtNome.Text, txtDesconto.Text, txtTipoCai.Text, _
        IIf(chkAvista.Value = 1, "S", "N"), txtParcelas.Text, txtIntervalo.Text, txtTipoPag.tag, _
        IIf(chkGerencial.Value = 1, "S", "N"), Left(cmbxTipoPagNfe.Text, 2)) Then
        
        MsgBox "Forma de pagamento GRAVADA!", vbInformation
        Unload Me
    Else
        MsgBox "erro", vbExclamation
    End If
End Sub
Sub altera()
    If FormasPag.GravaFormasPag(Me.tag, txtNome.Text, txtDesconto.Text, txtTipoCai.Text, _
        IIf(chkAvista.Value = 1, "S", "N"), txtParcelas.Text, txtIntervalo.Text, txtTipoPag.tag, _
        IIf(chkGerencial.Value = 1, "S", "N"), Left(cmbxTipoPagNfe.Text, 2)) Then
        
        MsgBox "Forma de pagamento alterada!", vbInformation
    Else
        MsgBox "erro", vbExclamation
    End If
End Sub
Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub

Private Sub Form_Load()
    bActive = False
    PreencheTipoPagNFe cmbxTipoPagNfe
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select tf.*, np.nome as nomepag, tt.cod_tip,tt.nome_tip from tab_formaspag as tf left join tab_tipopag as tt "
    strSQL = strSQL & " on tt.cod_tip=tf.cod_tip left join nfe_tpag as np on np.tpag=tf.tpag"
    
    strSQL = strSQL & " where cod_for='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Não encontrado!", vbExclamation
        End If
        txtNome.Text = !nome_for
        txtDesconto.Text = !desconto_for
        txtIntervalo.Text = !intervalo_for
        txtParcelas.Text = !parcelas_for
        txtTipoCai.Text = !tipo_cai
        P_LocalizaOperacao !tPag, cmbxTipoPagNfe
        
        txtTipoPag.Text = TiraNull(!nome_tip)
        txtTipoPag.tag = TiraNull(!cod_tip)
        
        If !avista_For = "S" Then
            chkAvista.Value = 1
        Else
            chkAvista.Value = 0
        End If
        
        If !gerencial_for = "S" Then
            chkGerencial.Value = 1
        Else
            chkGerencial.Value = 0
        End If
        
        .Close
    End With
    
End Sub
