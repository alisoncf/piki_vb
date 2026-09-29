VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInutilizarNFe 
   Caption         =   "Inutilizar NFe"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   5055
      Left            =   8400
      TabIndex        =   20
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8916
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Verificar"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtDe 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4560
      TabIndex        =   16
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtAte 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbxSerie 
      Height          =   315
      ItemData        =   "frmInutilizarNFe.frx":0000
      Left            =   1440
      List            =   "frmInutilizarNFe.frx":000D
      TabIndex        =   14
      Text            =   "cmbxSerie"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   6720
      TabIndex        =   12
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   615
      Left            =   5160
      TabIndex        =   11
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtJust 
      Height          =   765
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   7935
   End
   Begin VB.ComboBox cmbxModelo 
      Height          =   315
      ItemData        =   "frmInutilizarNFe.frx":0031
      Left            =   2640
      List            =   "frmInutilizarNFe.frx":003B
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   1710
   End
   Begin VB.TextBox txtAno 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtFinal 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtInicial 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      Height          =   195
      Left            =   4560
      TabIndex        =   18
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Até"
      Height          =   195
      Left            =   5760
      TabIndex        =   17
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblEmpresa 
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Justificativa:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Série:"
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Modelo"
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Número Final:"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número inicial:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmInutilizarNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Sub buscaEmpresa()
    Dim emp As sped_empresa.sped_empresa
    
    emp = sped_empresa.SelecionarSped_Empresa("")
    If emp.codigo <> "" Then
        InicializaEmpresa emp.codigo
        setarCertificado emp.codigo
        lblEmpresa.caption = emp.codigo & "-" & emp.nome
    End If
End Sub
Private Sub cmdConfirmar_Click()
    Dim strSQL As String
    
    
    
    
    If MsgBox("Confirma pedido de inutilização das NFe's " & txtInicial.text & " a " & txtFinal.text & "?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    strSQL = PedidoInut(strCNPJ_Emp, txtJust.text, txtAno.text, Left(cmbxModelo.text, 2), txtSerie.text, txtInicial.text, _
        txtFinal.text)
    
    
    
    
    If Left(strSQL, 4) <> "ERRO" Then
        If NFe_DAO.GravaNfePedidoInut(txtInicial.text, txtFinal.text, txtJust.text) Then
            ImprimeComprovanteNaoFiscalTEF strSQL, True
        End If
        
    Else
        MsgBox strSQL
    End If
    
    
End Sub

Private Sub cmdExportar_Click()

End Sub

Private Sub cmdConsultar_Click()
    buscarLacunas
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    buscaEmpresa
    bActive = True
    
End Sub
Sub ajustaGrid()
    
    
    ajustaGridTitulos grid, 3, "Nº Inicial,Nº Final,Data", "1000,1000,1000"
    
    grid.ColAlignment(2) = 1
End Sub
Sub buscarLacunas()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim proximo As String
    Dim anterior As String
    Dim inicio_lacuna As Long
    Dim fim_lacuna As Long
    On Error GoTo trata_Erro
    strSQL = "SELECT nfe.nNF,nfe.serie,nfe.modelo, nfe.cod_emp, nfe.dhEmi "
    strSQL = strSQL & " from NFe "
    strSQL = strSQL & " WHERE nfe.status=2"
    strSQL = strSQL & " AND nfe.dhEmi>=" & F_GravaData(txtDe.text)
    strSQL = strSQL & " and cod_emp=" & strCod_Emp
    strSQL = strSQL & " and serie=" & cmbxSerie.text
    strSQL = strSQL & " and modelo='" & Left(cmbxModelo.text, 2) & "' "
    strSQL = strSQL & " ORDER BY nnf "
    
    abreObjetoAdo rPesq, strSQL
    grid.rows = 1
    With rPesq
        anterior = !nNF - 1
        Do While Not .EOF
            
            If !nNF - 1 = anterior Then
                anterior = !nNF
                If inicio_lacuna <> 0 Then
                    fim_lacuna = !nNF - 2
                    grid.AddItem inicio_lacuna & Chr(9) & fim_lacuna & Chr(9) & FormataDataNfeParaNormal(!dhEmi)
                    inicio_lacuna = 0
                End If
            Else
                inicio_lacuna = anterior + 1
                anterior = !nNF
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_Erro:
P_MostraErro
End Sub
Private Sub Form_Load()
    bActive = False
    txtAno.text = Year(Date)
    txtDe.text = Date
    txtAte.text = Date
    ajustaGrid
    P_PreencheSerie cmbxSerie
     If strCod_Cai <> "" Then
        cmbxSerie.text = F_BuscaUmCampo("serie_nfe", "tab_caixas", "cod_cai='" & (strCod_Cai) & "'")
    End If
    cmbxModelo.ListIndex = 1
End Sub

Private Sub grid_DblClick()
    Dim de As String
    Dim ate As String
    If grid.row > 0 Then
        
        txtInicial.text = GridUtil.getText(grid, 0)
        txtFinal.text = GridUtil.getText(grid, 1)
        
    End If
End Sub

Private Sub lblEmpresa_Click()
    buscaEmpresa
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub


Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtInicial_GotFocus()
    P_SelecTexto txtInicial
End Sub

Private Sub txtInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtInicial_LostFocus()
    On Error Resume Next
    txtInicial.text = Format(txtInicial.text, "00000000#")
End Sub

Private Sub txtFinal_GotFocus()
    P_SelecTexto txtFinal
End Sub

Private Sub txtFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtFinal_LostFocus()
    On Error Resume Next
    txtFinal.text = Format(txtFinal.text, "00000000#")
End Sub

