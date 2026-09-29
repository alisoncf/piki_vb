VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFechaCaixa2 
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   6225
      TabIndex        =   42
      Top             =   7320
      Width           =   6255
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   270
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2295
      Left            =   120
      TabIndex        =   41
      Top             =   840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4048
      _Version        =   393216
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Sair"
      Height          =   735
      Left            =   8760
      Picture         =   "frmFechaCaixa2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar Caixa"
      Height          =   735
      Left            =   6960
      Picture         =   "frmFechaCaixa2.frx":054E
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   6480
      Width           =   10215
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   10185
      TabIndex        =   29
      Top             =   5400
      Width           =   10215
      Begin VB.TextBox txtTotalLancar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotalTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotalDiferenca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A lançar"
         Height          =   195
         Left            =   6120
         TabIndex        =   36
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Totalizado"
         Height          =   195
         Left            =   4800
         TabIndex        =   35
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença"
         Height          =   195
         Left            =   7440
         TabIndex        =   34
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   525
         TabIndex        =   33
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0073E1F4&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   10185
      TabIndex        =   21
      Top             =   4680
      Width           =   10215
      Begin VB.TextBox txtLancarCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   24
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotalCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDifChe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A lançar"
         Height          =   195
         Left            =   6120
         TabIndex        =   28
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Totalizado"
         Height          =   195
         Left            =   4800
         TabIndex        =   27
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença"
         Height          =   195
         Left            =   7440
         TabIndex        =   26
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimento em Cheque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   525
         TabIndex        =   25
         Top             =   240
         Width           =   2625
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   10185
      TabIndex        =   13
      Top             =   3960
      Width           =   10215
      Begin VB.TextBox txtLancarCartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotalCartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDifCartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A lançar"
         Height          =   195
         Left            =   6120
         TabIndex        =   20
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Totalizado"
         Height          =   195
         Left            =   4800
         TabIndex        =   19
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença"
         Height          =   195
         Left            =   7440
         TabIndex        =   18
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimento em Cartão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   525
         TabIndex        =   17
         Top             =   240
         Width           =   2490
      End
   End
   Begin VB.PictureBox pcDin 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   10185
      TabIndex        =   5
      Top             =   3240
      Width           =   10215
      Begin VB.TextBox txtDifDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotalDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtLancarDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimento em Dinheiro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   525
         TabIndex        =   12
         Top             =   120
         Width           =   2700
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença"
         Height          =   195
         Left            =   7440
         TabIndex        =   11
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Totalizado"
         Height          =   195
         Left            =   4800
         TabIndex        =   9
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A lançar"
         Height          =   195
         Left            =   6120
         TabIndex        =   7
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar"
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cmbxCaixa 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtData 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblObs 
      AutoSize        =   -1  'True
      Caption         =   "Observações"
      Height          =   195
      Left            =   240
      TabIndex        =   37
      Top             =   6240
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa:"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmFechaCaixa2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoResumo As ADODB.Recordset
Attribute adoResumo.VB_VarHelpID = -1
Dim WithEvents adoTotal As ADODB.Recordset
Attribute adoTotal.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmdPesquisar_Click()
    busca
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    buscaInfo
    If bActive Then Exit Sub
    busca
    
    bActive = True
End Sub
Sub buscaInfo()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from fecha_caixa where caixa='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and data=" & F_GravaData(txtData.Text)
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        strSQL = "Caixa Fechado em " & rPesq!data & " por " & rPesq!quem
        rPesq.Close
    Else
        strSQL = "Caixa não fechado para este dia"
    End If
    lblInfo.caption = strSQL
End Sub
Private Sub Form_Load()
    P_PreencheCaixas cmbxCaixa
    txtData.Text = Date
    Set adoResumo = New ADODB.Recordset
    Set adoTotal = New ADODB.Recordset
    If strCod_Cai <> "" Then
        P_LocalizaOperacao strCod_Cai, cmbxCaixa
    End If
End Sub

Private Sub txtData_GotFocus()
    P_SelecTexto txtData
End Sub
Sub calcula_diferenca()
    txtDifCartao.Text = FormatVal(CCur(txtTotalCartao.Text) - CCur(txtLancarCartao.Text))
    txtDifChe.Text = FormatVal(CCur(txtTotalCheque.Text) - CCur(txtLancarCheque.Text))
    txtDifDin.Text = FormatVal(CCur(txtTotalDin.Text) - CCur(txtLancarDin.Text))
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtData_LostFocus()
    P_AjustaData txtData
End Sub
Sub busca()
    Dim strSQL As String
    If txtData.Text = "" Then
        MsgBox "Informe a data", vbExclamation
        txtData.SetFocus
        Exit Sub
    End If
    If cmbxCaixa.Text = "" Then
        MsgBox "Informe o caixa", vbExclamation
        Exit Sub
    End If
    
    strSQL = "select efetivo,esp_Cai,sigla_cai,descricao_cai,sum(entrada_cai) as entrada,sum(saida_cai) as saida,"
    strSQL = strSQL & " sum(entrada_cai-saida_cai) as total "
    strSQL = strSQL & "  from " & TAB_CAIXA & " as tc inner join  " & TAB_ESPCAI & " as te on tc.esp_cai=te.sigla_cai "
    strSQL = strSQL & " where data_cai=" & F_GravaData(txtData.Text)
    strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " group by te.sigla_cai,te.descricao_cai "
    abreObjetoAdo adoResumo, strSQL
    
    carregar_grid
    totalizar
End Sub
Sub totalizar()
    Dim strSQL As String
    Dim cDin As Currency
    Dim cCheque As Currency
    Dim cCartao As Currency
    With adoResumo
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If !Esp_cai = "DIN" Or !Esp_cai = "RDI" Or !Esp_cai = "MOV" Then
                cDin = cDin + !total
            End If
            If !Esp_cai = "RCC" Or !Esp_cai = "RCC" Or !Esp_cai = "CRD" Or !Esp_cai = "CRT" Then
                cCartao = cCartao + !total
            End If
            If !Esp_cai = "CHP" Or !Esp_cai = "CHV" Or !Esp_cai = "RCH" Then
                cCheque = cCheque + !total
            End If
            
            
            .MoveNext
        Loop
        
        
    End With
    txtTotalCheque.Text = FormatVal(cCheque)
    txtTotalDin.Text = FormatVal(cDin)
    txtTotalCartao.Text = FormatVal(cCartao)
    txtTotalTotal.Text = FormatVal(cCartao + cDin + cCheque)
    calcula_diferenca
End Sub
Sub carregar_grid()
    grid.fontName = "Arial"
     AjustaGridData grid, 5, 1, 0, "Sigla,Descrição,Entrada,Saida,Saldo", _
     "sigla_cai,descricao_cai,entrada,saida,total", _
        "1000,3000,1000,1000,1000", adoResumo, "2,3,4", True
    Dim f As fonte
    
    
    f.cor_face = vbBlue
    
    f.negrito = True
        
    ColoreLinhaGridKey grid, "DIN", 0, f
    ColoreLinhaGridKey grid, "CRT", 0, f
    ColoreLinhaGridKey grid, "CRD", 0, f
        
    f.cor_face = &H40C0&
        ColoreLinhaGridKey grid, "MOV", 0, f
    f.cor_face = &HFF8080
    ColoreLinhaGridKey grid, "RDI", 0, f
    ColoreLinhaGridKey grid, "RCC", 0, f
    ColoreLinhaGridKey grid, "RCD", 0, f
    
End Sub

Private Sub txtLancarDin_GotFocus()
    P_SelecTexto txtLancarDin
End Sub

Private Sub txtLancarDin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtLancarDin_LostFocus()
    P_ValorLostFocus txtLancarDin
    calcula_diferenca
End Sub

Private Sub txtLancarCarta_GotFocus()
    P_SelecTexto txtLancarCartao
End Sub

Private Sub txtLancarCarta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtLancarCartao_LostFocus()
    P_ValorLostFocus txtLancarCartao
    calcula_diferenca
End Sub
Private Sub txtLancarCheque_GotFocus()
    P_SelecTexto txtLancarCheque
End Sub

Private Sub txtLancarCheque_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtLancarCheque_LostFocus()
    P_ValorLostFocus txtLancarCheque
    calcula_diferenca
End Sub
