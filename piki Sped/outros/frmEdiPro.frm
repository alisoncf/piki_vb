VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiPro 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Edição de Produtos"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   14100
   Icon            =   "frmEdiPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   14100
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pcMsg 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00B36A09&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   14070
      TabIndex        =   36
      Top             =   8775
      Width           =   14100
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5610
         TabIndex        =   37
         Tag             =   "!"
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.PictureBox pcBotoes 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00B0B0B0&
      ForeColor       =   &H80000008&
      Height          =   7425
      Left            =   12315
      ScaleHeight     =   7395
      ScaleWidth      =   1755
      TabIndex        =   17
      Top             =   0
      Width           =   1785
      Begin VB.CommandButton cmndbtnHtml 
         Caption         =   "&Impr. HTML"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3780
         Width           =   1575
      End
      Begin VB.ListBox lstBarCode 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkEst 
         Caption         =   "Ver Estoque"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "!"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkTodas 
         Caption         =   "Listar &Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "!"
         Top             =   0
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDelPro 
         Caption         =   "E&xcluir Produto"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":0997
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnCodBarra 
         Caption         =   "Emitir C.B&arra"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":0AE1
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4500
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnLista 
         Caption         =   "&List. Preços"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":106B
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Fe&char"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":11B5
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5820
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":12FF
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2460
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":1449
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Códigos de Barras:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14070
      TabIndex        =   6
      Top             =   7920
      Width           =   14100
      Begin VB.ComboBox cmbxICMS 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":17D3
         Left            =   12600
         List            =   "frmEdiPro.frx":17E3
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbxSt 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":1804
         Left            =   11160
         List            =   "frmEdiPro.frx":1814
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbxCriterio 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":1835
         Left            =   9480
         List            =   "frmEdiPro.frx":184E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   345
         Width           =   1575
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   7200
         MaxLength       =   40
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmbxBusca 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmEdiPro.frx":18B2
         Left            =   4320
         List            =   "frmEdiPro.frx":18CB
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":1915
         Left            =   5640
         List            =   "frmEdiPro.frx":1925
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":1952
         Left            =   120
         List            =   "frmEdiPro.frx":1954
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbxMarca 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":1956
         Left            =   2160
         List            =   "frmEdiPro.frx":1958
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS:"
         Height          =   195
         Left            =   12600
         TabIndex        =   53
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação Trib:"
         Height          =   195
         Left            =   11160
         TabIndex        =   4
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Critério:"
         Height          =   195
         Left            =   9480
         TabIndex        =   2
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   4320
         TabIndex        =   11
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   7200
         TabIndex        =   0
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   5640
         TabIndex        =   13
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblCat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1800
         TabIndex        =   35
         Tag             =   "!"
         Top             =   120
         Width           =   270
      End
      Begin VB.Label lblMar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   3840
         TabIndex        =   34
         Tag             =   "!"
         Top             =   120
         Width           =   270
      End
   End
   Begin VB.PictureBox pcDados 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   14070
      TabIndex        =   27
      Top             =   7425
      Width           =   14100
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4080
         ScaleHeight     =   495
         ScaleWidth      =   8535
         TabIndex        =   38
         Top             =   0
         Width           =   8535
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última Compra:"
            Height          =   195
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   1065
         End
         Begin VB.Label lblUltComp 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1200
            TabIndex        =   49
            Top             =   0
            Width           =   810
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor:"
            Height          =   195
            Left            =   2400
            TabIndex        =   48
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblNomeFor 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Fornecedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3360
            TabIndex        =   47
            Top             =   0
            Width           =   1740
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd Comprada:"
            Height          =   195
            Left            =   1440
            TabIndex        =   46
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label lblQtd 
            AutoSize        =   -1  'True
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2640
            TabIndex        =   45
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço Unitário:"
            Height          =   195
            Left            =   5160
            TabIndex        =   44
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label lblUnit 
            AutoSize        =   -1  'True
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6360
            TabIndex        =   43
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Embalagem:"
            Height          =   195
            Left            =   3360
            TabIndex        =   42
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lblEmb 
            AutoSize        =   -1  'True
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4320
            TabIndex        =   41
            Top             =   240
            Width           =   315
         End
         Begin VB.Label lblNumNot 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº:"
            Height          =   195
            Left            =   0
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   225
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   0
         TabIndex        =   30
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lblMarca 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   600
         TabIndex        =   29
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblCategoria 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame frProd 
      Caption         =   "&Produtos:"
      Height          =   5775
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8055
      Begin VB.Data dtPro 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   4320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEdiPro.frx":195A
         Height          =   5415
         Left            =   120
         OleObjectBlob   =   "frmEdiPro.frx":196E
         TabIndex        =   16
         Top             =   240
         Width           =   7815
      End
      Begin VB.Label lblList 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listadas:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   7320
         TabIndex        =   26
         Top             =   50
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmEdiPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    strSQL = "select num_not,entrada_not,tf.nome_for,ti.* "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti, "
    strSQL = strSQL & " tab_fornecedores as tf "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & dtPro.Recordset!cod_pro & "' "
    strSQL = strSQL & " order by entrada_not desc limit 1 "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            lblUltComp.caption = ""
            lblNomeFor.caption = ""
            lblNumNot.caption = ""
            lblQtd.caption = ""
            lblUnit.caption = ""
            lblEmb.caption = ""
        Else
            lblUltComp.caption = Format(!entrada_not)
            lblNomeFor.caption = !nome_for
            lblNumNot.caption = !num_not
            lblQtd.caption = !qtd_ite
            lblUnit.caption = !unit_ite
            lblEmb.caption = !qtd_emb
        End If
        .Close
    End With
    Exit Sub
trata_Erro:
'    P_MostraErro
End Sub

Sub P_TileMsg(msg As String)
    lblMsg.caption = msg
End Sub
Private Sub chkEst_Click()
    If chkEst.Value = 1 Then
        dbgrd.Columns(3).Visible = True
    Else
        dbgrd.Columns(3).Visible = False
    End If
End Sub

Private Sub chkTodas_Click()
    busca
End Sub

Private Sub cmbxBusca_Click()
    'If Me.Visible = True Then busca
End Sub

Private Sub cmbxDiv_Click()
    lblDiv.caption = F_BuscaUmCampo("Cod_div", "tab_divisoes", "nome_div='" & cmbxDiv.Text & "'")
    busca
End Sub

Private Sub cmbxDiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxDiv.ListIndex = -1
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub

Private Sub cmbxLoja_Click()
    Dim rPesq As Recordset
    Dim strSQL As String
    Dim iPos As Integer
    
    On Error GoTo trata_Erro
    strSQL = "select * from tab_lojas where cod_Loj='" & cmbxLoja.Text & "' "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Me.caption = "Edição de Produtos: " & !nome_loj & " - " & !Cidade_Loj
        .Close
    End With
    iPos = dtPro.Recordset.AbsolutePosition
    busca
    dtPro.Recordset.Move iPos
    dbgrd.SetFocus
trata_Erro:
End Sub

Private Sub cmbxLoja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxCriterio_Click()
    busca
End Sub

Private Sub cmbxICMS_Click()
    busca
End Sub

Private Sub cmbxICMS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxICMS.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_Click()
    lblMar.caption = F_BuscaUmCampo("Cod_mar", "tab_marcas", "nome_mar='" & cmbxMarca.Text & "'")
    busca
End Sub

Private Sub cmbxEsp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxEsp.ListIndex = -1
    End If
End Sub

Private Sub cmbxCat_Click()
    lblCat.caption = F_BuscaUmCampo("Cod_Cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
    
    busca
End Sub

Private Sub cmbxGenero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxGenero.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrd_Click()
    'If Me.Visible = True Then busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSt_Click()
    busca
End Sub

Private Sub cmbxSt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnCodBarra_Click()
    Dim fr As New frmGeraBarCode
    P_PreencheBarCode dtPro.Recordset!cod_pro, fr.lstBarCode
    fr.txtbxDesc.Text = dtPro.Recordset!nome_pro
    fr.txtbxPreco.Text = FormatVal(dtPro.Recordset!pV_pro)
    fr.Show 0
End Sub

Private Sub cmndbtnDelPro_Click()
    Dim strSQL As String
    With dtPro.Recordset
        If MsgBox("Confirma exclusão do Produto?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser excluído. Selecione um produto!", vbExclamation
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("cod_pro", "tab_itcmp", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "" Then
            MsgBox "Existem compras com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("cod_pro", "tab_itcompras", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "" Then
            MsgBox "Existem compras com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("cod_pro", "tab_itvendas", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "" Then
            MsgBox "Existem vendas com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("cod_pro", "tab_itsai", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "" Then
            MsgBox "Existem vendas com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("er_pro", "tab_produtos", "cod_pro = '" & !cod_pro & "'")
        If sAux > 0 Then
            MsgBox "Existe estoque com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        sAux = F_BuscaUmCampo("ef_pro", "tab_produtos", "cod_pro = '" & !cod_pro & "'")
        If sAux <> "0" Then
            MsgBox "Existe estoque com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
            Exit Sub
        End If
        strSQL = "delete from tab_produtos where cod_pro = '" & !cod_pro & "'"
        banco.Execute strSQL
        strSQL = "delete from tab_embalagens where cod_pro = '" & !cod_pro & "'"
        banco.Execute strSQL
        strSQL = "delete from tab_codbarras where cod_pro = '" & !cod_pro & "'"
        banco.Execute strSQL
    End With
    dtPro.Refresh
    MsgBox "Produto Excluído!", vbExclamation
End Sub

Private Sub cmndbtnEdit_Click()
    On Error Resume Next
    
    Dim fr As New frmCadPro
    
    If strCadprod_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    If dbgrd.Row = -1 Then
    Exit Sub
    End If
    
    With dtPro.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser Editado!", vbExclamation
            Exit Sub
        End If
        fr.Tag = !cod_pro
        fr.caption = !cod_pro & " - " & !nome_pro
        fr.cmbxMarca.Text = !nome_mar
        fr.cmbxCat.Text = !nome_cat
        fr.txtbxEstMin.Text = !estmin_pro
        If !balanca_pro = "S" Then
            fr.chkProd.Value = 1
        Else
            fr.chkProd.Value = 0
        End If
        fr.txtbxCodFab.Text = !codfab_pro
        'fr.txtbxCustoFiscal.Text = FormatVal(!c1_pro)
        
        'fr.txtbxEstoque.Text = FormatVal(!estoque_pro)
        fr.txtbxNome_pro.Text = !nome_pro
        fr.txtbxPreco.Text = FormatVal(!pV_pro)
        fr.txtbxCusto.Text = FormatVal(!cr_pro)
        'fr.txtbxCodBar = !codbarra_pro
        fr.txtbxUnidade.Text = !Unidade_pro
        fr.txtbxMargem.Text = FormatVal(!lucro_pro)
        fr.txtbxLocal = !Endereco_pro
        P_PreencheBarCode !cod_pro, fr.lstCodBarra
        P_PreencheEmbalagens !cod_pro, fr.lstEmb
        fr.cmbxMarca.Enabled = True
        fr.cmbxCat.Enabled = True
        fr.txtbxEstoque.Text = FormatVal(!er_pro)
        P_LocalizaOperacao !st_pro, fr.cmbxSt
        P_LocalizaOperacao !cod_imp, fr.cmbxICMS
        If !st_pro = "T" Then
            fr.cmbxICMS.Enabled = True
        End If
        fr.Show 0
        
    End With
End Sub

Private Sub cmndbtnFornadas_Click()
    frmDadosFornada.Tag = dtPro.Recordset!cod_pro
    frmDadosFornada.caption = "Fornadas: " & _
    Format(dtPro.Recordset!cod_pro, "00#\.###\-# ") & dtPro.Recordset!nome_pro
    frmDadosFornada.frControle.caption = dtPro.Recordset!nome_pro
    On Error Resume Next
    frmDadosFornada.Show 1
End Sub

Private Sub cmndbtnHtml_Click()
    P_DbgrdToHtml "LISTAGEM DE PRODUTOS", dbgrd, dtPro, 0
    
    
End Sub

Private Sub cmndbtnLista_Click()
If dbgrd.Row = -1 Then
        MsgBox "Não há itens na sua pesquisa!", vbExclamation
        Exit Sub
    End If
    With dtPro.Recordset
        
        If .RecordCount = 0 Then
            MsgBox "Não há itens na sua pesquisa!", vbExclamation
            Exit Sub
        End If
        P_ImprimeListagem dtPro, frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.Tag <> "" Then
            P_ImprimeListagem dtPro, Printer
            Printer.EndDoc
        End If
        Unload frmVisualiza
        .MoveFirst
    End With
    P_Status ""
End Sub

Private Sub cmndbtnNew_Click()
    If strCadprod_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    Dim fr As New frmCadPro
    fr.Show 0
        
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub dbgrd_DblClick()
    
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    ElseIf KeyCode = vbKeyC Then
        dbgrd.Col = 1
        Clipboard.SetText dbgrd.Text
    ElseIf KeyCode = vbKeyB Or KeyCode = vbKeyN Or KeyCode = vbKeyM Or _
                KeyCode = vbKeyG Or KeyCode = vbKeyF Then
        
            
            On Error Resume Next
            frmBusca.Show 1
            If frmBusca.Tag <> "" Then
                If KeyCode = vbKeyB Then
                    cmbxBusca.ListIndex = 1
                End If
                If KeyCode = vbKeyN Then
                    cmbxBusca.ListIndex = 0
                End If
                If KeyCode = vbKeyF Then
                    cmbxBusca.ListIndex = 2
                End If
                If KeyCode = vbKeyM Then
                    cmbxBusca.ListIndex = 3
                End If
                If KeyCode = vbKeyG Then
                    cmbxBusca.ListIndex = 4
                End If
                txtbxExpr.Text = frmBusca.txtbxExpr.Text
            End If
            Unload frmBusca
        
    End If
    If KeyCode = vbKeyI And Shift = 1 Then
        If strViscusto_usu <> "S" Then
            MsgBox "Você não tem permissão para isto!", vbCritical
            Exit Sub
        End If
        P_InfPro dtPro.Recordset!cod_pro
    End If
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        Clipboard.SetText dbgrd.Text
    End If
End Sub


Private Sub dbgrd_LostFocus()
   dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtPro_Reposition()
    On Error Resume Next
    With dtPro.Recordset
        If .RecordCount > 0 Then
            lblEst.caption = FormatVal(!Estoque)
            lblPreco.caption = FormatVal(!p2_pro)
            lblEstRes.caption = FormatVal(!r_pro)
            lblMarca.caption = !nome_mar
            lblCategoria.caption = !nome_cat
            txtbxST.Text = !s_pro
            P_PreencheBarCode !cod_pro, lstBarCode
            BuscaUltEntrada
        Else
            
        End If
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
        
    txtbxExpr.SetFocus
    cmbxLoja_Click
    If strViscusto_usu = "S" Then
        chkEst.Enabled = True
        chkEst_Click
    Else
        chkEst.Enabled = False
    End If
    txtbxExpr.SetFocus
    
    
End Sub

Private Sub Form_Load()
        
    dbgrd.Columns(3).Visible = False
    P_CentraForm Me
    dtPro.Connect = banco.Connect
    P_PreencheCateg cmbxCat
    P_PreencheMarca cmbxMarca
    
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0
    
    resizeForm frmEdiPro
    P_PreencheImpostos cmbxICMS
    'Busca
End Sub

Private Sub Form_Resize()
   posicionaControles
   P_CentraObejct lblMsg, Me, False
End Sub
Sub posicionaControles()
    pcBotoes.Left = Me.Width - pcBotoes.Width - 200
    frProd.Width = pcBotoes.Left - 200
    dbgrd.Width = frProd.Width - 200
    'pcBusca.Width = Me.Width - 200
    'pcBusca.Top = Me.Height - pcBusca.Height - 400 - pcMsg.Height
    'pcDados.Width = pcBusca.Width
    'pcDados.Top = pcBusca.Top - pcDados.Height - 50
    frProd.Height = pCdados.Top - 100
    dbgrd.Height = frProd.Height - 300
End Sub

Private Sub pcBotoes_GotFocus()
    SendKeys Chr(9)
End Sub

Private Sub pcBusca_GotFocus()
    SendKeys Chr(9)
End Sub

Sub busca()
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    If cmbxBusca.ListIndex = 0 Then
        sCampo = "tp.nome_pro"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sCampo = "tp.codfab_pro"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sCampo = "tb.codbarra_pro"
    ElseIf cmbxBusca.ListIndex = 3 Then
        sCampo = "tp.cod_pro"
    ElseIf cmbxBusca.ListIndex = 4 Then
        sCampo = "tg.nome_cat"
    ElseIf cmbxBusca.ListIndex = 6 Then
        sCampo = "tp.endereco_pro"
    Else
        sCampo = "te.nome_mar"
    End If
    
    If cmbxOrd.ListIndex = 1 Then
        sOrder = "tp.cod_pro,tp.nome_pro"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrder = "tp.codfab_pro,tp.nome_pro,tp.cod_pro"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrder = "tp.endereco_pro,tp.nome_pro,tp.cod_pro"
    Else
        sOrder = "tp.nome_pro,tp.cod_pro "
    End If
    strSQL = "select distinct if(ef_pro<0,' - ','') as sinal, tg.nome_cat, te.nome_mar, tp.*,((tp.pv_pro/tp.cr_pro-1)*100)-" & F_GravaDec(F_TaxasEntradas) & "  as liq, "
    strSQL = strSQL & " ((tp.pv_pro/tp.cr_pro-1)*100) as lb, "
    strSQL = strSQL & " (((tp.pv_pro/tp.cr_pro-1)*100)-" & F_GravaDec(F_TaxasEntradas) & ") * pv_pro/100  as lr, "
    strSQL = strSQL & " ((tp.pv_pro/tp.cr_pro-1)*100) * pv_pro/100  as bruto "
    strSQL = strSQL & " from tab_produtos as tp,tab_categorias as tg, tab_codbarras as tb, "
    strSQL = strSQL & " tab_marcas as te "
    If cmbxICMS.Text <> "" Then
        strSQL = strSQL & ", tab_impostos as ti where ti.cod_imp=tp.cod_imp  "
        strSQL = strSQL & " and st_pro='T' "
        strSQL = strSQL & " and ti.cod_imp='" & Left(cmbxICMS.Text, 3) & "' and "
        
    Else
        strSQL = strSQL & " where "
    End If
    strSQL = strSQL & sCampo & " like '"
    strSQL = strSQL & txtbxExpr.Text & "%' "
    strSQL = strSQL & " and tp.cod_pro=tb.cod_pro"
    
    strSQL = strSQL & " and tg.cod_cat=tp.cod_cat and te.cod_mar=tp.cod_mar "
    If chkTodas.Value = 0 Then
        strSQL = strSQL & " and tp.er_pro >0"
    End If
    
    If cmbxCat.ListIndex <> -1 Then
        strSQL = strSQL & " and tp.cod_cat='" & lblCat.caption & "'  "
    End If
    If cmbxMarca.ListIndex <> -1 Then
        strSQL = strSQL & " and tp.cod_mar='" & lblMar.caption & "'  "
    End If
    If cmbxSt.Text <> "" Then
        strSQL = strSQL & " and st_pro='" & Left(cmbxSt.Text, 1) & "' "
    End If
    If cmbxCriterio.ListIndex = 1 Then
        strSQL = strSQL & " and ef_pro<er_pro"
    ElseIf cmbxCriterio.ListIndex = 2 Then
        strSQL = strSQL & " and ef_pro>er_pro"
    ElseIf cmbxCriterio.ListIndex = 3 Then
        strSQL = strSQL & " and ef_pro<0"
    ElseIf cmbxCriterio.ListIndex = 4 Then
        strSQL = strSQL & " and er_pro<estmin_pro"
    ElseIf cmbxCriterio.ListIndex = 5 Then
        strSQL = strSQL & " and ef_pro>(er_pro*2)"
    ElseIf cmbxCriterio.ListIndex = 6 Then
        strSQL = strSQL & " and er_pro>(ef_pro*2)"
    End If
    
    
    strSQL = strSQL & " order by " & sOrder
    dtPro.RecordSource = strSQL
    dtPro.Refresh
    lblList.caption = "[Listadas..: " & dtPro.Recordset.RecordCount & _
        "   Cadastradas: " & F_BuscaUmCampo("count(*)", "tab_produtos", "") & "]"
    
End Sub

Private Sub txtbxExpr_GotFocus()
    P_TileMsg "Informe " & cmbxBusca.Text & " clique tecle enter..."
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        busca
        If dtPro.Recordset.RecordCount > 0 Then
            dbgrd.SetFocus
        End If
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dbgrd.SetFocus
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    P_TileMsg ""
    
End Sub
