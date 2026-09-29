VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTroca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troca de Produtos"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "frmTroca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnDesconto 
      Caption         =   "De&sconto:"
      Height          =   255
      Left            =   7800
      TabIndex        =   50
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtbxDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7800
      TabIndex        =   49
      Text            =   "0,00"
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1440
      ScaleHeight     =   705
      ScaleWidth      =   3825
      TabIndex        =   42
      Top             =   120
      Width           =   3855
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         Height          =   195
         Left            =   2400
         TabIndex        =   47
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblVendedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   2400
         TabIndex        =   45
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblDataVen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5400
      ScaleHeight     =   705
      ScaleWidth      =   2265
      TabIndex        =   35
      Top             =   120
      Width           =   2295
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Abater"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblAbater 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   40
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Acrescer:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblAcresc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   38
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final Previsto:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lblFinalPrev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   36
         Top             =   480
         Width           =   390
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   5745
      TabIndex        =   22
      Top             =   4800
      Width           =   5775
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "À Vista"
         Height          =   195
         Left            =   840
         TabIndex        =   34
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "À Prazo"
         Height          =   195
         Left            =   1680
         TabIndex        =   33
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financeira"
         Height          =   195
         Left            =   2520
         TabIndex        =   32
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final"
         Height          =   195
         Left            =   4560
         TabIndex        =   30
         Top             =   120
         Width           =   330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   405
      End
      Begin VB.Label lblInicial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblAvista 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   840
         TabIndex        =   27
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblAprazo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1680
         TabIndex        =   26
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblFinan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   2520
         TabIndex        =   25
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblDesconto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3600
         TabIndex        =   24
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblFinal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.TextBox txtbxQtd_alt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   3240
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data dtProd 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Msflxgrd 
      Height          =   2415
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4260
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Nº da Venda:"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      Begin VB.TextBox txtbxNum_Ven 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   8655
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnAdd 
         Caption         =   "&Adicionar"
         Height          =   495
         Left            =   7320
         Picture         =   "frmTroca.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
      Begin VB.ComboBox cmbxBuscar 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmTroca.frx":0E14
         Left            =   120
         List            =   "frmTroca.frx":0E24
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtbxProduto 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   9
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   375
         Left            =   4200
         Picture         =   "frmTroca.frx":0E55
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblRef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref:"
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
         Left            =   4200
         TabIndex        =   48
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6120
         TabIndex        =   16
         Top             =   120
         Width           =   405
      End
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
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
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   10
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label18 
         Caption         =   "&Buscar Por:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblNome_pro 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4680
         TabIndex        =   8
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   5
         Top             =   120
         Width           =   555
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   6000
      TabIndex        =   17
      Top             =   4680
      Width           =   2775
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Concluir"
         Height          =   615
         Left            =   120
         Picture         =   "frmTroca.frx":0F9F
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   1440
         Picture         =   "frmTroca.frx":10E9
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnAdd_Click()
    Dim strSQL As String
    Dim iCount As Integer
    If CCur(txtbxTotal.Text) = 0 Then
        MsgBox "Não foi informado um valor válido!", vbCritical
        txtbxQtd.SetFocus
        Exit Sub
    End If
    If F_InFlxgrd(msflxgrd, FormatCodigo(txtbxProduto.tag), 1) > -1 Then
        MsgBox "Este produto já existe"
        Exit Sub
    End If
    If cmndbtnAdd.tag = "" Then
        If MsgBox("Confirma inclusão do Produto " & txtbxProduto.Text & "?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "" & Chr(9) & FormatCodigo(txtbxProduto.tag) & Chr(9)
            strSQL = strSQL & lblProduto.caption & Chr(9) & txtbxQtd.Text & Chr(9)
            strSQL = strSQL & txtbxUnit.Text & Chr(9) & txtbxTotal.Text & Chr(9)
            strSQL = strSQL & txtbxTotal.Text & Chr(9) & txtbxQtd.Text & Chr(9) & "S"
            strSQL = strSQL & Chr(9) & lblRef.caption
            msflxgrd.AddItem strSQL
            msflxgrd.Row = msflxgrd.rows - 1
            For iCount = 0 To msflxgrd.Cols - 1
                
                msflxgrd.col = iCount
                msflxgrd.CellBackColor = vbInfoBackground
                msflxgrd.CellForeColor = vbBlue
            Next
            lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) + CCur(txtbxTotal.Text))
            lblFinalPrev.caption = FormatVal(CCur(lblAcresc.caption) - CCur(lblAbater.caption))
            lblFinalPrev.tag = lblFinalPrev.caption
        End If
    Else
    
    End If
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim p As Produto
    
    p = F_BuscaProdSaida(txtbxProduto.Text)
    
    txtbxProduto.tag = p.cod_pro
    If cmbxBuscar.ListIndex = 0 Then
        txtbxProduto.Text = p.codbarra_pro
    ElseIf cmbxBuscar.ListIndex = 1 Then
        txtbxProduto.Text = p.codfab_pro
    ElseIf cmbxBuscar.ListIndex = 2 Then
        txtbxProduto.Text = p.nome_pro
    ElseIf cmbxBuscar.ListIndex = 3 Then
        txtbxProduto.Text = p.cod_pro
    End If
    txtbxUnit.Text = FormatVal(p.p2_pro)
    lblProduto.caption = p.nome_pro
    lblRef.caption = Format(p.codfab_pro)
    txtbxQtd.SetFocus
End Sub
Sub P_InsereProd(sCod_Pro As String, cUnit As Currency, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    strSQL = F_BuscaUmCampo("max(seq_ite)", "" & TAB_ITVENDAS & "", "num_ven='" & Me.tag & "'")
    If strSQL = "" Then
        iIndex = 1
    Else
        iIndex = val(strSQL) + 1
    End If
    strSQL = "insert into " & TAB_ITVENDAS & "(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite)values("
    
    strSQL = strSQL & "'" & Me.tag & "', "
    strSQL = strSQL & "'" & sCod_Pro & "',"
    strSQL = strSQL & F_GravaDec(CStr(cUnit)) & ","
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & ","
    strSQL = strSQL & iIndex & ") "
    conn.Execute strSQL
    'RESERVANDO O ESTOQUE
    strSQL = "update tab_composicao set e2_pro=e2_pro - " & F_GravaDec(CStr(cQtd))
    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
    conn.Execute strSQL
        
End Sub



Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If F_VerificaAltera = False And F_VerificaNovoProduto = False Then
        MsgBox "Não há alterações na troca!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma conclusão da troca?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    conclui
End Sub
Sub conclui()
    Dim strSQL As String
    Dim sSeq As String
    Dim iCount As Integer
    Dim cQtd_Ant As Currency
    Dim cQtd_Atual As Currency
    Dim sPro As String
    Dim sCod_Pro As String
    Dim cFinal As Currency
    Dim sSeq_ite As String
    Dim sUnit_Ite As String
    Dim cAux As Currency
    Dim cIndice As Currency
    sSeq = MaxKeySQL("tab_trocas", "seq_tro", "", "T00000#")
    
    
    If CCur(lblFinalPrev.caption) > 0 Then
            frmTipoVenda.operacao = "Troca"
            frmTipoVenda.lblNome_cli.caption = lblCliente.caption
            frmTipoVenda.lblNome_cli.tag = lblCliente.tag
            frmTipoVenda.tag = txtbxNum_Ven.Text
            frmTipoVenda.lblnum_ven.caption = txtbxNum_Ven.Text
            frmTipoVenda.lblValor_Ven.caption = lblFinalPrev.caption
            frmTipoVenda.lblValorFinal.caption = lblFinalPrev.caption
            frmTipoVenda.lblVendedor.caption = strNome_Usu
            frmTipoVenda.txtbxDesconto.Text = FormatVal(txtbxDesconto.Text)
            frmTipoVenda.Show 1
            If frmTipoVenda.tag = "" Then
                MsgBox "Os Valores não foram informados corretamente!", vbExclamation
                
                Me.tag = ""
                Unload frmTipoVenda
                
                Exit Sub
            End If
            Unload frmTipoVenda
    ElseIf CCur(lblFinalPrev.caption) < 0 Then
        strSQL = F_GravaCai("D", Date, FormatVal(-1 * CCur(lblFinalPrev.caption)), "TROCA REF. VENDA " & txtbxNum_Ven.Text & "EM " & Date, txtbxNum_Ven.Text, "DIN", "", 0, 0, strCod_Cai, strCod_Tur)
        
    End If
    
    
    strSQL = "insert into tab_trocas("
    strSQL = strSQL & "Num_Ven, data_tro, cod_usu, seq_tro, abat_tro,acres_tro)values("
    strSQL = strSQL & "'" & txtbxNum_Ven.Text & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaDec(lblAbater.caption) & ", "
    strSQL = strSQL & F_GravaDec(lblAcresc.caption) & ") "
    conn.Execute strSQL
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 0: sPro = .Text
            .col = 1: sCod_Pro = F_SoNumero(.Text)
            .col = 3: cQtd_Atual = CCur(.Text)
            .col = 4: sUnit_Ite = (.Text)
            .col = 7: cQtd_Ant = CCur(.Text)
            .col = 6: cFinal = CCur(.Text)
            If sPro <> "" Then
                If cQtd_Ant <> cQtd_Atual Then
                    strSQL = "insert into tab_ittrocas(seq_tro,cod_pro,qtdant_itt,qtdatual_itt,unit_itt)values("
                    strSQL = strSQL & "'" & sSeq & "',"
                    strSQL = strSQL & "'" & sCod_Pro & "',"
                    strSQL = strSQL & F_GravaDec(cQtd_Ant) & ",  "
                    strSQL = strSQL & F_GravaDec(cQtd_Atual) & ",  "
                    strSQL = strSQL & F_GravaDec(cFinal) & ")  "
                    conn.Execute strSQL
                    strSQL = "update " & TAB_ITVENDAS & " set "
                    strSQL = strSQL & "qtd_ite=" & F_GravaDec(cQtd_Atual) & ","
                    If cQtd_Ant > cQtd_Atual Then
                        strSQL = strSQL & "final_ite=" & F_GravaDec(cFinal / cQtd_Ant * cQtd_Atual)
                    Else
                        cIndice = (cQtd_Ant * cFinal) / CCur(lblFinal.caption)
                        cAux = (cFinal / cQtd_Ant * cQtd_Atual)
                        strSQL = strSQL & "final_ite=" & F_GravaDec((CCur(lblFinalPrev.caption) * cIndice) + (CCur(lblFinal.caption) * cIndice))
                    End If
                    strSQL = strSQL & " where num_ven='" & txtbxNum_Ven.Text & "' "
                    strSQL = strSQL & " and seq_ite=" & sPro
                    conn.Execute strSQL
                Else
                    
                
                End If
                If cQtd_Ant > cQtd_Atual Then
                    strSQL = "update tab_composicao set e2_pro=e2_pro+" & F_GravaDec(cQtd_Ant - cQtd_Atual)
                    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
                    conn.Execute strSQL
                ElseIf cQtd_Ant < cQtd_Atual Then
                    strSQL = "update tab_composicao set e2_pro=e2_pro-" & F_GravaDec(cQtd_Atual - cQtd_Ant)
                    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
                    conn.Execute strSQL
                
                End If

                
            Else
                'inserir produto
                sSeq_ite = MaxKeySQL("" & TAB_ITVENDAS & "", "seq_ite", "num_ven='" & txtbxNum_Ven.Text & "'", "")
                strSQL = "insert into " & TAB_ITVENDAS & "(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite,descricao_ite,final_ite)values( "
                strSQL = strSQL & "'" & txtbxNum_Ven.Text & "',"
                .col = 1
                strSQL = strSQL & "'" & F_SoNumero(.Text) & "',"
                .col = 4
                strSQL = strSQL & F_GravaDec(.Text) & ","
                .col = 3
                strSQL = strSQL & F_GravaDec(.Text) & ","
                strSQL = strSQL & sSeq_ite & ","
                .col = 2
                strSQL = strSQL & "'" & .Text & "',"
                .col = 6
                strSQL = strSQL & F_GravaDec(.Text) & ")"
                conn.Execute strSQL
                .col = 3
                strSQL = "update tab_composicao set e2_pro=e2_pro-" & F_GravaDec(.Text)
                .col = 1
                strSQL = strSQL & " where cod_pro='" & F_SoNumero(.Text) & "' "
                conn.Execute strSQL
            End If
        Next
        MsgBox "Troca Concluída!", vbInformation
        imprimeTroca Printer
        Printer.EndDoc
        txtbxNum_Ven.Text = ""
        msflxgrd.rows = 1
    End With
    
    
End Sub
    
Function F_VerificaAltera() As Boolean
    Dim iCount As Integer
    Dim cQtd_Atual As Currency
    Dim cQtd_Anterior As Currency
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 3: cQtd_Atual = CCur(.Text)
            .col = 7: cQtd_Anterior = CCur(.Text)
            If cQtd_Anterior <> cQtd_Atual Then
                        
                F_VerificaAltera = True
                Exit Function
            End If
        Next
        F_VerificaAltera = False
    End With
End Function
Function F_VerificaNovoProduto() As Boolean
    Dim iCount As Integer
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 0
            If .Text = "" Then
                        
                F_VerificaNovoProduto = True
                Exit Function
            End If
        Next
        F_VerificaNovoProduto = False
    End With
End Function

Private Sub cmndbtnDesconto_Click()
    Dim sUsu As String
    
    If CCur(lblFinalPrev.caption) <= 0 Then
        MsgBox "Operação inválida!", vbCritical
        Exit Sub
    End If
    
    sUsu = F_SenhaCodBar
    If sUsu <> "" Then
        txtbxDesconto.Enabled = True
        cmndbtnDesconto.tag = sUsu
        txtbxDesconto.SetFocus
    End If
    
End Sub
Private Sub txtbxDesconto_GotFocus()
    ganhaFoco txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    Dim cDesc As Currency
    'PerdeFoco txtbxDesconto
    cDesc = CCur(F_BuscaUmCampo("desconto_usu", "" & TAB_DEFUSU & "", _
        "cod_usu='" & cmndbtnDesconto.tag & "'"))
    
    If CCur(txtbxDesconto.Text) > CCur(lblFinalPrev.caption) * cDesc / 100 Then
        MsgBox "Desconto muito alto!", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    
    End If
    lblFinalPrev.caption = FormatVal(lblFinalPrev.tag) - CCur(txtbxDesconto.Text)
    txtbxDesconto.Enabled = False
End Sub
Private Sub Form_Click()
    txtbxQtd_alt.Visible = False
End Sub
Sub imprimeTroca(obj As Object)
    On Error Resume Next
    Dim strSQL As String
    Dim sSeq As String
    Dim iCount As Integer
    Dim cQtd_Ant As Currency
    Dim cQtd_Atual As Currency
    Dim sPro As String
    Dim sCod_Pro As String
    Dim cFinal As Currency
    Dim sSeq_ite As String
    Dim sUnit_Ite As String
    P_CabRelTimbr "TROCA REFERENTE À VENDA " & txtbxNum_Ven.Text & " EM " & Date, obj
    P_ImprX obj, "Venda: " & txtbxNum_Ven.Text, 0
    P_ImprX obj, "" & lblCliente.caption, 30, False
    P_ImprX obj, "" & lblDataVen.caption, 0
    P_ImprX obj, "" & lblVendedor.caption, 45
    P_ImprX obj, "" & lblValor.caption, 100
    P_ImprX obj, "" & lblDesc.caption, 160, False
    P_ImprLinha obj
    P_ImprX obj, "Inicial: " & lblInicial.caption, 0
    P_ImprX obj, "À Vista: " & lblAvista.caption, 30
    P_ImprX obj, "A Prazo: " & lblAprazo.caption, 60
    P_ImprX obj, "Financeira: " & lblFinan.caption, 90
    
    
    P_ImprX obj, "Desconto: R$ " & lblDesconto.caption, 130
    P_ImprX obj, "Final: " & lblFinal.caption, 170, False
    P_ImprLinha obj, "="
    With msflxgrd
        obj.FontBold = True
        P_ImprX obj, "...", 0
        P_ImprX obj, "Cód", 5
        P_ImprX obj, "Descrição", 24
        P_ImprX obj, "Qtd", 120, True, True
        P_ImprX obj, "Unit", 135, True, True
        P_ImprX obj, "Total", 150, True, True
        P_ImprX obj, "Final", 165, True, True
        P_ImprX obj, "Inicial", 180, True, True
        P_ImprX obj, "Ref:", 182, False
        obj.FontBold = False
        P_ImprLinha obj
        For iCount = 1 To msflxgrd.rows - 1
            .Row = iCount
            .col = 0: sPro = .Text
            .col = 3: cQtd_Atual = CCur(.Text)
            .col = 7: cQtd_Ant = CCur(.Text)
            If cQtd_Atual <> cQtd_Ant Then
                obj.FontBold = True
                obj.ForeColor = vbBlue
            ElseIf sPro = "" Then
                obj.FontBold = True
                obj.ForeColor = vbRed
            Else
                obj.FontBold = False
                obj.ForeColor = vbBlack
            End If
            .col = 0: P_ImprX obj, .Text, 0
            .col = 1: P_ImprX obj, .Text, 5
            .col = 2: P_ImprX obj, .Text, 24
            .col = 3: P_ImprX obj, .Text, 120, True, True
            .col = 4: P_ImprX obj, .Text, 135, True, True
            .col = 5: P_ImprX obj, .Text, 150, True, True
            .col = 6: P_ImprX obj, .Text, 165, True, True
            .col = 7: P_ImprX obj, .Text, 180, True, True
            .col = 9: P_ImprX obj, .Text, 182, False
            obj.FontBold = False
            obj.ForeColor = vbBlack
        Next
        P_ImprLinha obj
        P_ImprX obj, "A Abater: " & lblAbater.caption, 10
        P_ImprX obj, "A Acrescer: " & lblAcresc.caption, 60
        P_ImprX obj, "Desconto: " & txtbxDesconto.Text, 110
        P_ImprX obj, "Final Previsto: " & lblFinalPrev.caption, 160, False
        P_ImprLinha obj, "_"
    End With
End Sub
Private Sub Form_Load()
    adc.ConnectionString = strConnect
    cmbxBuscar.ListIndex = 1
    With msflxgrd
        .FixedCols = 0
        .FixedRows = 1
        .Cols = 10
        .rows = 1
        .col = 0:  .Text = "..."
        .col = 1: .Text = "Cód"
        .col = 2: .Text = "Descrição"
        .col = 3: .Text = "Qtd"
        .col = 4: .Text = "Unit"
        .col = 5: .Text = "Total "
        .col = 6: .Text = "Final "
        .col = 7: .Text = "Qtd Ini."
        .col = 8: .Text = "*"
        .col = 9: .Text = "Ref."
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 400
        .ColWidth(9) = 2000
        
    End With
    
End Sub

Private Sub msflxgrd_Click()
    txtbxQtd_alt.Visible = False
        
End Sub

Private Sub msflxgrd_DblClick()
    If msflxgrd.col = 3 Then
        msflxgrd_KeyDown 13, 0
    End If
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        msflxgrd.col = 3
        txtbxQtd_alt.Left = msflxgrd.ColPos(3) + msflxgrd.Left
        txtbxQtd_alt.Top = msflxgrd.Top + (msflxgrd.RowHeight(msflxgrd.Row) * msflxgrd.Row)
        txtbxQtd_alt.Height = msflxgrd.RowHeight(msflxgrd.Row)
        txtbxQtd_alt.Width = msflxgrd.ColWidth(msflxgrd.col)
        txtbxQtd_alt.Text = msflxgrd.Text
        txtbxQtd_alt.Visible = True
        txtbxQtd_alt.SetFocus
    ElseIf KeyCode = vbKeyDelete Then
        
    End If
End Sub

Private Sub txtbxNum_Ven_GotFocus()
    ganhaFoco txtbxNum_Ven
End Sub

Private Sub txtbxNum_Ven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Sub BuscaVenda()
    'lblValorGerado.caption = "0,00"
    'lblItens.caption = "0"
    If F_Vazio(txtbxNum_Ven.Text) Then
        Exit Sub
    End If
    txtbxNum_Ven.Text = Format(F_SoNumero(txtbxNum_Ven.Text), "000000#")
    Me.tag = txtbxNum_Ven.Text
    busca
    If txtbxNum_Ven.Text = "" Then Exit Sub
    P_buscaitens "seq_ite"
        
End Sub
Sub P_buscaitens(sOrd As String)
    strSQL = "select tv.*,tp.codfab_pro,tp.nome_pro,tc.p2_pro,tc.c2_pro,tg.nome_cat,tg.nome_cat,"
    strSQL = strSQL & "te.nome_mar,te.cod_mar,ti.*,(ti.final_ite*ti.qtd_ite) as final,(ti.unit_ite*ti.qtd_ite) as total "
    strSQL = strSQL & " from tab_produtos as tp," & TAB_VENDAS & " as tv, tab_categorias as tg,"
    strSQL = strSQL & " tab_marcas as te, tab_Itvendas as ti,tab_composicao as tc "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and tp.cod_cat=tg.cod_cat "
    strSQL = strSQL & " and tc.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tp.cod_mar=te.cod_mar "
    strSQL = strSQL & " and ti.num_ven=tv.num_ven and tv.num_ven='" & Me.tag & "' "
    strSQL = strSQL & " order by " & sOrd
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    With dtProd.Recordset
        msflxgrd.rows = 1
        If .RecordCount = 0 Then Exit Sub
        
        Do While Not .EOF
            strSQL = !seq_ite & Chr(9) & FormatCodigo(!cod_pro) & Chr(9)
            strSQL = strSQL & !nome_pro & Chr(9) & (!qtd_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9) & FormatVal(!total) & Chr(9)
            strSQL = strSQL & FormatVal(!final_ite) & Chr(9) & !qtd_ite & Chr(9) & "N"
            strSQL = strSQL & Chr(9) & Format(!codfab_pro)
            msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        .MoveFirst
    End With
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    strSQL = "select tv.*,tf.avista_ven,tf.desconto_ven,tf.aprazo_ven,"
    strSQL = strSQL & "tf.valorfinal_ven,tf.finan_ven from " & TAB_VENDAS & " "
    strSQL = strSQL & " as tv,  " & TAB_VENFEC & "  as tf where "
    strSQL = strSQL & " tv.num_ven=tf.num_ven and tf.num_ven='" & txtbxNum_Ven.Text & "'"
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Esta venda não existe!", vbExclamation
            txtbxNum_Ven.Text = ""
            txtbxNum_Ven.SetFocus
            Exit Sub
        End If
        If !situacao_ven < "3" Then
            MsgBox "Não é possível haver devoluções nesta venda!" _
                & Chr(13) & "Ela ainda não se encontra finalizada!", vbExclamation
            txtbxNum_Ven.Text = ""
            txtbxNum_Ven.SetFocus
            Exit Sub
        End If
        lblCliente.caption = "Cliente: " & F_BuscaUmCampo("Nome_cli", "tab_clientes", _
            "cod_cli='" & !cod_cli & "'")
        lblCliente.tag = !cod_cli
        lblDataVen.caption = "Data: " & !data_ven
        lblValor.caption = "Valor: " & FormatVal(!valor_ven)
        lblDesc.caption = "Desc: " & Round(!desconto_ven / !valor_ven, 2) * 100 & " % "
        lblDesc.tag = Round(!desconto_ven / !valor_ven, 2) * 100
        
        lblVendedor.caption = "Vendedor: " & F_BuscaUmCampo("guerra_fun", "tab_funcionarios", _
            "cod_fun='" & !cod_fun & "'")
        lblInicial.caption = FormatVal(!valor_ven)
        lblAvista.caption = FormatVal(!avista_ven)
        lblAprazo.caption = FormatVal(!aprazo_ven)
        lblFinan.caption = FormatVal(!finan_ven)
        lblDesconto.caption = FormatVal(!desconto_ven)
        lblFinal.caption = FormatVal(!valorfinal_ven)
        lblAbater.caption = "0,00"
        lblAcresc.caption = "0,00"
        lblFinalPrev.caption = "0,00"
        lblFinalPrev.tag = "0,00"
        .Close
    End With
End Sub


Private Sub txtbxNum_Ven_LostFocus()
    BuscaVenda
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxQtd_alt_GotFocus()
    ganhaFoco txtbxQtd_alt
End Sub

Private Sub txtbxQtd_alt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cFinal As Currency
    Dim iQtd As Integer
    
    On Error GoTo trata_erro
    If KeyCode = vbKeyReturn Then
        txtbxQtd_alt_LostFocus
        If msflxgrd.Text = txtbxQtd_alt.Text Then
            txtbxQtd_alt.Visible = False
            Exit Sub
        End If
        
        If MsgBox("Confirma alteração da quantidade de " & msflxgrd.Text & _
            " para " & txtbxQtd_alt.Text, vbYesNo + vbCritical) = vbYes Then
            If txtbxQtd_alt.Text = 0 Then
                msflxgrd.col = 6
                cFinal = CCur(msflxgrd.Text)
            Else
                msflxgrd.col = 3
                iQtd = msflxgrd.Text
                If iQtd = 0 Then
                    iQtd = 1
                End If
                msflxgrd.col = 6
                cFinal = Round(CCur(msflxgrd.Text / iQtd) * (iQtd - CCur(txtbxQtd_alt.Text)), 2)
            End If
            msflxgrd.col = 3
            If cFinal = 0 Then
'                cFinal = ""
            End If
            If CCur(txtbxQtd_alt.Text) > CCur(msflxgrd.Text) Then
                
'                lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) + _
'                    ((CCur(txtbxQtd_alt.Text) - CCur(msflxgrd.Text)) * cFinal))
                 lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) + _
                    cFinal)
                
            Else
'                lblAcresc.caption = FormatVal(CCur(lblAcresc.caption) + _
'                    ((CCur(msflxgrd.Text) - CCur(txtbxQtd_alt.Text)) * cFinal))
                lblAbater.caption = FormatVal(CCur(lblAbater.caption) + _
                     cFinal)
            End If
            lblFinalPrev.caption = FormatVal(CCur(lblAcresc.caption) - CCur(lblAbater.caption))
            lblFinalPrev.tag = lblFinalPrev.caption
            msflxgrd.Text = txtbxQtd_alt.Text
            msflxgrd.CellBackColor = vbInfoBackground
            msflxgrd.CellForeColor = vbBlue
            txtbxQtd_alt.Visible = False
            
        End If
    End If ' parei aki
    If KeyCode = vbKeyEscape Then
        txtbxQtd_alt.Visible = False
        msflxgrd.SetFocus
    End If
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub txtbxQtd_alt_LostFocus()
    PerdeFoco txtbxQtd_alt, "###,##0"
End Sub

Private Sub txtbxQtd_GotFocus()
    ganhaFoco txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    PerdeFoco txtbxQtd, "###,##0"
    txtbxTotal.Text = FormatVal(CCur(txtbxQtd.Text) * CCur(txtbxUnit.Text))
End Sub
