VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVinculaDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vinculação de Documentos"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   13650
   Icon            =   "frmVinculaDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGrid 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6600
      TabIndex        =   33
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   11880
      Picture         =   "frmVinculaDoc.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   10440
      Picture         =   "frmVinculaDoc.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   240
      TabIndex        =   15
      Top             =   5880
      Width           =   6255
      Begin VB.TextBox txtBancoBol 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   36
         Text            =   "CEF-"
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkBol 
         Caption         =   "Gerar Nº Boleto"
         Height          =   375
         Left            =   1200
         TabIndex        =   35
         Top             =   960
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmdLancar 
         Caption         =   "Distribuir"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtBoleto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtbxParcelas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtbxVenc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxValorFinal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxDec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxAcr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblBancoBol 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Boleto"
         Height          =   195
         Left            =   4680
         TabIndex        =   28
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Parc."
         Height          =   195
         Left            =   4080
         TabIndex        =   31
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Data Base:"
         Height          =   195
         Left            =   3000
         TabIndex        =   26
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor Final:"
         Height          =   195
         Left            =   2040
         TabIndex        =   20
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "D&ecréscimo:"
         Height          =   195
         Left            =   1080
         TabIndex        =   18
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Acréscimo:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Ordenar Por:"
      Height          =   735
      Left            =   6840
      TabIndex        =   2
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton optnDoc 
         Caption         =   "&Documento"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Vencimento"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmndbtnBusca 
      Appearance      =   0  'Flat
      Caption         =   "&Buscar"
      Height          =   645
      Left            =   8640
      Picture         =   "frmVinculaDoc.frx":0C5E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid msflxGrd 
      Height          =   4575
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8070
      _Version        =   393216
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Cliente:"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.PictureBox pcFoco 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   40
      ScaleHeight     =   255
      ScaleWidth      =   13380
      TabIndex        =   24
      Top             =   960
      Width           =   13380
   End
   Begin MSFlexGridLib.MSFlexGrid gridParcelas 
      Height          =   1455
      Left            =   6600
      TabIndex        =   32
      Top             =   5880
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lblDistribuido 
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
      Left            =   6600
      TabIndex        =   34
      Top             =   7800
      Width           =   390
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total:"
      Height          =   195
      Left            =   6240
      TabIndex        =   14
      Top             =   5640
      Width           =   810
   End
   Begin VB.Label lblTotal 
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
      Left            =   7200
      TabIndex        =   13
      Top             =   5640
      Width           =   390
   End
   Begin VB.Label lblValorVinc 
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
      Left            =   3360
      TabIndex        =   12
      Top             =   5640
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Vinculado:"
      Height          =   195
      Left            =   2040
      TabIndex        =   11
      Top             =   5640
      Width           =   1155
   End
   Begin VB.Label lblParcVinc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   5880
      TabIndex        =   10
      Top             =   5640
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas Vinculadas:"
      Height          =   195
      Left            =   4320
      TabIndex        =   9
      Top             =   5640
      Width           =   1485
   End
   Begin VB.Label lblParcList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   8
      Top             =   5640
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas Listadas:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   5640
      Width           =   1290
   End
   Begin VB.Menu mnUtil 
      Caption         =   "&Parcelas"
      Begin VB.Menu mnVinc 
         Caption         =   "&Vincular Todos"
      End
      Begin VB.Menu mnDesvinc 
         Caption         =   "&Desvincular &Todos"
      End
   End
End
Attribute VB_Name = "frmVinculaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cCorSelect As Single
Private Sub cmndbtnTerc_Click()
    Dim c As Clientes.cliente
    c = Clientes.F_BuscaCliente(txtbxExpr.Text)
    If c.cod_cli <> "" Then
        txtbxExpr.Text = c.nome_Cli
        txtbxExpr.tag = c.cod_cli
         lblParcList.caption = "0"
        lblParcVinc.caption = "0"
        lblTotal.caption = "0"
        lblValorVinc.caption = "0"
        txtbxVenc.Text = Date
        txtbxParcelas.Text = 1
        
        busca
    
    End If
End Sub

Private Sub cmdLancar_Click()
    Dim i As Integer
    Dim sBol As String
    
    ajustaGridParcelas
    sBol = txtBoleto.Text
    
    For i = 1 To Int(txtbxParcelas.Text)
        If chkBol.Value = 1 Then
            sBol = "{Doc}:" & i
        End If
        strSQL = gridParcelas.rows & Chr(9) & DateAdd("M", i, txtbxVenc.Text) & Chr(9)
        strSQL = strSQL & CCur(txtbxValorFinal.Text / txtbxParcelas.Text) & Chr(9) & sBol
        gridParcelas.AddItem strSQL
    Next
    soma
End Sub

Private Sub cmndbtnBusca_Click()
    busca
End Sub
Sub MarcaTodos(bOpcao As Boolean)
    Dim i As Integer
    
    With msflxgrd
        busca
        If bOpcao = True Then
            For i = 1 To .rows - 1
                .row = i
                msflxgrd_DblClick
            Next
        End If
        
    End With
    
End Sub
Private Sub cmndbtnConfirmar_Click()
    If val(lblParcVinc.caption) <= 1 Then
        MsgBox "Não houve vinculação de Parcelas!", vbCritical
        Exit Sub
    End If
    If CCur(txtbxValorFinal.Text) = 0 Then
        MsgBox "O Valor Final está zerado!", vbCritical
        Exit Sub
    End If
    If txtbxParcelas.Text = "0" Then
        MsgBox "Informe um número de Parcelas válido!", vbCritical
        Exit Sub
    End If
    Vincula
End Sub
Sub Vincula()
    Dim strSQL As String
    Dim sSEQ_DOC As String
    Dim sSeq_Vinc As String
    Dim seq_itd As String
    Dim iCount As Integer
    Dim sNumDoc As String
    Dim vencimento As String
    Dim valor As String
    Dim boleto As String
    
    If gridParcelas.rows <= 1 Then
        MsgBox "O lançamento não está completo!", vbExclamation
        Exit Sub
        
    End If
    If CCur(lblDistribuido.caption) < CCur(lblValorVinc.caption) Then
        MsgBox "O valor não está totalmente distribuído", vbExclamation
        Exit Sub
    End If
    
    
    With msflxgrd
        sSeq_Vinc = MaxKeySQL("tab_vinc", "seq_vin", "", "000000#")
        sNumDoc = "V" & sSeq_Vinc
        sSEQ_DOC = F_GravaDoc("VINCULAÇÃO DE DOCUMENTOS", sNumDoc, _
            txtbxValorFinal.Text, Date, txtbxExpr.tag, "R", "", txtbxParcelas.Text, "F", "C", strPlaRecebDupl)
        
        If sSEQ_DOC <> "" Then
            For iCount = 1 To gridParcelas.rows - 1
                gridParcelas.row = iCount
                gridParcelas.col = 1: vencimento = gridParcelas.Text
                gridParcelas.col = 2: valor = gridParcelas.Text
                gridParcelas.col = 3:
                If chkBol.Value = 1 Then
                    gridParcelas.Text = txtBancoBol.Text & sSEQ_DOC & "-" & iCount
                End If
                gridParcelas.col = 3: boleto = gridParcelas.Text
                
                
                strSQL = F_GravaItDoc(sSEQ_DOC, valor, _
                     vencimento, vencimento, CStr(iCount), a_receber_avulso, boleto, "1", "000", "", "")
                If strSQL = "" Then
                    MsgBox "!erro", vbCritical
                    Exit Sub
                End If
            Next
        End If
        For iCount = 1 To .rows - 1
            
            .row = iCount:
            .col = 1
            If .Text = "*" Then
                
                .col = 0
                sSeq_Vinc = MaxKeySQL("tab_vinc", "seq_vin", "", "000000#")
                strSQL = "update  " & TAB_ITDOC & "  set reman_itd=0 where seq_itd='" & .Text & "'"
                conn.Execute strSQL
                
                .col = 4
                valor = .Text
            
                .col = 0
                seq_itd = .Text
                
                strSQL = "insert into tab_vinc(seq_vin,seq_itd,data_vin,"
                strSQL = strSQL & "valor_vin,acres_vin,decres_vin,venc_vin,"
                strSQL = strSQL & "parc_vin,seq_doc)values("
                strSQL = strSQL & "'" & sSeq_Vinc & "',"
                strSQL = strSQL & "'" & seq_itd & "',"
                strSQL = strSQL & F_GravaData(Date) & ","
                strSQL = strSQL & F_GravaDec(valor) & ","
                strSQL = strSQL & F_GravaDec(txtbxAcr.Text) & ","
                strSQL = strSQL & F_GravaDec(txtbxDec.Text) & ","
                strSQL = strSQL & F_GravaData(txtbxVenc.Text) & ","
                strSQL = strSQL & txtbxParcelas.Text & ","
                strSQL = strSQL & "'" & sSEQ_DOC & "')"
                conn.Execute strSQL
            End If
        Next
        MsgBox "Foi criado o documento " & sNumDoc, vbInformation
        lblParcList.caption = "0"
        lblParcVinc.caption = "0"
        lblTotal.caption = "0"
        lblValorVinc.caption = "0"
        txtbxVenc.Text = Date
        txtbxParcelas.Text = 1
        busca
        
        P_DocsVinculados sSEQ_DOC
        
        
    End With
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Sub soma()
    Dim i As Integer
    Dim cTotal As Currency
    Dim c As Integer
    Dim l As Integer
    c = gridParcelas.col
    l = gridParcelas.row
    For i = 1 To gridParcelas.rows - 1
        gridParcelas.row = i
        gridParcelas.col = 2
        cTotal = cTotal + CCur(gridParcelas.Text)
        
    Next
    lblDistribuido.caption = FormatVal(cTotal)
    gridParcelas.col = c
    gridParcelas.row = l
    gridParcelas.SetFocus
End Sub
Sub ajustaGridParcelas()
    With gridParcelas
        .Cols = 4
        .rows = 1
        .col = 0: .Text = "*"
        .col = 1: .Text = "Vencimento"
        .col = 2: .Text = "Valor"
        .col = 3: .Text = "Boleto"
        .ColWidth(3) = 2000
        
        
        
        .rows = 1
    End With
End Sub
Private Sub Form_Load()
    txtbxVenc.Text = Date
    cCorSelect = &HFFFF&
    With msflxgrd
        .Cols = 9
        .FixedCols = 0
        .rows = 1
        .col = 0: .Text = "Seq."
        .col = 1: .Text = "#"
        .col = 2: .Text = "Documento"
        .col = 3: .Text = "Vencimento"
        .col = 4: .Text = "Reman"
        .col = 5: .Text = "Parc."
        .col = 6: .Text = "Inicial"
        .col = 7: .Text = "Seq.Doc."
        .col = 8: .Text = "Juros"
        .ColWidth(0) = 0
        .ColWidth(1) = 200
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 800
        .ColWidth(5) = 500
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 800
    End With
End Sub

Private Sub gridParcelas_DblClick()
    txtGrid.Left = gridParcelas.CellLeft + gridParcelas.Left
    txtGrid.Top = gridParcelas.CellTop + gridParcelas.Top
    txtGrid.width = gridParcelas.CellWidth
    txtGrid.Height = gridParcelas.CellHeight
    txtGrid.Text = gridParcelas.Text
    txtGrid.Visible = True
    txtGrid.SetFocus
    If gridParcelas.col = 1 Then
        
    ElseIf gridParcelas.col = 2 Then
    
    
    ElseIf gridParcelas.col = 3 Then
    
    End If
End Sub

Private Sub gridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        gridParcelas_DblClick
    End If
End Sub

Private Sub mnDesvinc_Click()
    MarcaTodos False
End Sub

Private Sub mnVinc_Click()
   MarcaTodos True
End Sub

Sub CalcValorFinal()
    txtbxValorFinal.Text = FormatVal(CCur(lblValorVinc.caption) + _
        CCur(txtbxAcr.Text) - CCur(txtbxDec.Text))
End Sub
Private Sub msflxgrd_DblClick()
    Dim iCount As Integer
    With msflxgrd
        .col = 1
        If .Text = "" Then
            .Text = "*"
            .col = 4
            lblValorVinc.caption = CCur(lblValorVinc.caption) + CCur(.Text)
            lblParcVinc.caption = Int(lblParcVinc.caption) + 1
            For iCount = 0 To .Cols - 1
                .col = iCount
                .CellBackColor = cCorSelect
            Next
            .col = 8
            txtbxAcr.Text = CCur(txtbxAcr.Text) + .Text
            CalcValorFinal
        Else
            .Text = ""
            .col = 4
            lblValorVinc.caption = CCur(lblValorVinc.caption) - CCur(.Text)
            lblParcVinc.caption = Int(lblParcVinc.caption) - 1
            For iCount = 0 To .Cols - 1
                .col = iCount
                .CellBackColor = vbWindowBackground
            Next
            .col = 8
            txtbxAcr.Text = CCur(txtbxAcr.Text) - .Text
            CalcValorFinal
        End If
    End With
End Sub

Private Sub msflxGrd_GotFocus()
    pcFoco.BackColor = vbBlue
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        msflxgrd_DblClick
    End If
End Sub

Private Sub msflxGrd_LostFocus()
    pcFoco.BackColor = vbWhite
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnUtil
    End If
End Sub

Private Sub msflxgrd_RowColChange()
    Dim c As Long
    With msflxgrd
        
        pcFoco.Height = .CellHeight
        pcFoco.Top = .CellTop + .Top
        
    End With
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtBoleto_GotFocus()
    P_SelecTexto txtBoleto
End Sub

Private Sub txtBoleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtBoleto_LostFocus()
    P_AjustaTexto txtBoleto
End Sub

Private Sub txtbxAcr_GotFocus()
    ganhaFoco txtbxAcr
End Sub

Private Sub txtbxAcr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxAcr_LostFocus()
    PerdeFoco txtbxAcr
    CalcValorFinal
End Sub

Private Sub txtbxDec_GotFocus()
    ganhaFoco txtbxDec
End Sub

Private Sub txtbxDec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDec_LostFocus()
    PerdeFoco txtbxDec
    CalcValorFinal
End Sub

Private Sub txtbxExpr_GotFocus()
    ganhaFoco txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
    If KeyCode = vbKeyDown Then
        
        msflxgrd.SetFocus
    End If
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim cTotal As Currency
    strSQL = "select td.*,ti.*,tc.nome_cli,((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as juros "
    strSQL = strSQL & "from " & TAB_CLIENTES & " as tc,"
    strSQL = strSQL & " " & TAB_DOC & "  as td,  " & TAB_ITDOC & "  as ti, tab_parametros as tp "
    strSQL = strSQL & " where ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " and td.cod_ter=tc.cod_Cli "
    strSQL = strSQL & " and ti.reman_itd >0 and td.excluido_doc = 'N' "
    strSQL = strSQL & " and tc.cod_cli=" & txtbxExpr.tag & ""
    strSQL = strSQL & " and tipo_doc='R'"
    If optnDoc.Value Then
        strSQL = strSQL & " order by num_doc"
    Else
        strSQL = strSQL & " order by venc_itd "
    End If
    abreObjetoAdo rPesq, strSQL
    With rPesq
        msflxgrd.rows = 1
        Do While Not .EOF
            strSQL = !seq_itd & Chr(9) & "" & Chr(9)
            strSQL = strSQL & !num_doc & Chr(9)
            strSQL = strSQL & !venc_itd & Chr(9)
            strSQL = strSQL & FormatVal(!reman_itd) & Chr(9)
            strSQL = strSQL & !parcela_itd & "/" & !parcelas_doc & Chr(9)
            strSQL = strSQL & FormatVal(!valor_itd) & Chr(9)
            strSQL = strSQL & !seq_doc & Chr(9)
            strSQL = strSQL & FormatVal(!juros) & Chr(9)
            msflxgrd.AddItem strSQL
            cTotal = cTotal + !reman_itd
            .MoveNext
        Loop
        lblTotal.caption = FormatVal(cTotal)
        lblParcList.caption = .RecordCount
        .Close
    End With
    msflxgrd_RowColChange
End Sub

Private Sub txtbxParcelas_GotFocus()
    ganhaFoco txtbxParcelas
End Sub

Private Sub txtbxParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxParcelas_LostFocus()
    PerdeFoco txtbxParcelas, "##0"
End Sub

Private Sub txtbxVenc_GotFocus()
    ganhaFoco txtbxVenc
End Sub

Private Sub txtbxVenc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxVenc_LostFocus()
    P_AjustaData txtbxVenc
End Sub

Private Sub txtGrid_GotFocus()
    P_SelecTexto txtGrid
End Sub

Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
    
End Sub

Private Sub txtGrid_LostFocus()
    gridParcelas.Text = txtGrid.Text
    txtGrid.Text = ""
    soma
    txtGrid.Visible = False
    
    gridParcelas.SetFocus
End Sub
