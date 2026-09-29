VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVinculaDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vinculação de Documentos"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmVinculaDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   8880
      Picture         =   "frmVinculaDoc.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   7440
      Picture         =   "frmVinculaDoc.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   4695
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Parc."
         Height          =   195
         Left            =   4080
         TabIndex        =   28
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
      Picture         =   "frmVinculaDoc.frx":0B5E
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
      Width           =   9975
      _ExtentX        =   17595
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
      ScaleWidth      =   10260
      TabIndex        =   24
      Top             =   960
      Width           =   10260
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
      Caption         =   "&Utilitários"
      Visible         =   0   'False
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
    Dim c As Clientes.Cliente
    c = Clientes.F_BuscaCliente("C", txtbxExpr.Text)
    If c.cod_cli <> "" Then
        txtbxExpr.Text = c.nome_Cli
        txtbxExpr.tag = c.cod_cli
        busca
    
    End If
End Sub

Private Sub cmndbtnBusca_Click()
    busca
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
    Dim iCount As Integer
    Dim sNumDoc As String
    With msflxgrd
        sSeq_Vinc = MaxKeySQL("tab_vinc", "seq_vin", "", "000000#")
        sNumDoc = "V" & sSeq_Vinc
      '  sSEQ_DOC = F_GravaDoc("VINCULAÇÃO DE DOCUMENTOS", sNumDoc, _
            txtbxValorFinal.Text, Date, txtbxExpr.tag, "R", "", txtbxParcelas.Text, "", "", "")
        If sSEQ_DOC <> "" Then
            For iCount = 1 To Int(txtbxParcelas.Text)
                'strSQL = F_GravaItDoc(sSEQ_DOC, CStr(CCur(txtbxValorFinal.Text) / CCur(txtbxParcelas.Text)), _
                    F_CalculaVenc(txtbxVenc.Text, iCount, 30), txtd     iCount )
                If strSQL = "" Then
                    MsgBox "!erro", vbCritical
                    Exit Sub
                End If
            Next
        End If
        For iCount = 1 To .rows - 1
            
            .Row = iCount:
            .col = 1
            If .Text = "*" Then
                .col = 0
                sSeq_Vinc = MaxKeySQL("tab_vinc", "seq_vin", "", "000000#")
                strSQL = "update  " & TAB_ITDOC & "  set reman_itd=0 where seq_itd='" & .Text & "'"
                conn.Execute strSQL
            
                strSQL = "insert into tab_vinc(seq_vin,seq_itd,data_vin,"
                strSQL = strSQL & "valor_vin,acres_vin,decres_vin,venc_vin,"
                strSQL = strSQL & "parc_vin,seq_doc)values("
                strSQL = strSQL & "'" & sSeq_Vinc & "',"
                strSQL = strSQL & "'" & .Text & "',"
                strSQL = strSQL & F_GravaData(Date) & ","
                strSQL = strSQL & F_GravaDec(txtbxValorFinal.Text) & ","
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
    End With
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        
    End If
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

Private Sub mnDesvinc_Click()
    '
End Sub

Private Sub mnVinc_Click()
    '
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
    Dim rPesq As New ADODB.Recordset
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
