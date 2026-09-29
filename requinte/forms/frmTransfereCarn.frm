VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTransfereCarn 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Posição de Carnês na Cobrança"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pcInfo 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   11625
      TabIndex        =   18
      Top             =   6180
      Width           =   11655
      Begin VB.TextBox txtbxContaTra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmndbtnContaTra 
         Height          =   285
         Left            =   3360
         Picture         =   "frmTransfereCarn.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblQtdSelec 
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
         Left            =   3720
         TabIndex        =   23
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Carnês:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3720
         TabIndex        =   22
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Para Transferência:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1860
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11625
      TabIndex        =   12
      Top             =   6915
      Width           =   11655
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "Transf Carnês"
         Height          =   615
         Left            =   120
         Picture         =   "frmTransfereCarn.frx":00F9
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEmiteCarne 
         Caption         =   "Emitir Carnê"
         Height          =   615
         Left            =   7320
         Picture         =   "frmTransfereCarn.frx":0255
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "Baixar Parcelas"
         Height          =   615
         Left            =   5520
         Picture         =   "frmTransfereCarn.frx":0351
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   3720
         Picture         =   "frmTransfereCarn.frx":07F9
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   9120
         Picture         =   "frmTransfereCarn.frx":0C06
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Seguir Parcela"
         Height          =   615
         Left            =   1920
         Picture         =   "frmTransfereCarn.frx":0D71
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5175
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9128
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11625
      TabIndex        =   10
      Top             =   0
      Width           =   11655
      Begin VB.ComboBox cmbxTipoItd 
         Height          =   315
         ItemData        =   "frmTransfereCarn.frx":12FB
         Left            =   5280
         List            =   "frmTransfereCarn.frx":1302
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   285
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da Parcela:"
         Height          =   195
         Left            =   5280
         TabIndex        =   7
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança:"
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   840
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Height          =   6855
      Left            =   0
      OleObjectBlob   =   "frmTransfereCarn.frx":1313
      TabIndex        =   25
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmTransfereCarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxTipoItd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnCliente_Click()
    Dim c As ContaCob
    c = Cobranca.F_BuscaContaCob(txtbxContaCob.text)
    If c.cod_cob <> "" Then
        txtbxContaCob.text = c.nome_cob
        txtbxContaCob.tag = c.cod_cob
        busca
    End If
End Sub

Private Sub cmndbtnContaTra_Click()
    Dim c As ContaCob
    c = Cobranca.F_BuscaContaCob(txtbxContaTra.text)
    If c.cod_cob <> "" Then
        txtbxContaTra.text = c.nome_cob
        txtbxContaTra.tag = c.cod_cob
    End If

End Sub

Private Sub cmndbtnEmiteCarne_Click()
    With msflxgrd
        .col = 2
        If .row > 0 And .text <> "" Then
            imprimeCarne .text
        End If
    End With
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTransf_Click()
    'Permissao_usu.cob_transferir_carne = True
    'If Permissao_usu.cob_transferir_carne = False Then
    '    MsgBox "Você não tem permissão para este acesso!", vbExclamation
    '    Exit Sub
    'End If
    If txtbxContaTra.tag = "" Then
        MsgBox "Escolha a conta para transferência!", vbExclamation
        txtbxContaTra.SetFocus
        Exit Sub
        
    End If
    If MsgBox("Confirma transferência dos Carnês para a conta de cobrança " & _
        txtbxContaTra.text & "?", vbYesNo + vbQuestion) = vbYes Then
    End If
    If Int(lblQtdSelec.caption) < 1 Then
        MsgBox "Não há Carnês selecionados!", vbInformation
    End If
    transfere
End Sub
Sub transfere()
    Dim strSQL As String
    Dim sSEQ_DOC As String
    Dim sSeq_ite As String
    
    Dim iCount As Integer
    With msflxgrd
        For iCount = 1 To .rows - 1
            .col = 0
            If .text <> "" Then
                .col = 1
                sSeq_ite = .text
                .row = iCount
                strSQL = Cobranca.F_gravaCobranca(sSeq_ite, txtbxContaTra.tag)
            End If
        Next
    End With
    MsgBox "Carnês Transferidos com sucesso!", vbInformation
    busca
End Sub
Private Sub Form_Load()
    resizeForm Me
    AjustaGrid
End Sub
Sub AjustaGrid()
    With msflxgrd
        .FixedCols = 1
        .FixedRows = 1
        .rows = 1
        .Cols = 10
        .col = 0: .text = ""
        .col = 1: .text = "seqitd"
        .col = 2: .text = "seqdoc"
        .col = 3: .text = "Cód.Cli"
        .col = 4: .text = "Cliente"
        .col = 5: .text = "Vencimento"
        .col = 6: .text = "Valor"
        .col = 7: .text = "Parc."
        .col = 8: .text = "Conta Cob."
        .col = 9: .text = "T"
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(4) = 3000
        .ColWidth(7) = 800
        .ColWidth(9) = 1200
    End With
End Sub
Private Sub Form_Resize()
    msflxgrd.width = pcTop.width
    msflxgrd.height = pcInfo.top - msflxgrd.top - 100
    
    
End Sub

Sub selecionaLinha()
    With msflxgrd
        .col = 0
        If .text = "" Then
            .CellFontName = "verdana"
            .CellFontSize = 7
            .CellFontBold = True
            .text = "Selecionado"
            .CellBackColor = vbYellow
            lblQtdSelec.caption = Int(lblQtdSelec.caption) + 1
            
        Else
            .text = ""
            .CellBackColor = vbWhite
            lblQtdSelec.caption = Int(lblQtdSelec.caption) - 1
        End If
    End With
End Sub

Private Sub msflxgrd_DblClick()
    selecionaLinha
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
End Sub

Private Sub txtbxContaCob_GotFocus()
    P_SelecTexto txtbxContaCob
End Sub

Private Sub txtbxContaCob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub

Private Sub txtbxContaCob_LostFocus()
    P_AjustaTexto txtbxContaCob
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDataFin
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select ti.*,tc.cod_cli,td.parcelas_doc,tc.nome_cli,td.data_doc,td.valor_doc, td.esp_doc,tcb.cod_cob,tcb.nome_cob "
    strSQL = strSQL & " from tab_clientes as tc, tab_doc as td, tab_itdoc as ti,tab_contacob as tcb "
    strSQL = strSQL & " where tc.cod_cli=td.cod_ter "
    strSQL = strSQL & " and ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " and ti.cod_cob=tcb.cod_cob "
    strSQL = strSQL & " and td.tipo_doc='R' "
    'strsql = strsql & " and
    If txtbxContaCob.tag <> "" Then
        strSQL = strSQL & " and ti.cod_cob='" & txtbxContaCob.tag & "' "
    End If
    If cmbxTipoItd.ListIndex > 0 Then
        strSQL = strSQL & " and tipo_itd='" & Left(cmbxTipoItd.text, 1) & "' "
    End If
    If txtbxDataFin.text <> "" Then
        strSQL = strSQL & " and venc_itd<=" & F_GravaData(txtbxDataFin.text)
    End If
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and venc_itd>=" & F_GravaData(txtbxDataIni.text)
    End If
    strSQL = strSQL & " and reman_itd>0"
    abreObjetoAdo rPesq, strSQL
    msflxgrd.rows = 1
    
    With rPesq
        Do While Not .EOF
            strSQL = "" & Chr(9)
            strSQL = strSQL & !seq_itd & Chr(9)
            strSQL = strSQL & !seq_Doc & Chr(9)
            strSQL = strSQL & !cod_cli & Chr(9)
            strSQL = strSQL & !nome_Cli & Chr(9)
            strSQL = strSQL & !venc_itd & Chr(9)
            strSQL = strSQL & !reman_itd & Chr(9)
            strSQL = strSQL & !parcela_itd & "/" & !parcelas_doc & Chr(9)
            strSQL = strSQL & !nome_cob & Chr(9)
            strSQL = strSQL & !operacao_itd
            msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid msflxgrd, 1
     
End Sub
Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
