VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelDocContaCob 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Documentos com Conta Cobrança"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   6315
      TabIndex        =   16
      Top             =   7560
      Width           =   6375
   End
   Begin VB.CommandButton cmndbtnRelocar 
      Caption         =   "Relocar Selecionados"
      Height          =   495
      Left            =   7680
      Picture         =   "frmRelDocContaCob.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnTodos 
      Caption         =   "Todos"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   10440
      Picture         =   "frmRelDocContaCob.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnImpr 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   8160
      Picture         =   "frmRelDocContaCob.frx":04F5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6615
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11668
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "Buscar"
      Height          =   615
      Left            =   6600
      Picture         =   "frmRelDocContaCob.frx":0902
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtbxDataIni 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtbxdataFin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtbxContaCob 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecionados:"
      Height          =   195
      Left            =   10080
      TabIndex        =   13
      Top             =   7560
      Width           =   1005
   End
   Begin VB.Label lblSelect 
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
      Left            =   11280
      TabIndex        =   12
      Top             =   7560
      Width           =   120
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listados"
      Height          =   195
      Left            =   10080
      TabIndex        =   11
      Top             =   7800
      Width           =   585
   End
   Begin VB.Label lblListados 
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
      Left            =   11280
      TabIndex        =   10
      Top             =   7800
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&De:"
      Height          =   195
      Left            =   3840
      TabIndex        =   2
      Tag             =   "!"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Até:"
      Height          =   195
      Left            =   5040
      TabIndex        =   4
      Tag             =   "!"
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conta Cobrança"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
   Begin VB.Menu mnOpt 
      Caption         =   "&Opções"
      Visible         =   0   'False
      Begin VB.Menu mnDetalhes 
         Caption         =   "Detalhes de Memória da Cobrança"
      End
   End
End
Attribute VB_Name = "frmRelDocContaCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBusca_Click()
    buscaCob
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnImpr_Click()
    P_MsflxToHtml "Posição de Documentos em Conta Cobrança", msflxgrd, "Conta: " & txtbxContaCob.Text, _
        ""
End Sub

Sub transfere()
    Dim strSQL As String
    Dim iCount As Integer
    Dim Anterior As String
    Dim sSeqItd As String
    Dim tipo As String
    If txtbxContaCob.tag = "" Then
        MsgBox "Escolha uma conta cobrança!", vbExclamation
        Exit Sub
    End If
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 7
            tipo = .Text
            .Col = 0
            If .Text = "S" Then
                .Col = 1
                doc.P_gravaDocHist .Text, txtbxContaCob.tag, "000", "2", "2"
                doc.P_gravaDocHist .Text, "000", txtbxContaCob.tag, "2", "2"
                
            End If
        Next
        MsgBox "Transferidos com sucesso!", vbInformation
    
    End With
    busca
    
    
End Sub

Private Sub cmndbtnRelocar_Click()
    transfere
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTodos_Click()
    If cmndbtnTodos.tag = "" Then
        'marcaTodos True
        cmndbtnTodos.tag = "T"
        cmndbtnTodos.caption = "&Nenhum"
    Else
        'marcaTodos False
        cmndbtnTodos.tag = ""
        cmndbtnTodos.caption = "&Todos"
    End If
End Sub

Private Sub Form_Load()
    ajustaGrid
End Sub
Sub ajustaGrid()
    With msflxgrd
        .Cols = 8
        .rows = 1
        .Col = 0: .Text = ""
        .Col = 1: .Text = "Id.Doc"
        .Col = 2: .Text = "Número"
        .Col = 3: .Text = "Tipo"
        .Col = 4: .Text = "Valor"
        .Col = 5: .Text = "Parcelas"
        .Col = 6: .Text = "Cliente"
        .Col = 7: .Text = "TipoCob"
        
        .ColWidth(0) = 200
        .ColWidth(6) = 3000
        
    End With
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sMsg As String
    msflxgrd.rows = 1
    strSQL = "select distinct td.*,th.data_doc as data_cob,tc.nome_cli from tab_clientes as tc, "
    strSQL = strSQL & " tab_doc as td,tab_itdoc as ti,tab_dochist as th "
    strSQL = strSQL & " where tc.cod_cli=td.cod_ter  and ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " and tipo_doc='R' "
    strSQL = strSQL & " and tipo_ter='C' "
    strSQL = strSQL & " and td.seq_doc=th.seq_doc "
    strSQL = strSQL & " and th.seq_doc=ti.seq_doc "
    
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and left(th.data_doc,10)>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and left(th.data_doc,10)<=" & F_GravaData(txtbxdataFin.Text)
    End If
    strSQL = strSQL & " and cod_cob='" & txtbxContaCob.tag & "' "
    strSQL = strSQL & " and cobin_doc='" & txtbxContaCob.tag & "' "
    strSQL = strSQL & " order by num_doc,th.data_doc "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rPesq
        lblListados.caption = .RecordCount
        Do While Not .EOF
            sMsg = Chr(9) & !seq_doc & Chr(9) & !num_doc & Chr(9) & !esp_Doc & Chr(9) & FormatVal(!valor_doc) & Chr(9) & !parcelas_doc & Chr(9) & !nome_Cli & Chr(9) & !data_cob
            msflxgrd.AddItem sMsg
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid msflxgrd, 1
End Sub

Private Sub mnDetalhes_Click()
    msflxgrd.Col = 1
    
    frmMemoriaCob.tag = msflxgrd.Text
    frmMemoriaCob.Show 1
End Sub

Private Sub msflxgrd_DblClick()
     With msflxgrd
        .Col = 0
        If .Text = "" Then
            
            ZebraLinha msflxgrd, msflxgrd.Row
            .Col = 0
            .CellBackColor = vbBlue
            .CellForeColor = vbBlue
            .Text = "S"
            lblSelect.caption = Int(lblSelect.caption) + 1
        Else
            .Text = ""
            ZebraLinha msflxgrd, msflxgrd.Row, vbWhite
            .CellBackColor = vbWhite
            lblSelect.caption = Int(lblSelect.caption) - 1
        End If
        
    End With
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
    
        Me.PopupMenu mnOpt
    End If
End Sub

Private Sub txtbxContaCob_GotFocus()
    P_SelecTexto txtbxContaCob
End Sub

Private Sub txtbxContaCob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCob
    End If
End Sub
Sub buscaCob()
    Dim c As funcionario.funcionario
    c = F_SelectCobrador(txtbxContaCob.Text)
    If c.cod_fun <> "" Then
        txtbxContaCob.tag = c.cod_fun
        txtbxContaCob.Text = c.nome_fun
        
    End If
    
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxDataFin_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxDataFin_LostFocus()
    P_AjustaTexto txtbxdataFin
End Sub

