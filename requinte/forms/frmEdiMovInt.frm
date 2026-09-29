VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEdiMovInt 
   Caption         =   "Edição de Movimento Interno em Aberto"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13200
   Icon            =   "frmEdiMovInt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   13200
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13170
      TabIndex        =   9
      Top             =   5850
      Width           =   13200
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   11160
         Picture         =   "frmEdiMovInt.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1965
         TabIndex        =   13
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmEdiMovInt.frx":1275
         Left            =   120
         List            =   "frmEdiMovInt.frx":127F
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   225
         Width           =   1695
      End
      Begin VB.TextBox txtbxDest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6960
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdDestinos 
         Caption         =   "Destinos"
         Height          =   285
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnBuscaDest 
         Height          =   285
         Left            =   9480
         Picture         =   "frmEdiMovInt.frx":12A3
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   315
         Left            =   5880
         TabIndex        =   22
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   315
         Left            =   4800
         TabIndex        =   21
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Descrição:"
         Height          =   315
         Left            =   1965
         TabIndex        =   20
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Destino:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6960
         TabIndex        =   18
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   13200
      TabIndex        =   1
      Top             =   6585
      Width           =   13200
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11160
         Picture         =   "frmEdiMovInt.frx":162D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   7320
         Picture         =   "frmEdiMovInt.frx":1798
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   9120
         Picture         =   "frmEdiMovInt.frx":18E2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnPrint 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   5520
         Picture         =   "frmEdiMovInt.frx":1A2C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   1920
         Picture         =   "frmEdiMovInt.frx":1B76
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Apenas movimentos em aberto."
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Clique aki para listar todos os movimentos"
         Top             =   120
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Gerar Arq. Transf."
         Height          =   615
         Left            =   3720
         Picture         =   "frmEdiMovInt.frx":1CC0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9763
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmEdiMovInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
End Sub

Private Sub chkTodos_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmdDestinos_Click()
    P_CadDestMovInt
End Sub

Private Sub cmndbtnBuscaDest_Click()
    Dim m As DestinoMovInt
    m = F_BuscaDestMovInt(txtbxDest.text)
    If m.cod_Des <> "" Then
        txtbxDest.text = m.nome_des
        txtbxDest.tag = m.cod_Des
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    strSQL = "select tm.*,td.nome_des,tu.guerra_fun from tab_movint as tm,tab_funcionarios as tu, "
    strSQL = strSQL & " tab_destmovint as td "
    
    If chkTodos.Value = 1 Then
        strSQL = strSQL & " where situacao_mov=1 "
    Else
        strSQL = strSQL & " where situacao_mov<3 "
    End If
    strSQL = strSQL & " and td.cod_des=tm.cod_Des "
    strSQL = strSQL & "and tm.cod_usu=tu.cod_fun "
    If txtbxCliente.text <> "" Then
        strSQL = strSQL & " and desc_Mov like '%" & txtbxCliente.text & "%' "
        
    End If
    If txtbxDe.text <> "" Then
        strSQL = strSQL & " and data_mov>=" & F_GravaData(txtbxDe.text)
    End If
    If txtbxAte.text <> "" Then
        strSQL = strSQL & " and data_mov<=" & F_GravaData(txtbxAte.text)
    End If
    strSQL = strSQL & " order by num_mov desc "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        msflxgrd.rows = 1
        Do While Not .EOF
            strSQL = !num_mov & Chr(9)
            strSQL = strSQL & !tipo_mov & Chr(9)
            strSQL = strSQL & !nome_des & Chr(9)
            strSQL = strSQL & FormatVal(!valor_mov) & Chr(9)
               cTotal = cTotal + !valor_mov
               
            strSQL = strSQL & !data_mov & Chr(9)
            
            strSQL = strSQL & !guerra_fun & Chr(9)
            strSQL = strSQL & !situacao_mov & Chr(9)
            
            msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid msflxgrd, 1
    'lblTotal.caption = FormatVal(cTotal)
End Sub

Private Sub cmndbtnEdit_Click()
    If msflxgrd.row <= 0 Then
        Exit Sub
    End If
    With msflxgrd
        .col = 6
        If .text <> "1" Then
            MsgBox "Não é possível Editar um movimento fechado!", vbExclamation
            Exit Sub
        End If
    End With
    msflxgrd.col = 0
    Dim fr As New frmMovInt
    fr.tag = msflxgrd.text
    fr.Show 0
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    If msflxgrd.row <= 0 Then
        Exit Sub
    End If
    
    With msflxgrd
        .col = 6
        If .text <> "1" Then
            MsgBox "Não é possível excluir um movimento fechado!", vbExclamation
            Exit Sub
        End If
    End With
    
    
    
    msflxgrd.col = 0
    
    
    
    If MsgBox("Confirma Exclusão do Movimento Interno: " & msflxgrd.text & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "delete from tab_movint where num_mov='" & msflxgrd.text & "' "
    Conn.Execute strSQL
    strSQL = "delete from tab_itmov where num_mov='" & msflxgrd.text & "' "
    Conn.Execute strSQL
    cmndbtnBuscar_Click
End Sub

Private Sub cmndbtnNovo_Click()
    Dim fr As New frmMovInt
    fr.Show 0
End Sub

Private Sub cmndbtnPrint_Click()
    If msflxgrd.row <= 0 Then
        Exit Sub
    End If
    msflxgrd.col = 0
    ImprMovInt msflxgrd.text
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVendaAtacado_Click()
    Dim sCod_Cli As String
    With msflxgrd
        If .rows = 1 Then Exit Sub
        If strContrpre_usu <> "S" Then
            MsgBox "Você não tem permissão para venda em atacado!", vbCritical
            Exit Sub
        End If
        .col = 6
        If CDate(.text) < Date Then
            If MsgBox("O Orçamento está vencido! Deseja Continuar?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
        If MsgBox("Confirma Transferência do Orçamento para Venda em Atacado?", vbYesNo + vbQuestion) = vbYes Then
            .col = 1
            sCod_Cli = msflxgrd.text
            .col = 0
            transfereAtacado .text, sCod_Cli
        End If
    End With
End Sub
Sub transfereAtacado(num_orc As String, cod_cli As String)
    Dim strSQL As String
    Dim sNum_ven As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    
    sNum_ven = F_GeraNovaVenda(cod_cli)
    strSQL = "select * from tab_itorc where num_orc='" & num_orc & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
        strSQL = "insert into tab_itvendas(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite,final_ite)values("
        strSQL = strSQL & "'" & sNum_ven & "', "
        strSQL = strSQL & "'" & !cod_pro & "',"
        strSQL = strSQL & F_GravaDec(!unit_ite) & ","
        strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
        strSQL = strSQL & .AbsolutePosition + 1 & ","
    
        strSQL = strSQL & F_GravaDec(!unit_ite * !qtd_ite) & ") "
        Conn.Execute strSQL
        'RESERVANDO O ESTOQUE
        strSQL = "update tab_produtos set er_pro=er_pro - " & F_GravaDec(!qtd_ite)
        strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
        Conn.Execute strSQL
        
        
            
            .MoveNext
        Loop
        .Close
    End With
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    frmVenda.operacao = "B"
    frmVenda.tag = sNum_ven
    On Error Resume Next
    frmVenda.Atacado = True
    frmVenda.Show 0, mdiFrm
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnTransf_Click()
    msflxgrd.col = 6
    If msflxgrd.text = "1" Then
        MsgBox "Não pode gerar este movimento pois ele não se encontra disponível para transferência!", vbExclamation
        Exit Sub
    End If
    msflxgrd.col = 0
    P_GeraArquivoTransfMovInt msflxgrd.text, ""
End Sub

Private Sub Form_Activate()
    
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Load()
    resizeForm Me
    With msflxgrd
        .rows = 1
        .FixedCols = 0
        .Cols = 9
        .col = 0: .text = "Nº."
        .col = 1: .text = "Tipo:"
        .col = 2: .text = "Desc/Dest.:"
        .col = 3: .text = "Valor:"
        .col = 4: .text = "Data:"
        .col = 5: .text = "Usuário:"
        .col = 6: .text = "Status:"
        .col = 7: .text = "Protocolo:"
        .col = 8: .text = "Data Baixa:"
        .ColWidth(0) = 1000
        .ColWidth(1) = 1200
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1500
        
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    PosicionaGradeWidth msflxgrd, Me
    PosicionaGradeHeigth msflxgrd, pcBusca
    posicionaADireita cmndbtnSair, pcBot
    
End Sub

Private Sub msflxgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        msflxgrd_DblClick
    End If
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxCliente_GotFocus()
    P_SelecTexto txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim c As cliente
        c = F_BuscaCliente(txtbxCliente.text)
        If c.cod_cli <> "" Then
            txtbxCliente.text = c.nome_Cli
            txtbxCliente.tag = c.cod_cli
            Sendkeys Chr(9)
        Else
            txtbxCliente.text = ""
            txtbxCliente.tag = ""
            txtbxCliente.SetFocus
        End If
    End If
End Sub

Private Sub txtbxCliente_LostFocus()
    P_AjustaTexto txtbxCliente
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub

