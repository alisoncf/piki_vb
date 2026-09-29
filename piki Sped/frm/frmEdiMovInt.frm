VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEdiMovInt 
   Caption         =   "Edição de Movimento Interno em Aberto"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   Icon            =   "frmEdiMovInt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   11925
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11895
      TabIndex        =   22
      Top             =   0
      Width           =   11925
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   7200
         Picture         =   "frmEdiMovInt.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2085
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmEdiMovInt.frx":1075
         Left            =   120
         List            =   "frmEdiMovInt.frx":107F
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   345
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   6000
         TabIndex        =   6
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   4920
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Descrição:"
         Height          =   195
         Left            =   2040
         TabIndex        =   2
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   11925
      TabIndex        =   17
      Top             =   6915
      Width           =   11925
      Begin VB.CommandButton cmndbtnImprVenda 
         Caption         =   "Impr. Venda"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4920
         Picture         =   "frmEdiMovInt.frx":10A3
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   6480
         Picture         =   "frmEdiMovInt.frx":11ED
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   10320
         Picture         =   "frmEdiMovInt.frx":1337
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   8040
         Picture         =   "frmEdiMovInt.frx":14A2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnPrint 
         Caption         =   "&Impr. Custo"
         Height          =   615
         Left            =   3360
         Picture         =   "frmEdiMovInt.frx":15EC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiMovInt.frx":1736
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Apenas movimentos em aberto."
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Clique aki para listar todos os movimentos"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total...:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   960
         TabIndex        =   20
         Top             =   0
         Width           =   315
      End
      Begin VB.Label lblTotalCusto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   4680
         TabIndex        =   19
         Top             =   0
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Custo...:"
         Height          =   195
         Left            =   3480
         TabIndex        =   18
         Top             =   0
         Width           =   990
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5055
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8916
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

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    On Error GoTo trata_Erro
    strSQL = "select num_mov,desc_mov,valor_mov,tipo_mov,situacao_mov,data_mov,tu.guerra_fun from tab_movint as tm,tab_funcionarios as tu "
    
    If chkTodos.Value = 1 Then
        strSQL = strSQL & " where situacao_mov=1 "
    Else
        strSQL = strSQL & " where situacao_mov<3 "
    End If
    strSQL = strSQL & "and tm.cod_usu=tu.cod_fun "
    If txtbxCliente.Text <> "" Then
        strSQL = strSQL & " and desc_Mov like '%" & txtbxCliente.Text & "%' "
        
    End If
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and data_mov>=" & F_GravaData(txtbxDe.Text)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and data_mov<=" & F_GravaData(txtbxAte.Text)
    End If
    strSQL = strSQL & " order by num_mov desc "
    rPesq.Open strSQL, conn_client
    With rPesq
        msflxgrd.rows = 1
        Do While Not .EOF
            strSQL = !num_mov & Chr(9)
            strSQL = strSQL & !tipo_mov & Chr(9)
            strSQL = strSQL & !Desc_mov & Chr(9)
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
    lblTotal.caption = FormatVal(cTotal)
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmndbtnEdit_Click()
    If msflxgrd.row <= 0 Then
        Exit Sub
    End If
    With msflxgrd
        .col = 6
        If .Text <> "1" Then
            MsgBox "Não é possível Editar um movimento fechado!", vbExclamation
            Exit Sub
        End If
    End With
    msflxgrd.col = 0
    Dim fr As New frmMovInt
    fr.tag = msflxgrd.Text
    fr.Show 0
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    If msflxgrd.row <= 0 Then
        Exit Sub
    End If
    
    With msflxgrd
        .col = 6
        If .Text <> "1" Then
            MsgBox "Não é possível excluir um movimento fechado!", vbExclamation
            Exit Sub
        End If
    End With
    
    
    
    msflxgrd.col = 0
    
    
    
    If MsgBox("Confirma Exclusão do Movimento Interno: " & msflxgrd.Text & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "delete from tab_movint where num_mov='" & msflxgrd.Text & "' "
    conn.Execute strSQL
    strSQL = "delete from tab_itmov where num_mov='" & msflxgrd.Text & "' "
    conn.Execute strSQL
    cmndbtnBuscar_Click
End Sub

Private Sub cmndbtnImprVenda_Click()
    ImprMovInt Me.tag
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
    ImprMovInt msflxgrd.Text
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
        If CDate(.Text) < Date Then
            If MsgBox("O Orçamento está vencido! Deseja Continuar?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
        If MsgBox("Confirma Transferência do Orçamento para Venda em Atacado?", vbYesNo + vbQuestion) = vbYes Then
            .col = 1
            sCod_Cli = msflxgrd.Text
            .col = 0
            transfereAtacado .Text, sCod_Cli
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
    
    sNum_ven = F_GeraNovaVenda(cod_cli, strCod_Usu)
    strSQL = "select * from tab_itorc where num_orc='" & num_orc & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
        strSQL = "insert into " & TAB_ITVENDAS & "(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite,final_ite)values("
        strSQL = strSQL & "'" & sNum_ven & "', "
        strSQL = strSQL & "'" & !cod_pro & "',"
        strSQL = strSQL & F_GravaDec(!unit_ite) & ","
        strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
        strSQL = strSQL & .AbsolutePosition + 1 & ","
    
        strSQL = strSQL & F_GravaDec(!unit_ite * !qtd_ite) & ") "
        conn.Execute strSQL
        'RESERVANDO O ESTOQUE
        strSQL = "update " & TAB_PRODUTOS & " set er_pro=er_pro - " & F_GravaDec(!qtd_ite)
        strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
        conn.Execute strSQL
        
        
            
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
    MostraErro
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Load()
    resizeForm Me
    With msflxgrd
        .rows = 1
        .FixedCols = 0
        .Cols = 7
        .col = 0: .Text = "Nº."
        .col = 1: .Text = "Tipo:"
        .col = 2: .Text = "Desc/Dest.:"
        .col = 3: .Text = "Valor:"
        .col = 4: .Text = "Data:"
        .col = 5: .Text = "Usuário:"
        .col = 6: .Text = "Status:"
        .ColWidth(0) = 1000
        .ColWidth(1) = 1200
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 2500
        
    End With
End Sub

Private Sub Form_Resize()
    posicionaADireita cmndbtnSair, pcBot
    PosicionaGrade msflxgrd, Me, pcBot
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
    ganhaFoco txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxCliente_GotFocus()
    ganhaFoco txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim c As cliente
        c = F_BuscaCliente(txtbxCliente.Text)
        If c.cod_cli <> "" Then
            txtbxCliente.Text = c.nome_Cli
            txtbxCliente.tag = c.cod_cli
            TrocaTabKey KeyCode
        Else
            txtbxCliente.Text = ""
            txtbxCliente.tag = ""
            txtbxCliente.SetFocus
        End If
    End If
End Sub

Private Sub txtbxCliente_LostFocus()
    AjustaTexto txtbxCliente
End Sub

Private Sub txtbxDe_GotFocus()
    ganhaFoco txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub

