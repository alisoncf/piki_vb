VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAjustaEstFis 
   Caption         =   "Avaliação de Variação de Estoque"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   12990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSelecionado 
      Caption         =   "Ajustar Selecionado"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnTodos 
      Caption         =   "Ajustar Todos"
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox chkEst 
      Caption         =   "Apenas com Estoque"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NF Sem Cupom"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Venda"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compra"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   7560
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   11880
      _Version        =   393216
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cmbxInvent 
      BackColor       =   &H80000014&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Inventário de:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmAjustaEstFis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    ajustaGrid
    If cmbxInvent.Text = "" Then
        MsgBox "informe o inventário", vbExclamation
        cmbxInvent.SetFocus
        Exit Sub
        
    End If
    
    
    strSQL = "select tp.nome_pro,tp.ef_pro as e,ti.*,ti.ef_pro as ef from tab_invent as ti, " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and datahora_inv=" & F_GravaDataTempo(cmbxInvent.Text)
    If chkEst.Value = 1 Then
        strSQL = strSQL & " and ti.estoque_pro<>0 and ti.ef_pro <>0"
    End If
    strSQL = strSQL & " order by ti.cod_pro "
    rPesq.Open strSQL, conn_client
    
    With rPesq
        Do While Not .EOF
            strSQL = "" & Chr(9) & !cod_pro & Chr(9)
            strSQL = strSQL & !nome_pro & Chr(9)
            strSQL = strSQL & !ef & Chr(9)
            strSQL = strSQL & !e & Chr(9)
            
            msflxgrd.AddItem strSQL
            
            .MoveNext
            P_Status .AbsolutePosition + 1 & "/" & .RecordCount
        Loop
        .Close
    End With
    buscaCompra
    buscaSaidas
    buscaSaidasSemCupom
    P_Status
    
End Sub

Private Sub cmndbtnSelecionado_Click()
    muda_est msflxgrd.Row, msflxgrd.Row
End Sub

Private Sub cmndbtnTodos_Click()
    
    
    muda_est 1, msflxgrd.rows - 1
End Sub
Sub muda_est(ini As Integer, fim As Integer)
    Dim iCount As Integer
    Dim cod As String
    Dim qtd As String
    Dim a As Integer
    With msflxgrd
        For iCount = ini To fim
            .Row = iCount
            .col = 1: cod = .Text
            .col = 8
            qtd = .Text
            strSQL = "update " & TAB_PRODUTOS & " set ef_pro=" & F_GravaDec(qtd)
            strSQL = strSQL & " where cod_pro='" & cod & "'"
            conn.Execute strSQL, a
            P_Status ini & "/" & fim
        Next
        P_Status ""
        MsgBox "Modificado"
    End With
End Sub
Private Sub Command1_Click()
    buscaCompra
End Sub

Private Sub Command2_Click()
    buscaSaidas
End Sub

Private Sub Command3_Click()
    buscaSaidasSemCupom
End Sub

Private Sub Form_Load()
    P_CarregaInvent cmbxInvent
End Sub
Sub ajustaGrid()
    With msflxgrd
        .FixedCols = 0
        .rows = 1
        .Cols = 9
        .col = 0: .Text = "..."
        .col = 1: .Text = "Código"
        .col = 2: .Text = "Produto"
        .col = 3: .Text = "Inventário"
        .col = 4: .Text = "Est. Atual"
        .col = 5: .Text = "Comprado"
        .col = 6: .Text = "Vendido"
        .col = 7: .Text = "NF S/Cupom"
        .col = 8: .Text = "Sujerido"
        .ColWidth(2) = 2500
        
    End With
End Sub
Sub buscaCompra()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCod As String
    Dim i As Integer
    strSQL = "select cod_pro,sum(qtd_ite*qtd_emb) as total from tab_itcompras as ti, "
    strSQL = strSQL & " tab_compras as tc where tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and entrada_not>=" & F_GravaData(cmbxInvent.Text)
    strSQL = strSQL & " group by ti.cod_pro "
    strSQL = strSQL & " order by ti.cod_pro "
    rPesq.Open strSQL, conn_client
    With rPesq
        i = 1
        Do While Not .EOF
            P_Status "Achando compra " & .AbsolutePosition + 1 & "/" & .RecordCount
            i = achaGrid(!cod_pro, !total, 5, i)
            
            .MoveNext
        Loop
        .Close
    End With
    P_Status ""
End Sub
Sub buscaSaidas()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCod As String
    Dim i As Integer
    strSQL = "select cod_pro,sum(qtd_ite) as total from tab_itcup as ti, "
    strSQL = strSQL & " cupons as tc where tc.seq_cup=ti.seq_cup "
    strSQL = strSQL & " and data_cup>=" & F_GravaDataTempo(cmbxInvent.Text)
    strSQL = strSQL & " group by ti.cod_pro "
    strSQL = strSQL & " order by ti.cod_pro "
    rPesq.Open strSQL, conn_client
    With rPesq
        i = 1
        Do While Not .EOF
            P_Status "Achando Venda " & .AbsolutePosition + 1 & "/" & .RecordCount
            i = achaGrid(!cod_pro, !total, 6, i)
            
            .MoveNext
        Loop
        .Close
    End With
    P_Status ""
End Sub
Sub soma()
    Dim iCount As Integer
    Dim cSoma As Currency
    Dim cEst As Currency
    Dim cCup As Currency
    Dim cSai As Currency
    Dim cComp As Currency
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 3: cEst = CCur(.Text)
            
            .col = 5: If .Text = "" Then .Text = "0"
            cComp = CCur(.Text)
            .col = 6: If .Text = "" Then .Text = "0"
            cCup = CCur(.Text)
            .col = 7: If .Text = "" Then .Text = "0"
            cSai = CCur(.Text)
            .col = 8: .Text = cEst + cComp - cCup - cSai
            P_Status "Somando " & iCount & "/" & .rows - 1
        Next
    End With
    P_Status ""
End Sub
Sub buscaSaidasSemCupom()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCod As String
    Dim i As Integer
    strSQL = "select cod_pro,sum(qtd_its) as total from tab_itsainf as ti, "
    strSQL = strSQL & " tab_saidasnf as tc where tc.seq_sai=ti.seq_sai "
    strSQL = strSQL & " and emissao_sai>=" & F_GravaData(cmbxInvent.Text)
    strSQL = strSQL & " and seq_cup=''"
    strSQL = strSQL & " group by ti.cod_pro "
    strSQL = strSQL & " order by ti.cod_pro "
    rPesq.Open strSQL, conn_client
    With rPesq
        i = 1
        Do While Not .EOF
            P_Status "Achando Venda " & .AbsolutePosition + 1 & "/" & .RecordCount
            i = achaGrid(!cod_pro, !total, 7, i)
            
            .MoveNext
        Loop
        .Close
    End With
    P_Status ""
    soma
End Sub

Function achaGrid(cod As String, qtd As String, coluna As Integer, Optional ultimabusca As Integer = 1) As Integer
    Dim iCount As Integer
    With msflxgrd
        For iCount = ultimabusca To .rows - 1
            achaGrid = iCount
            .Row = iCount
            .col = 1
            If .Text = cod Then
                .col = coluna
                .Text = qtd
                Exit For
            End If
            If .Text > cod Then
                Exit For
            End If
            
        Next
    End With
End Function
