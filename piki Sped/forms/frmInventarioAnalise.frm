VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInventarioAnalise 
   Caption         =   "Análise de Inventário"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   12060
      TabIndex        =   9
      Top             =   5940
      Width           =   12090
      Begin VB.CommandButton export 
         Caption         =   "Exportar"
         Height          =   615
         Left            =   9720
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox pcComprou 
         Height          =   855
         Left            =   2280
         ScaleHeight     =   795
         ScaleWidth      =   1155
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSoma 
         Caption         =   "SOMA"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgrid 
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   2778
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12090
      TabIndex        =   0
      Top             =   0
      Width           =   12090
      Begin VB.TextBox txtAno 
         Height          =   285
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "busca"
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtInv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   12840
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbOrd 
         Height          =   315
         ItemData        =   "frmInventarioAnalise.frx":0000
         Left            =   2280
         List            =   "frmInventarioAnalise.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ano"
         Height          =   195
         Left            =   3840
         TabIndex        =   11
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventário:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   2280
         TabIndex        =   5
         Top             =   120
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmInventarioAnalise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim adoEntrada As New ADODB.Recordset
Dim adoSaida As New ADODB.Recordset
Sub busca2()
    Dim strSQL As String
    strSQL = "SELECT     inventario_it.codigo,"
    strSQL = strSQL & " inventario_it.nome,"
    strSQL = strSQL & " inventario_it.custo,"
    strSQL = strSQL & " inventario_it.q1,"
    strSQL = strSQL & " inventario_it.custo*q1 AS total "
    strSQL = strSQL & " ,(SELECT SUM(nfe_item.qCom)  from nfe_item LEFT JOIN nfe ON nfe_item.id_nfe=nfe.id "
    strSQL = strSQL & " where nfe.dhEmi>='2024-01-01' "
    strSQL = strSQL & " and nfe.dhEmi<='2024-12-31' "
    strSQL = strSQL & " AND nfe_item.cProd=inventario_it.codigo)as saida "
    strSQL = strSQL & " ,(SELECT SUM(sped_nf_itens.QTD)  from sped_nf_itens LEFT JOIN sped_nf ON sped_nf_itens.id_nf=sped_nf.id "
    strSQL = strSQL & " where sped_nf.DT_E_S >='2024-01-01' AND sped_nf.DT_E_S <='2024-12-31' and sped_nf_itens.COD_ITEM=inventario_it.codigo)"
    strSQL = strSQL & "  as entrada"
    strSQL = strSQL & " from"
    strSQL = strSQL & " inventario_it"
    strSQL = strSQL & " Where inventario = 33"
End Sub
Sub buscaEntrada()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select "
    strSQL = strSQL & " tp.cod_Pro, tp.nome_pro ,"
    strSQL = strSQL & " sum(i.qtd_ite) as qtd,sum(i.custo_ite*i.qtd_ite) as total"
    strSQL = strSQL & " from"
    strSQL = strSQL & " requinte.tab_produtos as tp inner join requinte.tab_itcompras as i on tp.cod_pro=i.cod_pro"
    strSQL = strSQL & " inner join requinte.tab_compras as n on n.seq_not=i.seq_not"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " date(n.emissao_not) between '" & txtAno.text & "-01-01' and '" & txtAno.text & "-12-31'"
    strSQL = strSQL & " group by tp.cod_pro,tp.nome_pro"
    strSQL = strSQL & " order by tp.cod_pro"
    FechaObjeto adoEntrada
    
    adoEntrada.Open strSQL, Conn_Client
     
     
     
     
End Sub
Sub buscaSaida()
    Dim strSQL As String
    strSQL = "select "
    strSQL = strSQL & " tp.cod_Pro, tp.nome_pro ,"
    strSQL = strSQL & " sum(i.qcom) as qtd,sum(i.qcom*i.vuncom) as total"
    strSQL = strSQL & " from    "
    strSQL = strSQL & " requinte.tab_produtos as tp inner join requinte.nfe_item as i on tp.cod_pro=i.cProd"
    strSQL = strSQL & " inner join requinte.nfe as n on n.id=i.id_nfe"
    strSQL = strSQL & " Where   "
    strSQL = strSQL & " date(left(n.dhEmi,10)) between '" & txtAno.text & "-01-01' and '" & txtAno.text & "-12-31'"
    strSQL = strSQL & " group by tp.cod_pro,tp.nome_pro"
    strSQL = strSQL & " order by tp.cod_pro "
    FechaObjeto adoSaida
     adoSaida.Open strSQL, Conn_Client
End Sub

Function encontrar_produto_ado(cod_pro, ado As ADODB.Recordset) As Integer
    Dim retorno As Integer
    retorno = -1 'não achou
    
    With ado
    If ado.AbsolutePosition <= -1 Then
        ado.MoveFirst
    End If
        If cod_pro = ado!cod_pro Then
            retorno = ado.AbsolutePosition
            GoTo retorno
        ElseIf cod_pro < ado!cod_pro Then
            Do While Not .BOF
                If cod_pro = ado!cod_pro Then
                    retorno = .AbsolutePosition
                    GoTo retorno
                End If
                
                .MovePrevious
            Loop
            
        ElseIf cod_pro > ado!cod_pro Then
            Do While Not .EOF
                If cod_pro = ado!cod_pro Then
                    retorno = .AbsolutePosition
                    GoTo retorno
                End If
                
                .MoveNext
            Loop
            
        End If
    End With
    
retorno:
    encontrar_produto_ado = retorno
    
End Function
Sub cruzaInfo()
    buscaEntrada
    buscaSaida
    Dim i As Integer
    Dim j As Integer
    Dim cod_pro As String
    Dim retorno As Single
    hgrid.Visible = False
    With hgrid
        For i = 1 To .rows - 1
            .row = i
            .col = 0
            cod_pro = .text
            retorno = encontrar_produto_ado(cod_pro, adoEntrada)
            hgrid.col = 6
            If retorno = -1 Then
                
                hgrid.text = "0"
            Else
                hgrid.text = adoEntrada!qtd
            End If
            
            .col = 0
            retorno = encontrar_produto_ado(cod_pro, adoSaida)
            hgrid.col = 7
            If retorno = -1 Then
                
                hgrid.text = "0"
            Else
                hgrid.text = adoSaida!qtd
            End If
            
        Next
    End With
    
    
    
    hgrid.Visible = True
    
End Sub
Sub ajustaGrid()
    GridUtil.ajustaGridTitulos hgrid, 10, "codigo,produto,QI,QE,CI,CE,entrada,saida"
End Sub

Sub busca(Optional ordem As String = "nome")
    ajustaGrid
    Dim strSQL As String
    Set ado = New ADODB.Recordset
    strSQL = "select p.cod_pro as codigo,p.nome_pro as produto,p.ef_pro, it.*,round(it.custo*it.q1,2) as total1,p.cf_pro  as cf "
    
    
    If ordem = "" Then
        If cmbOrd.ListIndex = 1 Then
        
        ElseIf cmbOrd.ListIndex = 2 Then
            ordem = "total1 desc"
         ElseIf cmbOrd.ListIndex = 3 Then
            ordem = "it.q1 desc, total1 desc"
         ElseIf cmbOrd.ListIndex = 4 Then
            ordem = "total1 desc,it.q1 desc"
        
        End If
    End If
    
    'If chkT2.Value = 1 Then
        strSQL = strSQL & ", custo*q2 as total2 "
'        grid.Columns(7).Visible = True
    'Else
    '    grid.Columns(7).Visible = False
    'End If
    strSQL = strSQL & " from inventario_it as it inner join inventario as i "
    strSQL = strSQL & " on i.id=it.inventario "
    'strSQL = strSQL & " left join inventario as i2 on i.id_ultinv=i2.id"
    'strSQL = strSQL & " left join inventario_it as it2 on i2.id=it2.inventario and it2.codigo=it.codigo"
    strSQL = strSQL & " inner join tab_produtos as p on p.cod_pro=it.codigo "
    
    strSQL = strSQL & " where it.inventario=" & txtInv.tag
    'If chkNeg.Value = 0 Then
        strSQL = strSQL & " and it.q1>0 "
    'End If
    
    If ordem <> "" Then
        strSQL = strSQL & " order by  p.cod_pro "
    End If
    
    abreObjetoAdo ado, strSQL
    hgrid.rows = 1
    hgrid.Visible = False
    With ado
        Do While Not .EOF
            strSQL = !codigo & Chr(9) & !produto & Chr(9) & !q1 & Chr(9) & !ef_pro & Chr(9) & !custo & Chr(9)
            strSQL = strSQL & !cF & Chr(9)
            hgrid.AddItem strSQL
            .MoveNext
        Loop
        
    End With
    hgrid.Visible = True
    
    cruzaInfo
    
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo trata_Erro
    Dim i As inventario.inventario
    Dim ord As String
    If cmbOrd.ListIndex = 1 Then
        ord = "nome_pro"
    ElseIf cmbOrd.ListIndex = 2 Then
    Else
    
    End If
    i = inventario.SelecionaInventario
    If i.id <> "" Then
        txtInv.text = i.data
        
        txtInv.tag = i.id
        'busca ord
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub Command1_Click()
    P_Status "carregando informações"
    busca
    P_Status "cruzando dados de entrada e saída"
    cruzaInfo
    P_Status
End Sub

Private Sub export_Click()
    ExportaMSFLXGridToCSV hgrid
End Sub

Private Sub Form_Load()
        Set ado = New ADODB.Recordset
        txtAno.text = Year(Date) - 1
End Sub

Private Sub Form_Resize()
    PosicionaGrade hgrid, Me, pcBot
End Sub
