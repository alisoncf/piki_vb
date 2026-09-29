VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelComissGeral 
   Caption         =   "Comissões de Funcionários"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14220
   Icon            =   "frmRelComissGeral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   14190
      TabIndex        =   13
      Top             =   8310
      Width           =   14220
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2295
      Left            =   0
      TabIndex        =   12
      Top             =   1200
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   4048
      _Version        =   393216
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
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   14220
      TabIndex        =   10
      Top             =   0
      Width           =   14220
      Begin VB.OptionButton optPivot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Em Pivot"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   16
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optGeral 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Geral"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cmbTipoComissao 
         Height          =   315
         ItemData        =   "frmRelComissGeral.frx":09AA
         Left            =   2640
         List            =   "frmRelComissGeral.frx":09BA
         TabIndex        =   14
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar"
         Height          =   735
         Left            =   9840
         Picture         =   "frmRelComissGeral.frx":09F6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   5760
         Picture         =   "frmRelComissGeral.frx":0AE0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   7800
         Picture         =   "frmRelComissGeral.frx":0E6A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   12000
         Picture         =   "frmRelComissGeral.frx":1277
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdFun 
         Height          =   285
         Left            =   4920
         Picture         =   "frmRelComissGeral.frx":136C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   0
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmRelComissGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmdFun_Click()
    buscaFun
End Sub
Sub pivot(rPesq As ADODB.Recordset)
    Dim codigo As String
    Dim tipo As String
    Dim i As Integer
    Dim iRow As Integer
    Dim total As Currency
    Dim j As Integer
    With grid
        .rows = 2
        .Cols = 6
        .row = 0
        .col = 0: .Text = "Código"
        .col = 1: .Text = "Nome"
        .col = 2: .Text = "Vendas"
        .col = 3: .Text = "Dif. Troca"
        .col = 4: .Text = "Dif. Troca Avulsa"
        .col = 5: .Text = "Total"
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .row = 1
        For i = 1 To .Cols - 1
            .col = i
            .Text = ""
        Next
        
        
        
    End With
    
    With rPesq
        Do While Not .EOF
            If !cod_fun <> codigo Then
                If grid.rows = 2 And codigo = "" Then
                    
                    grid.row = 1
                    
                    
                    grid.col = 0: grid.Text = !cod_fun
                Else
                    iRow = GridUtil.getRowByText(grid, !cod_fun, 0)
                    If iRow = -1 Then
                        grid.AddItem !cod_fun
                        grid.row = grid.rows - 1
                    Else
                        grid.row = iRow
                    End If
                End If
                grid.col = 1: grid.Text = !nome_fun
            Else
            End If
                If !tipo = "Vendas" Then
                    grid.col = 2
                ElseIf !tipo = "Diferença de Troca" Then
                    grid.col = 3
                Else
                    grid.col = 4
                End If
                
                grid.Text = FormatVal(!total)
                codigo = !cod_fun
            
            .MoveNext
        Loop
    End With
    With grid
        For j = 1 To .rows - 1
            
            total = 0
            .row = j
            For i = 2 To .Cols - 2
                .col = i
                If .Text = "" Then .Text = "0,00"
                
            
                total = total + CCur(.Text)
            
            Next
            .col = .Cols - 1
            .Text = FormatVal(total)
        Next
    End With
End Sub

Private Sub cmndbtnBuscar_Click()
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim comissao As Single
    Dim SQLTroca As String
    Dim SQLTrocaAvulsa As String
    Dim SQLVenda As String
    
    
    strSQL = "(select 'Vendas' as tipo, tca.nome_car,sum(tv.valor_ven) as total,"
    strSQL = strSQL & " (sum(valor_ven)*" & F_GravaDec(crcComissVen) & "/100) as comiss,"
        strSQL = strSQL & " tf.cod_fun,tf.guerra_fun,tf.nome_fun from tab_vendas as tv  "
        strSQL = strSQL & " inner join tab_funcionarios as tf on tv.abrevenda_usu=tf.cod_fun "
        strSQL = strSQL & " inner join tab_cargos as tca on tca.cod_car=tf.cod_car "
        
        strSQL = strSQL & " where 1>0 "
        
        If txtDe.Text <> "" Then
            strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtDe.Text)
        End If
        If txtAte.Text <> "" Then
            strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtAte.Text)
        End If
        If txtFun.tag <> "" Then
            strSQL = strSQL & " and abrevenda_usu='" & txtFun.tag & "' "
        End If
        strSQL = strSQL & " and tipo_operacao=1"
        
        strSQL = strSQL & " and situacao_ven=4 "
       
            strSQL = strSQL & " group by tf.cod_fun order by tf.guerra_fun,tipo)"
        '    dbgrd.Columns(2).Visible = True
        
        SQLVenda = strSQL
        
        strSQL = strSQL & " union "
        SQLTroca = SQLTroca & "(select  'Diferença de Troca' as tipo,tca.nome_car,sum(tt.acres_tro-tt.abat_tro-desconto_tro) as total,"
    SQLTroca = SQLTroca & " (sum(tt.acres_tro-tt.abat_tro-desconto_tro)*" & F_GravaDec(crcComissVen) & "/100) as comiss,"
    SQLTroca = SQLTroca & " tf.cod_fun,tf.guerra_fun,tf.nome_fun from tab_vendas as tv "
    SQLTroca = SQLTroca & " inner join tab_trocas as tt on tt.num_ven=tv.num_ven "
    SQLTroca = SQLTroca & " inner join tab_funcionarios as tf on tt.cod_usu=tf.cod_fun "
    SQLTroca = SQLTroca & " inner join tab_cargos as tca on tca.cod_car=tf.cod_car"
    SQLTroca = SQLTroca & " where 1 > 0 "
    
    If txtDe.Text <> "" Then
        SQLTroca = SQLTroca & " and data_tro>=" & F_GravaDataTempo(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        SQLTroca = SQLTroca & " and data_tro<=" & F_GravaDataTempo(txtAte.Text & " 23:59:59")
    End If
    
    If txtFun.Text <> "" Then
        SQLTroca = SQLTroca & " and tf.cod_fun='" & txtFun.tag & "' "
    End If
    
    
    SQLTroca = SQLTroca & " and situacao_ven=4 "
    SQLTroca = SQLTroca & " group by tf.cod_fun order by tf.guerra_fun, tipo) "
    
    strSQL = strSQL & SQLTroca
    
    strSQL = strSQL & " union "
        SQLTrocaAvulsa = SQLTrocaAvulsa & "(select  'Diferença de Troca Avulsa' as tipo,tca.nome_car,sum(tt.valor_out-tt.valor_in-desconto) as total,"
    SQLTrocaAvulsa = SQLTrocaAvulsa & " (sum(tt.valor_out-tt.valor_in-desconto)*" & F_GravaDec(crcComissVen) & "/100) as comiss,"
    SQLTrocaAvulsa = SQLTrocaAvulsa & " tf.cod_fun,tf.guerra_fun,tf.nome_fun from tab_trocaavulsa as tt "
    SQLTrocaAvulsa = SQLTrocaAvulsa & " inner join tab_funcionarios as tf on tt.vendedor=tf.cod_fun "
    SQLTrocaAvulsa = SQLTrocaAvulsa & " inner join tab_cargos as tca on tca.cod_car=tf.cod_car"
    SQLTrocaAvulsa = SQLTrocaAvulsa & " where tt.valor_out-tt.valor_in-desconto "
    
    If txtDe.Text <> "" Then
        SQLTrocaAvulsa = SQLTrocaAvulsa & " and data>=" & F_GravaDataTempo(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        SQLTrocaAvulsa = SQLTrocaAvulsa & " and data<=" & F_GravaDataTempo(txtAte.Text & " 23:59:59")
    End If
    
    If txtFun.Text <> "" Then
        SQLTrocaAvulsa = SQLTrocaAvulsa & " and tf.cod_fun='" & txtFun.tag & "' "
    End If
    
    
    
    SQLTrocaAvulsa = SQLTrocaAvulsa & " group by tf.cod_fun order by tf.guerra_fun, tipo) "
        
        strSQL = strSQL & SQLTrocaAvulsa
        
        
        If cmbTipoComissao.ListIndex = 1 Then
            strSQL = SQLVenda
        ElseIf cmbTipoComissao.ListIndex = 2 Then
            strSQL = SQLTroca
        ElseIf cmbTipoComissao.ListIndex = 3 Then
            strSQL = SQLTrocaAvulsa
        Else
        
        
        End If
        
        abreObjetoAdo rPesq, strSQL
        If optGeral.Value = True Then
            AjustaGridData grid, 5, 1, 1, "Tipo,Código,Funcionario,Valor Total,Comissão", _
            "tipo,cod_fun,guerra_fun,total,comiss", "2500,1000,3500,1200,1200", rPesq, "3,4", True
        
        Else
        
            pivot rPesq
        End If
        rPesq.Close
    
End Sub

Private Sub cmndbtnImprimir_Click()
    P_MsflxToHtml "Comissões " & txtDe.Text & " a " & txtAte.Text, grid, txtFun.Text, "", "3,4"
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    ExportaMSFLXGridToCSV grid
End Sub

Private Sub Form_Load()
    txtDe.Text = F_PrimDiaMes
    txtAte.Text = F_UltDiaMes
    Set ado = New ADODB.Recordset
    txtFun.Text = strNome_Usu
    txtFun.tag = strCod_Usu
    
    If Int(strNivel_Usu) >= nivel_usu.Gerente Then
        txtFun.Enabled = True
        cmdFun.Enabled = True
    Else
        txtFun.Enabled = False
        cmdFun.Enabled = False
    End If
    
    
    
End Sub
Sub buscaFun()
    Dim f As funcionario.funcionario
    
    f = funcionario.F_SelectFuncionario(txtFun.Text)
    If f.cod_fun <> "" Then
        
        txtFun.Text = f.guerra_fun
        txtFun.tag = f.cod_fun
        cmndbtnBuscar_Click
    End If
    
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub

Private Sub grid_DblClick()
    If optGeral.Value = True Then Exit Sub
    
    If optPivot.Value Then
        grid.col = 1
        txtFun.Text = grid.Text
        grid.col = 0
        txtFun.tag = grid.Text
    Else
        
        grid.col = 1
        
        txtFun.Text = grid.Text
        grid.col = 0
        txtFun.tag = grid.Text
    End If
    cmndbtnBuscar_Click
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub

Private Sub txtFun_Change()
    txtFun.tag = ""
End Sub

Private Sub txtFun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFun
    End If
End Sub
