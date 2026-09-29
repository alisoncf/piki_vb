VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDevolucao 
   Caption         =   "Devolução"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12660
   Icon            =   "frmDevolucao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12630
      TabIndex        =   20
      Top             =   7050
      Width           =   12660
      Begin VB.TextBox txtMotivo 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   120
         Width           =   5775
      End
      Begin VB.CommandButton cmbxSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11040
         Picture         =   "frmDevolucao.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "Confirmar Devolução"
         Height          =   615
         Left            =   9120
         Picture         =   "frmDevolucao.frx":1E65
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2400
         TabIndex        =   21
         Top             =   240
         Width           =   585
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
         Left            =   840
         TabIndex        =   26
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   12465
      TabIndex        =   5
      Top             =   3960
      Width           =   12495
      Begin VB.TextBox txtEmp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8640
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   11040
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDev 
         Caption         =   "Devolver"
         Height          =   375
         Left            =   9960
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtUnit 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtQtdDev 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   11
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid gridDev 
         Height          =   1575
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   2778
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp"
         Height          =   195
         Left            =   8640
         TabIndex        =   34
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   7560
         TabIndex        =   14
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário:"
         Height          =   195
         Left            =   6240
         TabIndex        =   12
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd"
         Height          =   195
         Left            =   5400
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Produtos a devolver"
         Height          =   195
         Left            =   5640
         TabIndex        =   19
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   300
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Produtos da Venda"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "codfab_pro"
         Caption         =   "Ref."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "desc_ite"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "qtd_ite"
         Caption         =   "Qtd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "unit_ite"
         Caption         =   "Unitário"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "desconto_ite"
         Caption         =   "Desc"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "final_ite"
         Caption         =   "Final"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "devolvido"
         Caption         =   "devolvido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6089,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   824,882
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12630
      TabIndex        =   0
      Top             =   0
      Width           =   12660
      Begin VB.TextBox txtVenda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscaVenda 
         Height          =   285
         Left            =   1560
         Picture         =   "frmDevolucao.frx":1FAF
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblVendedor 
         Height          =   255
         Left            =   9240
         TabIndex        =   32
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   9240
         TabIndex        =   31
         Top             =   120
         Width           =   690
      End
      Begin VB.Label lblData 
         Height          =   255
         Left            =   7200
         TabIndex        =   30
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   195
         Left            =   7200
         TabIndex        =   29
         Top             =   120
         Width           =   345
      End
      Begin VB.Label lblCliente 
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   2280
         TabIndex        =   27
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim num_ven As String
Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    lblCliente.caption = ado!nomecli_ven
    lblData.caption = ado!dataefet_ven
    lblVendedor.caption = F_BuscaUmCampo("guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & ado!abrevenda_usu & "'")
    
End Sub

Private Sub cmbxSair_Click()
    Unload Me
End Sub

Private Sub cmdBuscaVenda_Click()
    BuscaVenda
End Sub

Private Sub cmdDev_Click()
    '
    Dim strSQL As String
    Dim cDisponivel As Currency
    If F_InFlxgrd(gridDev, txtRef.Text, 1) > -1 Then
        MsgBox "Este produto já se encontra na lista de devoluções!", vbExclamation
        Exit Sub
    End If
    If IsNull(ado!devolvido) Then
        cDisponivel = 0
    Else
        cDisponivel = ado!devolvido
    End If
    
    cDisponivel = ado!qtd_ite - cDisponivel
    
    
    
    
    If CCur(txtQtdDev.Text) > cDisponivel Then
        MsgBox "A quantidade informada é maior que a quantidade disponível para devolução.", vbExclamation
        Exit Sub
    End If
    
    strSQL = ado!cod_pro & Chr(9)
    strSQL = strSQL & ado!codfab_pro & Chr(9)
    strSQL = strSQL & ado!desc_ite & Chr(9)
    strSQL = strSQL & txtQtdDev.Text & Chr(9)
    strSQL = strSQL & txtTotal.Text & Chr(9)
    strSQL = strSQL & txtEmp.Text & Chr(9)
    strSQL = strSQL & ado!seq_ite & Chr(9) & ado!qtd_ite & Chr(9) & ado!Desconto_ite
    gridDev.AddItem strSQL
    
    lblTotal.caption = FormatVal(CCur(lblTotal.caption) + CCur(txtTotal.Text))
    
End Sub

Private Sub cmdExcluir_Click()
    If gridDev.row > 0 Then
        gridDev.col = 4
        lblTotal.caption = FormatVal(CCur(lblTotal.caption) - CCur(gridDev.Text))
        If gridDev.row > 1 Then
            gridDev.RemoveItem gridDev.row
        Else
            gridDev.rows = 1
        End If
    End If
End Sub

Private Sub cmndbtnConfirmar_Click()
    Dim strSQL As String
    Dim i As Integer
    Dim seq As String
    Dim cod As String
    Dim qtd As String
    Dim total As String
    Dim emp As String
    Dim seq_ite As String
    Dim qtdOriginal As String
    Dim Desconto_ite As String
    
    Dim cProp As Currency
    
    
'    If F_Vazio(txtMotivo.Text) Then
'        MsgBox "Informe o motivo!", vbExclamation
'        Exit Sub
'    End If
    
    
    If CCur(lblTotal.caption) <= 0 Then
        MsgBox "Não há lançamentos na devolução!", vbExclamation
        Exit Sub
    End If
    
    seq = GravaDevolucao(Date, Time, strCod_Usu, lblTotal.caption, ado!num_ven, "", txtMotivo.Text)
    If seq = "" Then
        MsgBox "Erro na gravação da devolução!", vbExclamation
        Exit Sub
    End If
    With gridDev
        For i = 1 To .rows - 1
            .row = i
            .col = 0: cod = .Text
            .col = 3: qtd = .Text
            .col = 4: total = .Text
            .col = 5: emp = .Text
            
            If Not GravaItDev(seq, cod, qtd, total, emp) Then
                MsgBox "erro na gravação do item!", vbExclamation
                conn.Execute "delete from devolucoes where seq_dev=" & seq
                conn.Execute "delete from itdev where seq_dev=" & seq
                Exit Sub
            End If
            
        Next
        Dim result As Integer
        For i = 1 To .rows - 1
            .row = i
            .col = 0: cod = .Text
            .col = 3: qtd = .Text
            .col = 4: total = .Text
            .col = 5: emp = .Text
            .col = 6: seq_ite = .Text
            .col = 7: qtdOriginal = .Text
            .col = 6: Desconto_ite = .Text
            
            strSQL = "update estoque set er=er+" & F_GravaDec(qtd) & " where cod_pro='" & cod & "' "
            strSQL = strSQL & " and cod_emp='" & emp & "' "
            conn.Execute strSQL, result
            strSQL = "update " & TAB_PRODUTOS & "  set er_pro=er_pro+" & F_GravaDec(qtd) & " where cod_pro='" & cod & "'"
            
            conn.Execute strSQL, result
            
            strSQL = "update " & TAB_ITVENDAS & " set devolvido_ite=devolvido_ite+" & F_GravaDec(qtd)
            strSQL = strSQL & " where num_ven='" & num_ven & "' "
            
            
        
            
            strSQL = strSQL & " and seq_ite='" & seq_ite & "'"
            
            conn.Execute strSQL
            
            cProp = CCur(qtd) / qtdOriginal
            
            strSQL = "Update " & TAB_VENDAS & " set devolvido_ven=devolvido_ven+" & F_GravaDec(CCur(total))
            strSQL = strSQL & " where num_Ven='" & num_ven & "'"
            conn.Execute strSQL
            
            
        
        Next
        
    End With
    
    If strContaDevolucoes = "" Then
        strContaDevolucoes = "006"
    End If
    strSQL = F_GravaFinanceiro("D", Date, lblTotal.caption, "DEVOLUÇÃO " & txtVenda.Text & " " & lblCliente.caption, strContaDevolucoes, "S", _
     "21000", "")
    If strSQL = "" Then
        MsgBox "Erro na gravação do financeiro!", vbExclamation
    End If
    
    
    ImprimeDevolucao seq
    BuscaVenda
        
    
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        cmdBuscaVenda_Click
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    ajustaGrid
End Sub

Sub ajustaGrid()
    lblTotal.caption = "0,00"
    With gridDev
        .FixedCols = 0
        .rows = 1
        .Cols = 9
        .col = 0: .Text = "Código"
        .col = 1: .Text = "Cod.Fab"
        .col = 2: .Text = "Nome"
        .col = 3: .Text = "Qtd.Dev."
        .col = 4: .Text = "Total"
        .col = 5: .Text = "Emp"
        .col = 6: .Text = "Seq.Item"
        .col = 7: .Text = "Qtd Item"
        .col = 8: .Text = "Desc."
      
        .ColWidth(2) = 3500
    End With
End Sub

Private Sub grid_DblClick()
    editaItem
End Sub
Sub editaItem()
    With ado
        txtNome.Text = !desc_ite
        txtQtdDev.Text = "1"
        txtRef.Text = !codfab_pro
        txtUnit.Text = FormatVal(!unit_ite - (!Desconto_ite / !qtd_ite))
        txtEmp.Text = !cod_emp
        somaItem
        txtQtdDev.SetFocus
    End With
End Sub
Sub somaItem()
    txtTotal.Text = FormatVal(CCur(txtQtdDev.Text) * CCur(txtUnit.Text))
    
End Sub
Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grid_DblClick
    End If
End Sub

Private Sub txtMotivo_GotFocus()
    P_SelecTexto txtMotivo
End Sub

Private Sub txtMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtMotivo_LostFocus()
    AjustaTexto txtMotivo
End Sub

Private Sub txtQtdDev_GotFocus()
    P_SelecTexto txtQtdDev
End Sub

Private Sub txtQtdDev_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtQtdDev_LostFocus()
    P_ValorLostFocus txtQtdDev
    
    
    txtTotal.Text = FormatVal(CCur(txtQtdDev.Text) * CCur(txtUnit.Text))
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
    MsgBox "Escolha o produto na lista àcima.", vbInformation
End Sub

Private Sub txtVenda_GotFocus()
    P_SelecTexto txtVenda
End Sub

Private Sub txtVenda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BuscaVenda
    End If
End Sub
Sub BuscaVenda()
    Dim strSQL As String
    ajustaGrid
    
    num_ven = txtVenda.Text
    
    strSQL = "select abrevenda_usu,nomecli_ven,dataefet_ven,tp.codfab_pro,tv.data_ven,dataefet_ven,"
    strSQL = strSQL & "tc.cnpj_cli,tc.cpf_cli,"
    strSQL = strSQL & "(select sum(qtd_dev) from devolucoes as td inner join itdev as tid on td.seq_dev=tid.seq_dev "
    strSQL = strSQL & " where num_ven='" & txtVenda.Text & "' and tid.cod_pro=ti.cod_pro  ) as devolvido,"
    strSQL = strSQL & " ti.* from " & TAB_VENDAS & " as tv "
    strSQL = strSQL & " inner join " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on ti.cod_pro=tp.cod_pro "
    'strSQL = strSQL & " left join devolucoes as td on td.num_ven=tv.num_ven "
    'strSQL = strSQL & " left join itdev as tid on td.seq_dev=tid.seq_dev "
    strSQL = strSQL & " where tv.num_ven='" & txtVenda.Text & "' "
    strSQL = strSQL & " AND tv.situacao_ven='4' "
    
    
    abreObjetoAdo ado, strSQL, grid
    
End Sub
