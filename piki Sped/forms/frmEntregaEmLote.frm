VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEntregaEmLote 
   Caption         =   "Entrega"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10770
   Icon            =   "frmEntregaEmLote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLote 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cndExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   9480
      TabIndex        =   15
      Top             =   3960
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridItens 
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtRetorno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   4800
      TabIndex        =   12
      Text            =   "00:00"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtHora 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   3960
      TabIndex        =   10
      Text            =   "00:00"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   10770
      TabIndex        =   5
      Top             =   5850
      Width           =   10770
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar (F7)"
         Height          =   735
         Left            =   7800
         Picture         =   "frmEntregaEmLote.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair (Esc)"
         Height          =   735
         Left            =   9240
         Picture         =   "frmEntregaEmLote.frx":1351
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.TextBox txtObss 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   10455
   End
   Begin VB.TextBox txtEntregador 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdEntregador 
      Height          =   285
      Left            =   3480
      Picture         =   "frmEntregaEmLote.frx":18DF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   315
   End
   Begin VB.CommandButton cmdLote 
      Height          =   285
      Left            =   6960
      Picture         =   "frmEntregaEmLote.frx":1C2F
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote:"
      Height          =   195
      Left            =   6000
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblObs 
      AutoSize        =   -1  'True
      Caption         =   "Observação"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   4560
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retorno:"
      Height          =   195
      Left            =   4800
      TabIndex        =   11
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saída:"
      Height          =   195
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entregador"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   780
   End
End
Attribute VB_Name = "frmEntregaEmLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Public status As String
Dim vendas() As String
Public Sub AdicionaVenda(num_ven As String)
gridItens.Refresh
    If F_InFlxgrd(gridItens, num_ven, 1) = -1 Then
        BuscaVenda num_ven
        
        
    End If
End Sub


Private Sub cmdConfirmar_Click()
    Dim s As String
    Dim i As Integer
    If txtEntregador.tag = "" Then
        MsgBox "informe o entregador!", vbExclamation
        Exit Sub
    End If
    BuscaVendasGrid
    If status = "2" Then
        s = EncaminhaEntregaEmLote(vendas, txtEntregador.tag, txtData.Text & " " & txtHora.Text)
        If s <> "" Then
            MsgBox "Entrega encaminhada! Foi gerado o lote " & s, vbInformation
            
            Unload Me
        End If
    End If
    If status = "3" Then
        For i = 1 To UBound(vendas)
            If Not concluientrega(vendas(i - 1), txtEntregador.tag, txtData.Text & " " & txtRetorno.Text, txtObss.Text) Then
                MsgBox "erro na conclusão da entrega"
            End If
            
        Next
        
        Unload Me
        
            
    End If
    
End Sub
Sub BuscaVendasGrid()
    Dim i As Integer
    ReDim vendas(gridItens.rows - 1) As String
    For i = 1 To gridItens.rows - 1
        gridItens.row = i
        gridItens.col = 1
        vendas(i - 1) = gridItens.Text
    Next
    
    
End Sub
Private Sub cmdEntregador_Click()
    Dim f As funcionario.funcionario
    On Error Resume Next
    f = funcionario.F_SelectFuncionario(txtEntregador.Text)
    If f.cod_fun <> "" Then
        txtEntregador.Text = f.guerra_fun
        txtEntregador.tag = f.cod_fun
        
        txtHora.SetFocus
    End If
    
End Sub

Private Sub cmdNovo_Click()
    
End Sub

Private Sub cmdLote_Click()
    BuscaVenda "", txtLote.Text
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cndExcluir_Click()
  '  On Error Resume Next
    gridItens.RemoveItem gridItens.RowSel
    reorganizaGrid gridItens.RowHeight(0)
End Sub
Sub reorganizaGrid(alturaLinha As Single)
    Dim i As Integer
    
    
    With gridItens
        .col = 1
        .row = 1
        If .Text = "" Then
            .RemoveItem 1
        End If
        
        For i = 1 To .rows - 1
            .row = i
            
            .col = 0
            .Text = .row
            .RowHeight(i) = alturaLinha
        Next
    End With
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    For i = 1 To gridItens.rows - 1
        gridItens.col = 1
        gridItens.row = i
        MsgBox gridItens.Text & gridItens.RowHeight(i)
        gridItens.RowHeight(i) = 240
    Next
    gridItens.RowSizingMode = flexRowSizeAll
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    'BuscaVenda
    'buscaProd
    If status = "2" Then
        txtRetorno.Enabled = False

        txtEntregador.SetFocus

    Else
        txtRetorno.Enabled = True
        txtLote.Visible = True
        cmdLote.Visible = True
        lblLote.Visible = True
        txtRetorno.Text = Format(Time, "HH:mm")
        If Me.tag <> "" Then
            BuscaVenda "", Me.tag
        End If
        
        txtRetorno.SetFocus
    End If
    bActive = True
End Sub

Sub BuscaVenda(num_ven As String, Optional lote As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from " & TAB_VENDAS & " as tv left join historico_entrega AS h on tv.num_ven=h.num_Ven "
    strSQL = strSQL & " where 1 > 0 "
    
    If num_ven <> "" Then
        strSQL = strSQL & " and tv.num_ven='" & num_ven & "'"
    End If
    If lote <> "" And lote <> "0" Then
        strSQL = strSQL & " and lote='" & lote & "'"
    End If
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Venda não encontrada!", vbExclamation
            Exit Sub
        End If
        Do While Not .EOF
        
            gridItens.AddItem gridItens.rows - 1 & Chr(9) & !num_ven & Chr(9) & !nomecli_ven & Chr(9) & TiraNull(!endereco_ven) & Chr(9) & TiraNull(!obs_ven)
            .MoveNext
        Loop
    End With
    reorganizaGrid gridItens.RowHeight(0)
    
    End Sub
Sub buscaProd()
    Dim strSQL As String
    strSQL = "select  codfab_pro,desc_ite as produto, qtd_ite as quantidade from " & TAB_ITVENDAS & " as ti, "
    strSQL = strSQL & TAB_PRODUTOS & " as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and num_ven='" & txtVenda.Text & "' "
    'abreObjetoAdo ado, strSQL, DataGrid1
    
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmdConfirmar_Click
    End If
End Sub

Private Sub Form_Load()
    
    ajustaGrid
     
    txtHora.Text = Format(Time, "hh:mm")
    txtData.Text = Date
    Set ado = New ADODB.Recordset
    bActive = False
    
    
    
End Sub
Sub ajustaGrid()
    
    With gridItens
        .row = 0
        .Cols = 5
        .row = 0
        .col = (0): .Text = "Nº"
        .col = (1): .Text = "Venda"
        .col = (2): .Text = "Cliente"
        .col = (3): .Text = "Endereço"
        .col = (4): .Text = "Obs"
        .ColWidth(0) = 500
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 3000
        .ColWidth(4) = 3000
        
        
    End With
    
    
End Sub

Private Sub txtEntregador_Change()
    txtEntregador.tag = ""
End Sub

Private Sub txtEntregador_GotFocus()
    P_SelecTexto txtEntregador
End Sub

Private Sub txtEntregador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdEntregador_Click
    End If
End Sub

Private Sub txtEntregador_LostFocus()
    P_AjustaHora txtHora
End Sub

Private Sub txtHora_GotFocus()
    P_SelecTexto txtHora
End Sub

Private Sub txtHora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtHora_LostFocus()
    P_AjustaHora txtHora
End Sub

Private Sub txtRetorno_GotFocus()
    P_SelecTexto txtRetorno
End Sub

Private Sub txtRetorno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtRetorno_LostFocus()
    P_AjustaHora txtRetorno
End Sub

