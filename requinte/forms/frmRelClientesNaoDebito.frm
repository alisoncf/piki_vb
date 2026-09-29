VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelClientesNaoDebito 
   Caption         =   "Relatório de Pós-Venda"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   11820
      TabIndex        =   4
      Top             =   8820
      Width           =   11820
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2143
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pCtop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      Begin VB.TextBox txtDataVen 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   9240
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   10440
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   495
         Left            =   8160
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda desde"
         Height          =   195
         Left            =   5400
         TabIndex        =   13
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência:"
         Height          =   195
         Left            =   4320
         TabIndex        =   11
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Cliente"
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmRelClientesNaoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdCidade_Click()
    Dim c As Cidades.Cidade
    
    c = Cidades.F_BuscaCidades(txtCidade.text)
    If c.cod_cid <> "" Then
        txtCidade.text = c.nome_cid
        txtCidade.tag = c.cod_cid
    End If
    
    
End Sub

Private Sub cmdImprimir_Click()
    P_MsflxToHtml "Relatório de Pós Venda", grid, "", "Total: " & grid.rows - 1 & " Clientes"
End Sub

Private Sub cmdPesquisar_Click()
    Dim strSQL As String
    Dim codigo As String
    Dim sNotIn As String
    P_Status "Verificando todos os débitos... aguarde..."
    
    strSQL = " select tc.cod_cli, nome_cli,nome_cid,count(ti.seq_itd) as parcelas, "
    strSQL = strSQL & " min(ti.vencini_itd) as v_min, MAX(td.data_doc) AS ultimo_doc from tab_clientes as tc "
    strSQL = strSQL & " left join tab_setores as ts on ts.cod_set=tc.cod_set  "
    'strSQL = strSQL & " inner join tab_telefones as tf on tf.cod_cli = tc.cod_cli and tf.tipo_ter='C'"
    strSQL = strSQL & " left join tab_cidades as tcid on tcid.cod_cid=ts.cod_cid  left join tab_doc as td on td.cod_ter=tc.cod_cli "
    strSQL = strSQL & " and td.tipo_ter='C'  left join tab_itdoc as ti on td.seq_doc=ti.seq_doc  "
    strSQL = strSQL & " where cpf_cli <> '' and (ti.vencini_itd is null or ti.vencini_itd >=" & F_GravaData(txtDate.text) & ")"
    strSQL = strSQL & " and nome_cli like '" & txtNome.text & "%' "
    strSQL = strSQL & " and nome_cid like '" & txtCidade.text & "%' "
    
    If txtDataVen.text <> "" Then
        strSQL = strSQL & " and data_doc>=" & F_GravaData(txtDataVen.text)
    End If
    strSQL = strSQL & " group by tc.cod_cli having  ultimo_doc >= '" & Year(Date) - 3 & "-01-01'  AND parcelas<3 order by nome_cli "
    
    
    
    abreObjetoAdo ado, strSQL
    ajustaGrid
    
    grid.Visible = False
    With ado
        Do While Not .EOF
            strSQL = !cod_cli & Chr(9)
            strSQL = strSQL & !nome_cli & Chr(9)
            codigo = F_BuscaTel(!cod_cli, "C")
            strSQL = strSQL & codigo & Chr(9)
            strSQL = strSQL & !nome_cid
            If codigo <> "" Then
                grid.AddItem strSQL
            End If
            P_Status .AbsolutePosition & "/" & .RecordCount
            .MoveNext
        Loop
    End With
    grid.Visible = True
    
    
    P_Status ""
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    ajustaGrid
    txtDate.text = Date
End Sub
Sub ajustaGrid()
    With grid
        .rows = 1
        .Cols = 4
        .col = 0: .text = "Código"
        .col = 1: .text = "Cliente"
        .col = 2: .text = "Telefones"
        .col = 3: .text = "Cidade"
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 3500
        .ColWidth(2) = 3500
        .ColWidth(3) = 2500
        .ColAlignment(0) = 0
        .ColAlignment(1) = 0
        .ColAlignment(2) = 0
    End With
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub

Private Sub txtCidade_Change()
    txtCidade.tag = ""
End Sub

Private Sub txtCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCidade_Click
    End If
End Sub


Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtDataVen_GotFocus()
    P_SelecTexto txtDataVen
End Sub

Private Sub txtDataVen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDataVen_LostFocus()
    P_AjustaData txtDataVen
End Sub
