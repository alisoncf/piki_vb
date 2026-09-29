VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRelResumoRep 
   Caption         =   "Relatório de Resumo de Representante"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6855
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12091
      _Version        =   393216
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   11370
      TabIndex        =   0
      Top             =   0
      Width           =   11370
      Begin VB.CheckBox chkAtivos 
         Caption         =   "Apenas Representantes Ativos"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   4080
         Picture         =   "frmRelResumoRep.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   7920
         Picture         =   "frmRelResumoRep.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   9600
         Picture         =   "frmRelResumoRep.frx":0797
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   3480
         Picture         =   "frmRelResumoRep.frx":088C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRelResumoRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub Label1_Click()

End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim ultimoacerto As String
    Dim debitos As String
    ajustagrid
    grid.Visible = False
    strSQL = "select tr.cod_rep,tr.nome_rep,tr.limite_rep,ultimaconferencia, sum(reman_con*unit_con) as total "
    
    strSQL = strSQL & " from tab_representantes as tr inner join tab_Consignacoes as tc on tc.cod_rep=tr.cod_rep "
    
    
    
    
    strSQL = strSQL & " where status_rep='1' "
    If chkAtivos.Value = 1 Then
        strSQL = strSQL & " and excluido_rep='N' "
    End If
    
    
    
    strSQL = strSQL & " group by tr.cod_rep order by nome_rep  "
    
    
    abreObjetoAdo ado, strSQL
    
    With ado
        Do While Not .EOF
            ultimoacerto = F_BuscaUmCampo("max(data_ace)", "tab_acertos", "cod_rep=" & !cod_Rep & " and status_ace=2")
            debitos = TotalDebitos(!cod_Rep, "R")
            
            
            
            strSQL = .AbsolutePosition & Chr(9) & !cod_Rep & Chr(9) & !nome_rep & Chr(9)
            strSQL = strSQL & FormatVal(!total) & Chr(9) & debitos & Chr(9) & ultimoacerto & Chr(9) & !ultimaconferencia & Chr(9) & FormatVal(!limite_rep)
            grid.AddItem strSQL
            
            .MoveNext
        Loop
    End With
    
    ZebraGrid grid, 1
    grid.Visible = True
    
    'max(ta.data_ace) as ultimo_acerto
    'strSQL = strSQL & " sum(reman_itd) as debito, strSQL = strSQL & " and ti.venc_itd<" & F_GravaData(Date)
End Sub
Sub ajustagrid()
    With grid
        .Cols = 8
        .rows = 1
        .Col = 0: .Text = "..."
        .Col = 1: .Text = "Código"
        .Col = 2: .Text = "Nome"
        .Col = 3: .Text = "Em estoque"
        .Col = 4: .Text = "Débito"
        .Col = 5: .Text = "Último Acerto"
        .Col = 6: .Text = "Última Conf.:"
        .Col = 7: .Text = "Limite:"
        .ColWidth(0) = 600
        .ColWidth(1) = 600
        .ColWidth(2) = 2500
        .ColWidth(7) = 1500
    End With
End Sub
Private Sub cmndbtnBuscaRep_Click()
    buscaRep
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = representante.F_BuscaRep(txtbxRepresentante.Text)
    If r.cod_Rep <> "" Then
        txtbxRepresentante.Text = r.nome_rep
        txtbxRepresentante.tag = r.cod_Rep
        
        busca
        
    End If
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    P_MsflxToHtml "Resumo Consolidado de Representantes", grid, "Resumo Consolidado de Representantes em " & Date
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
    
End Sub

Private Sub txtbxRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaRep_Click
    End If
End Sub
