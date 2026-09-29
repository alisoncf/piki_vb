VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVerificaInconsistenciaCons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Verificar Inconsistências"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9128
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   11385
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.CommandButton cmndbtnAjusta 
         Caption         =   "Ajusta"
         Height          =   615
         Left            =   5880
         Picture         =   "frmVerificaInconsistenciaCons.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   4200
         Picture         =   "frmVerificaInconsistenciaCons.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   9600
         Picture         =   "frmVerificaInconsistenciaCons.frx":0D8C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   3480
         Picture         =   "frmVerificaInconsistenciaCons.frx":0E81
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lblLog 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   135
   End
End
Attribute VB_Name = "frmVerificaInconsistenciaCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnAjusta_Click()
    busca True
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnBuscaRep_Click()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxRepresentante.Text)
    txtbxRepresentante.Text = r.nome_rep
    txtbxRepresentante.tag = r.cod_Rep
    If txtbxRepresentante.tag = "" Then
        txtbxRepresentante.Text = "(Todos)"
    End If
    
    busca
End Sub
Sub busca(Optional ajusta As Boolean = False)
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    msflxgrd.rows = 1
    P_Status "Buscando Inconsistências..."
    strSQL = "select tp.nome_pro,te.*,sum(tc.reman_con) as restante from tab_estcons as te,"
    strSQL = strSQL & " tab_consignacoes As tc,tab_produtos as tp "
    strSQL = strSQL & " where"
    strSQL = strSQL & " tc.cod_pro = te.cod_pro"
    strSQL = strSQL & " and te.cod_rep=tc.cod_rep"
    strSQL = strSQL & " and tp.cod_pro=te.cod_pro "
    If txtbxRepresentante.tag <> "" Then
        strSQL = strSQL & " and tc.cod_Rep ='" & txtbxRepresentante.tag & "' "
    End If
    strSQL = strSQL & " group by tc.cod_pro,tc.cod_rep "
    strSQL = strSQL & " having restante <> te.estoque_pro"
    strSQL = strSQL & " order by te.cod_rep,te.cod_pro"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        lblLog.caption = .RecordCount & " Registros Encontrados"
        Do While Not .EOF
            If ajusta Then
                P_Status "Ajustando produto " & !nome_pro
                strSQL = "Update tab_estcons set estoque_pro=" & F_GravaDec(!restante)
                strSQL = strSQL & " where cod_rep='" & !cod_Rep & "' "
                strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
                Banco.Execute strSQL
            End If
            strSQL = !cod_Rep & Chr(9)
            strSQL = strSQL & !cod_pro & Chr(9) & !nome_pro & Chr(9)
            strSQL = strSQL & !estoque_pro & Chr(9) & !restante
            msflxgrd.AddItem strSQL
            .MoveNext
            
        Loop
        .Close
    End With
    ZebraGrid msflxgrd, 1
    P_Status ""
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ajustagrid
End Sub
Sub ajustagrid()
    With msflxgrd
        
        .rows = 1
        .Cols = 5
        .Col = 0: .Text = "Repr."
        .Col = 1: .Text = "Código"
        .Col = 2: .Text = "Produto"
        .Col = 3: .Text = "Estoque"
        .Col = 4: .Text = "Reman."
        .ColWidth(2) = 2500
        
    End With
End Sub
Private Sub txtbxRepresentante_Change()
    txtbxRepresentante.tag = ""
End Sub
