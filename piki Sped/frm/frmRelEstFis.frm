VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelEstFis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comparativo de Estoque"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5415
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9551
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordem:"
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   5640
      Width           =   1455
      Begin VB.OptionButton Option3 
         Caption         =   "Nome"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Déficit"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmndbtnBusca 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Produtos:"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   3015
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmRelEstFis.frx":0000
         Left            =   120
         List            =   "frmRelEstFis.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total...:"
      Height          =   195
      Left            =   8160
      TabIndex        =   9
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "frmRelEstFis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxBusca_Click()
    Busca
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnBusca_Click()
    Busca
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    
End Sub
Sub imprime(obj As Object)
    Dim strSQL As String
    Dim iCount As Integer
    P_CabRelTimbr "Listagem de Comparativo de Estoque em " & Date, obj
    With Msflxgrd
        .Row = 0
        .col = 0: P_ImprX obj, .Text, 0
        .col = 1: P_ImprX obj, .Text, 20
        .col = 2: P_ImprX obj, .Text, 130, True, True
        .col = 3: P_ImprX obj, .Text, 150, True, True
        .col = 4: P_ImprX obj, .Text, 170, False, True
        For iCount = 1 To .rows - 1
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, strNome_Emp, "Continua"
                .col = 0: P_ImprX obj, .Text, 0
                .col = 1: P_ImprX obj, .Text, 20
                .col = 2: P_ImprX obj, .Text, 130, True, True
                .col = 3: P_ImprX obj, .Text, 150, True, True
                .col = 4: P_ImprX obj, .Text, 170, False, True
            End If
            .Row = iCount
            .col = 0: P_ImprX obj, .Text, 0
            .col = 1: P_ImprX obj, .Text, 20
            .col = 2: P_ImprX obj, .Text, 130, True, True
            .col = 3: P_ImprX obj, .Text, 150, True, True
            .col = 4: P_ImprX obj, .Text, 170, False, True
        Next
        P_RodapeVis obj, strNome_Emp, "(Fim)"
    End With
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Sub Busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from " & TAB_PRODUTOS
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " where ef_pro<er_pro "
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " where ef_pro>er_pro "
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " where ef_pro<0"
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & " where er_pro<estmin_pro"
    ElseIf cmbxBusca.ListIndex = 4 Then
        strSQL = strSQL & " where er_pro<0"
    End If
    strSQL = strSQL & " order by nome_pro "
    rPesq.Open strSQL, conn_client
    Msflxgrd.rows = 1
    
    With rPesq
        Do While Not .EOF
        Msflxgrd.Visible = False
        P_Status "Buscando..."
            strSQL = FormatCodigo(!cod_pro) & Chr(9)
            strSQL = strSQL & !nome_pro & Chr(9)
            strSQL = strSQL & FormatVal(!ef_pro) & Chr(9)
            strSQL = strSQL & FormatVal(!er_pro) & Chr(9)
            strSQL = strSQL & FormatVal(!er_pro) - !ef_pro & Chr(9)
            strSQL = strSQL & FormatVal(!estmin_pro)
            Msflxgrd.AddItem strSQL
            .MoveNext
            
        Loop
        P_Status
        Msflxgrd.Visible = True
        lblTotal.caption = "Itens Listados....:" & .RecordCount
        .Close
    End With
    ZebraGrid Msflxgrd, 1
End Sub
Private Sub Form_Activate()
    'MudaAparencia Me
    Busca
End Sub

Private Sub Form_Load()
    Me.Top = 10
    Me.Left = 10
    cmbxBusca.ListIndex = 0
    With Msflxgrd
        .Cols = 6
        .FixedCols = 0
        .rows = 1
        .col = 0: .Text = "Código"
        .col = 1: .Text = "Descrição"
        .col = 2: .Text = "Est."
        .col = 3: .Text = "Qtd."
        .col = 4: .Text = "Dif."
        .col = 5: .Text = "Est.Mín."
        .ColWidth(0) = 1000
        .ColWidth(1) = 3500
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
    End With
    
    
End Sub

Private Sub optnDef_Click()
    Busca
End Sub

Private Sub optnSuper_Click()
    Busca
End Sub

Private Sub optnTodos_Click()
    Busca
End Sub
