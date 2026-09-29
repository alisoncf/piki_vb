VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBuscarTelefone 
   Caption         =   "Pesquisar por Telefone:"
   ClientHeight    =   6585
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10425
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7858
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   8520
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton cmdPequisar 
      Caption         =   "Pesquisar"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1110
   End
   Begin VB.Menu mnOpt 
      Caption         =   "&Opções"
      Begin VB.Menu mnEdiCli 
         Caption         =   "Editar Cliente"
      End
      Begin VB.Menu mnAbrirContas 
         Caption         =   "Abrir contas a receber"
      End
      Begin VB.Menu mnCopiarTexto 
         Caption         =   "Copiar Texto"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmBuscarTelefone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    busca
End Sub
Sub busca()
    
    Dim strSQL As String
    Dim numtel As String
    Dim rPesq As New ADODB.Recordset
    Dim i As Integer
    Dim nomecli As String
    Dim tipo As String
    numtel = ""
    AjustaGrid
    
    P_Status "Pesquisando..."
    For i = 1 To Len(txtNumero.Text)
        numtel = numtel & "%" & Mid(txtNumero, i, 1)
    
    Next
    numtel = numtel & "%"
    
    
    
    strSQL = "select * from tab_telefones where num_tel like '" & numtel & "' "
    
    
    abreObjetoAdo rPesq, strSQL
    grid.Visible = False
    With rPesq
        Do While Not .EOF
                    
            If !tipo_ter = "C" Then
                nomecli = F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
                tipo = "Cliente"
            ElseIf !tipo_ter = "F" Then
                nomecli = F_BuscaUmCampo("nome_for", "tab_fornecedores", "cod_for='" & !cod_cli & "'")
                tipo = "Fornecedor"
            ElseIf !tipo_ter = "U" Then
                nomecli = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & !cod_cli & "'")
                tipo = "Funcionario"
            End If
        
        
            strSQL = .AbsolutePosition & Chr(9) & !num_tel & Chr(9) & nomecli & Chr(9) & !tipo_ter & Chr(9) & !cod_cli
        
        
            grid.AddItem strSQL
        
            .MoveNext
        Loop
        
        .Close
    End With
    
    ZebraGrid grid, 1
    grid.Visible = True
    
    P_Status ""
End Sub

Private Sub cmdPequisar_Click()
    busca
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AjustaGrid
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        Clipboard.SetText grid.Text
    End If
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
    End If
    
    
    
End Sub

Private Sub mnAbrirContas_Click()
'
    grid.Col = 2
    frmEditaMov.txtbxTerceiro.Text = grid.Text
    grid.Col = 4
    frmEditaMov.Show 0

End Sub

Private Sub mnCopiarTexto_Click()
    Clipboard.SetText grid.Text
End Sub

Private Sub mnEdiCli_Click()
'

    

    grid.Col = 4
    
    frmCadCli.tag = grid.Text
    frmCadCli.Show 1
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
Sub AjustaGrid()
    With grid
        .Cols = 5
        .rows = 1
        .Col = 0: .Text = "..."
        .Col = 1: .Text = "Telefone"
        .Col = 2: .Text = "Nome"
        .Col = 3: .Text = "Tipo"
        .Col = 4: .Text = "Código"
        .ColWidth(2) = 3000
        .ColWidth(1) = 2000
        .ColAlignment(1) = 0
        
        
        
    End With
End Sub
