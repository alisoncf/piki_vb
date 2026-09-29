VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiClientes 
   Caption         =   "Clientes"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   Icon            =   "frmEdiClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   11505
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   11475
      TabIndex        =   14
      Top             =   0
      Width           =   11505
      Begin VB.TextBox txtbxUf 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtbxEnd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiClientes.frx":014A
         Left            =   120
         List            =   "frmEdiClientes.frx":015D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   7440
         Picture         =   "frmEdiClientes.frx":018D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblList 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   9960
         TabIndex        =   21
         Top             =   360
         Width           =   120
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   9960
         TabIndex        =   20
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Listados:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8880
         TabIndex        =   18
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrados:"
         Height          =   195
         Left            =   8880
         TabIndex        =   17
         Top             =   120
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Totais:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8160
         TabIndex        =   16
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         Height          =   195
         Left            =   4200
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço/Cidade:"
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/CPF"
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   11475
      TabIndex        =   9
      Top             =   7080
      Width           =   11505
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   615
         Left            =   5760
         Picture         =   "frmEdiClientes.frx":0517
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCheques 
         Caption         =   "Cheques"
         Height          =   615
         Left            =   8640
         Picture         =   "frmEdiClientes.frx":0601
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDoc 
         Caption         =   "Documentos"
         Height          =   615
         Left            =   7200
         Picture         =   "frmEdiClientes.frx":12CB
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox chkNegativo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clientes Negativos"
         Height          =   555
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   10080
         Picture         =   "frmEdiClientes.frx":1855
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   2880
         Picture         =   "frmEdiClientes.frx":199F
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   1440
         Picture         =   "frmEdiClientes.frx":23A1
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   4320
         Picture         =   "frmEdiClientes.frx":272B
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Data dtClientes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiClientes.frx":2B38
      Height          =   6015
      Left            =   0
      OleObjectBlob   =   "frmEdiClientes.frx":2B51
      TabIndex        =   2
      Top             =   840
      Width           =   9975
   End
End
Attribute VB_Name = "frmEdiClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBGrid1_Click()

End Sub

Private Sub chkNegativo_Click()
    cmndbtnLocaliza_Click
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmIncClientes
    If dtClientes.Recordset.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    strHora = Time
    fr.Tag = dtClientes.Recordset!cod_cli
    
    fr.Show 0
End Sub

Private Sub cmndbtnHist_Click()
    On Error Resume Next
    P_histCli dtClientes.Recordset!cod_cli
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    On Error GoTo trata_erro
    If chkNegativo.Value Then
        P_CabRelTimbr "LISTAGEM DE CLIENTES NEGATIVADOS", obj
    Else
        P_CabRelTimbr "LISTAGEM DE CLIENTES", obj
    End If
    
    P_ImprX obj, "CPF/CNPJ:", 5
    P_ImprX obj, "Nome:", 45
    P_ImprX obj, "Cidade:", 155, False
    P_line1 obj
    With dtClientes.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                
                P_RodapeVis obj, "continua..."
                P_AddPicVis obj
                If chkNegativo.Value Then
                    P_CabRelTimbr "LISTAGEM DE CLIENTES NEGATIVADOS", obj
                Else
                    P_CabRelTimbr "LISTAGEM DE CLIENTES", obj
                End If
            End If
            P_ImprX obj, !cpf_cli, 5
            P_ImprX obj, !nome_cli, 45
            P_ImprX obj, !cidade_cli, 155, False
            
            .MoveNext
        Loop
        P_RodapeVis obj, strNome_Emp
        P_AddPicVis obj
    End With
    Exit Sub
trata_erro:
    P_MostraErro ""
    
End Sub
Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    On Error GoTo trata_erro:
    strSQL = "select * from tab_clientes where "
    
    
    strSQL = strSQL & "(Nome_cli like '%" & txtbxExpr.Text & "%' "

    strSQL = strSQL & "or cod_cli like '" & txtbxExpr.Text & "%' "
    
    strSQL = strSQL & "or cpf_cli like '" & txtbxExpr.Text & "%' "

    strSQL = strSQL & "or rg_cli like '" & txtbxExpr.Text & "%') "
    
    strSQL = strSQL & " and cod_cli <> 0 " 'venda a consumidor
    If txtbxEnd.Text <> "" Then
        strSQL = strSQL & " and (endereco_cli like '%" & txtbxEnd.Text & "%' "
        strSQL = strSQL & " or cidade_cli like '%" & txtbxEnd.Text & "%' "
        strSQL = strSQL & " or bairro_cli like '%" & txtbxEnd.Text & "%') "
    End If
    If Trim(txtbxUf.Text) <> "" Then
        strSQL = strSQL & " and uf_cli='" & txtbxUf.Text & "'"
    End If
    If chkNegativo.Value = 1 Then
        strSQL = strSQL & " and status_cli=2"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrderBy = "cod_cli"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderBy = "Nome_Cli,cpf_Cli"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderBy = "CPF_CLI"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderBy = "Cidade_Cli,uf_Cli,NOme_Cli"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderBy = "UF_CLI,Cidade_CLi"
    End If
    strSQL = strSQL & " order by " & sOrderBy
    dtClientes.RecordSource = strSQL
    dtClientes.Refresh
    If dtClientes.Recordset.RecordCount = 0 Then
        MsgBox "Não foi encontrado nenhum cliente na sua busca!", vbExclamation
    End If
    lblTotal.caption = F_BuscaUmCampo("count(cod_Cli)", "tab_clientes", "")
    lblList.caption = dtClientes.Recordset.RecordCount
    Exit Sub
trata_erro:
    P_MostraErro
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmIncClientes
    fr.Show 0
'    If frmIncClientes.Visible = False Then
'        frmIncClientes.Visible = True
'    End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    dtClientes.Recordset.Close
    Unload Me
    
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    ElseIf KeyCode = vbKeyC And Shift = vbCtrlMask Then
        P_HistCheCPF dtClientes.Recordset!cpf_cli
    End If
    
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    txtbxExpr.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    'Me.Width = 10530
    'Me.Height = 6810
    
    dtClientes.Connect = banco.Connect
    
    cmbxOrd.ListIndex = 0
    cmndbtnLocaliza_Click
    'txtbxExpr.SetFocus
    resizeForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Height = pcBot.Top - dbgrd.Top - 100
    dbgrd.Width = Me.Width - dbgrd.Left - 200
    
    cmndbtnSair.Left = Me.Width - cmndbtnSair.Width - 400
    
    
End Sub

Private Sub txtbxEnd_Change()
    cmndbtnLocaliza_Click
End Sub

Private Sub txtbxEnd_GotFocus()
    P_SelecTexto txtbxEnd
End Sub

Private Sub txtbxEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxEnd_LostFocus()
    P_AjustaTexto txtbxEnd
End Sub

Private Sub txtbxExpr_Change()
    cmndbtnLocaliza_Click
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
    If KeyCode = vbKeyDown Then
        dbgrd.SetFocus
    End If
    
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    
End Sub

Private Sub txtbxUf_GotFocus()
    P_SelecTexto txtbxUf
End Sub

Private Sub txtbxUf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUf_LostFocus()
    P_AjustaTexto txtbxUf
    cmndbtnLocaliza_Click
End Sub
