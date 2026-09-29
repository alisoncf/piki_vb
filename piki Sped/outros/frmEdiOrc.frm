VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEdiOrc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Orçamentos"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnF10 
      Caption         =   "&Atacado F10"
      Height          =   615
      Left            =   1800
      Picture         =   "frmEdiOrc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnExcl 
      Caption         =   "E&xcluir"
      Height          =   615
      Left            =   4920
      Picture         =   "frmEdiOrc.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnVendaAtacado 
      Caption         =   "&Atacado"
      Height          =   615
      Left            =   240
      Picture         =   "frmEdiOrc.frx":327C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   8040
      Picture         =   "frmEdiOrc.frx":33C6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Height          =   615
      Left            =   6480
      Picture         =   "frmEdiOrc.frx":3510
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnPrint 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   3360
      Picture         =   "frmEdiOrc.frx":365A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   3855
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6800
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   6120
         Picture         =   "frmEdiOrc.frx":37A4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox txtbxAte 
         Height          =   285
         Left            =   4920
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxCliente 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   4920
         TabIndex        =   5
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   3840
         TabIndex        =   3
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmEdiOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As Recordset

    
    
    
    strSQL = "select * from tab_orc where situacao_orc=1 "
    If txtbxCliente.Text <> "" Then
        strSQL = strSQL & " and nomecli_orc like '%" & txtbxCliente.Text & "%' "
        
    End If
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and left(data_orc,10)>=" & F_GravaData(txtbxDe.Text)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and left(data_orc,10)<=" & F_GravaData(txtbxAte.Text)
    End If
    strSQL = strSQL & " order by num_orc desc "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        msflxGrd.rows = 1
        Do While Not .EOF
            strSQL = !num_orc & Chr(9) & !cod_cli & Chr(9)
            strSQL = strSQL & !nomecli_orc & Chr(9)
            strSQL = strSQL & FormatVal(!valor_orc) & Chr(9)
            strSQL = strSQL & FormatVal(!Desconto_orc) & Chr(9)
            strSQL = strSQL & !data_orc & Chr(9)
            strSQL = strSQL & !validade_orc & Chr(9)
            strSQL = strSQL & !cod_usu & Chr(9)
            strSQL = strSQL & !situacao_orc & Chr(9)
            strSQL = strSQL & !obs_orc
            msflxGrd.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    zebraGrid msflxGrd, 1
End Sub

Private Sub cmndbtnEdit_Click()
    If msflxGrd.Row <= 0 Then
        Exit Sub
    End If
    msflxGrd.Col = 0
    Dim fr As New frmOrcamento
    fr.Tag = msflxGrd.Text
    fr.Show 0
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    If msflxGrd.Row <= 0 Then
        Exit Sub
    End If
    msflxGrd.Col = 0
    If MsgBox("Confirma Exclusão do Orçamento: " & msflxGrd.Text & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "delete from tab_orc where num_orc='" & msflxGrd.Text & "' "
    Banco.Execute strSQL
    strSQL = "delete from tab_itorc where num_orc='" & msflxGrd.Text & "' "
    Banco.Execute strSQL
    cmndbtnBuscar_Click
End Sub

Private Sub cmndbtnF10_Click()
    Dim sCod_Cli As String
    With msflxGrd
        If .rows = 1 Then Exit Sub
        If strContrpre_usu <> "S" Then
            MsgBox "Você não tem permissão para venda em atacado!", vbCritical
            Exit Sub
        End If
        .Col = 6
        If CDate(.Text) < Date Then
            If MsgBox("O Orçamento está vencido! Deseja Continuar?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
        If MsgBox("Confirma Transferência do Orçamento para Venda em Atacado?", vbYesNo + vbQuestion) = vbYes Then
            .Col = 1
            sCod_Cli = msflxGrd.Text
            .Col = 0
            transfereAtacado .Text, sCod_Cli, True
        End If
    End With
End Sub

Private Sub cmndbtnPrint_Click()
    If msflxGrd.Row <= 0 Then
        Exit Sub
    End If
    msflxGrd.Col = 0
    ImprOrc msflxGrd.Text
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnVendaAtacado_Click()
    Dim sCod_Cli As String
    With msflxGrd
        If .rows = 1 Then Exit Sub
        If strContrpre_usu <> "S" Then
            MsgBox "Você não tem permissão para venda em atacado!", vbCritical
            Exit Sub
        End If
        .Col = 6
        If CDate(.Text) < Date Then
            If MsgBox("O Orçamento está vencido! Deseja Continuar?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
        If MsgBox("Confirma Transferência do Orçamento para Venda em Atacado?", vbYesNo + vbQuestion) = vbYes Then
            .Col = 1
            sCod_Cli = msflxGrd.Text
            .Col = 0
            transfereAtacado .Text, sCod_Cli, False
        End If
    End With
End Sub
Sub transfereAtacado(num_orc As String, cod_cli As String, f10 As Boolean)
    Dim strSQL As String
    Dim sNum_ven As String
    Dim rPesq As Recordset
    On Error GoTo trata_erro
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    
    sNum_ven = F_GeraNovaVenda(cod_cli, strCod_Usu)
    strSQL = "select * from tab_itorc where num_orc='" & num_orc & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
        strSQL = "insert into tab_itvendas(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite,final_ite)values("
        strSQL = strSQL & "'" & sNum_ven & "', "
        strSQL = strSQL & "'" & !cod_pro & "',"
        strSQL = strSQL & F_GravaDec(!unit_ite) & ","
        strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
        strSQL = strSQL & .AbsolutePosition + 1 & ","
    
        strSQL = strSQL & F_GravaDec(!unit_ite * !qtd_ite) & ") "
        Banco.Execute strSQL
        'RESERVANDO O ESTOQUE
        If f10 = False Then
            strSQL = "update tab_produtos set er_pro=er_pro - " & F_GravaDec(!qtd_ite)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Banco.Execute strSQL
        End If
        
            
            .MoveNext
        Loop
        .Close
    End With
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    frmVenda.operacao = "B"
    frmVenda.Tag = sNum_ven
    On Error Resume Next
    frmVenda.Atacado = True
    frmVenda.Show 0, mdiFrm
    If f10 Then
        frmVenda.pcBaixa.Visible = True
        
    End If
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Load()
    txtbxDe.Text = Date
    txtbxAte.Text = Date
    With msflxGrd
        .rows = 1
        .FixedCols = 0
        .Cols = 10
        .Col = 0: .Text = "Nº."
        .Col = 1: .Text = "Cód."
        .Col = 2: .Text = "Cliente:"
        .Col = 3: .Text = "Valor:"
        .Col = 4: .Text = "Desconto:"
        .Col = 5: .Text = "Data:"
        .Col = 6: .Text = "Validade:"
        .Col = 7: .Text = "Usuário:"
        .Col = 8: .Text = "Situação:"
        .Col = 9: .Text = "Condições Pgto:"
        .ColWidth(0) = 1000
        .ColWidth(1) = 800
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 2500
        .ColWidth(6) = 1200
        .ColWidth(7) = 1500
        .ColWidth(8) = 1200
        .ColWidth(9) = 2500
    End With
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxCliente_GotFocus()
    P_SelecTexto txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim c As tCliente
        c = F_BuscaCliente("C", txtbxCliente.Text)
        If c.cod_cli <> "" Then
            txtbxCliente.Text = c.nome_Cli
            txtbxCliente.Tag = c.cod_cli
            SendKeys Chr(9)
        Else
            txtbxCliente.Text = ""
            txtbxCliente.Tag = ""
            txtbxCliente.SetFocus
        End If
    End If
End Sub

Private Sub txtbxCliente_LostFocus()
    P_AjustaTexto txtbxCliente
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub
