VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaDevConsign 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Devolução:"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5535
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9763
      _Version        =   393216
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "data_dev"
         Caption         =   "Data"
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
         DataField       =   "seq_dev"
         Caption         =   "seq"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "F12 - Sair"
      Height          =   615
      Left            =   3360
      Picture         =   "frmListaDevConsign.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnImpr 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   1800
      Picture         =   "frmListaDevConsign.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnVisualiza 
      Caption         =   "Visualizar"
      Height          =   615
      Left            =   120
      Picture         =   "frmListaDevConsign.frx":0578
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   1575
   End
   Begin VB.PictureBox pcProdAcerto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   4905
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   4905
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      Begin VB.Label lblRepresentante 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Representante"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   2700
      End
   End
End
Attribute VB_Name = "frmListaDevConsign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
    Dim p As Produtos.ProdutoConsig
    'p = Produtos.F_BuscaProdConsign(txtbxNomeCli.tag, txtbxProdAcerto.Text)
    
Private Sub cmndbtnConfirmar_Click()
    If ado.RecordCount = 0 Then
        MsgBox "Não há o que ser devolvido!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma devolução dos itens acima?", vbYesNo + vbQuestion) = vbYes Then
        confirma
    End If
End Sub
Sub confirma()
    Dim strSQL As String
    Dim sSEQ As String
    Dim seq_ite As String
    sSEQ = MaxKeySQL("tab_devconsign", "seq_dev", "", "000000#")
    
    strSQL = "insert into tab_devconsign(seq_dev,data_dev,cod_rep)values("
    strSQL = strSQL & "'" & sSEQ & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & lblRepresentante.tag & "')"
    
    
    With ado
        Do While Not .EOF
            strSQL = "insert into tab_itdevconsign(seq_Dev,seq_ite,cod_pro,qtd_dev,seq_con)values("
            strSQL = strSQL & "'" & sSEQ & "',"
            strSQL = strSQL & .AbsolutePosition + 1 & ","
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & F_GravaDec(!qtd_dev) & ","
            strSQL = strSQL & "'" & !seq_con & "') "
            
            
            
            
            Banco.Execute strSQL
            strSQL = "update tab_consignacoes set reman_con=reman_con-" & F_GravaDec(!qtd_dev) & " "
            strSQL = strSQL & " where seq_con='" & !seq_con & "'"
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            Banco.Execute strSQL
            
            
            .MoveNext
        Loop
        strSQL = "delete from tab_devconsignabertas where cod_rep='" & lblRepresentante.tag & "'"
        Banco.Execute strSQL
        MsgBox "Devolução Concluída"
        
    End With
End Sub

Private Sub cmndbtnExclItem_Click()
    
    
    If MsgBox("Deseja excluir o item " & ado!nome_pro & "?", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    End If
    excluiItem
End Sub
Sub excluiItem()
    On Error GoTo trata_erro
    Dim strSQL As String
    strSQL = "delete from tab_devconsignabertas where seq_dev='" & ado!seq_dev & "'"
    Banco.Execute strSQL
    buscaDevAbertas
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Private Sub cmndbtnInserir_Click()
    If txtbxProdAcerto.tag = "" Then
        MsgBox "Informe o produto que deseja devolver e  tecle <ENTER>", vbInformation
        txtbxProdAcerto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxQtdAcerto.Text) <= 0 Then
        MsgBox "Informe um valor possível para a quantidade!", vbExclamation
        txtbxQtdAcerto.SetFocus
        Exit Sub
    End If
    gravaItem
    
    
End Sub
Sub gravaItem()
    Dim strSQL As String
    Dim sSEQ As String
    sSEQ = MaxKeySQL("tab_devconsignabertas", "seq_dev", "", "000000#")
    
    
    
    strSQL = "insert into tab_devconsignabertas(seq_dev,data_dev,cod_pro,qtd_dev,cod_rep,seq_con)values("
    strSQL = strSQL & "'" & sSEQ & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & txtbxProdAcerto.tag & "',"
    strSQL = strSQL & F_GravaDec(txtbxQtdAcerto.Text) & ","
    strSQL = strSQL & "'" & lblRepresentante.tag & "',"
    strSQL = strSQL & "'" & lblSeqConAcerto.caption & "')"
    Banco.Execute strSQL
    
    buscaDevAbertas
    
    
End Sub
Sub buscaDevAbertas()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select tp.nome_pro,tp.ref_pro,td.* from tab_devconsignabertas as td, tab_produtos as tp"
    strSQL = strSQL & " where tp.cod_pro=td.cod_pro "
    strSQL = strSQL & " and cod_rep='" & lblRepresentante.tag & "'"
    strSQL = strSQL & " order by seq_dev "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = "select * from tab_devconsign where cod_rep='" & lblRepresentante.tag & "' "
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and left(data_dev,10) >=" & F_GravaData(txtbxDe.Text)
    End If
    
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and left(data_dev,10) <=" & F_GravaData(txtbxAte.Text)
    End If
    strSQL = strSQL & " order by data_dev,seq_dev "
    
    abreObjetoAdo ado, strSQL, grid
End Sub

Private Sub cmndbtnImpr_Click()
    On Error GoTo trata_erro
    imprimeDevConsign ado!seq_dev
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    lblRepresentante.caption = F_BuscaUmCampo("Nome_rep", "tab_representantes", "cod_rep='" & lblRepresentante.tag & "'")
    cmndbtnBuscar_Click
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

    
Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    txtbxDe.Text = F_PrimDiaMes(Date)
    txtbxAte.Text = F_UltDiaMes(Date)
    
    
End Sub




Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub




Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub
