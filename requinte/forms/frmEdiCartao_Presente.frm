VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCartao_Presente 
   Caption         =   "Transações de Cartão-Presente"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   Icon            =   "frmEdiCartao_Presente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   14400
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   14400
      TabIndex        =   13
      Top             =   7110
      Width           =   14400
      Begin VB.CommandButton cmdBaixa 
         Caption         =   "Forçar baixa"
         Height          =   855
         Left            =   5880
         Picture         =   "frmEdiCartao_Presente.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Selecionar"
         Height          =   855
         Left            =   2040
         Picture         =   "frmEdiCartao_Presente.frx":0156
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   855
         Left            =   3960
         Picture         =   "frmEdiCartao_Presente.frx":0225
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   855
         Left            =   120
         Picture         =   "frmEdiCartao_Presente.frx":036F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   9720
         Picture         =   "frmEdiCartao_Presente.frx":07E9
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   855
         Left            =   12480
         Picture         =   "frmEdiCartao_Presente.frx":0B73
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Im&primir"
         Height          =   855
         Left            =   7800
         Picture         =   "frmEdiCartao_Presente.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   14400
      TabIndex        =   0
      Top             =   0
      Width           =   14400
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   7440
         Picture         =   "frmEdiCartao_Presente.frx":1075
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtDataFim 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frmEdiCartao_Presente.frx":13FF
         Left            =   5880
         List            =   "frmEdiCartao_Presente.frx":140F
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   195
         Left            =   3480
         TabIndex        =   7
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   5880
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid grade 
      Height          =   6135
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   10821
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
      Caption         =   "Transações Encontradas"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "numero"
         Caption         =   "Número"
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
         DataField       =   "valor"
         Caption         =   "Valor"
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
         DataField       =   "data"
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
      BeginProperty Column04 
         DataField       =   "vencimento"
         Caption         =   "Vencimento"
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
         DataField       =   "nome"
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
      BeginProperty Column06 
         DataField       =   "contato"
         Caption         =   "Contato"
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
      BeginProperty Column07 
         DataField       =   "num_ven"
         Caption         =   "Venda"
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
      BeginProperty Column08 
         DataField       =   "status"
         Caption         =   "Status"
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
            ColumnWidth     =   585,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2280,189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2580,095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1620,284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   615,118
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiCartao_Presente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean

Private Sub cmdBaixa_Click()
    Dim strSQL As String
    On Error GoTo trata_erro:
    
    If ado!status <> "1" Then
        MsgBox "Este status não permite baixa!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente forçar a baixa sem indicar venda?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para lançamentos em Caixa", vbExclamation
        Exit Sub
    End If
    Cartao_Presente.BaixaCartaoPresente ado!id, "0", ""
    
    
    strSQL = F_GravaCai("C", Date, ado!valor, "BAIXA CARTÃO PRESENTE " & ado!nome, "", "MOV", strCod_Cai, ado!valor, ado!valor, strCod_Cai, _
        strCod_Tur, "000", "S")
    ado.Requery
    Exit Sub
    
trata_erro:
P_MostraErro
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo trata_erro
    
    ImprimeCartao_Presente ado!id
    
    Exit Sub
trata_erro:
    
End Sub

Private Sub cmdNovo_Click()
    frmCartaoPresente.Show 1
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub
Sub busca()
    On Error GoTo trata_erro
    Dim strSQL As String
    strSQL = "select c.* from cartao_presente as c "
    strSQL = strSQL & " where 1>0 "
    If txtDataIni.Text <> "" Then
        strSQL = strSQL & " and left(data,10)>=" & F_GravaData(txtDataIni.Text)
    End If
    If txtDataFim.Text <> "" Then
        strSQL = strSQL & " and left(data,10)<=" & F_GravaData(txtDataFim.Text)
    End If
    If cmbStatus.Text <> "" Then
        strSQL = strSQL & " and status='" & Left(cmbStatus.Text, 1) & "'"
    End If
    
    If txtNome.Text <> "" Then
        strSQL = strSQL & " and (nome like '% " & txtNome.Text & "%' or presenteado like '%" & txtNome.Text & "%') "
    End If
    If txtNumero.Text <> "" Then
        strSQL = strSQL & " and numero='" & txtNumero.Text & "' "
    End If
    
    strSQL = strSQL & " order by id "
    
    
    
    abreObjetoAdo ado, strSQL, grade
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_erro

    frmCartaoPresente.tag = ado!id
    frmCartaoPresente.Show 1
trata_erro:
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo trata_erro
    Me.tag = ado!id
    Me.Hide
trata_erro:
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    busca
    If Me.tag = "" Then
        cmdSelect.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    cmbStatus.ListIndex = 0
    bActive = False
End Sub

Private Sub txtDataIni_GotFocus()
    P_SelecTexto txtDataIni
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDataIni_LostFocus()
    P_AjustaData txtDataIni
End Sub
Private Sub txtDataFim_GotFocus()
    P_SelecTexto txtDataFim
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDataFim_LostFocus()
    P_AjustaData txtDataFim
End Sub

Private Sub txtNumero_GotFocus()
    P_SelecTexto txtNumero
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNumero_LostFocus()
    P_AjustaTexto txtNumero
End Sub

Private Sub txtNome_GotFocus()
    P_SelecTexto txtNome
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNome_LostFocus()
    P_AjustaTexto txtNome
End Sub
