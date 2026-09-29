VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmClienteSpc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cliente SPC"
   ClientHeight    =   5505
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4260
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "data_spc"
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
         DataField       =   "doc_spc"
         Caption         =   "Doc"
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
         DataField       =   "valor_spc"
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
         DataField       =   "baixa_spc"
         Caption         =   "Baixa"
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
         DataField       =   "cod_usu"
         Caption         =   "Usuário"
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
         DataField       =   "dataacao_spc"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmndbtnBaixa 
      Caption         =   "Baixar Bloqueio"
      Height          =   735
      Left            =   2040
      Picture         =   "frmClienteSpc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Frame frBloq 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   8535
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "Bloquear"
         Height          =   615
         Left            =   6960
         Picture         =   "frmClienteSpc.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo Doc:"
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1215
         Begin VB.OptionButton optnParc 
            Caption         =   "Parcela"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optnCh 
            Caption         =   "Cheque"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox txtbxValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtbxDoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxData 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   285
         Left            =   3600
         Picture         =   "frmClienteSpc.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcela:"
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.CommandButton cmndbtnLista 
      Caption         =   "Lista"
      Height          =   735
      Left            =   120
      Picture         =   "frmClienteSpc.frx":11DE
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   735
      Left            =   6960
      Picture         =   "frmClienteSpc.frx":133F
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1695
   End
   Begin VB.ComboBox cmbxTipoTer 
      Height          =   315
      ItemData        =   "frmClienteSpc.frx":14AA
      Left            =   120
      List            =   "frmClienteSpc.frx":14B4
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtbxCpf 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      MaxLength       =   50
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtbxRazao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CPF:"
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
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome:"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
   Begin VB.Menu mnAcao 
      Caption         =   "&Ação"
      Begin VB.Menu mnDesbloq 
         Caption         =   "&Desbloquear"
      End
      Begin VB.Menu mnBloq 
         Caption         =   "&Bloquear"
      End
      Begin VB.Menu div 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmClienteSpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub chkBloq_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If strCadcli_usu <> "S" Then
        MsgBox "Não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    If chkBloq.Value = 1 Then
        
    End If
End Sub
Sub baixa_bloq()
    If MsgBox("Deseja realmente Baixar o bloqueio?", vbYesNo + vbQuestion) = vbYes Then
            'if strcadu
            P_BaixaClienteSpc txtbxData.tag
            
            busca
        End If
End Sub

Private Sub cmndbtnBaixa_Click()
    baixa_bloq
End Sub

Private Sub cmndbtnBuscar_Click()
    If optnCh.Value Then
        buscaCheque
    Else
        buscaItDoc
    End If
    
End Sub
Sub buscaCheque()
    Dim c As cheque
    c = Cheques.F_BuscaChequeTer(txtbxDoc.Text)
    If c.seq_che <> "" Then
        txtbxDoc.Text = c.seq_che
        txtbxValor.Text = c.Valor_che
        
    End If
End Sub
Sub buscaItDoc()
    Dim i As itDoc
    i = F_BuscaItDoc(txtbxDoc.Text)
    If i.seq_itd <> "" Then
        txtbxDoc.Text = i.seq_itd
        txtbxValor.Text = i.reman_itd
        Me.Show
    End If
End Sub
Private Sub cmndbtnConfirmar_Click()
    If CCur(txtbxValor.Text) <= 0 Then
        MsgBox "Informe um valor válido!", vbExclamation
        txtbxValor.SetFocus
        Exit Sub
    End If
    If txtbxRazao.tag = "" Then
        MsgBox "Escolha primeiro Cliente/Representante.", vbExclamation
        txtbxRazao.SetFocus
        Exit Sub
    End If
    If txtbxDoc.Text = "" Then
        MsgBox "Informe o documento!", vbExclamation
        txtbxDoc.SetFocus
        Exit Sub
        
    End If
    If txtbxData.Text = "" Then
        MsgBox "Informe a data de bloqueio!", vbExclamation
        txtbxData.SetFocus
        Exit Sub
    End If
    'If chkBloq.Value = 0 Then
    '    MsgBox "Você deve primeiro bloquear antes de criar um desbloqueio!", vbExclamation
    '
    '    Exit Sub
    'End If
    If MsgBox("Confirma operação?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    Clientes.P_GravaClienteSpc txtbxRazao.tag, txtbxData.Text, txtbxValor.Text, txtbxDoc.Text, IIf(optnParc.Value, "P", "C"), IIf(cmbxTipoTer.ListIndex = 1, "R", "C")
    busca
    
    
End Sub

Private Sub cmndbtnLista_Click()
    On Error Resume Next
    frmRelClienteSpc.Show 1
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_DblClick()
    With adc.Recordset
    On Error Resume Next
    If !baixa_spc = "" Or IsNull(!baixa_spc) Then
        chkBloq.Value = 1
                
        txtbxValor.Text = FormatVal(!valor_spc)
        txtbxDoc.Text = !doc_spc
        txtbxData.Text = !data_spc
        txtbxData.tag = !seq_spc
    End If
    End With
End Sub

Private Sub dbgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnAcao
    End If
End Sub

Private Sub dt_Reposition()
    On Error Resume Next
    With adc.Recordset
        If IsNull(!baixa_spc) Then
            cmndbtnBaixa.Enabled = True
            mnDesbloq.Enabled = True
        Else
            
            cmndbtnBaixa.Enabled = False
            mnDesbloq.Enabled = False
        End If
        
    End With
    
    dbgrd_DblClick
    
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    carrega
    bActive = True
End Sub

Private Sub Form_Load()
    txtbxData.Text = Date
    
End Sub

Private Sub mnBloq_Click()
    '
End Sub

Private Sub mnDesbloq_Click()
    '
End Sub

Private Sub mnSair_Click()
    cmndbtnSair_Click
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub

Private Sub txtbxDoc_GotFocus()
    P_SelecTexto txtbxDoc
End Sub

Private Sub txtbxDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDoc_LostFocus()
    P_AjustaTexto txtbxDoc
End Sub

Private Sub txtbxRazao_GotFocus()
    P_SelecTexto txtbxRazao
End Sub

Private Sub txtbxRazao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaTer
    End If
End Sub
Sub buscaTer()
    If cmbxTipoTer.ListIndex = 1 Then
        buscaRep
    Else
        buscaCLi
    End If
    
End Sub
Sub buscaCLi()
    Dim c As cliente
    c = Clientes.F_BuscaCliente(txtbxRazao.Text)
    If c.cod_cli <> "" Then
        txtbxRazao.Text = c.nome_Cli
        txtbxRazao.tag = c.cod_cli
        txtbxCpf.Text = c.cnpj_Cli
        busca
    End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = representante.F_BuscaRep(txtbxRazao.Text)
    If r.cod_Rep <> "" Then
        txtbxRazao.Text = r.nome_rep
        txtbxRazao.tag = r.cod_Rep
        txtbxCpf.Text = r.cpf_rep
        busca
    End If
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_clientespc where cod_cli='" & txtbxRazao.tag & "' "
    
    strSQL = strSQL & " order by seq_spc"
    
    dt.RecordSource = strSQL
    
    dt.Refresh
    With dt.Recordset
        Do While Not .EOF
            If !baixa_spc = "" Or IsNull(!baixa_spc) Then
                'chkBloq.Value = 1
                txtbxValor.Text = FormatVal(!valor_spc)
                txtbxDoc.Text = !doc_spc
                txtbxData.Text = !data_spc
                txtbxData.tag = !seq_spc
            End If
            .MoveNext
        Loop
    End With
End Sub
Sub carrega()
    Dim strSQL As String
    Dim c As cliente
    
    c = Clientes.F_BuscaClienteCodigo(Me.tag)
    If c.cod_cli <> "" Then
        txtbxRazao.Text = c.nome_Cli
        txtbxRazao.tag = c.cod_cli
        txtbxCpf.Text = c.cnpj_Cli
        busca
    End If
    
    
    
End Sub

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    P_ValorLostFocus txtbxValor
End Sub
