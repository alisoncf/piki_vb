VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiBan 
   Caption         =   "Edição de Bancos"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   9150
   Icon            =   "frmEdiBan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9150
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   9150
      TabIndex        =   3
      Top             =   6000
      Width           =   9150
      Begin VB.CommandButton cmdSel 
         Caption         =   "Selecionar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiBan.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiBan.frx":066B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiBan.frx":09F5
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   7680
         Picture         =   "frmEdiBan.frx":0B3F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.TextBox txtbxBusca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   8775
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiBan.frx":108D
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Bancos"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cod_bco"
         Caption         =   "Código"
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
         DataField       =   "nome_bco"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5444,788
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   540
   End
End
Attribute VB_Name = "frmEdiBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdSel_Click()
    
    Me.Hide
End Sub

Private Sub cmndbtnEdit_Click()
    If Me.tag = "" Then
        With adc.Recordset
            frmCadBan.tag = !cod_bco
            frmCadBan.txtbxCod.Text = !cod_bco
            frmCadBan.txtbxCod.Enabled = False
            frmCadBan.txtbxNome.Text = !nome_bco
            frmCadBan.Show 1
        End With
    Else
        Me.Hide
    End If
End Sub

Private Sub cmndbtnNew_Click()
    frmCadBan.Show 1
    
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
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
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmdSel_Click
    End If
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
    
End Sub

Private Sub Form_Activate()
    
    If Me.tag <> "" Then
    cmndbtnEdit.caption = "&Selecionar"
    P_Busca
    End If
    
    If Me.tag = "" Then
        cmdSel.Enabled = False
    Else
        cmdSel.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    Set ado = New ADODB.Recordset
    
    
End Sub
Sub P_Busca()
    
    strSQL = "select * from tab_bancos "
    strSQL = strSQL & " where cod_bco='" & F_AjustaTexto(txtbxBusca.Text) & "'"
    strSQL = strSQL & " or nome_bco like '%" & F_AjustaTexto(txtbxBusca.Text) & "%'"
    strSQL = strSQL & " order by nome_bco"
    FechaObjeto ado
    
    ado.Open strSQL, conn_client
    
    Set dtgrd.DataSource = ado
        
    If ado.RecordCount = 1 Then
        If Me.tag <> "" Then
            On Error Resume Next
            Me.Hide
        End If
    Else
    
    End If
End Sub
Private Sub txtbxBusca_Change()
    P_Busca
End Sub

Private Sub txtbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_Busca
    End If
    If KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
End Sub
