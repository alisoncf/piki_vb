VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiUmCampo 
   Caption         =   "Edição de Categorias"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8160
   Icon            =   "frmEdiUmCampo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   8160
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   5280
      Width           =   8175
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   6240
         Picture         =   "frmEdiUmCampo.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   2160
         Picture         =   "frmEdiUmCampo.frx":02B5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiUmCampo.frx":03FF
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Bindings        =   "frmEdiUmCampo.frx":0789
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
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
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         Picture         =   "frmEdiUmCampo.frx":079B
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblIns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edição de Categorias"
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
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   2385
      End
   End
End
Attribute VB_Name = "frmEdiUmCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sTab As String
Dim sCod As String
Dim sCampo As String
Dim sNome As String
Dim sMask As String
Public operacao As operacao_um_campo
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadUmCampo
    If ado.RecordCount > 0 Then
        fr.tag = ado!cod
        fr.txtbxNome.Text = ado!campo
        fr.operacao = Me.operacao
        fr.Show 0
    Else
        MsgBox "Não há o que ser Editado!", vbExclamation
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadUmCampo
    fr.operacao = Me.operacao
    fr.Show 0, mdiFrm
    
End Sub

Private Sub cmndbtnSair_Click()
    If IsLoad("frmCadPro") Then
        'P_PreencheCateg frmCadPro.cmbxCat
        'P_PreencheMarca frmCadPro.cmbxMarca
    End If
    Unload Me
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    
    
    If operacao = "1" Then
        sNome = "Seções"
        sTab = "tab_secoes"
        sCod = "cod_sec"
        sCampo = "Nome_sec"
        sMask = "00#"
    ElseIf operacao = "2" Then
        sNome = "Princípios Ativos"
        sTab = "tab_principio"
        sCod = "cod_pri"
        sCampo = "Nome_pri"
        sMask = "00#"
    ElseIf operacao = "3" Then
        sNome = "Categorias"
        sTab = "tab_categorias"
        sCod = "cod_cat"
        sCampo = "Nome_cat"
        sMask = "00#"
    ElseIf operacao = "4" Then
        sNome = "Tipos de Serviço"
        sTab = "tab_tipos"
        sCod = "cod_tip"
        sCampo = "Nome_tip"
        sMask = "00#"
    ElseIf operacao = "5" Then
        sNome = "Destinos de Movimento Interno"
        sTab = "tab_destmovint"
        sCod = "cod_Des"
        sCampo = "Nome_des"
        sMask = "#"
    End If
    
    Me.caption = "Edição de " & sNome
    lblIns.caption = "Edição de " & sNome
    dtg.caption = "Lista de " & sNome
    
    busca
    
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    'Me.Height = 3870
    'Me.Width = 3570
    Set ado = New ADODB.Recordset
    P_CentraForm Me
    
    
End Sub
Sub busca()
    Dim strSQL As String
    If operacao = "1" Then
        strSQL = secao.F_SelectSecao
        
    End If
    If operacao = "2" Then
        strSQL = principio.F_SelectPrincipio
        
    End If
    If operacao = "5" Then
        strSQL = MovInt.F_SelectDest
        
    End If
    FechaObjeto ado
    ado.Open strSQL, Conn_Client
'    dtg.Columns(0).DataField = "cod"
'    dtg.Columns(0).DataField = "cod"
    Set dtg.DataSource = ado
        
        
        
End Sub
