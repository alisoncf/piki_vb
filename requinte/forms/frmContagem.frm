VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmContagem 
   Caption         =   "Contagem"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   12945
      TabIndex        =   15
      Top             =   7920
      Width           =   13005
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   2143
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
      ColumnCount     =   8
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
         DataField       =   "cod_pro"
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
      BeginProperty Column02 
         DataField       =   "nome_pro"
         Caption         =   "Produto"
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
         DataField       =   "qtd_ant"
         Caption         =   "Qtd Anterior"
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
         DataField       =   "Qtd"
         Caption         =   "Qtd Nova"
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
      BeginProperty Column06 
         DataField       =   "hora"
         Caption         =   "hora"
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
         DataField       =   "guerra_fun"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3075,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   0
      Width           =   13005
      Begin VB.CommandButton cmndbtnEdita 
         Caption         =   "Lançar"
         Height          =   615
         Left            =   10200
         Picture         =   "frmContagem.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtPro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11520
         Picture         =   "frmContagem.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   8880
         Picture         =   "frmContagem.frx":047F
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   7560
         Picture         =   "frmContagem.frx":088C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmContagem.frx":0C16
         Left            =   5280
         List            =   "frmContagem.frx":0C20
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   4920
         Picture         =   "frmContagem.frx":0C32
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   2520
         TabIndex        =   13
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1320
         TabIndex        =   12
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   5280
         TabIndex        =   10
         Top             =   120
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmContagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = "select tp.cod_pro,nome_pro,er_pro,ef_pro, tc.*,tf.guerra_fun from tab_produtos as tp"
    strSQL = strSQL & " left join contagem as tc on "
    
    strSQL = strSQL & " tp.cod_pro=tc.cod_Pro "
    strSQL = strSQL & " left join tab_funcionarios as tf on tc.cod_usu=tf.cod_fun "
    
    
    
    strSQL = strSQL & " where 1>0 "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and data>=" & F_GravaData(txtDe.Text)
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and data<=" & F_GravaData(txtAte.Text)
    End If
    
    If txtPro.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & txtPro.tag & "' "
    End If
    
    abreObjetoAdo ado, strSQL, grid
    
End Sub


Private Sub cmndbtnBuscaRep_Click()
    buscapro
End Sub
Sub buscapro()
    Dim p As Produto
    
    p = Produtos.F_BuscaProdSaida(txtPro.Text)
    If p.cod_pro <> "" Then
        txtPro.Text = p.nome_pro
        txtPro.tag = p.cod_pro
        
        
        
    End If
    
    
End Sub

Private Sub cmndbtnEdita_Click()
    frmContagemNova.Show 1
    busca
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DbgrdToHtml "Contagens", grid, ado
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    resizeForm Me
    Set ado = New ADODB.Recordset
    
    txtDe.Text = F_PrimDiaMes(Date)
    txtAte.Text = F_UltDiaMes(Date)
    
    
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub


