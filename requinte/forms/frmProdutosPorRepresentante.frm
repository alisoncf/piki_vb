VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProdutosPorRepresentante 
   Caption         =   "Produtos por Representante"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14295
   Icon            =   "frmProdutosPorRepresentante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   14295
   Begin MSDataGridLib.DataGrid grid 
      Height          =   6615
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   11668
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Produtos"
      ColumnCount     =   9
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "nome_pro"
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
      BeginProperty Column02 
         DataField       =   "cod_rep"
         Caption         =   "Rep"
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
         DataField       =   "nome_rep"
         Caption         =   "Nome Rep."
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
         DataField       =   "pv_pro"
         Caption         =   "Preço"
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
         DataField       =   "estoque_pro"
         Caption         =   "Estoque"
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
         DataField       =   "nome_sec"
         Caption         =   "Seção"
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
         DataField       =   "nome_sub"
         Caption         =   "Subseção"
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
         DataField       =   "pq_pro"
         Caption         =   "Queima"
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
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3750,236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3105,071
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EF9C7C&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   14265
      TabIndex        =   0
      Top             =   7335
      Width           =   14295
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   12120
         Picture         =   "frmProdutosPorRepresentante.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   9960
         Picture         =   "frmProdutosPorRepresentante.frx":02B5
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   735
         Left            =   8280
         Picture         =   "frmProdutosPorRepresentante.frx":06C2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxRep 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtbxSub 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtbxSec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmProdutosPorRepresentante.frx":07BB
         Left            =   6120
         List            =   "frmProdutosPorRepresentante.frx":07D1
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSec 
         Height          =   285
         Left            =   2160
         Picture         =   "frmProdutosPorRepresentante.frx":081C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnSub 
         Height          =   285
         Left            =   2160
         Picture         =   "frmProdutosPorRepresentante.frx":0BA6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtbxProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmndbtnProd 
         Height          =   285
         Left            =   7560
         Picture         =   "frmProdutosPorRepresentante.frx":0F30
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnRep 
         Height          =   285
         Left            =   5520
         Picture         =   "frmProdutosPorRepresentante.frx":12BA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   2640
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto:"
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblSubCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Subseção:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblSecCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seção:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   6120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Label lblFone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   6960
      Width           =   13815
   End
End
Attribute VB_Name = "frmProdutosPorRepresentante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = F_SelectProdutosPorRepresentante(txtbxProd.tag, txtbxRep.tag, _
        txtbxSec.tag, txtbxSub.tag)
    abreObjetoAdo ado, strSQL, grid
    
End Sub

Private Sub cmndbtnImpr_Click()
    Dim extra As String
    extra = "Seção:" & txtbxSec.Text & " :.: Subseção:" & txtbxSub.Text
    extra = extra & " :.: Produto: " & txtbxProd.tag & " :.: Repr: " & txtbxRep.tag
    P_DtgrdToHtml "Listagem de Produtos por Representante", grid, ado, extra, "5"
End Sub

Private Sub cmndbtnProd_Click()
    Dim p As Produto
    p = Produtos.F_BuscaProdSaida(txtbxProd.Text)
    If p.cod_pro <> "" Then
        txtbxProd.Text = p.nome_pro
        txtbxProd.tag = p.cod_pro
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub cmndbtnRep_Click()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxRep.Text)
    If r.cod_Rep <> "" Then
        txtbxRep.Text = r.nome_rep
        txtbxRep.tag = r.cod_Rep
        cmndbtnBuscar_Click
    End If
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSec_Click()
    Dim s As secao.secao
    s = secao.F_BuscaSec(txtbxSec.Text)
    If s.cod_sec <> "" Then
        txtbxSec.Text = s.Nome_Sec
        txtbxSec.tag = s.cod_sec
        cmndbtnBuscar_Click
    End If
End Sub


Private Sub cmndbtnSub_Click()
    Dim s As subsec
    If txtbxSec.tag = "" Then
        MsgBox "Escolha primeiro a seção!", vbExclamation
        txtbxSec.SetFocus
        Exit Sub
    End If
    s = subsecAO.F_BuscaSubSec(txtbxSub.Text, txtbxSub.Text)
    If s.cod_sub <> "" Then
        txtbxSec.Text = s.nome_sub
        txtbxSec.tag = s.cod_sub
        
    End If
End Sub

Private Sub dtPro_Reposition()

End Sub

Private Sub Form_Activate()
    Form_Resize
    If Me.tag <> "" Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    cmbxOrd.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    
    
End Sub

Private Sub txtbxProd_Change()
    txtbxProd.tag = ""
End Sub

Private Sub txtbxProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnProd_Click
    End If
End Sub

Private Sub txtbxRep_Change()
    txtbxRep.tag = ""
End Sub

Private Sub txtbxRep_GotFocus()
    P_SelecTexto txtbxRep
End Sub

Private Sub txtbxSec_Change()
    txtbxSec.tag = ""
End Sub

Private Sub txtbxSec_GotFocus()
    P_SelecTexto txtbxSec
End Sub

Private Sub txtbxSec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSec_Click
    End If
End Sub

Private Sub txtbxSub_Change()
    txtbxSub.tag = ""
End Sub

Private Sub txtbxSub_GotFocus()
    P_SelecTexto txtbxSub
End Sub

Private Sub txtbxSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSub_Click
    End If
End Sub
