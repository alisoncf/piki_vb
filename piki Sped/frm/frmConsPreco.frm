VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConsPreco 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Consulta de Preços de Mercadorias"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14970
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsPreco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   14970
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "frmConsPreco.frx":038A
      ScaleHeight     =   735
      ScaleWidth      =   14970
      TabIndex        =   15
      Top             =   0
      Width           =   14970
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal de Consulta de Preços"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   1080
         TabIndex        =   16
         Top             =   120
         Width           =   5625
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmConsPreco.frx":AF12
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.PictureBox pcPr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   240
      ScaleHeight     =   7785
      ScaleWidth      =   10425
      TabIndex        =   1
      Top             =   960
      Width           =   10455
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         Picture         =   "frmConsPreco.frx":CBDC
         ScaleHeight     =   735
         ScaleWidth      =   4575
         TabIndex        =   21
         Top             =   2760
         Width           =   4575
         Begin VB.Label lblCodBar 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0000000000000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   630
            Left            =   240
            TabIndex        =   22
            Tag             =   "!"
            Top             =   0
            Width           =   3705
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   120
         Picture         =   "frmConsPreco.frx":133A1
         ScaleHeight     =   750
         ScaleWidth      =   9750
         TabIndex        =   18
         Top             =   5520
         Width           =   9750
         Begin VB.Label lblCategoria 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Left            =   240
            TabIndex        =   20
            Tag             =   "!"
            Top             =   120
            Width           =   105
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   120
         Picture         =   "frmConsPreco.frx":1A547
         ScaleHeight     =   750
         ScaleWidth      =   9750
         TabIndex        =   17
         Top             =   4080
         Width           =   9750
         Begin VB.Label lblMarca 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Left            =   240
            TabIndex        =   19
            Tag             =   "!"
            Top             =   120
            Width           =   105
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   5280
         Picture         =   "frmConsPreco.frx":216ED
         ScaleHeight     =   735
         ScaleWidth      =   4575
         TabIndex        =   13
         Top             =   2760
         Width           =   4575
         Begin VB.Label lblPreço 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   630
            Left            =   240
            TabIndex        =   14
            Tag             =   "!"
            Top             =   0
            Width           =   990
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   120
         Picture         =   "frmConsPreco.frx":28170
         ScaleHeight     =   750
         ScaleWidth      =   9750
         TabIndex        =   11
         Top             =   1440
         Width           =   9750
         Begin VB.Label lblNomePro 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   120
            TabIndex        =   12
            Tag             =   "!"
            Top             =   120
            Width           =   9360
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         Picture         =   "frmConsPreco.frx":2F316
         ScaleHeight     =   735
         ScaleWidth      =   4575
         TabIndex        =   9
         Top             =   360
         Width           =   4575
         Begin VB.TextBox txtbxExpr 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000000&
            Height          =   435
            Left            =   120
            MaxLength       =   40
            TabIndex        =   10
            Top             =   120
            Width           =   4215
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   120
         Picture         =   "frmConsPreco.frx":35D99
         ScaleHeight     =   915
         ScaleWidth      =   2220
         TabIndex        =   2
         Top             =   6840
         Width           =   2220
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informe aqui o produto:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   2910
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Barras:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   2250
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5280
         TabIndex        =   5
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   5040
         Width           =   1305
      End
   End
   Begin VB.PictureBox pcGrid 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8250
      Left            =   11115
      ScaleHeight     =   8220
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   735
      Visible         =   0   'False
      Width           =   3855
      Begin MSDataGridLib.DataGrid dbgrd 
         Bindings        =   "frmConsPreco.frx":3A1A9
         Height          =   2895
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5106
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
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4394,835
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmConsPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnBusca_Click()
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    AjustaTexto txtbxExpr
    For iCount = 1 To 3

        strSQL = "select distinct tg.nome_cat, tb.codbarra_pro, te.nome_mar, tp.* "
     
        strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp,tab_categorias as tg, tab_codbarras as tb, "
        strSQL = strSQL & TAB_MARCAS & " as te "
        If iCount = 1 Then
            strSQL = strSQL & " where tb.codbarra_pro = '"
            strSQL = strSQL & txtbxExpr.Text & "' "
        ElseIf iCount = 2 Then
            strSQL = strSQL & " where tp.cod_pro = '"
            strSQL = strSQL & txtbxExpr.Text & "' "
        Else
            strSQL = strSQL & " where nome_pro like '%"
            strSQL = strSQL & txtbxExpr.Text & "%' "
        End If
        strSQL = strSQL & " and tp.cod_pro=tb.cod_pro"
        strSQL = strSQL & " and tg.cod_cat=tp.cod_cat and te.cod_mar=tp.cod_mar "
        strSQL = strSQL & " order by nome_pro "
        dtPro.RecordSource = strSQL
        dtPro.Refresh
        If dtPro.Recordset.RecordCount > 0 Then
            ganhaFoco txtbxExpr
            txtbxExpr.SetFocus
            Exit For
        End If
    Next
    If dtPro.Recordset.RecordCount > 1 Then
        pcGrid.Visible = True
        Form_Resize
    Else
        pcGrid.Visible = False
    End If
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 1
End Sub

Private Sub dtPro_Reposition()
    With dtPro.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        lblNomePro.caption = !nome_pro
        lblPreço.caption = "R$ " & Format(!pv_pro, "##,###,###0.00")
        lblMarca.caption = !nome_mar
        lblCategoria.caption = !nome_cat
        lblCodBar.caption = !codbarra_pro
    End With
End Sub

Private Sub Form_Activate()
    ''MudaAparencia Me
    txtbxExpr.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        
    End If
End Sub

Private Sub Form_Load()
    adc.ConnectionString = strConnect
    'P_CentraForm Me
'    cmbxBusca.ListIndex = 2
'    cmbxOrd.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Height = pcGrid.Height - dbgrd.Top - 100
    P_CentraObejct pcPr, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If App.EXEName = "piki_pesq" Then
        P_Fim
    End If
End Sub

Private Sub Label4_Click()
    P_Fim
End Sub

Private Sub txtbxExpr_GotFocus()
    'GanhaFoco txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        cmndbtnBusca_Click
        
        dbgrd.SetFocus
        txtbxExpr.SetFocus
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dbgrd.SetFocus
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    'AjustaTexto txtbxExpr
End Sub
