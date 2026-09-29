VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiPlanoContas 
   Caption         =   "Edição de Plano de Contas (Centros de Custo)"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   11835
   Icon            =   "frmEdiPlanoContas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   11835
   Begin MSDataGridLib.DataGrid grade 
      Height          =   5415
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9551
      _Version        =   393216
      Appearance      =   0
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "tipo_pla"
         Caption         =   "Tipo"
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
         DataField       =   "cod_pla"
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
         DataField       =   "nome_pla"
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
         BeginProperty Column00 
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   5265,071
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11805
      TabIndex        =   6
      Top             =   0
      Width           =   11835
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipo:"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton optnDeb 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Débito"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optnCred 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Crédito"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar:"
         Height          =   195
         Left            =   1920
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   11805
      TabIndex        =   5
      Top             =   6495
      Width           =   11835
      Begin VB.CommandButton cmdSelec 
         Caption         =   "Selecionar"
         Height          =   735
         Left            =   3720
         Picture         =   "frmEdiPlanoContas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnRel 
         Caption         =   "&Relatório"
         Enabled         =   0   'False
         Height          =   735
         Left            =   7320
         Picture         =   "frmEdiPlanoContas.frx":0791
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   9960
         Picture         =   "frmEdiPlanoContas.frx":08ED
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1920
         Picture         =   "frmEdiPlanoContas.frx":0A58
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnExcluir 
         Caption         =   "Desab."
         Enabled         =   0   'False
         Height          =   735
         Left            =   5520
         Picture         =   "frmEdiPlanoContas.frx":0BA2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiPlanoContas.frx":0CEC
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmEdiPlanoContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tipo_pla  As String
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmdSelec_Click()
    If Me.tag <> "" Then
        Me.tag = ado!cod_pla
        Me.Hide
    Else
        MsgBox "Não está em modo de seleção.", vbExclamation
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    With dtPla.Recordset
   
        Dim fr As New frmCadPlano
        
        fr.tag = !cod_pla
        fr.txtbxNOme.Text = !nome_pla
        If !tipo_pla = "C" Then
            fr.optnCred.Value = True
        Else
            fr.optnDeb.Value = True
        End If
        fr.optnCred.Enabled = False
        fr.optnDeb.Enabled = False
        fr.Show 1
    End With
End Sub

Private Sub cmndbtnNovo_Click()
    Dim fr As New frmCadPlano
    fr.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub dbgrdPla_DblClick()
    If dtPla.Recordset.RecordCount > 0 Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dbgrdPla_GotFocus()
    dbgrdPla.MarqueeStyle = 3
End Sub

Private Sub dbgrdPla_HeadClick(ByVal ColIndex As Integer)
    busca
End Sub

Private Sub dbgrdPla_KeyDown(KeyCode As Integer, Shift As Integer)
        
        If KeyCode = vbKeyReturn Then
        dbgrdPla_DblClick
    End If
End Sub

Private Sub dbgrdPla_LostFocus()
    dbgrdPla.MarqueeStyle = 0
End Sub

Private Sub dtPla_Reposition()
    If dtPla.Recordset.RecordCount > 0 Then
        cmndbtnEdit.Enabled = True
        If F_BuscaUmCampo("Cod_Cai", "tab_caixas", " cod_pla='" & _
            dtPla.Recordset!cod_pla & "'") = "" Then
            cmndbtnExcluir.Enabled = True
        End If
    Else
        cmndbtnEdit.Enabled = False
        cmndbtnExcluir.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
   If bActive Then Exit Sub
   
        busca
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
End Sub
Sub busca()
    Dim strSQL As String
    If tipo_pla = "" Then
        tipo_pla = "C"
    End If
    strSQL = "select * from tab_planocontas "
    strSQL = strSQL & " where cod_pla='" & txtbxExpr.Text & "' "
    strSQL = strSQL & " or nome_pla like '%" & txtbxExpr.Text & "%' "
    strSQL = strSQL & " and tipo_pla='" & tipo_pla & "'"
        
    abreObjetoAdo ado, strSQL, grade
End Sub

Private Sub grade_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmdSelec_Click
    End If
End Sub

Private Sub grade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grade_DblClick
    End If
End Sub

Private Sub optnCred_Click()
    If optnCred.Value Then
        tipo_pla = "C"
    Else
        tipo_pla = "D"
    End If
End Sub

Private Sub txtbxExpr_Change()
    busca
End Sub
