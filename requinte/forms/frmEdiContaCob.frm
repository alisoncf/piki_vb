VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiContaCob 
   Caption         =   "Contas de Cobrança"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   Icon            =   "frmEdiContaCob.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   9180
      TabIndex        =   7
      Top             =   4470
      Width           =   9210
      Begin VB.TextBox txtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   3255
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         DataField       =   "cod_cob"
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
         DataField       =   "nome_cob"
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
         DataField       =   "excluido_cob"
         Caption         =   "Excluído"
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
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3630,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275,024
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9210
      TabIndex        =   0
      Top             =   4965
      Width           =   9210
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiContaCob.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiContaCob.frx":2B2C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "Excluir"
         Height          =   615
         Left            =   3480
         Picture         =   "frmEdiContaCob.frx":2C76
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   7080
         Picture         =   "frmEdiContaCob.frx":2DC0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   5280
         Picture         =   "frmEdiContaCob.frx":2F2B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmEdiContaCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmndbtnEdit_Click()
    On Error Resume Next
    frmCadCob.tag = ado!cod_cob
    frmCadCob.Show 1
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    If strExcluirFin_Usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    If F_BuscaUmCampo("cod_cob", TAB_ITDOC, "cod_cob='" & ado!cod_cob & "' and reman_itd>0") <> "" Then
        MsgBox "Essa conta possui carnês", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma exclusão?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "update tab_itdoc set cod_cob='000' where cod_cob='" & ado!cod_cob & "'"
        Conn.Execute strSQL
        strSQL = "delete from tab_contacob where cod_cob='" & ado!cod_cob & "'"
        Conn.Execute strSQL
        MsgBox "Excluido!", vbInformation
        ado.Requery
    End If
    
    
End Sub

Private Sub cmndbtnImpr_Click()
    P_DbgrdToHtml "Cobradores", grid, ado
End Sub

Private Sub cmndbtnNew_Click()
    frmCadCob.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade grid, Me, pcBusca
    
    posicionaADireita cmndbtnSair, Me
    
    
End Sub

Private Sub txtBuscar_Change()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "select * from tab_contacob where nome_cob like '%" & txtBuscar.Text & "%' "
    
    abreObjetoAdo ado, strSQL, grid
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
