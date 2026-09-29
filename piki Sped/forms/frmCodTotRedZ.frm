VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCodTotRedZ 
   Caption         =   "sped_totalizador_redz"
   ClientHeight    =   4680
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   8100
   Begin MSDataGridLib.DataGrid grid 
      Height          =   3135
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5530
      _Version        =   393216
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   8040
      TabIndex        =   1
      Top             =   3825
      Width           =   8100
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   690
         Left            =   5760
         Picture         =   "frmCodTotRedZ.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   690
         Left            =   3600
         Picture         =   "frmCodTotRedZ.frx":016B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   690
         Left            =   1800
         Picture         =   "frmCodTotRedZ.frx":02B5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   690
         Left            =   0
         Picture         =   "frmCodTotRedZ.frx":03FF
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   6480
      TabIndex        =   0
      Top             =   3960
      Width           =   1080
   End
End
Attribute VB_Name = "frmCodTotRedZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmndbtnEdit_Click()
    With ado
        frmCodTotGravar.Tag = !id
        frmCodTotGravar.txtCodigo.Text = !codigo
        frmCodTotGravar.txtDescricao.Text = !Descricao
        frmCodTotGravar.txtId.Text = !id
        frmCodTotGravar.txtNome.Text = !nome
        P_LocalizaOperacao !cod_imp, frmCodTotGravar.cmbxImposto
        frmCodTotGravar.Show 1
    End With
End Sub

Private Sub Form_Activate()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select t.*,i.* from sped_totalizador_redz as t left join tab_impostos as i "
    strSQL = strSQL & " on t.cod_imp=i.cod_imp"
    strSQL = strSQL & " where 1>0"
    abreObjetoAdo ado, strSQL, grid
    
End Sub
Private Sub Form_Load()
        Set ado = New ADODB.Recordset
            
    
    
    
End Sub
