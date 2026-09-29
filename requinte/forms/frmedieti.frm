VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiEti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Etiquetas"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10398
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "cod_eti"
         Caption         =   "..."
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
         DataField       =   "nome_eti"
         Caption         =   "Descrição"
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
         DataField       =   "padrao_eti"
         Caption         =   "Padrão"
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
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2715,024
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmndbtnSelect 
      Caption         =   "&Selecionar"
      Height          =   735
      Left            =   1560
      Picture         =   "frmedieti.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   4440
      Picture         =   "frmedieti.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Height          =   735
      Left            =   3000
      Picture         =   "frmedieti.frx":06D4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnNew 
      Caption         =   "&Novo"
      Height          =   735
      Left            =   120
      Picture         =   "frmedieti.frx":081E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
   End
End
Attribute VB_Name = "frmEdiEti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_erro:

    With ado
        frmCadEti.tag = !cod_eti
        
        frmCadEti.Show 1
    End With
    Exit Sub
trata_erro:
    P_MostraErro
    
End Sub

Private Sub cmndbtnNew_Click()
    frmCadEti.Show 1
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSelect_Click()
    Me.tag = "x"
    Me.Hide
End Sub

Private Sub Form_Activate()
    busca
    If Me.tag = "" Then
        cmndbtnSelect.Visible = False
    Else
        cmndbtnSelect.Visible = True
    End If
End Sub

Private Sub Form_Load()
    
    Set ado = New ADODB.Recordset
        
End Sub
Sub busca(Optional sOrder As String = "nome_eti")
    Dim strSQL As String
    strSQL = "select * from tab_etiquetas "
    If sOrder <> "" Then
        strSQL = strSQL & " order by " & sOrder
    End If
   abreObjetoAdo ado, strSQL, grid
End Sub
