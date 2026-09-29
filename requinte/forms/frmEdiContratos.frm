VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiContratos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contratos"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid grid 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
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
      Caption         =   "Contratos"
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
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   5160
      Picture         =   "frmEdiContratos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Height          =   615
      Left            =   1560
      Picture         =   "frmEdiContratos.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnNew 
      Caption         =   "&Novo"
      Height          =   615
      Left            =   120
      Picture         =   "frmEdiContratos.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "frmEdiContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmContrato
    On Error GoTo trata_erro
    If IsLoadTag("frmContrato", adc.Recordset!seq_con) Then
        MsgBox "Este Contrato já está sendo editado em outra janela!", vbExclamation
        Exit Sub
    End If
    fr.tag = adc.Recordset!seq_con
    fr.Show 0
    Exit Sub
trata_erro:
    P_MostraErro
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmContrato
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_contratos order by seq_con"
    
    abreObjetoAdo ado, strSQL, grid
End Sub
Private Sub Form_Load()

    P_CentraForm Me
    busca
End Sub
