VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVisualizaTro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualização de Troca"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "frmVisualizaTro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   8355
   Begin MSDBGrid.DBGrid dbgrdProd 
      Bindings        =   "frmVisualizaTro.frx":08CA
      Height          =   1980
      Left            =   120
      OleObjectBlob   =   "frmVisualizaTro.frx":08DF
      TabIndex        =   0
      Top             =   120
      Width           =   8115
   End
   Begin VB.CommandButton cmndbtnFecha 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   7
      Top             =   2400
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   405
   End
   Begin VB.Label lblAbat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label lblAcr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   2400
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vlr Abatido:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   795
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vlt Acrescido"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   915
   End
End
Attribute VB_Name = "frmVisualizaTro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnFecha_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    busca
End Sub

Private Sub Form_Load()
   ' dtProd.Connect = Banco.Connect
    'P_CentraForm Me
End Sub
Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select tp.nome_pro,tp.ref_pro, ti.* from tab_produtos as tp,"
    strSQL = strSQL & "tab_ittrocas as ti,tab_trocas as tt  "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tt.seq_tro=ti.seq_tro "
    strSQL = strSQL & " and tt.seq_tro='" & Me.tag & "' "
    strSQL = strSQL & " order by tp.nome_pro "
    
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
