VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelAnaliseProduto 
   Caption         =   "Análise de Produtos"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4095
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7223
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14940
      TabIndex        =   0
      Top             =   8085
      Width           =   14970
      Begin VB.CommandButton cmdPesq 
         Caption         =   "Pesquisar"
         Height          =   615
         Left            =   13560
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnSubSec 
         Height          =   285
         Left            =   7200
         Picture         =   "frmRelAnaliseProduto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmRelAnaliseProduto.frx":038A
         Left            =   9840
         List            =   "frmRelAnaliseProduto.frx":039A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtbxFab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         MaxLength       =   40
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtbxSecao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtbxSubSec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         MaxLength       =   40
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnSecao 
         Height          =   285
         Left            =   4560
         Picture         =   "frmRelAnaliseProduto.frx":03C7
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnFabr 
         Height          =   285
         Left            =   9480
         Picture         =   "frmRelAnaliseProduto.frx":0751
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblSubCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Subseção:"
         Height          =   195
         Left            =   4920
         TabIndex        =   14
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lblSecCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seção:"
         Height          =   195
         Left            =   2400
         TabIndex        =   13
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   9840
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante/Marca:"
         Height          =   195
         Left            =   7560
         TabIndex        =   10
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   14760
         TabIndex        =   9
         Tag             =   "!"
         Top             =   0
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmRelAnaliseProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub Form_Activate()
    resizeForm Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    ajustagrid
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade grid, Me, pcBusca
    
    
    
    
    
End Sub
Sub ajustagrid()
    With grid
        .rows = 1
        .Cols = 6
        .Col = 0: .Text = "Código"
        .Col = 1: .Text = "Nome"
        
        
        
        
    End With
End Sub

Private Sub SSTab1_DblClick()

End Sub
