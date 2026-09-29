VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBaixarCSV 
   Caption         =   "Escolha um CSV para Baixa Automática"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Text            =   "1"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4048
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
   Begin VB.CommandButton cmdArquivo 
      Caption         =   "Buscar Arquivo"
      Height          =   855
      Left            =   240
      Picture         =   "frmBaixarCSV.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Linhas do Rodapé"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Linhas de Cabeçalho:"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblArquivo 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "frmBaixarCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdArquivo_Click()
    Dim Arquivo As String
    f = BuscaNomeArquivoExterno("*.csv")
    
    If f <> "" Then
        lblArquivo.caption = f
        Carrega lblArquivo.caption
    End If
    
    
    
    
End Sub
Sub Carrega(Arquivo As String)
    Dim conteudo As String
    conteudo = LeArquivoExterno(Arquivo)
    
    
    
End Sub
