VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVincListaDocumentos 
   Caption         =   "Lista de Documentos Vinculados"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9795
   Icon            =   "frmVincListaDocumentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dtgrd 
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7858
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
      Caption         =   "Parcelas Vinculadas"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "seq_itd"
         Caption         =   "Seq.Itd"
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
         DataField       =   "venc_itd"
         Caption         =   "Vencimento"
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
         DataField       =   "valor_itd"
         Caption         =   "Valor"
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
         DataField       =   "valor_vin"
         Caption         =   "Valor Vinculado"
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
         DataField       =   "data_vin"
         Caption         =   "Data"
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
         DataField       =   "hist_orig"
         Caption         =   "Histórico Original"
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
         DataField       =   "doc_original"
         Caption         =   "Doc. Original"
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
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2745,071
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   9765
      TabIndex        =   2
      Top             =   5550
      Width           =   9795
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   7920
         Picture         =   "frmVincListaDocumentos.frx":2CFA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.TextBox txtDoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdBusca 
      Height          =   285
      Left            =   1200
      Picture         =   "frmVincListaDocumentos.frx":2E65
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seq. Doc"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmVincListaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdBusca_Click()
    busca
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = F_SelectDocsVinculados(txtDoc.Text)
    
    abreObjetoAdo ado, strSQL, dtgrd
End Sub
Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
End Sub
