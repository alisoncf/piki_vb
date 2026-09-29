VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelCliAniv 
   Caption         =   "Relatório de Clientes Aniversariantes"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquiar"
      Height          =   615
      Left            =   2160
      Picture         =   "frmRelCliAniv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   615
      Left            =   3960
      Picture         =   "frmRelCliAniv.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Detalhes"
      Height          =   615
      Left            =   5760
      Picture         =   "frmRelCliAniv.frx":0797
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   9000
      Picture         =   "frmRelCliAniv.frx":0881
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8493
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cod_cli"
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
         DataField       =   "nome_cli"
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
         DataField       =   "aniv"
         Caption         =   "Aniversário"
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
         DataField       =   "Endereco_cli"
         Caption         =   "Endereço"
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
         DataField       =   "nome_set"
         Caption         =   "Bairro"
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
         DataField       =   "nome_cid"
         Caption         =   "Cidade"
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
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtAte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   10
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&De:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Tag             =   "!"
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Até:"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Tag             =   "!"
      Top             =   240
      Width           =   285
   End
End
Attribute VB_Name = "frmRelCliAniv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1





Private Sub cmdPesquisar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tc.nome_cli,concat(right(datanasc_cli,2)," / ", mid(datanasc_cli,6,2)) as aniv,datanasc_cli,"
    strSQL = strSQL & " nome_set,nome_cid from tab_clientes as tc left join tab_setores as ts on ts.cod_set=tc.cod_set "
    strSQL = strSQL & " left join tab_cidades as tc on tc.cod_cid=ts.cod_cid "
    strSQL = strSQL & " where right(datanasc_cli,5) >= " & txt
    
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Clientes Aniversariantes", grid, ado
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
    
    
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte, "dd/mm"
End Sub


Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe, "dd/mm"
End Sub
