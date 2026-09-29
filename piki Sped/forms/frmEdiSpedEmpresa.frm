VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiSpedEmpresa 
   Caption         =   "Sped Empresas"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14265
   Icon            =   "frmEdiSpedEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7140
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   14265
      TabIndex        =   4
      Top             =   6165
      Width           =   14265
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiSpedEmpresa.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   735
         Left            =   2040
         Picture         =   "frmEdiSpedEmpresa.frx":10C15
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSelec 
         Caption         =   "Selecionar"
         Height          =   735
         Left            =   3960
         Picture         =   "frmEdiSpedEmpresa.frx":11151
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   12360
         Picture         =   "frmEdiSpedEmpresa.frx":115D8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   14265
      TabIndex        =   0
      Top             =   5460
      Width           =   14265
      Begin VB.TextBox txtPesquisar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   9128
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
      Caption         =   "Participantes"
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "codigo"
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
         DataField       =   "NOME"
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
         DataField       =   "CNPJ"
         Caption         =   "CNPJ"
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
         DataField       =   "apelido"
         Caption         =   "Apelido"
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
         DataField       =   "fantasia"
         Caption         =   "Fantasia"
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
         DataField       =   "cod_mun"
         Caption         =   "Município"
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
         DataField       =   "ENDERECO"
         Caption         =   "ENDERECO"
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
      BeginProperty Column07 
         DataField       =   "NUMERO"
         Caption         =   "NUMERO"
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
      BeginProperty Column08 
         DataField       =   "COMPLEMENTO"
         Caption         =   "COMPLEMENTO"
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
      BeginProperty Column09 
         DataField       =   "BAIRRO"
         Caption         =   "BAIRRO"
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
      BeginProperty Column10 
         DataField       =   "cod_cli"
         Caption         =   "cod_cli"
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
      BeginProperty Column11 
         DataField       =   "cod_for"
         Caption         =   "cod_for"
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
      BeginProperty Column12 
         DataField       =   "caminho_certificado"
         Caption         =   "caminho_certificado"
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
      BeginProperty Column13 
         DataField       =   "senha_certificado"
         Caption         =   "senha"
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
      BeginProperty Column14 
         DataField       =   "idcsc"
         Caption         =   "idCSC"
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
      BeginProperty Column15 
         DataField       =   "token"
         Caption         =   "token"
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
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3495,118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2055,118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1755,213
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1620,284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiSpedEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdEdit_Click()
    On Error GoTo trata_Erro:
    frmCadsped_empresa.tag = ado!codigo
    frmCadsped_empresa.Show 1
    Unload frmCadsped_empresa
    ado.Requery
trata_Erro:
    
End Sub

Private Sub cmdNovo_Click()
        frmCadsped_empresa.Show 1
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub cmdSelec_Click()
    
    Me.Hide
End Sub

Private Sub Form_Activate()
    busca
    If Me.tag <> "" Then
        cmdSelec.Visible = True
    Else
        cmdSelec.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub




Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from sped_empresa "
    strSQL = strSQL & " where ( nome like '" & txtPesquisar.text & "%' "
    strSQL = strSQL & " or codigo='" & txtPesquisar.text & "' "
    strSQL = strSQL & " or apelido='" & txtPesquisar.text & "') "
    abreObjetoAdo ado, strSQL, grid
    
    If Me.tag <> "" And ado.RecordCount = 1 Then
        cmdSelec_Click
    End If
    
trata_Erro:
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, pcTop, pcTop
    posicionaADireita cmdSair, pcBot
End Sub

Private Sub grid_DblClick()
    If Me.tag <> "" Then
        cmdSelec_Click
    Else
        cmdEdit_Click
    End If
End Sub

Private Sub txtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
