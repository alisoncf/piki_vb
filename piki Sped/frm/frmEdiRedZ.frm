VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiRedZ 
   Caption         =   "Reduções Z"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10935
   Icon            =   "frmEdiRedZ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid grid 
      Height          =   2055
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
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
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   10905
      TabIndex        =   9
      Top             =   4755
      Width           =   10935
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   690
         Left            =   7320
         Picture         =   "frmEdiRedZ.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   690
         Left            =   120
         Picture         =   "frmEdiRedZ.frx":25C4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   690
         Left            =   1920
         Picture         =   "frmEdiRedZ.frx":294E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Excluir"
         Height          =   690
         Left            =   3720
         Picture         =   "frmEdiRedZ.frx":2A98
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   690
         Left            =   9120
         Picture         =   "frmEdiRedZ.frx":2BE2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEfetivar 
         Caption         =   "&Efetivar"
         Height          =   690
         Left            =   5520
         Picture         =   "frmEdiRedZ.frx":2D4D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcBuscar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10905
      TabIndex        =   0
      Top             =   4020
      Width           =   10935
      Begin VB.TextBox txtCaixa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiRedZ.frx":2E97
         Left            =   3840
         List            =   "frmEdiRedZ.frx":2EA7
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Pesquisar"
         Height          =   570
         Left            =   5520
         Picture         =   "frmEdiRedZ.frx":2EDE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   0
         Width           =   435
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&De:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   3840
         TabIndex        =   5
         Top             =   0
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmEdiRedZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmdBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = SpedCupom.BuscaSpedRedZ(txtDe.Text, txtAte.Text, txtCaixa.Text)
    
    abreObjetoAdo ado, strSQL, grid
    
    
End Sub

Private Sub cmndbtnCancelar_Click()
    On Error GoTo trata_erro:
    Dim strSQL As String
    With ado
        If MsgBox("Deseja excluir o movimento selecionado?", vbYesNo + vbCritical) = vbYes Then
            strSQL = "delete from sped_redz where id=" & !id
            conn.Execute strSQL
            ado.Requery
        End If
    End With
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_erro
    With ado
        frmGravaRedZ.tag = !id
        frmGravaRedZ.Show 1
        ado.Requery
    End With
    
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    On Error Resume Next
    frmGravaRedZ.Show 1
    ado.Requery
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtCaixa.Text = strCod_Cai
    Set ado = New ADODB.Recordset
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBuscar
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Private Sub txtDe_GotFocus()
    ganhaFoco txtDe
End Sub
Private Sub txtAte_GotFocus()
    ganhaFoco txtAte
End Sub
