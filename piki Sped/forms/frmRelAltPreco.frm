VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelAltPreco 
   Caption         =   "Relatório de Alteração de Preços"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14025
   Icon            =   "frmRelAltPreco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   13995
      TabIndex        =   16
      Top             =   5940
      Width           =   14025
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   3360
         Picture         =   "frmRelAltPreco.frx":0F2A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   720
         Left            =   1800
         Picture         =   "frmRelAltPreco.frx":1074
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   10560
         Picture         =   "frmRelAltPreco.frx":14A3
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "Voltar Preço"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelAltPreco.frx":19F1
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Lista de Produtos Alterados"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "cod_pro"
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
         DataField       =   "nome_pro"
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
         DataField       =   "antigo_pro"
         Caption         =   "Pç Antigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "novo_pro"
         Caption         =   "Pç Novo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "data_alt"
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
         DataField       =   "guerra_fun"
         Caption         =   "Usuário"
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
         DataField       =   "status_alt"
         Caption         =   "Status"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2715,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   585,071
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   14025
      TabIndex        =   0
      Top             =   0
      Width           =   14025
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   11040
         Picture         =   "frmRelAltPreco.frx":1B3B
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdProd 
         Height          =   285
         Left            =   5400
         Picture         =   "frmRelAltPreco.frx":1EE6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdCat 
         Height          =   285
         Left            =   7920
         Picture         =   "frmRelAltPreco.frx":2291
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Height          =   285
         Left            =   10440
         Picture         =   "frmRelAltPreco.frx":263C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marcar:"
         Height          =   195
         Left            =   8280
         TabIndex        =   11
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   5760
         TabIndex        =   8
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E:"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entre:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmRelAltPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdExportar_Click()
    ExportaGridToCSV dtg, ado
    
End Sub

Private Sub cmdPesquisar_Click()
    busca
    
End Sub

Private Sub cmdProd_Click()
    buscaProd
End Sub
Sub buscaProd()

    Dim p As Produto
    p = Produtos.F_BuscaProdSaida(txtProd.Text)
    If p.cod_pro <> "" Then
        txtProd.Text = p.nome_pro
        txtProd.tag = p.cod_pro
        busca
    End If
End Sub
Sub busca()
    Dim strSQL As String
    
    strSQL = "select tf.guerra_fun,tp.nome_pro,ta.* from " & TAB_PRODUTOS & " as tp inner join tab_altpreco as ta on tp.cod_pro=ta.cod_pro "
    strSQL = strSQL & " left join tab_funcionarios as tf on ta.cod_usu=tf.cod_fun "
    
    strSQL = strSQL & " where 1 > 0 "
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and data_alt >=" & F_GravaDataTempo(txtDe.Text & " 00:00:00")
    End If
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and data_alt <=" & F_GravaDataTempo(txtAte.Text & " 23:59:00")
    End If
    If txtProd.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & txtProd.tag & "'"
    End If
    strSQL = strSQL & " order by  data_alt desc, tp.nome_pro"
    abreObjetoAdo ado, strSQL, dtg
    
    
    
End Sub

Private Sub cmdVoltar_Click()
    On Error GoTo trata_erro
    Dim strSQL As String
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
        
    End If
    With ado
        If MsgBox("Deseja alterar o preço do produto: " & !noem_pro & "?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update " & TAB_PRODUTOS & " set pv_pro=" & F_GravaDec(!antigo_pro) & " where cod_pro='" & !cod_pro & "'"
            conn.Execute strSQL
            P_AltPreco !cod_pro, !antigo_pro, !novo_pro
            MsgBox "Preço alterado!", vbInformation
            busca
        End If
    End With
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Relatório de Alteração de Preços", dtg, ado, txtDe.Text & " - " & txtAte.Text
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    txtDe.Text = Date
    txtAte.Text = Date
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    posicionaADireita cmndbtnSair, pcBot
    PosicionaGrade dtg, Me, pcBot
    
End Sub

Private Sub txtDe_GotFocus()
    ganhaFoco txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
Private Sub txtAte_GotFocus()
    ganhaFoco txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub
