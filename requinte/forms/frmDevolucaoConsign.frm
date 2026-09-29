VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDevolucaoConsign 
   Caption         =   "Devolução:"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   18150
   Icon            =   "frmDevolucaoConsign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   18150
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   18120
      TabIndex        =   21
      Top             =   7635
      Width           =   18150
      Begin VB.CommandButton cmdSoma 
         Caption         =   "Soma"
         Height          =   495
         Left            =   2280
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdZerar 
         Caption         =   "Zerar Estoque"
         Height          =   495
         Left            =   3600
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnExclItem 
         Caption         =   "F4 - Excluir Item"
         Height          =   855
         Left            =   7080
         Picture         =   "frmDevolucaoConsign.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "F6 - Cancelar Tudo"
         Height          =   855
         Left            =   8880
         Picture         =   "frmDevolucaoConsign.frx":1254
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "F7 - Confirmar"
         Height          =   855
         Left            =   10680
         Picture         =   "frmDevolucaoConsign.frx":361B
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "F12 - Sair"
         Height          =   855
         Left            =   14280
         Picture         =   "frmDevolucaoConsign.frx":3EE5
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnListaDev 
         Caption         =   "F7 - Lista Devoluções"
         Height          =   855
         Left            =   12480
         Picture         =   "frmDevolucaoConsign.frx":4050
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Atual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblAtual 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1200
         TabIndex        =   24
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1200
         TabIndex        =   23
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox pcProdAcerto 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   18120
      TabIndex        =   9
      Top             =   735
      Width           =   18150
      Begin VB.CheckBox chkRetif 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Retificadora (F6)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8400
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnInserir 
         Caption         =   "Inserir"
         Height          =   255
         Left            =   7320
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxProdAcerto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtbxQtdAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5520
         TabIndex        =   5
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtbxUnitAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxReman 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblCorrecao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Correção"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8280
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5520
         TabIndex        =   4
         Top             =   0
         Width           =   300
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6120
         TabIndex        =   6
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reman:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4920
         TabIndex        =   2
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência:"
         ForeColor       =   &H00CF614E&
         Height          =   195
         Left            =   1920
         TabIndex        =   16
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblRefAcerto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   2880
         TabIndex        =   15
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Lançamento:"
         ForeColor       =   &H00CF614E&
         Height          =   195
         Left            =   4920
         TabIndex        =   14
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label lblDataConAcerto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   6360
         TabIndex        =   13
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seq. Con:"
         ForeColor       =   &H00CF614E&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   705
      End
      Begin VB.Label lblSeqConAcerto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   120
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   18120
      TabIndex        =   10
      Top             =   0
      Width           =   18150
      Begin VB.CommandButton cmdLembrete 
         Caption         =   "Lembrete"
         Height          =   615
         Left            =   13920
         Picture         =   "frmDevolucaoConsign.frx":413A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   8055
      End
      Begin VB.CommandButton cmndbtnBusca 
         Height          =   360
         Left            =   8160
         Picture         =   "frmDevolucaoConsign.frx":45B4
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos com  Diferença de Preço"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9840
         TabIndex        =   27
         Top             =   0
         Width           =   3285
      End
      Begin VB.Label lblMudanca 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "----"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   9000
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Height          =   5415
      Left            =   120
      TabIndex        =   28
      Top             =   1680
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9551
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
      Caption         =   "Produtos a serem Transferidos"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ref_pro"
         Caption         =   "Ref"
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
         Caption         =   "Produto"
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
         DataField       =   "qtd_dev"
         Caption         =   "Qtd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "seq_con"
         Caption         =   "Nº Consign."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "data_con"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Unit_con"
         Caption         =   "Vlr. Cons."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "pv_pro"
         Caption         =   "Vlr Atual"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "totalcons"
         Caption         =   "Total Cons."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "totalatual"
         Caption         =   "Total Atual"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4034,835
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDevolucaoConsign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
    Dim p As Produtos.ProdutoConsig
    'p = Produtos.F_BuscaProdConsign(txtbxNomeCli.tag, txtbxProdAcerto.Text)
    
Private Sub chkRetif_Click()
    If chkRetif.Value = 1 Then
        pcProdAcerto.BackColor = vbRed
        lblCorrecao.Visible = True
    Else
        pcProdAcerto.BackColor = vbWhite
        lblCorrecao.Visible = False
    End If
End Sub

Private Sub cmdLembrete_Click()
    buscaLembrete txtbxRepresentante.tag, True
End Sub

Private Sub cmdSoma_Click()
    soma
End Sub

Private Sub cmdZerar_Click()
    If MsgBox("Deseja Realmente Devolver todos os produtos deste representante?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    ado.Requery
    If ado.RecordCount > 0 Then
        MsgBox "Existem devoluções em aberto!", vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    Devolver_Todos
    
End Sub

Private Sub cmndbtnBusca_Click()
    buscaRep
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxRepresentante.Text)
    If r.cod_Rep <> "" Then
        txtbxRepresentante.tag = r.cod_Rep
        txtbxRepresentante.Text = r.nome_rep
        buscaDevAbertas
        BuscaMensagem "R", r.cod_Rep
        
        buscaLembrete r.cod_Rep
        
        
        
        txtbxProdAcerto.SetFocus
    End If
End Sub

Private Sub cmndbtnCancel_Click()
    If MsgBox("Deseja excluir todos os itens?", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    End If
    excluiTodos
End Sub

Private Sub cmndbtnConfirmar_Click()
    soma
    If ado.RecordCount = 0 Then
        MsgBox "Não há o que ser devolvido!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma devolução dos itens acima?", vbYesNo + vbQuestion) = vbYes Then
        confirma
    End If
End Sub
Sub confirma()
    Dim strSQL As String
    Dim sSeq As String
    Dim seq_ite As String
    
    If chkRetif.Value = 1 Then
    
        If strCorrigeEstCons_Usu <> "S" Then
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    
        strSQL = "Você optou por efetuar uma correção no estoque!" & Chr(13)
        strSQL = strSQL & "Certifique-se que esta operação é necessária e todos os dados estão corretos!" & Chr(13)
        strSQL = strSQL & "Deseja realmente continuar?"
        If MsgBox(strSQL, vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
    End If
    If F_senhaReturnUsu() = "" Then
        Exit Sub
    End If
    If chkRetif.Value = 0 Then
        If Not F_SenhaRep(txtbxRepresentante.tag) Then
            Exit Sub
        End If
    End If
    
    sSeq = MaxKeySQL("tab_devconsign", "seq_dev", "", "000000#")
    
    
   
    
    
    strSQL = "insert into tab_devconsign(seq_dev,data_dev,cod_rep,cod_usu,retif_dev)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & txtbxRepresentante.tag & "',"
    strSQL = strSQL & "'" & strCod_Usu & "', "
    If chkRetif.Value = 1 Then
        strSQL = strSQL & "'S') "
    Else
        strSQL = strSQL & "'N') "
    End If
    Conn.Execute strSQL
    
    With ado
        .MoveFirst
        Do While Not .EOF
            strSQL = "insert into tab_itdevconsign(seq_Dev,seq_ite,cod_pro,qtd_dev,seq_con)values("
            strSQL = strSQL & "'" & sSeq & "',"
            strSQL = strSQL & .AbsolutePosition + 1 & ","
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & F_GravaDec(!qtd_dev) & ","
            strSQL = strSQL & "'" & !seq_con & "') "
            
            
            
            
            Conn.Execute strSQL
            strSQL = "update tab_consignacoes set reman_con=reman_con-" & F_GravaDec(!qtd_dev) & ", "
            strSQL = strSQL & " reservado_con = reservado_con - " & F_GravaDec(!qtd_dev)
            strSQL = strSQL & " where seq_con='" & !seq_con & "'"
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and cod_rep='" & txtbxRepresentante.tag & "' "
            Conn.Execute strSQL
            
            
             
            
            strSQL = "update tab_estcons set estoque_pro=estoque_pro-" & F_GravaDec(!qtd_dev) & " "
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and cod_rep='" & txtbxRepresentante.tag & "' "
            Conn.Execute strSQL
                
                
            strSQL = "update tab_produtos set"
            If chkRetif.Value = 0 Then
                strSQL = strSQL & " er_pro=er_pro+" & F_GravaDec(!qtd_dev) & ","
            End If
            strSQL = strSQL & " r_pro=r_pro-" & F_GravaDec(!qtd_dev)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
                
                
            .MoveNext
        Loop
        strSQL = "delete from tab_devconsignabertas where cod_rep='" & txtbxRepresentante.tag & "'"
        Conn.Execute strSQL
        MsgBox "Devolução Concluída", vbInformation
        If lblMudanca.caption <> "" Or lblMudanca.caption <> "0" Or lblMudanca.caption <> "----" Then
            MsgBox "Atenção houve " & lblMudanca.caption & " com diferenças de preços!", vbExclamation
            
        End If
        
        imprimeDevConsign sSeq
        
    End With
End Sub

Sub Devolver_Todos()
    Dim strSQL As String
    Dim sSeq As String
    Dim seq_ite As String
    Dim rPesq As New ADODB.Recordset
    If chkRetif.Value = 1 Then
    
        If strCorrigeEstCons_Usu <> "S" Then
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            Exit Sub
        End If
    
        strSQL = "Você optou por efetuar uma correção no estoque!" & Chr(13)
        strSQL = strSQL & "Certifique-se que esta operação é necessária e todos os dados estão corretos!" & Chr(13)
        strSQL = strSQL & "Deseja realmente continuar?"
        If MsgBox(strSQL, vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
    End If
    If F_senhaReturnUsu() = "" Then
        Exit Sub
    End If
    If chkRetif.Value = 0 Then
        If Not F_SenhaRep(txtbxRepresentante.tag) Then
            Exit Sub
        End If
    End If
    
    sSeq = MaxKeySQL("tab_devconsign", "seq_dev", "", "000000#")
    
    
    strSQL = "select  reman_con-reservado_con as restante,"
    strSQL = strSQL & " er_pro-r_pro as er, ts.nome_sec, tu.nome_sub,tu.comiss_sub,tu.descmax_sub, "
    
    strSQL = strSQL & "tp.*,tc.* "
    strSQL = strSQL & " from tab_produtos as tp, "
    strSQL = strSQL & " tab_subsec as tu,tab_secoes as ts,tab_consignacoes as tc  "
    strSQL = strSQL & " where tp.cod_sub=tu.cod_sub "
    strSQL = strSQL & " and tu.cod_sec=ts.cod_sec "
    strSQL = strSQL & " and tc.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tc.cod_rep='" & txtbxRepresentante.tag & "' "
    strSQL = strSQL & " and reman_con>0"
    
    
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    With rPesq
        
        If .RecordCount > 0 Then
            strSQL = "insert into tab_devconsign(seq_dev,data_dev,cod_rep,cod_usu,retif_dev)values("
            strSQL = strSQL & "'" & sSeq & "',"
            strSQL = strSQL & F_GravaDataTempo(Now) & ","
            strSQL = strSQL & "'" & txtbxRepresentante.tag & "',"
            strSQL = strSQL & "'" & strCod_Usu & "', "
            If chkRetif.Value = 1 Then
                strSQL = strSQL & "'S') "
            Else
                strSQL = strSQL & "'N') "
            End If
            Conn.Execute strSQL
        End If
        
        Do While Not .EOF
            strSQL = "insert into tab_itdevconsign(seq_Dev,seq_ite,cod_pro,qtd_dev,seq_con)values("
            strSQL = strSQL & "'" & sSeq & "',"
            strSQL = strSQL & .AbsolutePosition + 1 & ","
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & F_GravaDec(!reman_con) & ","
            strSQL = strSQL & "'" & !seq_con & "') "
            
            
            
            
            Conn.Execute strSQL
            strSQL = "update tab_consignacoes set reman_con=reman_con-" & F_GravaDec(!reman_con) & ", "
            strSQL = strSQL & " reservado_con = reservado_con - " & F_GravaDec(!reman_con)
            strSQL = strSQL & " where seq_con='" & !seq_con & "'"
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and cod_rep='" & txtbxRepresentante.tag & "' "
            Conn.Execute strSQL
            
            
             
            
            strSQL = "update tab_estcons set estoque_pro=estoque_pro-" & F_GravaDec(!reman_con) & " "
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and cod_rep='" & txtbxRepresentante.tag & "' "
            Conn.Execute strSQL
                
                
            strSQL = "update tab_produtos set"
            If chkRetif.Value = 0 Then
                strSQL = strSQL & " er_pro=er_pro+" & F_GravaDec(!reman_con) & ","
            End If
            strSQL = strSQL & " r_pro=r_pro-" & F_GravaDec(!reman_con)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
                
                
            .MoveNext
        Loop
        strSQL = "delete from tab_devconsignabertas where cod_rep='" & txtbxRepresentante.tag & "'"
        Conn.Execute strSQL
        MsgBox "Devolução Concluída", vbInformation
        If lblMudanca.caption <> "" Or lblMudanca.caption <> "0" Or lblMudanca.caption <> "----" Then
            MsgBox "Atenção houve " & lblMudanca.caption & " com diferenças de preços!", vbExclamation
            
        End If
        
        imprimeDevConsign sSeq
        
    End With
End Sub



Private Sub cmndbtnExclItem_Click()
    
    
    If MsgBox("Deseja excluir o item " & ado!nome_pro & "?", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    End If
    excluiItem
End Sub
Sub excluiItem()
    On Error GoTo trata_Erro
    Dim strSQL As String
    With ado
        strSQL = "update tab_consignacoes set reservado_con=reservado_con-" & F_GravaDec(!qtd_dev) & ""
        strSQL = strSQL & " where seq_con='" & !seq_con & "' "
        strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
        
        Conn.Execute strSQL
        
        strSQL = "delete from tab_devconsignabertas where seq_dev='" & ado!seq_dev & "'"
        Conn.Execute strSQL
    End With
    buscaDevAbertas
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub excluiTodos()
    On Error GoTo trata_Erro
    Dim strSQL As String
    With ado
        .MoveFirst
        Do While Not .EOF
            strSQL = "update tab_consignacoes set reservado_con=reservado_con-" & F_GravaDec(!qtd_dev) & ""
            strSQL = strSQL & " where seq_con='" & !seq_con & "' "
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            
            Conn.Execute strSQL
            
            strSQL = "delete from tab_devconsignabertas where seq_dev='" & ado!seq_dev & "'"
            Conn.Execute strSQL
            .MoveNext
        Loop
        
    End With
    buscaDevAbertas
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnInserir_Click()
    If txtbxProdAcerto.tag = "" Then
        MsgBox "Informe o produto que deseja devolver e  tecle <ENTER>", vbInformation
        txtbxProdAcerto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxQtdAcerto.Text) <= 0 Then
        MsgBox "Informe um valor possível para a quantidade!", vbExclamation
        txtbxQtdAcerto.SetFocus
        Exit Sub
    End If
    
    If F_BuscaUmCampo("bloqimpr", "tab_representantes", "cod_rep='" & txtbxRepresentante.tag & "'") = "S" Then
        MsgBox "O representante está com a devolução bloqueada!", vbExclamation
        Exit Sub
    End If
    
    
    gravaItem
    
    
End Sub
Sub gravaItem()
    Dim strSQL As String
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_devconsignabertas", "seq_dev", "", "000000#")
    
    
    
    strSQL = "insert into tab_devconsignabertas(seq_dev,data_dev,cod_pro,qtd_dev,cod_rep,seq_con,cod_usu)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & txtbxProdAcerto.tag & "',"
    strSQL = strSQL & F_GravaDec(txtbxQtdAcerto.Text) & ","
    strSQL = strSQL & "'" & txtbxRepresentante.tag & "',"
    strSQL = strSQL & "'" & lblSeqConAcerto.caption & "',"
    strSQL = strSQL & "'" & strCod_Usu & "')"
    Conn.Execute strSQL
        
    strSQL = "update tab_consignacoes set reservado_con=reservado_con+" & F_GravaDec(txtbxQtdAcerto.Text) & ""
    strSQL = strSQL & " where seq_con='" & lblSeqConAcerto.caption & "' "
    strSQL = strSQL & " and cod_pro='" & txtbxProdAcerto.tag & "' "
    Conn.Execute strSQL
    buscaDevAbertas
    txtbxProdAcerto.SetFocus
    
End Sub
Sub buscaDevAbertas()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select tp.nome_pro,unit_con*qtd_dev as totalcons,pv_pro*qtd_dev as totalatual, tp.pv_pro,tp.ref_pro,tc.*,td.* from tab_devconsignabertas as td, "
    strSQL = strSQL & " tab_produtos as tp,tab_consignacoes as tc "
    strSQL = strSQL & " where tp.cod_pro=td.cod_pro and tc.seq_con=td.seq_con "
    strSQL = strSQL & " and td.cod_rep='" & txtbxRepresentante.tag & "'"
    strSQL = strSQL & " and td.cod_pro=tc.cod_pro "
    strSQL = strSQL & " order by seq_dev desc"
    
    
    abreObjetoAdo ado, strSQL, dtg
    
    
    
End Sub
Sub soma()
    Dim cTotal As Currency
    Dim cAtual As Currency
    Dim i As Integer
  '  On Error Resume Next
    With ado
    If .RecordCount > 0 Then
                .MoveFirst
            End If
        Do While Not .EOF
            
            cTotal = cTotal + !totalcons
            cAtual = cAtual + !totalatual
            If !unit_con <> !pv_pro Then
                i = i + 1
            End If
            
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    lblTotal.caption = FormatVal(cTotal)
    lblAtual.caption = FormatVal(cAtual)
    lblMudanca.caption = i
    lblMudanca.Visible = True
End Sub
Private Sub cmndbtnListaDev_Click()
    P_ListaDevConsign txtbxRepresentante.tag
    
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If strEfetuarDevCons_Usu <> "S" Then
        MsgBox "Não tem permissão para efetuar devolução!", vbExclamation
        Exit Sub
    End If
    
    
    If txtbxRepresentante.tag = "" Then
        txtbxRepresentante.SetFocus
    End If
    buscaDevAbertas
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

    'If p.reman_con > 0 Then
    '    txtbxProdAcerto.Text = StrConv(p.pro.nome_pro, vbProperCase)
'        txtbxProdAcerto.tag = p.pro.cod_pro
'        lblRefAcerto.caption = p.pro.ref_pro
'        txtbxReman.tag = p.reman_con - F_QtdProdutoLancadoAcerto(p.pro.cod_pro)
'        txtbxReman.Text = txtbxReman.tag
'        txtbxQtdAcerto.Text = "1"
'        txtbxUnitAcerto.Text = FormatVal(p.unit_con)
'        txtbxUnitAcerto.tag = FormatVal(p.unit_con)
'        lblDataConAcerto.caption = p.data_con
'        lblSeqConAcerto.caption = p.seq_con
'        txtbxQtdAcerto.SetFocus
'
'    End If

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    Me.top = 10
    Me.Left = 10
    
End Sub

Private Sub pcBotao_Click(Index As Integer)
    If Index = 3 Then
        
    ElseIf Index = 5 Then
    
    ElseIf Index = 6 Then
    
    ElseIf Index = 7 Then
        Unload Me
    Else
    
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dtg.width = Me.width - dtg.Left - 300
    dtg.height = pcBot.top - dtg.top - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ado.Close
End Sub

Private Sub txtbxProdAcerto_GotFocus()
    P_SelecTexto txtbxProdAcerto
End Sub

Private Sub txtbxProdAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProdAcerto
        
    End If
End Sub
Sub buscaProdAcerto()
    
    Dim p As Produtos.ProdutoConsig
    p = Produtos.F_BuscaProdConsign(txtbxRepresentante.tag, txtbxProdAcerto.Text)
    
    If p.reman_con > 0 Then
        txtbxProdAcerto.Text = StrConv(p.pro.nome_pro, vbProperCase)
        txtbxProdAcerto.tag = p.pro.cod_pro
        lblRefAcerto.caption = p.pro.ref_pro
        txtbxReman.tag = p.reman_con '- F_QtdProdutoLancadoDevAberta(p.pro.cod_pro, txtbxRepresentante.tag, p.seq_con)
        txtbxReman.Text = txtbxReman.tag
        txtbxQtdAcerto.Text = "1"
        txtbxUnitAcerto.Text = FormatVal(p.unit_con)
        txtbxUnitAcerto.tag = FormatVal(p.unit_con)
        lblDataConAcerto.caption = p.data_con
        lblSeqConAcerto.caption = p.seq_con
        txtbxQtdAcerto.SetFocus
        
    Else
        MsgBox "Este produto não possui remanescente para ser devolvido", vbExclamation
                
    End If
    
    
    
End Sub


Private Sub txtbxQtdAcerto_GotFocus()
    P_SelecTexto txtbxQtdAcerto
End Sub

Private Sub txtbxQtdAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtdAcerto_LostFocus()
    P_ValorLostFocus txtbxQtdAcerto
    If CCur(txtbxQtdAcerto.Text) > CCur(txtbxReman.Text) Then
        MsgBox "A quantidade que está lançando é maior que a restante!", vbExclamation
        txtbxQtdAcerto.Text = txtbxReman.Text
        txtbxQtdAcerto.SetFocus
        
    End If
End Sub



Private Sub txtbxRepresentante_Change()
    'txtbxRepresentante.tag = ""
End Sub

Private Sub txtbxRepresentante_GotFocus()
    P_SelecTexto txtbxRepresentante
End Sub

Private Sub txtbxRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBusca_Click
    End If
End Sub
