VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelDevProd 
   Caption         =   "Relatório de Devoluções por Produto"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   Icon            =   "frmRelDevProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   9630
   Begin MSDataGridLib.DataGrid grid 
      Height          =   6015
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10610
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   1905
      TabIndex        =   10
      Top             =   4800
      Width           =   1935
      Begin VB.Label lblItens 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   840
         TabIndex        =   14
         Top             =   600
         Width           =   105
      End
      Begin VB.Label lblTotalGerado 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   840
         TabIndex        =   13
         Top             =   120
         Width           =   375
      End
      Begin VB.Label label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gerado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Itens:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelDevProd.frx":08CA
         Left            =   120
         List            =   "frmRelDevProd.frx":08DA
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   945
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelDevProd.frx":0903
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelDevProd.frx":11CD
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelDevProd.frx":14D7
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmRelDevProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tp.nome_pro,tp.cod_pro,tp.ref_pro,tp.unidade_pro,tc.seq_pro,tc.er_pro as estoque,"
    strSQL = strSQL & " sum(td.qtd_dev) as qtd,sum(td.gerado_dev) as gerado "
    strSQL = strSQL & " from tab_produtos as tp, tab_devolucoes as td, "
    strSQL = strSQL & " tab_vendas as tv,tab_estoque as tc where tv.num_ven=td.num_ven"
    strSQL = strSQL & " and tp.cod_pro=td.cod_pro "
    strSQL = strSQL & " and tc.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and td.seq_pro=tc.seq_pro "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " AND left(td.data_dev,10)>= " & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " AND left(td.data_dev,10)<= " & F_GravaData(txtbxDAtaFIn.Text)
    End If
    strSQL = strSQL & " group by td.cod_pro "
    If cmbxOrderBY.ListIndex = 0 Then
        strSQL = strSQL & "order by nome_pro"
    ElseIf cmbxOrderBY.ListIndex = 1 Then
        strSQL = strSQL & "order by tp.cod_pro"
    ElseIf cmbxOrderBY.ListIndex = 2 Then
        strSQL = strSQL & "order by gerado desc"
    ElseIf cmbxOrderBY.ListIndex = 3 Then
        strSQL = strSQL & "order by qtd desc"
    End If
    abreObjetoAdo ado, strSQL, grid
    'soma
End Sub
Sub soma()
    Dim cQtd As Currency
    Dim cGerado As Currency
    With ado
        Do While Not .EOF
            cQtd = cQtd + !qtd
            cGerado = cGerado + !gerado
            .MoveNext
        Loop
        lblItens.caption = cQtd
        lblTotalGerado.caption = FormatVal(cGerado)
        If .RecordCount > 0 Then .MoveFirst
    End With
    
End Sub

Private Sub cmndbtnImprimir_Click()
    MsgBox "Comming Soon ;-)", vbExclamation
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.height = 6855
    Me.width = 9750
    Set ado = New ADODB.Recordset
    
    cmbxOrderBY.ListIndex = 0
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
