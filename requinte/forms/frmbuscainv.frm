VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaInv 
   Caption         =   "Selecionar"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8100
   Icon            =   "frmbuscainv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUltInv 
      Caption         =   "Ult Inv."
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "delete"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Abrir"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecionar"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "data"
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
      BeginProperty Column02 
         DataField       =   "total"
         Caption         =   "Valor"
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
      BeginProperty Column03 
         DataField       =   "de"
         Caption         =   "De"
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
         DataField       =   "ate"
         Caption         =   "Até"
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
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBuscaInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdDel_Click()
    Dim strSQL As String
    If MsgBox("deletar?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "delete from inventario where id=" & ado!id
    Conn.Execute strSQL
    
    strSQL = "delete from inventario_it where inventario=" & ado!id
    Conn.Execute strSQL
    
    ado.Requery
    
    
End Sub

Private Sub cmdEditar_Click()
    On Error GoTo trata_Erro
    With ado
        frmEditarInvent.tag = !id
        frmEditarInvent.txtAte.Text = Format(!ate)
        frmEditarInvent.txtDe.Text = Format(!de)
        frmEditarInvent.lblId.tag = !data
        frmEditarInvent.Show 1
        ado.Requery
    End With
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdGerar_Click()
    Dim id As String
    Dim strSQL As String
    If MsgBox("deseja gerar um inventário agora?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    id = inventario.geraInventario(Now, Date, Date)
    If id <> "" Then
        strSQL = "insert into inventario_it(inventario,codigo,nome,custo,preco,q1,q2,r)"
        strSQL = strSQL & "select '" & id & "', cod_pro,nome_pro,cr_pro,pv_pro,ef_pro,er_pro,r_pro from tab_produtos"
        Conn.Execute strSQL
    End If
    ado.Requery
    
End Sub

Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmdUltInv_Click()
    Dim i As inventario.inventario
    Dim strSQL As String
    i = inventario.SelecionaInventario
    If i.id <> "" Then
        If MsgBox("Confirma alteração da ligação deste inventário com o inventário anterior (" & i.data & ")", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update inventario set id_ultinv=" & i.id
            strSQL = strSQL & " where id=" & ado!id
            Conn.Execute strSQL
        End If
        
    End If
    
    
End Sub

Private Sub Command1_Click()
    Me.tag = "y"
    Me.Hide
End Sub


Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
End Sub
Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select i.*,sum(it.custo*it.q1) as total  from inventario as i inner join "
    strSQL = strSQL & " inventario_it as it on i.id=it.inventario group by it.inventario order by data desc "
    abreObjetoAdo ado, strSQL, grid
    Exit Sub
trata_Erro:
P_MostraErro
End Sub
