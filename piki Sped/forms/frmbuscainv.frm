VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaInv 
   Caption         =   "Selecionar ou Gerar Inventário"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9960
   Icon            =   "frmbuscainv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9960
      TabIndex        =   4
      Top             =   4560
      Width           =   9960
      Begin VB.CommandButton cmdGerar 
         Caption         =   "Gerar"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Selecionar"
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Abrir"
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   5880
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUltInv 
         Caption         =   "Ult Inv."
         Enabled         =   0   'False
         Height          =   495
         Left            =   6960
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   8400
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtEmpresa 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6376
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
      ColumnCount     =   6
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "cod_emp"
         Caption         =   "Empresa"
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
         DataField       =   "total"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#.##0,00"
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
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEmpresa 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   960
      Picture         =   "frmbuscainv.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lblEmpresa 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label blEmpresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   660
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
    conn.Execute strSQL
    
    strSQL = "delete from inventario_it where inventario=" & ado!id
    conn.Execute strSQL
    
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

Private Sub cmdEmpresa_Click()
    buscaEmpresa
End Sub
Sub buscaEmpresa()
    Dim e As sped_empresa.sped_empresa
    e = sped_empresa.BuscaSpedEmpresa("")
    If e.codigo <> "" Then
        txtEmpresa.Text = e.codigo
        txtEmpresa.tag = e.codigo
        lblEmpresa.caption = e.nome
    End If
End Sub
Private Sub cmdGerar_Click()
    Dim id As String
    Dim strSQL As String
    If MsgBox("deseja gerar um inventário agora?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If txtEmpresa.tag = "" Then
    
        MsgBox "Informe a empresa", vbAbortRetryIgnore
        Exit Sub
    End If
    id = inventario.geraInventario(Now, Date, Date, txtEmpresa.tag)
    If id <> "" Then
        strSQL = "insert into inventario_it(inventario,codigo,nome,custo,preco,q1,q2,r)"
        strSQL = strSQL & "select '" & id & "', tp.cod_pro,tp.nome_pro,tp.cr_pro,tp.pv_pro,e.ef,e.er,reservado from " & TAB_PRODUTOS
        strSQL = strSQL & " as tp inner join estoque as e on tp.cod_pro=e.cod_pro "
        strSQL = strSQL & " where e.cod_emp='" & txtEmpresa.tag & "' "
        conn.Execute strSQL
    End If
    
    busca
    
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
            conn.Execute strSQL
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
    txtEmpresa.Text = strCod_Emp
    txtEmpresa.tag = strCod_Emp
    lblEmpresa.caption = strNome_Emp
End Sub
Sub busca()
    Dim strSQL As String
    P_Status "Pesquisando..."
    strSQL = "select i.*,sum(it.custo*it.q1) as total,i.cod_emp  from inventario as i inner join "
    strSQL = strSQL & " inventario_it as it on i.id=it.inventario "
    strSQL = strSQL & " where cod_emp='" & txtEmpresa.tag & "' "
    strSQL = strSQL & " group by it.inventario order by data desc "
    abreObjetoAdo ado, strSQL, grid
    P_Status ""
    
End Sub
