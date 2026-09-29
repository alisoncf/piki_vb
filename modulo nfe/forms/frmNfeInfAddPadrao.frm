VERSION 5.00
Begin VB.Form frmNfeInfAddPadrao 
   Caption         =   "Informação Adicional de NFe - Textos Padrões"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975
   Icon            =   "frmNfeInfAddPadrao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCFOP 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   1230
   End
   Begin VB.TextBox txtTipoIcms 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   1230
   End
   Begin VB.TextBox txtCredICMS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Text            =   "0,00"
      Top             =   4560
      Width           =   1230
   End
   Begin VB.TextBox txtTexto 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   9375
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   9945
      TabIndex        =   0
      Top             =   5790
      Width           =   9975
      Begin VB.CommandButton cmdUltimo 
         Caption         =   ">>"
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdProx 
         Caption         =   ">"
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "<"
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrimeiro 
         Caption         =   "<<"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   8520
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   7320
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   6120
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Alterar"
         Height          =   495
         Left            =   4920
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   3720
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Label lblEmpresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label lblId 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   720
      TabIndex        =   15
      Top             =   120
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Crédito ICMS"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   930
   End
   Begin VB.Label lblT 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de ICMS"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "CFOP"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   3720
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texto"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   405
   End
End
Attribute VB_Name = "frmNfeInfAddPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bNova As Boolean
Private Sub ado_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Carrega
End Sub
Sub Carrega()
    On Error GoTo trata_Erro
    With ado
        If .RecordCount > 0 Then
            txtCFOP.Text = TiraNull(!CFOP)
            txtCredICMS.Text = FormatVal(!credicms)
            txtTexto.Text = !Texto
            txtTipoIcms.Text = TiraNull(!tipo_icms)
            lblId.caption = !id
            lblEmpresa.caption = "Empresa: " & !cod_emp
        Else
            MsgBox "Não foi encontrada nenhuma informação complementar para essa empresa. Crie uma nova!", vbExclamation
        End If
    End With
    Exit Sub
trata_Erro:
    MsgBox "erro"
End Sub

Private Sub cmdAnterior_Click()
    On Error Resume Next
    ado.MovePrevious
    Carrega
End Sub

Private Sub cmdEdit_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If bNova Then
    
        strSQL = "insert into nfe_infaddpadrao (texto,cfop,tipo_icms,credicms,cod_emp)values("
        strSQL = strSQL & "'" & txtTexto.Text & "',"
        strSQL = strSQL & " '" & txtCFOP.Text & "',"
        strSQL = strSQL & " '" & txtTipoIcms.Text & "',"
        strSQL = strSQL & " " & F_GravaDec(txtCredICMS.Text) & ","
        strSQL = strSQL & "'" & strCod_Emp & "') "
        
        conn.Execute strSQL
    
    
        bNova = False
        cmdEdit.caption = "Alterar"
        cmdNovo.Enabled = True
        cmdAnterior.Enabled = True
        cmdPrimeiro.Enabled = True
        cmdProx.Enabled = True
        cmdUltimo.Enabled = True
        busca
        Carrega
    Else
    
    
        strSQL = "update nfe_infaddpadrao set texto='" & txtTexto.Text & "',"
        strSQL = strSQL & " cfop='" & txtCFOP.Text & "',"
        strSQL = strSQL & " tipo_icms='" & txtTipoIcms.Text & "',"
        strSQL = strSQL & " credicms=" & F_GravaDec(txtCredICMS.Text) & " "
        strSQL = strSQL & " where id=" & ado!id
        conn.Execute strSQL
    End If
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdNovo_Click()
    bNova = True
    cmdEdit.caption = "Gravar"
    cmdNovo.Enabled = False
    cmdAnterior.Enabled = False
    cmdPrimeiro.Enabled = False
    cmdProx.Enabled = False
    cmdUltimo.Enabled = False
End Sub

Private Sub cmdPrimeiro_Click()
    On Error Resume Next
    ado.MoveFirst
    Carrega
End Sub

Private Sub cmdProx_Click()
    On Error Resume Next
    ado.MoveNext
    Carrega
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdUltimo_Click()
    On Error Resume Next
    ado.MoveLast
    Carrega
End Sub

Private Sub Form_Activate()
    busca
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = "select * from nfe_infaddpadrao where cod_emp ='" & strCod_Emp & "' order by id "
    Set ado = New ADODB.Recordset
    abreObjetoAdo ado, strSQL
    Carrega
        
    
    
End Sub

