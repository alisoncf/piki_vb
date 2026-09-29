VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMensagens 
   Caption         =   "Mensagens pra você"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNova 
      Caption         =   "Nova"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CheckBox chkLidas 
      Caption         =   "Apenas não lidas"
      Height          =   195
      Left            =   5880
      TabIndex        =   5
      Top             =   0
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtMsg 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3240
      Width           =   7335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Responder"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5318
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMensagens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkLidas_Click()
    busca
End Sub

Private Sub cmdNova_Click()
    frmMensagem.Show 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    ajustagrid
End Sub
Sub ajustagrid()
    With grid
        .Cols = 4
        .rows = 1
        .row = 0
        .Col = 0: .Text = "..."
        .Col = 1: .Text = "De"
        .Col = 2: .Text = "Mensagem"
        .Col = 3: .Text = "em"
        .ColWidth(2) = 2000
    End With

End Sub

Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select tm.*,tf.guerra_fun from mensagens as tm left join tab_funcionarios as tf "
    strSQL = strSQL & " on tf.cod_fun=tm.de where para='" & strCod_Usu & "' "
    If chkLidas.Value = 1 Then
        strSQL = strSQL & " and status=1 "
    End If
    
    abreObjetoAdo rPesq, strSQL
    ajustagrid
    With rPesq
        Do While Not .EOF
            strSQL = .AbsolutePosition & Chr(9) & !de & "(" & !guerra_fun & ")" & Chr(9) & !mensagem & Chr(9) & !data
            grid.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid grid, 1
    
End Sub

