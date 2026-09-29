VERSION 5.00
Begin VB.Form frmPosologia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Posologia"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnExcluir 
      Caption         =   "F9 - E&xcluir"
      Height          =   615
      Left            =   120
      Picture         =   "frmPosologia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "ESC - &Sair"
      Height          =   615
      Left            =   6120
      Picture         =   "frmPosologia.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   6120
      Picture         =   "frmPosologia.frx":02B5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnGravar 
      Caption         =   "F7 - Gravar"
      Height          =   615
      Left            =   4560
      Picture         =   "frmPosologia.frx":1FAF
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtbxPosologia 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   7455
   End
   Begin VB.Label lblNomePro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Produto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   525
   End
   Begin VB.Label lblSeqIte 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   330
   End
   Begin VB.Label lblNumVen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "frmPosologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnExcluir_Click()
    Dim strSQL As String
    If Me.tag <> "" Then
        If MsgBox("Deseja Excluir a posologia?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "delete from tab_posologia where num_ven='" & lblNumVen.caption & "' "
            strSQL = strSQL & " and seq_ite=" & lblSeqIte.caption
            Banco.Execute strSQL
            Unload Me
        End If
        
    End If
End Sub

Private Sub cmndbtnGravar_Click()
    If txtbxPosologia.Text = "" Then
            MsgBox "Informe a descrição da posologia!", vbExclamation
            txtbxPosologia.SetFocus
            Exit Sub
        End If
    If Me.tag = "" Then
        If Not Venda.F_AdicionaPosologia(lblNumVen.caption, lblSeqIte.caption, txtbxPosologia.Text) Then
            MsgBox "Erro na gravação!"
            Exit Sub
        Else
            Unload Me
        End If
    Else
        If Not Venda.F_AlteraPosologia(lblNumVen.caption, lblSeqIte.caption, txtbxPosologia.Text) Then
            MsgBox "Erro na alteração!"
            Exit Sub
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    carrega
    
    
End Sub
Sub carrega()
    Dim strSQL As String
    
    Dim rPesq As Recordset
    strSQL = "select tp.nome_pro from tab_produtos as tp, tab_itvendas as ti "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.num_ven='" & lblNumVen.caption & "' "
    strSQL = strSQL & " and  ti.seq_ite=" & lblSeqIte.caption
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
        Else
            lblNomePro.caption = !nome_pro
        End If
        .Close
    End With
    strSQL = "select * from tab_posologia where num_ven='" & lblNumVen.caption & "' "
    strSQL = strSQL & " and seq_ite =" & lblSeqIte.caption
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            
            Me.tag = "x"
            txtbxPosologia.Text = !descricao_pos
        End If
    End With
    
    txtbxPosologia.SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnGravar_Click
    ElseIf KeyCode = vbKeyF9 Then
        cmndbtnExcluir_Click
    End If
End Sub

Private Sub txtbxPosologia_LostFocus()
    P_AjustaTexto txtbxPosologia
End Sub
