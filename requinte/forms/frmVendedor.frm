VERSION 5.00
Begin VB.Form frmVendedor 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe o Vendedor:"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frmVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxVendedor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblNome_For 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Vendedor:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        ' MsgBox("Deseja realmente não informar o vendedor?", vbYesNo + vbCritical) = vbYes Then
            Me.Tag = ""
            Me.Hide
        'End If
    
    End If
End Sub

Private Sub txtbxVendedor_GotFocus()
    P_SelecTexto txtbxVendedor
End Sub

Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select tf.nome_Fun,tf.cod_fun from tab_funcionarios as tf,"
    strSQL = strSQL & "tab_Cargos as  tc where tf.cod_car=tc.cod_car"
    strSQL = strSQL & " and tf.cod_fun='" & Left(txtbxVendedor.Text, 3) & "' "
    strSQL = strSQL & " and tc.venda_car='S' "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Vendedor não encontrado!", vbCritical
            Me.Tag = ""
            lblNome_For.caption = ""
            txtbxVendedor.SetFocus
        Else
            Me.Tag = !cod_fun
            lblNome_For.caption = !nome_fun
            Me.Hide
        End If
    .Close
    End With
End Sub
