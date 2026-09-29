VERSION 5.00
Begin VB.Form frmEdiItemInvent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editar Item do Inventário"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Sair"
      Height          =   495
      Left            =   6840
      Picture         =   "frmEdiItemInvent.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnConf 
      Caption         =   "&Confirmar"
      Height          =   495
      Left            =   5160
      Picture         =   "frmEdiItemInvent.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtbxTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Text            =   "0,00"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtbxCusto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "0,00"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtbxQtd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custo:"
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   870
   End
   Begin VB.Label lblNomePro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nome do produto"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   600
   End
   Begin VB.Label lblInvent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventário:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmEdiItemInvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub cmndbtnConf_Click()
    Dim strSQL As String
    If MsgBox("Confirma Gravação?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "update tab_invent set ef_pro=" & F_GravaDec(txtbxQtd.Text) & ","
    strSQL = strSQL & " cf_pro=" & F_GravaDec(txtbxCusto.Text)
    strSQL = strSQL & " where cod_pro='" & lblNomePro.tag & "'"
    strSQL = strSQL & " and datahora_inv=" & F_GravaDataTempo(lblInvent.caption) & " "
    conn.Execute strSQL
'    MsgBox "Produto Alterado!", vbInformation
    Unload Me
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    calctotal
End Sub

Private Sub txtbxCusto_GotFocus()
    ganhaFoco txtbxCusto
End Sub

Private Sub txtbxCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCusto_LostFocus()
    PerdeFoco txtbxCusto
    calctotal
End Sub
Sub calctotal()
    On Error Resume Next
    txtbxTotal.Text = FormatVal(CCur(txtbxCusto.Text) * CCur(txtbxQtd.Text))
End Sub
Private Sub txtbxQtd_GotFocus()
    ganhaFoco txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    PerdeFoco txtbxQtd
    calctotal
End Sub
