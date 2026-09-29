VERSION 5.00
Begin VB.Form frmCadTipoPag 
   Caption         =   "Cadastro de Tipo de Pagamento"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   Icon            =   "frmCadTipoPag.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   5895
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   2070
      Width           =   6735
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   4920
         Picture         =   "frmCadTipoPag.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   3360
         Picture         =   "frmCadTipoPag.frx":10DD8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CÛdigo:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DescriÁ„o:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "frmCadTipoPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdSave_Click()
    If strAltformaspag_usu <> "S" Then
        MsgBox "Voc  n„o tem permiss„o para esse acesso", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        If GravaTipoPag(txtCodigo.Text, txtDescricao.Text) Then
            MsgBox "Tipo gravado!", vbInformation
        Else
            MsgBox "erro"
        End If
    Else
        If strAltformaspag_usu <> "S" Then
            MsgBox "Voc  n„o tem permiss„o para esse acesso", vbExclamation
            Exit Sub
        End If
    
        If AlteraTipoPag(txtCodigo.Text, txtDescricao.Text) Then
            MsgBox "Tipo gravado!", vbInformation
        Else
            MsgBox "erro"
        End If
    End If
    
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    P_AjustaTexto txtCodigo
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDescricao_LostFocus()
    P_AjustaTexto txtDescricao
End Sub
