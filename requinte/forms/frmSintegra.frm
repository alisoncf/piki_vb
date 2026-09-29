VERSION 5.00
Begin VB.Form frmSintegra 
   Caption         =   "Sintegra"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar Arquivo"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox chk75 
      Caption         =   "Tipo 75"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chk74 
      Caption         =   "Tipo 74"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chk10 
      Caption         =   "Tipo 10"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtInv 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "C:\"
      Top             =   2040
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtDataI 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   285
      Left            =   6120
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
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
      Left            =   5520
      TabIndex        =   13
      Top             =   720
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Valor Total:"
      Height          =   195
      Left            =   4560
      TabIndex        =   12
      Top             =   720
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventário:"
      Height          =   195
      Left            =   4560
      TabIndex        =   11
      Top             =   0
      Width           =   750
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho do Arquivo a ser gerado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final:"
      Height          =   195
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial:"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referência:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   825
   End
End
Attribute VB_Name = "frmSintegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
        Dim i As inventario.inventario
    Dim ord As String
    If cmbOrd.ListIndex = 1 Then
        ord = "nome_pro"
    ElseIf cmbOrd.ListIndex = 2 Then
    Else
    
    End If
    i = inventario.SelecionaInventario
    If i.id <> "" Then
        txtInv.Text = i.Data
        
        txtInv.tag = i.id
        'busca  ord
    End If

End Sub

Private Sub cmdGerar_Click()
    
    If chk74.Value = 1 Then
        If txtInv.tag = "" Then
            MsgBox "Para registro tipo 74 é necessário informar o inventário!", vbExclamation
            Exit Sub
        End If
    End If
    If chk75.Value = 1 Then
        If txtInv.tag = "" Then
            MsgBox "Para registro tipo 75 é necessário informar o inventário!", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    
End Sub
