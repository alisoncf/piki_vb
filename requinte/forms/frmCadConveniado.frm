VERSION 5.00
Begin VB.Form frmCadConveniado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Conveniados"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxFone 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   16
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox txtbxMatricula 
      Height          =   285
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   22
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox txtbxCNPJ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      MaxLength       =   18
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtbxEndereco 
      Height          =   285
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox txtbxBairro 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   10
      Top             =   1920
      Width           =   4935
   End
   Begin VB.TextBox txtbxCidade 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   12
      Text            =   "Anápolis"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.ComboBox cmbxUF 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtbxCEP 
      Height          =   285
      Left            =   1560
      MaxLength       =   14
      TabIndex        =   18
      Text            =   "75.000-000"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtbxDataNasc 
      Height          =   285
      Left            =   4200
      MaxLength       =   40
      TabIndex        =   20
      Top             =   3000
      Width           =   990
   End
   Begin VB.TextBox txtbxNome_Con 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox txtbxConv 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Height          =   255
      Left            =   6240
      Picture         =   "frmCadConveniado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fecha&r"
      Height          =   735
      Left            =   4680
      Picture         =   "frmCadConveniado.frx":00F9
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Sal&var"
      Height          =   735
      Left            =   2760
      Picture         =   "frmCadConveniado.frx":0243
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4440
      Width           =   1815
   End
   Begin VB.ComboBox cmbxStatus 
      Height          =   315
      ItemData        =   "frmCadConveniado.frx":038D
      Left            =   1560
      List            =   "frmCadConveniado.frx":0397
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Fone:"
      Height          =   195
      Left            =   990
      TabIndex        =   15
      Top             =   2640
      Width           =   405
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Matr. Funcional:"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Width           =   1140
   End
   Begin VB.Label lblCpf 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&CPF:"
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
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Endereço:"
      Height          =   255
      Left            =   285
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Bairro:"
      Height          =   195
      Left            =   930
      TabIndex        =   9
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ci&dade:"
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&UF:"
      Height          =   195
      Left            =   5400
      TabIndex        =   13
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CE&P:"
      Height          =   195
      Left            =   1020
      TabIndex        =   17
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data Nasc:"
      Height          =   195
      Left            =   3240
      TabIndex        =   19
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conveniado:"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Convênio:"
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Left            =   885
      TabIndex        =   23
      Top             =   3720
      Width           =   495
   End
End
Attribute VB_Name = "frmCadConveniado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim c As Cliente
    c = F_BuscaCliente(txtbxConv.Text, True)
    If c.cod_cli <> "" Then
        txtbxConv.Text = c.nome_Cli
        txtbxConv.tag = c.cod_cli
        
        
        
        
    End If
End Sub

Private Sub txtbxCliente_Change()

End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select tc.nome_cli,tcv.* from tab_clientes as tc, tab_conveniados as tcv "
    strSQL = strSQL & " where tc.cod_cli=tcv.cod_cli "
    strSQL = strSQL & " and seq_con=" & Me.tag & " "
    strSQL = strSQL & " and tc.cod_cli='" & txtbxConv.tag & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            txtbxBairro.Text = !bairro_con
            txtbxCEP.Text = !cep_con
            txtbxCidade.Text = !cidade_con
            txtbxCNPJ.Text = !cpf_con
            txtbxConv.Text = !nome_Cli
            txtbxConv.tag = !cod_cli
            cmndbtnBuscar.Enabled = False
            txtbxConv.Enabled = False
            txtbxEndereco.Text = !endereco_con
            txtbxFone.Text = !fone_con
            txtbxMatricula.Text = !matricula_con
            txtbxNome_Con.Text = !nome_con
            P_LocalizaOperacao !status_con, cmbxStatus
            P_LocalizaOperacao !uf_con, cmbxUF
            
        Else
            MsgBox "Não há dados!", vbExclamation
            Unload Me
        End If
    End With
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If Me.tag = "" Then
        strSQL = conveniado.F_GravaConveniado(txtbxConv.tag, txtbxNome_Con.Text, txtbxEndereco.Text, txtbxBairro.Text, _
            cmbxUF.Text, txtbxFone.Text, txtbxCNPJ.Text, txtbxMatricula.Text)

        If strSQL <> "" Then
            If MsgBox("Conveniado gravado com sucesso, Deseja gravar outro?", vbYesNo + vbQuestion) = vbYes Then
            Else
            End If
        Else
            MsgBox "Deu erro!", vbExclamation
        End If
            
    End If
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        carrega
    End If
End Sub

Private Sub txtbxBairro_GotFocus()
    P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBairro_LostFocus()
    P_AjustaTexto txtbxBairro
End Sub

Private Sub txtbxCEP_GotFocus()
    P_SelecTexto txtbxCEP
End Sub

Private Sub txtbxCEP_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCEP_LostFocus()
    P_CEPLostFocus txtbxCEP
End Sub

Private Sub txtbxCidade_GotFocus()
    P_SelecTexto txtbxCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCidade_LostFocus()
    P_AjustaTexto txtbxCidade
End Sub

Private Sub txtbxCNPJ_GotFocus()
    P_SelecTexto txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    P_FormataCPF txtbxCNPJ
End Sub

Private Sub txtbxDataNasc_GotFocus()
    P_SelecTexto txtbxDataNasc
End Sub

Private Sub txtbxDataNasc_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataNasc_LostFocus()
    P_AjustaData txtbxDataNasc
End Sub

Private Sub txtbxEndereco_GotFocus()
    P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEndereco_LostFocus()
    P_AjustaTexto txtbxEndereco
End Sub

Private Sub txtbxFone_GotFocus()
    P_SelecTexto txtbxFone
End Sub

Private Sub txtbxFone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFone_LostFocus()
    P_AjustaTelefone txtbxFone
End Sub

Private Sub txtbxMatricula_GotFocus()
    P_SelecTexto txtbxMatricula
End Sub

Private Sub txtbxMatricula_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMatricula_LostFocus()
    P_AjustaTexto txtbxMatricula
End Sub

Private Sub txtbxNome_Con_GotFocus()
    P_SelecTexto txtbxNome_Con
End Sub

Private Sub txtbxNome_Con_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Con_LostFocus()
    P_AjustaTexto txtbxNome_Con
End Sub
