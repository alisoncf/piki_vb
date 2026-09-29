VERSION 5.00
Begin VB.Form frmAlteraVendaDadosCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alterar Dados do Cliente"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   Icon            =   "frmAlteraVendaDadosCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVendedor 
      Enabled         =   0   'False
      Height          =   405
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Esc - &Fecha&r"
      Height          =   855
      Left            =   5760
      Picture         =   "frmAlteraVendaDadosCliente.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "F7 - Sal&var"
      Height          =   855
      Left            =   3480
      Picture         =   "frmAlteraVendaDadosCliente.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtCPF 
      Height          =   405
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtTel 
      Height          =   405
      Left            =   840
      TabIndex        =   7
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtNome 
      Height          =   405
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   6855
   End
   Begin VB.TextBox txtNumVen 
      Enabled         =   0   'False
      Height          =   405
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vendedor:"
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CPF:"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Telefone:"
      Height          =   195
      Left            =   30
      TabIndex        =   6
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   285
      TabIndex        =   4
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Venda:"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   240
      Width           =   510
   End
End
Attribute VB_Name = "frmAlteraVendaDadosCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    
    
    If Trim(txtNome.Text) = "" Then
        MsgBox "Informe o nome!", vbExclamation
        txtNome.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTel.Text) = "" Then
        MsgBox "Informe o telefone!", vbExclamation
        txtTel.SetFocus
        Exit Sub
    End If
    
    Vendas.P_AtualizaVendaCliente Me.tag, txtNome.tag, txtNome.Text, txtTel.Text, txtCPF.Text
    
    
    
   Unload Me
    
End Sub

Private Sub Form_Activate()
    busca
End Sub

Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select tv.*, tf.guerra_fun from " & TAB_VENDAS & " as tv left join tab_funcionarios as tf "
    strSQL = strSQL & " on tv.abrevenda_usu=tf.cod_fun"
    strSQL = strSQL & " where tv.num_ven='" & Me.tag & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            txtNumVen.Text = !num_ven
            txtNome.tag = !cod_cli
            txtNome.Text = !nomecli_ven
            txtCPF.Text = Format(!cpfcli_ven)
            txtNumVen.Text = !num_ven
            txtTel.Text = !telefonecli_ven
            txtVendedor.Text = !guerra_fun
            txtVendedor.tag = !abrevenda_usu
        End If
        .Close
    End With
    
    
    
    
End Sub
