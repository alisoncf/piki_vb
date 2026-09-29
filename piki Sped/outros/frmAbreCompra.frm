VERSION 5.00
Begin VB.Form frmAbreCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura de Compra"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnEscolher 
      Caption         =   "&Escolher Fornecedor"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbreCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEscolher_Click()
    Dim sNum_Ven As String
    'On Error Resume Next
    frmSelectCli.Show 1
    If frmSelectCli.Tag = "" Then
        MsgBox "Não foi selecionado nenhum cliente!", vbExclamation
        Unload frmSelectCli
    Else
        If MsgBox("Confirma Abertura? ", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        Else
            
        End If
        Dim fr As New frmVenda
        With frmSelectCli.dtBusca.Recordset
            sNum_Ven = F_GeraNovaVenda(!cod_cli, "0001")
            fr.lblAbert.caption = Now
            fr.lblNome_Cli.caption = !nome_cli
            fr.lblNome_Cli.Tag = !cod_cli
            fr.caption = "Venda Nº:" & sNum_Ven & " Cliente:" & !cod_cli & "-" & !nome_cli
            fr.lblVendedor.caption = "vendedor"
            fr.Tag = sNum_Ven
            Unload frmSelectCli
            Me.Hide
            fr.Show 0
            Unload Me
        End With
    End If
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

