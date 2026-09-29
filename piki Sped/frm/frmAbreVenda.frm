VERSION 5.00
Begin VB.Form frmAbreVenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura de Venda"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   Icon            =   "frmAbreVenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   2955
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   480
      Picture         =   "frmAbreVenda.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   480
      Picture         =   "frmAbreVenda.frx":0694
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   480
      Picture         =   "frmAbreVenda.frx":0AD6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnCons 
      Caption         =   "      Venda A  &Consumidor"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmndbtnEscolher 
      Caption         =   "&Escolher Cliente"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmAbreVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnCons_Click()
    Dim fr As New frmVenda
    
    
        sNum_Ven = F_GeraNovaVenda("00000", strCod_Usu)
        fr.lblHora.caption = Now
        fr.lblNome_Cli.caption = "Venda a Consumidor"
        fr.lblNome_Cli.Tag = "00000"
        fr.caption = "Venda Nº:" & sNum_Ven & " Cliente:" & "00000" & "-Venda a Consumidor" & " [Vendedor " & strNome_Usu & "]"
        fr.lblNome_Ven.caption = strNome_Usu
        fr.lblCPF.Visible = False
        fr.lblCPFCap.Visible = False
        fr.Tag = sNum_Ven

        Me.Hide
        fr.Show 0
        Unload Me
    
End Sub

Private Sub cmndbtnEscolher_Click()
    Dim sNum_Ven As String
    On Error Resume Next
    frmSelectCli.Label1.caption = "&Cliente:"
    frmSelectCli.Show 1
    If frmSelectCli.Tag = "" Then
        MsgBox "Não foi selecionado nenhum cliente!", vbExclamation
        Unload frmSelectCli
    Else
        If MsgBox("Confirma Abertura da Venda? ", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        Else
            
        End If
        Dim fr As New frmVenda
        With frmSelectCli.dtBusca.Recordset
            sNum_Ven = F_GeraNovaVenda(!cod, strCod_Usu)
            fr.lblHora.caption = Now
            fr.lblNome_Cli.caption = !Nome
            fr.lblNome_Cli.Tag = !cod
            fr.caption = "Venda Nº:" & sNum_Ven & " Cliente:" & !cod & "-" & !Nome
            fr.caption = fr.caption & Chr(9) & "[Vendedor: " & strNome_Usu & "]"
            fr.lblNome_Ven.caption = strNome_Usu
            If Format(!tipo_cli) = "J" Then
                fr.lblCPFCap.caption = "CPF:"
                fr.lblCPF.caption = Format(!cpf_cli)
            Else
                fr.lblCPFCap.caption = "CNPJ:"
                fr.lblCPF.caption = Format(!cpf_cli)
            End If
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

Private Sub Form_Load()
    MudaAparencia Me
End Sub
