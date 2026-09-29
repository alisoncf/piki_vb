VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmValoresServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de Valores de Serviço:"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmValoresServ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Valores:"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtbxComissao_Ser 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Tag             =   "!"
         Text            =   "1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtbxPreco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtbxDescricao_Ite 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label lblCom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissão.(%):"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblPreco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         Height          =   195
         Left            =   3240
         TabIndex        =   9
         Top             =   840
         Width           =   300
      End
      Begin VB.Label lblCod_Pro 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   2760
         TabIndex        =   2
         Tag             =   "!"
         Top             =   -45
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód:"
         Height          =   195
         Left            =   2280
         TabIndex        =   1
         Top             =   0
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Distribuição da Execução:"
      Height          =   1575
      Left            =   4440
      TabIndex        =   11
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmndbtnExcl 
         Height          =   375
         Left            =   3720
         Picture         =   "frmValoresServ.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   390
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnOk 
         Height          =   375
         Left            =   3315
         Picture         =   "frmValoresServ.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtbxPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox cmbxFun 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid msflxGrd 
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   795
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1296
         _Version        =   393216
         FixedCols       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Perc. (%)"
         Height          =   195
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Funcionário:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar:"
      Height          =   375
      Left            =   6840
      TabIndex        =   19
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblComissTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00 "
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Tag             =   "!"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A distribuir:"
      Height          =   195
      Left            =   1440
      TabIndex        =   24
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label lblDistribuido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00 "
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Tag             =   "!"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distribuído(%):"
      Height          =   195
      Left            =   3360
      TabIndex        =   22
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00 "
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Tag             =   "!"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   1560
      Width           =   405
   End
End
Attribute VB_Name = "frmValoresServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnCOnfirma_Click()

    
    Me.Hide
End Sub
Sub P_GravaServ()
    Dim strSQL As String
    Dim iCount As Integer
    With msflxGrd
        For iCount = 1 To msflxGrd.Row - 1
            .Col = 0
            .Row = iCount
            strSQL = "insert into tab_funserv(cod_fun,num_ven,valorcomiss_ser,item_ser,"
            strSQL = strSQL & "percent_ser)Values("
            strSQL = strSQL & "'" & Left(.Text, 2) & "',"
            
        Next
        
    End With
End Sub
Private Sub cmndbtnExcl_Click()
    If msflxGrd.Rows = 1 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
        Exit Sub
    Else
        If msflxGrd.Rows = 2 Then
            msflxGrd.Rows = 1
        Else
            msflxGrd.RemoveItem msflxGrd.Row
        End If
    End If
    
End Sub

Private Sub cmndbtnOk_Click()
    Dim cPerc As Currency
    Dim cDistr As Currency
    cPerc = CCur(txtbxPerc.Text)
    cDistr = CCur(lblDistribuido.caption)
     If CCur(txtbxPerc.Text) + CCur(lblDistribuido) > 100 Then
        MsgBox "Valor maior que o restante!", vbExclamation
        txtbxPerc.SetFocus
        Exit Sub
    End If
    If CCur(txtbxPerc.Text) > 100 Then
        MsgBox "Valor Discrepante!", vbExclamation
        txtbxPerc.SetFocus
        Exit Sub
    End If
    If CCur(txtbxPerc.Text) = 0 Then
        MsgBox "Valor Inválido!", vbExclamation
        txtbxPerc.SetFocus
        Exit Sub
    End If
    msflxGrd.AddItem cmbxFun.Text & Chr(9) & txtbxPerc.Text
    lblDistribuido.caption = FormatVal(CCur(lblDistribuido.caption) + CCur(txtbxPerc.Text))
    
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    P_PreencheFunc cmbxFun
    
    P_CentraForm Me, 0
    
    With msflxGrd
        .Cols = 2
        .Rows = 1
        .ColWidth(0) = 2000
        .ColWidth(1) = 800
        .Col = 0
        .Text = "nome"
        .Col = 1
        .Text = "Perc. %"
        .ColAlignment(0) = 1
    End With
End Sub

Private Sub Label8_Click()

End Sub

Private Sub txtbxComissao_Change()

End Sub

Private Sub txtbxComissao_Ser_GotFocus()
    P_SelecTexto txtbxComissao_Ser
End Sub

Private Sub txtbxComissao_Ser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxComissao_Ser_LostFocus()
    P_ValorLostFocus txtbxComissao_Ser
End Sub

Private Sub txtbxDescricao_Ite_GotFocus()
    P_SelecTexto txtbxDescricao_Ite
End Sub

Private Sub txtbxDescricao_Ite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub
Private Sub txtbxDescricao_Ite_LostFocus()
    P_AjustaTexto txtbxDescricao_Ite
End Sub

Private Sub txtbxPerc_GotFocus()
    P_SelecTexto txtbxPerc
End Sub

Private Sub txtbxPerc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPerc_LostFocus()
    P_ValorLostFocus txtbxPerc
End Sub

Private Sub txtbxPreco_GotFocus()
    P_SelecTexto txtbxPreco
End Sub

Private Sub txtbxPreco_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPreco_LostFocus()
    P_ValorLostFocus txtbxPreco
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
    txtbxQtd.Tag = txtbxQtd.Text
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd
    lblComissTotal.caption = FormatVal((CCur(txtbxComissao_Ser.Text) / 100) * CCur(txtbxPreco.Text) * CCur(txtbxQtd.Text))
    lblTotal = FormatVal(CCur(txtbxPreco.Text) * CCur(txtbxQtd.Text))
    If CCur(txtbxQtd.Text) <> CCur(txtbxQtd.Tag) Then
        lblDistribuido.caption = "0,00"
        msflxGrd.Rows = 1
    End If
End Sub
