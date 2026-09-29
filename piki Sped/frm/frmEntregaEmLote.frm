VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEntregaEmLote 
   Caption         =   "Entrega"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10770
   Icon            =   "frmEntregaEmLote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcDados 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   10575
      TabIndex        =   18
      Top             =   600
      Width           =   10575
      Begin VB.TextBox txtComprador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   29
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   5175
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   21
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtFone 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9120
         TabIndex        =   20
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtObs_Ven 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   10335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprador"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5400
         TabIndex        =   30
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   525
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5400
         TabIndex        =   25
         Top             =   840
         Width           =   450
      End
      Begin VB.Label lblFone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   9120
         TabIndex        =   24
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblObsEntrega 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observação:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   915
      End
   End
   Begin VB.TextBox txtRetorno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   6480
      TabIndex        =   17
      Text            =   "00:00"
      Top             =   240
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   240
      TabIndex        =   15
      Top             =   3000
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtHora 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Text            =   "00:00"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   10770
      TabIndex        =   9
      Top             =   5850
      Width           =   10770
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar (F7)"
         Height          =   735
         Left            =   7800
         Picture         =   "frmEntregaEmLote.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair (Esc)"
         Height          =   735
         Left            =   9240
         Picture         =   "frmEntregaEmLote.frx":1351
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.TextBox txtObss 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   10455
   End
   Begin VB.TextBox txtVenda 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox pcEntrega 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1440
      Picture         =   "frmEntregaEmLote.frx":18DF
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   240
      Width           =   240
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Pesquisar Venda"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.TextBox txtEntregador 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdEntregador 
      Height          =   285
      Left            =   5160
      Picture         =   "frmEntregaEmLote.frx":22E1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   315
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Observação"
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   4560
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retorno:"
      Height          =   195
      Left            =   6480
      TabIndex        =   16
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   1800
      TabIndex        =   12
      Top             =   0
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saída:"
      Height          =   195
      Left            =   5640
      TabIndex        =   6
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entregador"
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   780
   End
End
Attribute VB_Name = "frmEntregaEmLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Public status As String

Private Sub cmdConfirmar_Click()
    If txtEntregador.tag = "" Then
        MsgBox "informe o entregador!", vbExclamation
        Exit Sub
    End If
    If status = "2" Then
        If EncaminhaEntrega(txtVenda.Text, txtEntregador.tag, txtData.Text & " " & txtHora.Text) Then
            MsgBox "Entrega encaminhada!", vbInformation
            Unload Me
        End If
    End If
    If status = "3" Then
        If concluientrega(txtVenda.Text, txtEntregador.tag, txtData.Text & " " & txtRetorno.Text, txtObss.Text) Then
            MsgBox "Entrega CONCLUÍDA!", vbInformation
            ado.Requery
            Unload Me
        Else
            MsgBox "erro"
        End If
    End If
End Sub

Private Sub cmdEntregador_Click()
    Dim f As funcionario.funcionario
    On Error Resume Next
    f = funcionario.F_SelectFuncionario(txtEntregador.Text)
    If f.cod_fun <> "" Then
        txtEntregador.Text = f.guerra_fun
        txtEntregador.tag = f.cod_fun
        
        txtHora.SetFocus
    End If
    
End Sub

Private Sub cmdNovo_Click()
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    BuscaVenda
    buscaProd
    If status = "2" Then
        txtRetorno.Enabled = False
        pcDados.BackColor = vbYellow
        txtEntregador.SetFocus

    Else
        txtRetorno.Enabled = True
        pcDados.BackColor = vbWhite
        txtRetorno.Text = Format(Time, "HH:mm")
        txtRetorno.SetFocus
    End If
    bActive = True
End Sub
Sub BuscaVenda()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from " & TAB_VENDAS & " where num_ven='" & txtVenda.Text & "' "
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Venda não encontrada!", vbExclamation
            Exit Sub
        End If
        txtCliente.Text = !nomecli_ven
        txtComprador.Text = Format(!comprador_ven)
        txtBairro.Text = Format(!bairro_ven)
        txtEnd.Text = Format(!endereco_ven)
        txtFone.Text = Format(!fone_ven)
        txtObs_Ven.Text = Format(!obs_ven)
        
        .Close
    End With
    
    strSQL = "select * from historico_entrega where num_ven='" & txtVenda.Text & "' "
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount > 0 Then
            If Not IsNull(!entregador) Then
                If !entregador <> "" Then
                    txtEntregador.Text = F_BuscaUmCampo("guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & !entregador & "'")
                    txtEntregador.tag = !entregador
                End If
            End If
            If Not IsNull(!hora_env) Then
                txtHora.Text = Format(!hora_env)
            End If
            If Not IsNull(!hora_conf) Then
                txtRetorno.Text = Format(!hora_conf)
            End If
            txtObs_Ven.Text = Format(!obs_entrega)
            
        End If
        .Close
    End With
    buscaProd
End Sub
Sub buscaProd()
    Dim strSQL As String
    strSQL = "select  codfab_pro,desc_ite as produto, qtd_ite as quantidade from " & TAB_ITVENDAS & " as ti, "
    strSQL = strSQL & TAB_PRODUTOS & " as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and num_ven='" & txtVenda.Text & "' "
    abreObjetoAdo ado, strSQL, DataGrid1
    
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmdConfirmar_Click
    End If
End Sub

Private Sub Form_Load()
    
    
    
    txtHora.Text = Format(Time, "hh:mm")
    txtData.Text = Date
    Set ado = New ADODB.Recordset
    bActive = False
    
    
    
End Sub

Private Sub txtEntregador_GotFocus()
    P_SelecTexto txtEntregador
End Sub

Private Sub txtEntregador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdEntregador_Click
    End If
End Sub

Private Sub txtEntregador_LostFocus()
    P_AjustaHora txtHora
End Sub

Private Sub txtHora_GotFocus()
    P_SelecTexto txtHora
End Sub

Private Sub txtHora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtHora_LostFocus()
    P_AjustaHora txtHora
End Sub

Private Sub txtRetorno_GotFocus()
    P_SelecTexto txtRetorno
End Sub

Private Sub txtRetorno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtRetorno_LostFocus()
    P_AjustaHora txtRetorno
End Sub

