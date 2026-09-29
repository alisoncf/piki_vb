VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelCarnes 
   Caption         =   "Relatório de Carnês"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16485
   Icon            =   "frmRelCarnes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   16485
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcTotal 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EF9C7C&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   16455
      TabIndex        =   52
      Top             =   6675
      Width           =   16485
      Begin VB.CheckBox chkAberto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Parc. em Aberto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkProrr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Parcelas Prorrogadas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   64
         Top             =   240
         Width           =   2775
      End
      Begin VB.PictureBox pcTrab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2880
         ScaleHeight     =   825
         ScaleWidth      =   4905
         TabIndex        =   60
         Top             =   0
         Width           =   4935
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Trabalho:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   45
            Width           =   675
         End
         Begin VB.Label lblNomeOrg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Orgão"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   840
            TabIndex        =   62
            Top             =   0
            Width           =   525
         End
         Begin VB.Label lblEnd 
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   4740
         End
      End
      Begin VB.CheckBox chkPrej 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Prejuízo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   59
         Top             =   480
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   7920
         ScaleHeight     =   825
         ScaleWidth      =   1305
         TabIndex        =   56
         Top             =   0
         Width           =   1335
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefones:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   58
            Top             =   0
            Width           =   1110
         End
         Begin VB.Label lblTel 
            BackStyle       =   0  'Transparent
            Caption         =   "0000000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1365
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   9360
         ScaleHeight     =   825
         ScaleWidth      =   705
         TabIndex        =   53
         Top             =   0
         Width           =   735
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vias:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   55
            Top             =   0
            Width           =   525
         End
         Begin VB.Label lblVias 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   135
         End
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   67
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   1320
         TabIndex        =   66
         Top             =   540
         Width           =   405
      End
   End
   Begin VB.Data dtNot 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   16485
      TabIndex        =   11
      Top             =   0
      Width           =   16485
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   600
         Left            =   14520
         Picture         =   "frmRelCarnes.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   60
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmRelCarnes.frx":0F8C
         Left            =   14520
         List            =   "frmRelCarnes.frx":0F9C
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmRelCarnes.frx":0FC8
         Left            =   2400
         List            =   "frmRelCarnes.frx":0FD5
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   21
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtbxTipoCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   60
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxSeqItd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   60
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbxEspDoc 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmRelCarnes.frx":1007
         Left            =   10560
         List            =   "frmRelCarnes.frx":1017
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtbxRepr 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   12480
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxHist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   60
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxGrupo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9360
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtbxOrgao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11520
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmndbtnGru 
         Height          =   285
         Left            =   11160
         Picture         =   "frmRelCarnes.frx":1038
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   8880
         Picture         =   "frmRelCarnes.frx":13C2
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnRep 
         Enabled         =   0   'False
         Height          =   285
         Left            =   14040
         Picture         =   "frmRelCarnes.frx":174C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnOrg 
         Height          =   285
         Left            =   14040
         Picture         =   "frmRelCarnes.frx":1AD6
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnTipoCob 
         Height          =   285
         Left            =   7920
         Picture         =   "frmRelCarnes.frx":1E60
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnCob 
         Height          =   285
         Left            =   10200
         Picture         =   "frmRelCarnes.frx":21EA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico:"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Tag             =   "!"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   49
         Tag             =   "!"
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Doc."
         Height          =   195
         Left            =   2160
         TabIndex        =   48
         Tag             =   "!"
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar Por:"
         Height          =   195
         Left            =   14520
         TabIndex        =   47
         Tag             =   "!"
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   2400
         TabIndex        =   46
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cobrança:"
         Height          =   195
         Left            =   6000
         TabIndex        =   45
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label lblTipoCob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança:"
         Height          =   195
         Left            =   8280
         TabIndex        =   44
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lblTerCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   4680
         TabIndex        =   43
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Doc:"
         Height          =   195
         Left            =   4320
         TabIndex        =   42
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Parc:"
         Height          =   195
         Left            =   5160
         TabIndex        =   41
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblCPF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   195
         Left            =   8160
         TabIndex        =   40
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie de Doc:"
         Height          =   195
         Left            =   10560
         TabIndex        =   39
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label lblRepr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representante:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   12480
         TabIndex        =   38
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   3360
         TabIndex        =   37
         Tag             =   "!"
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblGru 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Clientes:"
         Height          =   195
         Left            =   9360
         TabIndex        =   36
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label lblOrgao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Órgão:"
         Height          =   195
         Left            =   11520
         TabIndex        =   35
         Top             =   600
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   16425
      TabIndex        =   0
      Top             =   7530
      Width           =   16485
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   595
         Left            =   5040
         Picture         =   "frmRelCarnes.frx":2574
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   595
         Left            =   14760
         Picture         =   "frmRelCarnes.frx":2981
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   595
         Left            =   0
         Picture         =   "frmRelCarnes.frx":2ACB
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   595
         Left            =   1680
         Picture         =   "frmRelCarnes.frx":2C15
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   595
         Left            =   11520
         Picture         =   "frmRelCarnes.frx":2F9F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   595
         Left            =   3360
         Picture         =   "frmRelCarnes.frx":30E9
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnCarne 
         Caption         =   "Emitir &Carnê"
         Height          =   595
         Left            =   6720
         Picture         =   "frmRelCarnes.frx":3233
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnContrato 
         Caption         =   "Emitir Contrato"
         Height          =   595
         Left            =   8400
         Picture         =   "frmRelCarnes.frx":332F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEditer 
         Caption         =   "&Editar Terceiro"
         Height          =   595
         Left            =   13080
         Picture         =   "frmRelCarnes.frx":342B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbrnPrejuizo 
         Caption         =   "Prejuízo"
         Height          =   595
         Left            =   10080
         Picture         =   "frmRelCarnes.frx":3587
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelCarnes.frx":3B11
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "frmRelCarnes.frx":3B25
      TabIndex        =   68
      Top             =   1320
      Width           =   16215
   End
End
Attribute VB_Name = "frmRelCarnes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub chkAberto_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub chkProrr_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_Click()
    If Not bActive Then Exit Sub
    If cmbxBuscar.ListIndex = 3 Then
        dbgrd.Columns(6).caption = "Boletos"
        dbgrd.Columns(6).Visible = True
        dbgrd.Columns(6).DataField = "num_bol"
    ElseIf cmbxBuscar.ListIndex = 2 Then
        txtbxExpr.Text = "0"
    Else
        dbgrd.Columns(6).Visible = False
    End If
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxEspDoc_Click()
    If Left(cmbxEspDoc.Text, 1) = "A" Then
        cmndbtnRep.Enabled = True
        txtbxRepr.Enabled = True
        lblRepr.Enabled = True
        txtbxRepr.SetFocus
    Else
        cmndbtnRep.Enabled = False
        txtbxRepr.Enabled = False
        lblRepr.Enabled = False
    End If
End Sub

Private Sub cmbxOrder_Click()
    If Not bActive Then Exit Sub
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    If cmbxTipo.ListIndex = 0 Then
        cmbxTipoTer.ListIndex = 0
        cmbxTipo.Enabled = True
    Else
        cmbxTipo.ListIndex = -1
        cmbxTipo.Enabled = False
    End If
    If Not bActive Then Exit Sub
    
    
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub


Private Sub cmndbtnBaixar_Click2()
    Dim cJuros As Currency
    Dim iCount As Integer
    Dim jCount As Integer
    On Error GoTo trata_Erro
    With dtNot.Recordset
        cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
        
        frmBaixaDocReq.tag = !cod_ter
        
                
        frmBaixaDocReq.txtbxCliente.Text = !nome
        frmBaixaDocReq.txtbxCliente.tag = !Codigo
        frmBaixaDocReq.txtbxJuros.Text = FormatVal(cJuros)
        frmBaixaDocReq.txtbxReman.Text = frmListaDoc.lblTotal.caption
        frmBaixaDocReq.txtbxValor_Par.Text = frmListaDoc.lblTotal.caption
        'frmBaixaDocReq.lblParc.caption = "Parcelas: " & frmListaDoc.msflxgrd.rows
        frmBaixaDocReq.txtbxTotalBaixa.Text = "0,00"
        frmBaixaDocReq.cmndbtnCOnfirma.tag = "R"
        frmBaixaDocReq.lblCPF.caption = lblCPF.caption
        frmBaixaDocReq.msflxParcelas.Cols = frmListaDoc.msflxgrd.Cols
        frmBaixaDocReq.msflxParcelas.rows = frmListaDoc.msflxgrd.rows + 1
        For iCount = 0 To frmListaDoc.msflxgrd.rows - 1
            frmBaixaDocReq.msflxParcelas.Row = iCount + 1
            frmListaDoc.msflxgrd.Row = iCount
            For jCount = 0 To frmListaDoc.msflxgrd.Cols - 1
                frmBaixaDocReq.msflxParcelas.Col = jCount
                frmListaDoc.msflxgrd.Col = jCount
                frmBaixaDocReq.msflxParcelas.Text = frmListaDoc.msflxgrd.Text
            Next
            
        Next
        frmBaixaDocReq.Show 1
        Unload frmBaixaDocReq
        Unload frmListaDoc
        dtNot.Refresh
    End With
trata_Erro:
    P_MostraErro
    Exit Sub
End Sub

Private Sub cmbxTipoTer_Click()
    lblTerCap.caption = "&" & cmbxTipoTer.Text
    If Left(cmbxTipoTer.Text, 1) = "C" Then
        txtbxGrupo.Enabled = True
        txtbxOrgao.Enabled = True
        lblOrgao.Enabled = True
        lblGru.Enabled = True
        txtbxGrupo.Enabled = True
        txtbxOrgao.Enabled = True
    Else
        txtbxGrupo.Enabled = False
        txtbxOrgao.Enabled = False
        lblOrgao.Enabled = False
        lblGru.Enabled = False
        txtbxGrupo.Enabled = False
        txtbxOrgao.Enabled = False
    
    End If
End Sub

Private Sub cmndbrnPrejuizo_Click()
    On Erro GoTo trata_Erro
    
    With dtNot.Recordset
        If Not F_senha() Then
            Exit Sub
        End If
        If strExcldoc_usu <> "S" Then
            MsgBox "Você não tem permissão para esta operação!", vbExclamation
            Exit Sub
        End If
        If MsgBox("Confirma Inclusão do Carnê nos Prejuízos?", vbYesNo + vbQuestion) = vbYes Then
            If !prejuizo_itd = "S" Then
                strSQL = "Update tab_itdoc set prejuizo_itd='N' where seq_itd='" & !seq_itd & "'"
                Banco.Execute strSQL
            Else
                strSQL = "Update tab_itdoc set prejuizo_itd='S' where seq_itd='" & !seq_itd & "'"
                Banco.Execute strSQL
            End If
        End If
        cmndbtnBuscar_Click
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnBaixar_Click()
    On Error GoTo trata_Erro
    If strBaixadoc_usu <> "S" Then
        MsgBox "Não tem permissão pra baixar documento!", vbExclamation
        Exit Sub
    End If
    With dtNot.Recordset
        frmBaixaDocReq.txtbxCliente.Text = !nome
        
        frmBaixaDocReq.txtbxCliente.tag = !Codigo
        frmBaixaDocReq.lblCPF.caption = !CPF
        frmBaixaDocReq.txtbxSeq.Text = !seq_doc
        frmBaixaDocReq.Show 1
        
        Unload frmBaixaDocReq
    End With
trata_Erro:
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    
    
    P_Status "Somando..."
    
    strSQL = "select sum(ti.reman_itd) as reman, td.*,tc.nome_cli from tab_clientes as tc, tab_doc as td, tab_itdoc as ti "
    strSQL = strSQL & " where tc.cod_cli=td.cod_cli"
    
    
    
    soma
    P_Status
    bActive = True
    Exit Sub
trata_Erro:
    P_Status
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnCarne_Click()
    imprimeCarne dtNot.Recordset!seq_doc
End Sub

Private Sub cmndbtnCob_Click()
    buscaContaCob
End Sub
Sub buscaContaCob()
    Dim c As ContaCob
    c = F_BuscaContaCob(txtbxContaCob.Text)
    If c.cod_cob <> "" Then
        txtbxContaCob.Text = c.nome_cob
        txtbxContaCob.tag = c.cod_cob
    End If
    
    
End Sub

Private Sub cmndbtnContrato_Click()
    'imprimeContratoHTML dtNot.Recordset!seq_doc
    imprimeContrato dtNot.Recordset!seq_doc
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEdiLan
    On Error GoTo trata_Erro
    
    
    With dtNot.Recordset
        If .RecordCount < 0 Then
            Exit Sub
        End If
        fr.tag = !seq_itd
        
        fr.Show 1
        RefreshDataPosition dtNot
        Exit Sub
        
        fr.lblData.caption = !data_doc
        fr.lblNUM_Doc.caption = !num_doc
        fr.lblParc.caption = !parcela_itd & "/" & !parcelas_doc
        fr.lblSeq_Doc.caption = !seq_doc
        fr.lblSeq_Itd.caption = !seq_itd
        fr.lblValor.caption = FormatVal(!valor_itd)
        fr.txtbxHist.Text = !historico_doc
        fr.txtbxNome.Text = !nome
        fr.txtbxNome.tag = !Codigo
        fr.txtbxVenc.Text = !venc_itd
        fr.txtbxValor.Text = FormatVal(!reman_itd)
        fr.txtbxBoleto.Text = !boleto_itd
        fr.txtbxClass.tag = !cod_pla
        fr.txtbxClass.Text = !nome_pla
        fr.txtbxContaCob.Text = !nome_cob
        fr.txtbxContaCob.tag = !cod_cob
        fr.txtbxTipoCob.Text = !nome_tip
        fr.txtbxTipoCob.tag = !cod_tip
        fr.lblTipo.caption = "Crédito"
        fr.Show 1
        dtNot.Refresh
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnExcl_Click()
    If strExcldoc_usu <> "S" Then
        MsgBox "Você não pode excluir este documento!", vbExclamation
        Exit Sub
    End If
    
    With dtNot.Recordset
        If !num_ven = "" Then
            If F_BuscaUmCampo("reman_itd", "tab_itdoc", "reman_itd<valor_itd and seq_doc='" & !seq_doc & "'") = "" Then
                If MsgBox("Deseja realmente excluir o documento e suas parcelas?", vbYesNo + vbCritical) = vbYes Then
                    strSQL = "delete from tab_doc where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    strSQL = "delete from tab_itdoc where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    dtNot.Refresh
                End If
            Else
                MsgBox "Este documento não pode ser excluído por que já teve uma parcela baixada!", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Este documento não pode ser excluído pois se refere a uma venda!", vbExclamation
            
        End If
    End With
    
    
    
    
    
    
End Sub

Private Sub cmndbtnGru_Click()
    Dim g As grupoCli
    g = F_BuscaGrupoCli
    txtbxGrupo.Text = g.nome_gru
    txtbxGrupo.tag = g.cod_gru
    
End Sub

Private Sub cmndbtnImpr_Click()
    P_DbgrdToHtml "Contas a Receber de " & txtbxDataIni.Text & " Até " & txtbxdataFin.Text, dbgrd, dtNot, _
    "Impresso: " & Now & " por " & strNome_Usu, "5,6"
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
End Sub

Private Sub cmndbtnOrg_Click()
    Dim o As orgao.orgao
    o = F_BuscaOrgao(txtbxOrgao.Text, txtbxGrupo.tag)
    
    txtbxOrgao.Text = o.nome_org
    txtbxOrgao.tag = o.cod_org
    
    
    
End Sub

Private Sub cmndbtnRep_Click()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxRepr.Text)
        If r.cod_Rep <> "" Then
            txtbxRepr.Text = r.nome_rep
            txtbxRepr.tag = r.cod_Rep
            
        End If
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.BorderStyle = 0
End Sub

Private Sub cmndbtnTerc_Click()
    
    
        If cmbxTipoTer.ListIndex = 0 Then
            buscaCLi
        ElseIf cmbxTipoTer.ListIndex = 2 Then
            buscaFun
        Else
            buscaRep
        End If
    
    
    cmndbtnBuscar_Click
End Sub
Sub buscaFun()
    Dim c As funcionario.funcionario
    c = funcionario.F_SelectFuncionario(txtbxTerceiro.Text)
        If c.cod_fun <> "" Then
            txtbxTerceiro.Text = c.nome_fun
            txtbxTerceiro.tag = c.cod_fun
            lblCPF.caption = c.cpf_fun
        End If
End Sub

Sub buscaCLi()
    Dim c As Cliente
    c = F_BuscaCliente(txtbxTerceiro.Text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.Text = c.nome_Cli
            txtbxTerceiro.tag = c.cod_cli
            lblCPF.caption = c.cnpj_Cli
        End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
        f = F_SelectForn(txtbxTerceiro.Text)
        If f.cod_for <> "" Then
            txtbxTerceiro.Text = f.nome_for
            txtbxTerceiro.tag = f.cod_for
            lblCPF.caption = f.cnpj_for
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.Text)
        If r.cod_Rep <> "" Then
            txtbxTerceiro.Text = r.nome_rep
            txtbxTerceiro.tag = r.cod_Rep
            lblCPF.caption = r.cpf_rep
        End If
    
End Sub


Private Sub cmndbtnTerceiro_Click()
    
End Sub

Private Sub cmndbtnTipoCob_Click()
    buscaTipoCob
End Sub
Sub buscaTipoCob()
    Dim t As tipocob
    t = F_BuscaTipoCob(txtbxTipoCob.Text)
    If t.cod_tip <> "" Then
        txtbxTipoCob.Text = t.nome_tip
        txtbxTipoCob.tag = t.cod_tip
    End If
End Sub


'Private Sub dbgrd_DblClick1()
'        Dim cJuros As Currency
'        With dtNot.Recordset
'        Dim sJur As String
'        If !reman_itd <= 0 Then
'            Exit Sub
'        End If
'        On Error Resume Next
'        cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
'        If cJuros < 0 Then cJuros = 0
'
'        If IsLoad("frmListaDoc") Then
'            If frmListaDoc.tag <> Me.Name And frmListaDoc.msflxGrd.rows > 0 Then
'                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
'                Unload frmListaDoc
'                Exit Sub
'
'            End If
'            If frmListaDoc.lblCliente.tag <> "" And _
'                frmListaDoc.lblCliente.tag <> !cod_ter Then
'                MsgBox "Você só pode selecionar parcelas do mesmo Terceiro!", vbExclamation
'                Exit Sub
'            End If
'            frmListaDoc.lblCliente.tag = !cod_ter
'            frmListaDoc.lblCliente.caption = !Nome
'
'            If F_InFlxgrd(frmListaDoc.msflxGrd, !seq_itd, 0) = -1 Then
'                frmListaDoc.msflxGrd.AddItem !seq_itd & Chr(9) & FormatVal(!reman_itd) & Chr(9) & FormatVal(cJuros) _
'                & Chr(9) & !num_doc & Chr(9) & !cod_pla & Chr(9) & FormatVal(!reman_itd) & Chr(9) & !seq_doc
'            End If
'            frmListaDoc.soma
'        Else
'
'            frmListaDoc.tag = Me.Name
'            frmListaDoc.lblCliente.tag = !cod_ter
'            frmListaDoc.lblCliente.caption = !Nome
'            frmListaDoc.caption = "Documentos"
'            frmListaDoc.msflxGrd.Cols = 7
'            frmListaDoc.msflxGrd.rows = 0
'            frmListaDoc.msflxGrd.AddItem !seq_itd & Chr(9) & FormatVal(!reman_itd) & Chr(9) & FormatVal(cJuros) _
'                & Chr(9) & !num_doc & Chr(9) & !cod_pla & Chr(9) & FormatVal(!reman_itd) & Chr(9) & !seq_doc
'            frmListaDoc.Show 0, mdiFrm
'            frmListaDoc.soma
'        End If
'
'    End With
'
'End Sub

Private Sub dbgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        mnSelec_Click
    End If
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dbgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnParcela, , x + dbgrd.Left, y + dbgrd.Top
    End If
End Sub

Private Sub dbgrdChe_GotFocus()
    dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChe_LostFocus()
    dbgrdChe.MarqueeStyle = 0
End Sub

Private Sub dtNot_Reposition()
    Dim strSQL As String
    On Error Resume Next
    If bActive = False Then Exit Sub
    With dtNot.Recordset
        
        If .RecordCount > 0 Then
            If !reman_itd = 0 Then
                cmndbtnEdit.Enabled = False
                cmndbtnBaixar.Enabled = False
            Else
                cmndbtnEdit.Enabled = True
                cmndbtnBaixar.Enabled = True
            End If
            
        End If
        If !tipo_doc = "R" Then
            cmndbtnCarne.Enabled = True
        Else
            cmndbtnCarne.Enabled = False
        End If
        lblTel.caption = F_BuscaTel(!cod_ter, !tipo_ter)
        buscaBaixas
        buscaOrg
        buscaVias
    End With
End Sub
Sub buscaVias()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error Resume Next
    lblVias.caption = F_BuscaUmCampo("count(*)", "tab_docimpress", "seq_doc='" & dtNot.Recordset!seq_doc & "'")
    
    
    
End Sub
Sub buscaOrg()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro:
    strSQL = orgao.F_SelectOrgaoCli(dtNot.Recordset!cod_ter)
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            lblNomeOrg.caption = !nome_org
            lblEnd.caption = !endereco_org & " " & !nome_set & " " & !nome_cid & "/" & !uf_cid & "(" & !fone_org & ")"
            
        Else
        
        End If
        .Close
    End With
    
    Exit Sub
trata_Erro:
End Sub
Sub buscaBaixas()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim r2 As Recordset
    On Error GoTo trata_Erro
    Dim sMsg As String
    
    With dtNot.Recordset
        sMsg = "Parcelas" & Chr(13)
        strSQL = "select * from tab_baixaitd2 where seq_itd='" & !seq_itd & "' "
'        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
'        With rPesq
'            Do While Not .EOF
'                sMsg = sMsg & "Data: " & !data_bai & "Cobrança: " & !cod_cob & " Vlr Baixado: " & FormatVal(!valor_bai) & " Usu: " & !cod_usu & Chr(13)
'               ' sMsg = sMsg & "Formas Pagamento" & Chr(13)
'                strSQL = "select * from tab_baixafin where seq_cab='" & !seq_cab & "' "
'                Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
'                With rPesq
'                    Do While Not .EOF
'                        sMsg = sMsg & "Espécie: " & !esp_bai & "Valor: " & FormatVal(!valor_bai)
'
'                        .MoveNext
'                    Loop
'                    .Close
'                End With
'                .MoveNext
'            Loop
'            .Close
'        End With
'        txtbxBaixas.Text = sMsg
    dtBaixas.RecordSource = strSQL
    dtBaixas.Refresh
    End With
    Exit Sub
trata_Erro:
End Sub
Private Sub Form_Activate()
    MudaAparencia Me
    If bActive = True Then
        P_Status
        Exit Sub
    End If
    
    cmndbtnBuscar_Click
    bActive = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Sub soma()
    Dim cTotal As Currency
    On Error GoTo trata_Erro
    bActive = False
    With dtNot.Recordset
        dbgrd.caption = .RecordCount & " Documentos Listados"
        
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !reman_itd
            .MoveNext
        Loop
        .MoveFirst
        
        
        
    End With
    lblTotal.caption = FormatVal(cTotal)
    bActive = True
    Exit Sub
trata_Erro:
    lblTotal.caption = "0,00"
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    On Error Resume Next
    bActive = False
    resizeForm Me
    txtbxDataIni.Text = Date
    
    txtbxdataFin.Text = Date
    dtNot.Connect = Banco.Connect
    cmbxTipo.ListIndex = 0
    cmbxOrder.ListIndex = 0
    dtNot.Connect = Banco.Connect
    dtBaixas.Connect = Banco.Connect
    cmbxBuscar.ListIndex = 1
    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    cmbxTipoTer.ListIndex = 0
    'cmndbtnBuscar_Click
    
End Sub

Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Width = Me.Width - dbgrd.Left - 300
    dbgrdChe.Top = pcTotal.Top - dbgrdChe.Height - 100
    dbgrd.Height = pcTotal.Top - dbgrd.Top - 100
    dbgrdChe.Width = dbgrd.Width
    
End Sub

Private Sub mnBaixar_Click()
    cmndbtnBaixar_Click
End Sub

Private Sub mnBaixas_Click()
    On Error Resume Next
    frmRelCarnesRecebidos.txtbxSeqItd.Text = dtNot.Recordset!seq_ite
    frmRelCarnesRecebidos.tag = "x"
    frmRelCarnesRecebidos.Show 0
    frmRelCarnesRecebidos.SetFocus

End Sub

Private Sub mnEdiPar_Click()
    cmndbtnEdit_Click
End Sub

Private Sub mnExclDoc_Click()
    cmndbtnExcl_Click
End Sub

Private Sub mnImpr_Click()
    cmndbtnImpr_Click
End Sub

Private Sub mnImprCarne_Click()
    cmndbtnCarne_Click
End Sub

Private Sub mnImprContr_Click()
    cmndbtnContrato_Click
End Sub

Private Sub mnNova_Click()
    mnNova_Click
End Sub

Private Sub mnSair_Click()
    Unload Me
End Sub

Private Sub mnSelec_Click()
    If Me.tag <> "" Then
        Me.tag = "X"
        Me.Hide
    End If
End Sub

Private Sub mnSelecPar_Click()
    dbgrd_DblClick
End Sub

Private Sub mnSPC_Click()
    On Error Resume Next
    P_ParcelaSpc dtNot.Recordset!cod_ter, dtNot.Recordset!seq_itd, dtNot.Recordset!reman_itd
End Sub

Private Sub mnVisAcertoVenda_Click()
    On Error GoTo trata_Erro:
    With dtNot.Recordset
        If !esp_Doc = "A" Then
            P_VisualizaAcerto !num_ven
        ElseIf !esp_Doc = "V" Then
            ImprVenda !num_ven
        End If
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub mnVisDoc_Click()
    On Error GoTo trata_Erro
    With dtNot.Recordset
        txtbxSeqDoc.Text = !seq_doc
        
        txtbxSeqItd.Text = ""
        'txtbx.Text = ""
        txtbxdataFin.Text = ""
        txtbxDataIni.Text = ""
        cmndbtnBuscar_Click
        txtbxSeqDoc.Text = ""
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub pcTop_GotFocus()
    SendKeys Chr(9)
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
End Sub

Private Sub txtbxdatafin_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxdatafin_LostFocus()
    P_AjustaData txtbxdataFin
    If txtbxdataFin.Text <> "" Then
        
    End If
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()

End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        cmndbtnTerc_Click
    End If
End Sub


Private Sub txtbxGrupo_Change()
    txtbxGrupo.tag = ""
End Sub

Private Sub txtbxGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnGru_Click
    End If
End Sub

Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub

Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    P_AjustaTexto txtbxHist
End Sub
Private Sub txtbxNumDoc_GotFocus()
    P_SelecTexto txtbxNumDoc
End Sub

Private Sub txtbxNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNumDoc_LostFocus()
    P_AjustaTexto txtbxNumDoc
End Sub

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    P_ValorLostFocus txtbxValor
End Sub
Private Sub txtbxRepr_Change()
    txtbxRepr.tag = ""
End Sub

Private Sub txtbxRepr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnRep_Click
    End If
End Sub

Private Sub txtbxSeqDoc_GotFocus()
    P_SelecTexto txtbxSeqDoc
End Sub

Private Sub txtbxSeqDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxSeqDoc_LostFocus()
    txtbxSeqDoc.Text = Format(txtbxSeqDoc.Text, "000000#")
End Sub

Private Sub txtbxSeqItd_GotFocus()
    P_SelecTexto txtbxSeqItd
End Sub

Private Sub txtbxSeqItd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSeqItd_LostFocus()
    txtbxSeqItd.Text = Format(txtbxSeqItd.Text, "000000#")
End Sub

Private Sub txtbxTerceiro_Change()
    txtbxTerceiro.tag = ""
    lblCPF.caption = ""
    lblCPF.tag = ""

End Sub

Private Sub txtbxTerceiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxTipoCob_Change()
    txtbxTipoCob.tag = ""
End Sub


