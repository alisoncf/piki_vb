VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form frmMigraDados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerar Dados Para Migração"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmd 
      Left            =   960
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstb 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Gerar Dados"
      TabPicture(0)   =   "frmMigraDados.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdGerar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "optnProduto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optnCliente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtTabelas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Captar Dados"
      TabPicture(1)   =   "frmMigraDados.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdExecutar"
      Tab(1).Control(1)=   "txtData"
      Tab(1).Control(2)=   "txtTab"
      Tab(1).Control(3)=   "txtEmp"
      Tab(1).Control(4)=   "txtArq"
      Tab(1).Control(5)=   "cmdCapProd"
      Tab(1).Control(6)=   "cmdArq"
      Tab(1).Control(7)=   "Label4"
      Tab(1).Control(8)=   "Label3"
      Tab(1).Control(9)=   "Label2"
      Tab(1).ControlCount=   10
      Begin VB.CommandButton cmdExecutar 
         Caption         =   "Executar"
         Height          =   495
         Left            =   -74880
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtData 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74160
         TabIndex        =   15
         Top             =   1920
         Width           =   5415
      End
      Begin VB.TextBox txtTab 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74160
         TabIndex        =   14
         Top             =   1560
         Width           =   5415
      End
      Begin VB.TextBox txtEmp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74160
         TabIndex        =   13
         Top             =   1200
         Width           =   5415
      End
      Begin VB.TextBox txtArq 
         Height          =   285
         Left            =   -72840
         TabIndex        =   8
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtTabelas 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton cmdCapProd 
         Caption         =   "Carregar"
         Height          =   495
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optnCliente 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optnProduto 
         Caption         =   "Produto"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "Gerar"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdArq 
         Caption         =   "..."
         Height          =   255
         Left            =   -68880
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   195
         Left            =   -74880
         TabIndex        =   12
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tabelas:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   10
         Top             =   1200
         Width           =   660
      End
   End
   Begin VB.TextBox txtbxCaminho 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "frmMigraDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdArq_Click()
    Dim f As String
    cmd.ShowOpen
    txtArq.Text = cmd.FileName
        
    
End Sub

Private Sub cmdCapProd_Click()
    Dim Linha As String
    Dim iLinha As Integer
    Dim Cont As Integer
    Dim marcador1 As Integer
    Dim marcador2 As Integer
    Dim Tabelas As String
    On Error GoTo trata_Erro
    Open txtArq.Text For Input As #1
    
    Do While Not EOF(1)
        Cont = Cont + 1
        Line Input #1, Linha
        If InStr(1, Linha, "<tabelas>") > 0 Then
            marcador1 = InStr(1, Linha, "<tabelas>")
            marcador2 = InStr(1, Linha, "</tabelas>")
            Tabelas = Mid(Linha, marcador1 + 9, marcador2)
            txtTab.Text = Tabelas
            
            
        End If
        If InStr(1, Linha, "<empresa>") > 0 Then
            marcador1 = InStr(1, Linha, "<empresa>")
            marcador2 = InStr(1, Linha, "</empresa>")
            Tabelas = Mid(Linha, marcador1 + 9, marcador2)
            txtEmp.Text = Tabelas
            
            
        End If
        If InStr(1, Linha, "<data>") > 0 Then
            marcador1 = InStr(1, Linha, "<data>")
            marcador2 = InStr(1, Linha, "</data>")
            Tabelas = Mid(Linha, marcador1 + 6, marcador2 - 7)
            txtData.Text = Tabelas
            
            
        End If
    
    Loop
    
    cmdExecutar.caption = "<" & Cont & "> Executar"
    cmdExecutar.BackColor = vbGreen
    Close #1
    Exit Sub
trata_Erro:
    
    P_MostraErro
    On Error Resume Next
      Close #1
End Sub

Private Sub cmdExecutar_Click()
    If MsgBox("Confirma execução da ação?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    migra.AtualizaDados txtArq.Text
    
End Sub

Private Sub cmdGerar_Click()
    migra.MigraTabela txtTabelas.Text, txtbxCaminho.Text
    shell "notepad " & txtbxCaminho.Text
End Sub

Private Sub Form_Load()
    txtbxCaminho.Text = App.path & "\produtos.xml"
End Sub

Private Sub optnProduto_Click()
    txtbxCaminho.Text = App.path & "\produtos.xml"
    
    txtTabelas.Text = "tab_produtos|tab_fornecedores|tab_secoes|tab_subsec|tab_codbarras"
End Sub
