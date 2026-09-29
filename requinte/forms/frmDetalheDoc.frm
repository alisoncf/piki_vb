VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDetalheDoc 
   Caption         =   "Detalhe do Documento"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   12225
      TabIndex        =   11
      Top             =   5175
      Width           =   12285
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   11040
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox pcLocal 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   240
      Width           =   12015
      Begin VB.TextBox txtEmpresa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtParcelas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Seq Doc:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Parcelas:"
         Height          =   195
         Left            =   3720
         TabIndex        =   3
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   5760
         TabIndex        =   2
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   2160
         TabIndex        =   1
         Top             =   0
         Width           =   870
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1720
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmDetalheDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public json As String

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    monta_grade
End Sub
Sub monta_grade()
    Dim id_doc As String
    Dim c As New Collection
    Dim parcelas As String
    Dim i As Integer
    Dim Linha As String
    If json = "" Then Exit Sub
    txtNumDoc.Text = GetValorJson("num_doc", json)
    txtSeqDoc.Text = GetValorJson("seq_doc", json)
    txtCliente.Text = GetValorJson("cpf_cli", json) & " - " & GetValorJson("nome_cli", json)
    txtParcelas.Text = GetValorJson("parcelas_doc", json)
    txtEmpresa.Text = GetValorJson("cod_emp", json)
    txtData.Text = Format(GetValorJson("data_doc", json), "dd/mm/yyyy")
    id_doc = GetValorJson("id", json)
    ajustaGridTitulos grade, 4, "id,vencimento,valor,remanescente"
    
    parcelas = APIBuscarItDoc(id_doc)
    
    
    
    Set c = JSONUtil.ParseJsonIntoList(parcelas)
    
    For i = 1 To c.Count
        Linha = GetValorJson("id", c(i)) & Chr(9)
        Linha = Linha & GetValorJson("venc_itd", c(i)) & Chr(9)
        Linha = Linha & GetValorJson("valor", c(i)) & Chr(9)
        Linha = Linha & GetValorJson("reman_itd", c(i))
        grade.AddItem Linha
    Next
    
End Sub

Private Sub Form_Resize()
    PosicionaGradeWidth grade, Me
    PosicionaGradeHeigth grade, pcBot
End Sub
