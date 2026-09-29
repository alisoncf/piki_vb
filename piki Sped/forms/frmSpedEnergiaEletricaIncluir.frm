VERSION 5.00
Begin VB.Form frmSpedEnergiaEletricaIncluir 
   Caption         =   "Lançamento de NF de Energia Elétrica/ Água /Gás"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   10950
      TabIndex        =   52
      Top             =   5310
      Width           =   10950
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   7920
         Picture         =   "frmSpedEnergiaEletricaIncluir.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   9360
         Picture         =   "frmSpedEnergiaEletricaIncluir.frx":0487
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   6720
      ScaleHeight     =   4995
      ScaleWidth      =   4035
      TabIndex        =   35
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtVl_Cofins 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   51
         Text            =   "0,00"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtVl_Pis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   49
         Text            =   "0,00"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtVl_ICms_St 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   47
         Text            =   "0,00"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtVl_Bc_ICMS_St 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   45
         Text            =   "0,00"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtVl_Icms 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   43
         Text            =   "0,00"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtVlBc_Icms 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   41
         Text            =   "0,00"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtVl_Da 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   39
         Text            =   "0,00"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtVl_Serv_Nt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   37
         Text            =   "0,00"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Valor COFINS:"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   2760
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Valor PIS:"
         Height          =   195
         Left            =   240
         TabIndex        =   48
         Top             =   2400
         Width           =   705
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Valor do ICMS ST:"
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Valor Acumulado da BC ICMS ST:"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Valor do ICMS"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Valor Acumulado da BC ICMS"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   960
         Width           =   2115
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Valor das Despesas Acessórias:"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Valor Serv. Não Trib. ICSM:"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1965
      End
   End
   Begin VB.TextBox txtVl_Forn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      MaxLength       =   12
      TabIndex        =   33
      Text            =   "0,00"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtVl_desc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   31
      Text            =   "0,00"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtVl_Doc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   12
      TabIndex        =   29
      Text            =   "0,00"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipo:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton optGas 
         Caption         =   "Gás"
         Height          =   195
         Left            =   1920
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optAgua 
         Caption         =   "Água"
         Height          =   195
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optEnergia 
         Caption         =   "Energia"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbxCodGruTensao 
      Height          =   315
      ItemData        =   "frmSpedEnergiaEletricaIncluir.frx":0A15
      Left            =   360
      List            =   "frmSpedEnergiaEletricaIncluir.frx":0A46
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4440
      Width           =   5895
   End
   Begin VB.ComboBox cmbxTIpoLigacao 
      Height          =   315
      ItemData        =   "frmSpedEnergiaEletricaIncluir.frx":0C53
      Left            =   2880
      List            =   "frmSpedEnergiaEletricaIncluir.frx":0C63
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox txtSubser 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   19
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   14
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtSerie 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtDT_E_S 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtDT_Doc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox cmbxCodCons 
      Height          =   315
      ItemData        =   "frmSpedEnergiaEletricaIncluir.frx":0C96
      Left            =   360
      List            =   "frmSpedEnergiaEletricaIncluir.frx":0C98
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3720
      Width           =   2415
   End
   Begin VB.ComboBox cmbxSituacao 
      Height          =   315
      ItemData        =   "frmSpedEnergiaEletricaIncluir.frx":0C9A
      Left            =   2160
      List            =   "frmSpedEnergiaEletricaIncluir.frx":0C9C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox txtCodMod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtForn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1680
      Width           =   5655
   End
   Begin VB.ComboBox cmbxInd_Oper 
      Height          =   315
      ItemData        =   "frmSpedEnergiaEletricaIncluir.frx":0C9E
      Left            =   360
      List            =   "frmSpedEnergiaEletricaIncluir.frx":0CA8
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnLocaliza 
      Height          =   285
      Left            =   6000
      Picture         =   "frmSpedEnergiaEletricaIncluir.frx":0CC4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr Fornecido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2760
      TabIndex        =   34
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Desc"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   32
      Top             =   2760
      Width           =   780
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Doc"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   30
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo Tensão:"
      Height          =   195
      Left            =   360
      TabIndex        =   24
      Top             =   4200
      Width           =   1065
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Ligação:"
      Height          =   195
      Left            =   2880
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subsérie:"
      Height          =   195
      Left            =   3480
      TabIndex        =   20
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Número:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1320
      TabIndex        =   18
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Série:"
      Height          =   195
      Left            =   2400
      TabIndex        =   17
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "E&ntrada:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5280
      TabIndex        =   16
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "&Emissão: "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4200
      TabIndex        =   15
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Consumo:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situação NF:"
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo NF:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Participante(Fornecedor):"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Indicador de Operação:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1680
   End
End
Attribute VB_Name = "frmSpedEnergiaEletricaIncluir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxCodCons_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub cmbxCodGruTensao_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub cmbxTIpoLigacao_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub cmdNovo_Click()
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    If GravaSpedNfEnergia("0", "1", txtCodMod.Text, txtForn.tag, Left(cmbxSituacao.Text, 2), txtSerie.Text, _
        txtSubser.Text, Left(cmbxCodCons.Text, 2), txtNumero.Text, txtDT_Doc.Text, txtDT_E_S.Text, _
        txtVl_Doc, txtVl_desc.Text, txtVl_Forn.Text, txtVl_Serv_Nt.Text, "0,00", txtVl_Da.Text, _
        txtVlBc_Icms.Text, txtVl_Icms.Text, txtVl_ICms_St.Text, "", txtVl_Pis.Text, txtVl_Cofins.Text, _
        Left(cmbxTIpoLigacao.Text, 1), Left(cmbxTIpoLigacao.Text, 2)) Then
        
        MsgBox "Gravado com sucesso!", vbExclamation
        txtNumero.Text = ""
        cmbxInd_Oper.SetFocus
        
    End If
        
End Sub
Sub altera()
    If AlteraSpedNfEnergia(Me.tag, "0", "1", txtForn.tag, txtCodMod.Text, Left(cmbxSituacao.Text, 2), txtSerie.Text, _
        txtSubser.Text, Left(cmbxCodCons.Text, 2), txtNumero.Text, txtDT_Doc.Text, txtDT_E_S.Text, _
        txtVl_Doc, txtVl_desc.Text, txtVl_Forn.Text, txtVl_Serv_Nt.Text, "0,00", txtVl_Da.Text, _
        txtVlBc_Icms.Text, txtVl_Icms.Text, txtVl_ICms_St.Text, "", txtVl_Pis.Text, txtVl_Cofins.Text, _
        Left(cmbxTIpoLigacao.Text, 1), Left(cmbxTIpoLigacao.Text, 2)) Then
        
        MsgBox "Gravado com sucesso!", vbExclamation
        txtNumero.Text = ""
        cmbxInd_Oper.SetFocus
        
    End If
End Sub
Private Sub cmndbtnLocaliza_Click()
      Dim f As participante
    f = SpedParticipante.F_BuscaParticipante(txtForn.Text)
    If f.codigo <> "" Then
            txtForn.Text = f.nome
            
            txtForn.tag = f.codigo
            
            txtbxNum_Not.SetFocus
    End If
End Sub

Private Sub Form_Load()
    P_PreencheTipoCOnsumoEnergia cmbxCodCons
    P_PreencheSituacaoDoc cmbxSituacao
End Sub

Private Sub optAgua_Click()
    P_PreencheTipoCOnsumoAgua cmbxCodCons
End Sub

Private Sub optEnergia_Click()
    P_PreencheTipoCOnsumoEnergia cmbxCodCons
End Sub

Private Sub optGas_Click()
    P_PreencheTipoCOnsumoEnergia cmbxCodCons
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtDT_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtDT_E_S_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtSubser_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_Bc_ICMS_St_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_Cofins_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_Da_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_Forn_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxEmissao_Not_Change()

End Sub

Private Sub txtbxEmissao_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxEntrada_Not_Change()

End Sub

Private Sub txtbxEntrada_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxNum_Not_Change()

End Sub

Private Sub txtbxNum_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxSerie_Not_Change()

End Sub

Private Sub txtbxSerie_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxSerie_Not_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtCodMod_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtForn_GotFocus()
    GanhaFoco txtForn
End Sub

Private Sub txtForn_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
    
End Sub

Private Sub txtVl_desc_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_Icms_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_ICms_St_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_Pis_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVl_Serv_Nt_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtVlBc_Icms_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub
