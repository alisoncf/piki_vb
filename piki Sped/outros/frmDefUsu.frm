VERSION 5.00
Begin VB.Form frmDefUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definições de Usuário"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frmDefUsu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbxExclVenda 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":08CA
      Left            =   2040
      List            =   "frmDefUsu.frx":08D4
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   4800
      Width           =   855
   End
   Begin VB.ComboBox cmbxDesconto 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":08E2
      Left            =   5640
      List            =   "frmDefUsu.frx":08F5
      Sorted          =   -1  'True
      TabIndex        =   53
      Text            =   "cmbxDesconto"
      Top             =   4800
      Width           =   855
   End
   Begin VB.ComboBox cmbxContrPre 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":090D
      Left            =   5640
      List            =   "frmDefUsu.frx":0917
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   3600
      Picture         =   "frmDefUsu.frx":0925
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   240
      Picture         =   "frmDefUsu.frx":0CAF
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5400
      Width           =   3015
   End
   Begin VB.ComboBox cmbxCadFormasPag 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1239
      Left            =   5640
      List            =   "frmDefUsu.frx":1243
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox cmbxAltPro 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1251
      Left            =   5640
      List            =   "frmDefUsu.frx":125B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cmbxFechaMes 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1269
      Left            =   5640
      List            =   "frmDefUsu.frx":1273
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   4080
      Width           =   855
   End
   Begin VB.ComboBox cmbxFechaCai 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1281
      Left            =   5640
      List            =   "frmDefUsu.frx":128B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   3720
      Width           =   855
   End
   Begin VB.ComboBox cmbxVisFin 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1299
      Left            =   5640
      List            =   "frmDefUsu.frx":12A3
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox cmbxVisRel 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":12B1
      Left            =   5640
      List            =   "frmDefUsu.frx":12BB
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   3000
      Width           =   855
   End
   Begin VB.ComboBox cmbxVisCusto 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":12C9
      Left            =   5640
      List            =   "frmDefUsu.frx":12D3
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   2640
      Width           =   855
   End
   Begin VB.ComboBox cmbxAltFormasPag 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":12E1
      Left            =   5640
      List            =   "frmDefUsu.frx":12EB
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox cmbxCadCCO 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":12F9
      Left            =   5640
      List            =   "frmDefUsu.frx":1303
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox cmbxCadCli 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1311
      Left            =   5640
      List            =   "frmDefUsu.frx":131B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox cmbxCadUsu 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1329
      Left            =   5640
      List            =   "frmDefUsu.frx":1333
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cmbxAltPrecos 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1341
      Left            =   5640
      List            =   "frmDefUsu.frx":134B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox cmbxCadPro 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1359
      Left            =   2040
      List            =   "frmDefUsu.frx":1363
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   4440
      Width           =   855
   End
   Begin VB.ComboBox cmbxExclDoc 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1371
      Left            =   2040
      List            =   "frmDefUsu.frx":137B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4080
      Width           =   855
   End
   Begin VB.ComboBox cmbxExclCai 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1389
      Left            =   2040
      List            =   "frmDefUsu.frx":1393
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3720
      Width           =   855
   End
   Begin VB.ComboBox cmbxEdiCai 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":13A1
      Left            =   2040
      List            =   "frmDefUsu.frx":13AB
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox cmbxBaixarChe 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":13B9
      Left            =   2040
      List            =   "frmDefUsu.frx":13C3
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3000
      Width           =   855
   End
   Begin VB.ComboBox cmbxBaixarDoc 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":13D1
      Left            =   2040
      List            =   "frmDefUsu.frx":13DB
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.ComboBox cmbxLanChe 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":13E9
      Left            =   2040
      List            =   "frmDefUsu.frx":13F3
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox cmbxLanDoc 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1401
      Left            =   2040
      List            =   "frmDefUsu.frx":140B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox cmbxLanCai 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1419
      Left            =   2040
      List            =   "frmDefUsu.frx":1423
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox cmbxAbreSaida 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1431
      Left            =   2040
      List            =   "frmDefUsu.frx":143B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox cmbxEfetNota 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1449
      Left            =   2040
      List            =   "frmDefUsu.frx":1453
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cmbxAbreNota 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1461
      Left            =   2040
      List            =   "frmDefUsu.frx":146B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox cmbxAbreVenda 
      Height          =   315
      ItemData        =   "frmDefUsu.frx":1479
      Left            =   2040
      List            =   "frmDefUsu.frx":1483
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&8 - Excluir Venda"
      Height          =   195
      Left            =   240
      TabIndex        =   57
      Top             =   4860
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&7 - Desconto:(%)"
      Height          =   195
      Left            =   3840
      TabIndex        =   52
      Top             =   4860
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&6 - Controlar Preços"
      Height          =   195
      Left            =   3840
      TabIndex        =   50
      Top             =   4500
      Width           =   1485
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&9 - Cadastrar Frms Pgto"
      Height          =   195
      Left            =   3840
      TabIndex        =   36
      Top             =   1980
      Width           =   1740
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&4 - Alterar Produto"
      Height          =   195
      Left            =   3840
      TabIndex        =   26
      Top             =   180
      Width           =   1365
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&5 - Fechar Mês"
      Height          =   195
      Left            =   3840
      TabIndex        =   48
      Top             =   4140
      Width           =   1155
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&4 - Fechar Caixa"
      Height          =   195
      Left            =   3840
      TabIndex        =   46
      Top             =   3780
      Width           =   1245
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&3 - Vis. Relat. Financ."
      Height          =   195
      Left            =   3840
      TabIndex        =   44
      Top             =   3420
      Width           =   1605
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&2 - Vis. Relat. de Est."
      Height          =   195
      Left            =   3840
      TabIndex        =   42
      Top             =   3060
      Width           =   1575
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&1 - Visualizar Custos"
      Height          =   195
      Left            =   3840
      TabIndex        =   40
      Top             =   2700
      Width           =   1500
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2&0 - Alterar. Formas Pgto"
      Height          =   195
      Left            =   3840
      TabIndex        =   38
      Top             =   2340
      Width           =   1740
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&8 - Cadastrar C. Corr."
      Height          =   195
      Left            =   3840
      TabIndex        =   34
      Top             =   1620
      Width           =   1560
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&7 - Cadastrar Clientes"
      Height          =   195
      Left            =   3840
      TabIndex        =   32
      Top             =   1260
      Width           =   1590
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&6 - Cadastrar Usuários"
      Height          =   195
      Left            =   3840
      TabIndex        =   30
      Top             =   900
      Width           =   1650
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&5 - Alterar Preços"
      Height          =   195
      Left            =   3840
      TabIndex        =   28
      Top             =   540
      Width           =   1305
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&3 - Cadastrar Produto"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   4500
      Width           =   1590
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&2-  Excluir Doc."
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   4140
      Width           =   1170
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&1 - Excluir Caixa"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   3780
      Width           =   1215
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1&0 - Editar Caixa"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   3420
      Width           =   1155
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&9 - Baixar Cheque"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   3060
      Width           =   1260
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&8 - Baixar Doc."
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2700
      Width           =   1050
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&7 - Fazer Lanç. Cheque"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2340
      Width           =   1665
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&6 - Fazer Lanç. Doc"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1980
      Width           =   1410
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&5 - Fazer Lanç. Caixa"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1620
      Width           =   1500
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&4 - Abrir Cupom Fiscal"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&3 - Efetivar Nota Fiscal"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   900
      Width           =   1605
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&2 - Abrir Nota Fiscal"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   540
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&1 - Abrir Venda"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1050
   End
End
Attribute VB_Name = "frmDefUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxEscVend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxDesconto_LostFocus()
    P_ValorLostFocus cmbxDesconto
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If MsgBox("Confirma alteração das definições de usuário?", vbYesNo + vbQuestion) = vbYes Then
        grava
    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_defusu where cod_usu = '" & Me.Tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            Exit Sub
        End If
        P_LocalizaOperacao !abrevenda_usu, cmbxAbreVenda
        'P_LocalizaOperacao !efetvenda_usu, cmbxEfetvenda
        'P_LocalizaOperacao !baixavenda_usu, cmbxBaixarVenda
        'P_LocalizaOperacao !efetped_usu, cmbxEfetPed
        'P_LocalizaOperacao !abreped_usu, cmbxAbrPed
        P_LocalizaOperacao !abrenota_usu, cmbxAbreNota
        P_LocalizaOperacao !efetnota_usu, cmbxEfetNota
        P_LocalizaOperacao !abresaida_usu, cmbxAbreSaida
        P_LocalizaOperacao !lancaixa_usu, cmbxLanCai
        P_LocalizaOperacao !landoc_usu, cmbxLanDoc
        P_LocalizaOperacao !lanche_usu, cmbxLanChe
        P_LocalizaOperacao !baixadoc_usu, cmbxBaixarDoc
        P_LocalizaOperacao !baixache_usu, cmbxBaixarChe
        P_LocalizaOperacao !edicai_usu, cmbxEdiCai
        P_LocalizaOperacao !exclcai_usu, cmbxExclCai
        P_LocalizaOperacao !excldoc_usu, cmbxExclDoc
        P_LocalizaOperacao !cadprod_usu, cmbxCadPro
        P_LocalizaOperacao !altprod_usu, cmbxAltPro
        P_LocalizaOperacao !altpreco_usu, cmbxAltPrecos
        P_LocalizaOperacao !cadusu_usu, cmbxCadUsu
        P_LocalizaOperacao !cadcli_usu, cmbxCadCli
        P_LocalizaOperacao !cadcco_usu, cmbxCadCCO
        P_LocalizaOperacao !cadformaspag_usu, cmbxCadFormasPag
        P_LocalizaOperacao !exclvenda_usu, cmbxExclVenda
        P_LocalizaOperacao !altformaspag_usu, cmbxAltFormasPag
        
        P_LocalizaOperacao !viscusto_usu, cmbxVisCusto
        P_LocalizaOperacao !visrelest_usu, cmbxVisRel
        P_LocalizaOperacao !visrelfin_usu, cmbxVisFin
        P_LocalizaOperacao !fechacai_usu, cmbxFechaCai
        P_LocalizaOperacao !fechames_usu, cmbxFechaMes
        P_LocalizaOperacao !contrpre_usu, cmbxContrPre
        cmbxDesconto.Text = !desconto_usu
        .Close
    End With
End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "delete from tab_defusu where cod_usu='" & Me.Tag & "'"
    Banco.Execute strSQL
    strSQL = "insert into tab_defusu(abrevenda_usu,"
    
    
    
    
    strSQL = strSQL & "abrenota_usu,"
    strSQL = strSQL & "efetnota_usu,"
    strSQL = strSQL & "abresaida_usu,"
    strSQL = strSQL & "lancaixa_usu,"
    strSQL = strSQL & "landoc_usu,"
    strSQL = strSQL & "lanche_usu,"
    strSQL = strSQL & "baixadoc_usu,"
    strSQL = strSQL & "baixache_usu,"
    strSQL = strSQL & "edicai_usu,"
    strSQL = strSQL & "exclcai_usu,"
    strSQL = strSQL & "excldoc_usu,"
    strSQL = strSQL & "cadprod_usu,"
    strSQL = strSQL & "altprod_usu,"
    strSQL = strSQL & "altpreco_usu,"
    strSQL = strSQL & "cadusu_usu,"
    strSQL = strSQL & "cadcli_usu,"
    strSQL = strSQL & "cadcco_usu,"
    strSQL = strSQL & "cadformaspag_usu,"
    strSQL = strSQL & "altformaspag_usu,"
    strSQL = strSQL & "exclvenda_usu,"
    strSQL = strSQL & "viscusto_usu,"
    strSQL = strSQL & "visrelest_usu,"
    strSQL = strSQL & "visrelfin_usu,"
    strSQL = strSQL & "fechacai_usu,"
    strSQL = strSQL & "fechames_usu,"
    strSQL = strSQL & "contrpre_usu,"
    strSQL = strSQL & "desconto_usu,"
    strSQL = strSQL & "cod_usu)values("
    strSQL = strSQL & "'" & Left(cmbxAbreVenda.Text, 1) & "', "
    



    strSQL = strSQL & "'" & Left(cmbxAbreNota.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxEfetNota.Text, 1) & "', "
    strSQL = strSQL & "'" & Left(cmbxAbreSaida.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanCai.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanDoc.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanChe.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixarDoc.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixarChe.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxEdiCai.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclCai.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclDoc.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadPro.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltPro.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltPrecos.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadUsu.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadCli.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadCCO.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadFormasPag.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltFormasPag.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclVenda.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisCusto.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisRel.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisFin.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxFechaCai.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxFechaMes.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxContrPre.Text, 1) & "',"
    strSQL = strSQL & F_GravaDec(cmbxDesconto.Text) & ","
    strSQL = strSQL & "'" & Me.Tag & "') "
    Banco.Execute strSQL
    MsgBox "Definições de Usuário Gravadas!", vbInformation
    Unload Me
End Sub
Sub altera()
    
End Sub
