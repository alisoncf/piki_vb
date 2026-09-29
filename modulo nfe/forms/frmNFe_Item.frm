VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNFe_Item 
   Caption         =   "Detalhe do Produto"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13440
   Icon            =   "frmNFe_Item.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstb 
      Height          =   3375
      Left            =   120
      TabIndex        =   47
      Top             =   2760
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "ICMS"
      TabPicture(0)   =   "frmNFe_Item.frx":2CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblT"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label24"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTipoIcms"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbxTipoICMS"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbOrigem"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "gridValores"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExcluir"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAdd"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAlterar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtGrid"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "PIS"
      TabPicture(1)   =   "frmNFe_Item.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "COFINS"
      TabPicture(2)   =   "frmNFe_Item.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "IPI"
      TabPicture(3)   =   "frmNFe_Item.frx":2D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.TextBox txtGrid 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   9840
         TabIndex        =   58
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   8880
         TabIndex        =   57
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   10680
         TabIndex        =   56
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   11880
         TabIndex        =   55
         Top             =   2880
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid gridValores 
         Height          =   2175
         Left            =   8880
         TabIndex        =   52
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3836
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.ComboBox cmbOrigem 
         Height          =   315
         ItemData        =   "frmNFe_Item.frx":2D6A
         Left            =   120
         List            =   "frmNFe_Item.frx":2D89
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   600
         Width           =   8535
      End
      Begin VB.ComboBox cmbxTipoICMS 
         Height          =   315
         ItemData        =   "frmNFe_Item.frx":30C5
         Left            =   120
         List            =   "frmNFe_Item.frx":30C7
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1200
         Width           =   8535
      End
      Begin VB.Label lblTipoIcms 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo_Icms"
         Height          =   195
         Left            =   7920
         TabIndex        =   54
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Valores:"
         Height          =   195
         Left            =   8880
         TabIndex        =   53
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblT 
         Caption         =   "Tipo de ICMS"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   13440
      TabIndex        =   41
      Top             =   6360
      Width           =   13440
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   11640
         TabIndex        =   42
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox pcProd 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2625
      ScaleWidth      =   13065
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.TextBox txtInfAdProd 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   120
         TabIndex        =   61
         Top             =   2160
         Width           =   11175
      End
      Begin VB.CommandButton cmndAlterar 
         Caption         =   "Alterar"
         Height          =   495
         Left            =   11400
         TabIndex        =   46
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtCEST 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   44
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdPesqProd 
         Height          =   255
         Left            =   7440
         Picture         =   "frmNFe_Item.frx":30C9
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         Width           =   255
      End
      Begin VB.ComboBox cmbxIndTot 
         Height          =   315
         ItemData        =   "frmNFe_Item.frx":343B
         Left            =   1200
         List            =   "frmNFe_Item.frx":3445
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtVOutro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Text            =   "0,00"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtVDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9120
         TabIndex        =   28
         Text            =   "0,00"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtVSeg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11280
         TabIndex        =   32
         Text            =   "0,00"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtVFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10200
         TabIndex        =   30
         Text            =   "0,00"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtVunTrib 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   24
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtQtrib 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   20
         Text            =   "0,000"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtQCom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   18
         Text            =   "0,000"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtUTrib 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtCEanTrib 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9480
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtVProd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtVUnCom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtUCom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtCFOP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtEXTIPI 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   40
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtNVE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   38
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtNCM 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11040
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCEAN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtxProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informações adicionais:"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label lblCST_Cofins 
         AutoSize        =   -1  'True
         Caption         =   "COFINS"
         Height          =   195
         Left            =   9720
         TabIndex        =   60
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblCST_Pis 
         AutoSize        =   -1  'True
         Caption         =   "PIS"
         Height          =   195
         Left            =   9240
         TabIndex        =   59
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEST"
         Height          =   195
         Left            =   7920
         TabIndex        =   45
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indicador de Total"
         Height          =   195
         Left            =   1200
         TabIndex        =   35
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. Outro"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. Desc."
         Height          =   195
         Left            =   9120
         TabIndex        =   27
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. Seg."
         Height          =   195
         Left            =   11280
         TabIndex        =   31
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. Frete."
         Height          =   195
         Left            =   10200
         TabIndex        =   29
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. Un. Trib."
         Height          =   195
         Left            =   6720
         TabIndex        =   23
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. Trib."
         Height          =   195
         Left            =   4320
         TabIndex        =   19
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. Com."
         Height          =   195
         Left            =   3120
         TabIndex        =   17
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un. Trib"
         Height          =   195
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN Trib."
         Height          =   195
         Left            =   9480
         TabIndex        =   7
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. Prod."
         Height          =   195
         Left            =   7920
         TabIndex        =   25
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. Un. Com."
         Height          =   195
         Left            =   5520
         TabIndex        =   21
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un. Com."
         Height          =   195
         Left            =   1200
         TabIndex        =   13
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXT IPI:"
         Height          =   195
         Left            =   6720
         TabIndex        =   39
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NVE:"
         Height          =   195
         Left            =   5520
         TabIndex        =   37
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM"
         Height          =   195
         Left            =   11040
         TabIndex        =   9
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN:"
         Height          =   195
         Left            =   7920
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmNFe_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id_nfe As String
Dim row As Integer
Dim bActive As Boolean

Private Sub cmbxIndTot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipoICMS_Click()
    buscaICMS
End Sub

Private Sub cmdAdd_Click()
    Dim strSQL As String
    Dim tabela As String
    Dim tabelAntiga As String
    Dim iCount As Integer
    Dim values As String
    Dim tipo As String
    Dim TipoICMS As String
    On Error GoTo trata_Erro
    If buscaStatusNFe(id_nfe) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    tabelaantiga = F_BuscaUmCampo("tipo_icms", "nfe_item_imposto", "id=" & Me.tag & "")
    
    If tabelaantiga = "20" Or tabelaantiga = "00" Or tabelaantiga = "60" Then
        tabelaantiga = "ICMS" & tabelaantiga
    End If
    If tabelaantiga <> "" Then
        strSQL = "delete from nfe_item_imposto_" & tabelaantiga
        strSQL = strSQL & " where id=" & Me.tag
        Conn.Execute strSQL
        
    End If
    
    
    
    
    tabela = BuscaTabelaICMS(Left(cmbxTipoICMS.text, 3))
    strSQL = "insert into " & tabela & "(id,id_nfe,"
    values = "(" & Me.tag & "," & id_nfe & ","
    

    
    
    
    
    
    For iCount = 1 To gridValores.rows - 1
        gridValores.row = iCount
        gridValores.col = 2
        tipo = gridValores.text
        
        If F_SoNumero(tipo) = tipo Then 'hex
            tipo = hex2ascii(tipo)
        End If
        gridValores.col = 0
        strSQL = strSQL & gridValores.text
        
        
        
        gridValores.col = 1
        If Left(tipo, 6) = "double" Then
            values = values & F_GravaDec(gridValores.text)
        ElseIf Left(tipo, 7) = "decimal" Then
            values = values & F_GravaDec(gridValores.text)
        ElseIf Left(tipo, 4) = "char" Then
            values = values & "'" & gridValores.text & "'"
        ElseIf Left(tipo, 4) = "text" Then
            values = values & "'" & gridValores.text & "'"
        ElseIf LCase(tipo) = "string" Then
            values = values & "'" & gridValores.text & "'"
        ElseIf LCase(tipo) = "double" Then
            values = values & F_GravaDec(gridValores.text)
        Else
            values = values & gridValores.text
        End If
        If iCount <> gridValores.rows - 1 Then
            strSQL = strSQL & ","
            values = values & ","
        End If
        
    Next
    strSQL = strSQL & ")" & "values" & values & ")"
    Conn.Execute strSQL
    TipoICMS = Left(cmbxTipoICMS.text, 3)
    If Len(Trim(TipoICMS)) = 2 Then
    
        strSQL = "update nfe_item_imposto set tipo_icms='ICMS" & Trim(Left(cmbxTipoICMS.text, 3)) & "' "
        strSQL = strSQL & " where id=" & Me.tag
        Conn.Execute strSQL
    Else
        strSQL = "update nfe_item_imposto set tipo_icms='ICMSSN" & Trim(Left(cmbxTipoICMS.text, 3)) & "' "
        strSQL = strSQL & " where id=" & Me.tag
        Conn.Execute strSQL
    End If
    MsgBox "Dados gravados", vbInformation
    
    Exit Sub
    
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdAlterar_Click()
    Dim strSQL As String
    If buscaStatusNFe(id_nfe) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    
    

End Sub

Private Sub cmdExcluir_Click()
    Dim strSQL As String
    Dim tabela As String
    If buscaStatusNFe(id_nfe) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    tabela = F_BuscaUmCampo("tabela", "nfe_tipo_icms", "")
    
    If MsgBox("Confirma exclusão do lançamento do ICMS para o produto?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from nfe_item_imposto_" & lblTipoIcms.caption & " where id=" & Me.tag
        Conn.Execute strSQL
        carrega
        
    End If
    
End Sub

Private Sub cmdPesqProd_Click()
    buscaProduto
End Sub
Sub buscaProduto()
    Dim p As produto
    p = F_BuscaProdSaida(txtxProd)
    If p.cod_pro <> "" Then
        txtCProd.text = p.cod_pro
        txtCean.text = p.ean_pro
        txtNCM.text = p.ncm_pro
        
    End If
End Sub
Private Sub cmdSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndAlterar_Click()
    Dim tipo_pis As String
    Dim tipo_icms As String
    Dim tipo_cofins As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
    
    If buscaStatusNFe(id_nfe) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    If Me.tag <> "" Then
        If AlteraNFe_Item(Me.tag, txtCProd.text, txtCean.text, txtxProd.text, txtNCM.text, txtNVE.text, _
            txtEXTIPI.text, txtCfop.text, txtUCom.text, txtQCom.text, txtVUnCom.text, txtVProd.text, _
            txtCEanTrib.text, txtUTrib.text, txtQtrib.text, txtVUnCom.text, txtVFrete.text, txtVSeg.text, txtVDesc.text, _
            txtVOutro.text, Left(cmbxIndTot.text, 1), txtCEST.text, txtInfAdProd.text) Then
            MsgBox "Dados alterados", vbInformation
        End If
        
    Else
        Me.tag = GravaNFe_Item(id_nfe, txtCProd.text, txtCean.text, txtxProd.text, txtNCM.text, txtNVE.text, _
            txtEXTIPI.text, txtCfop.text, txtUCom.text, txtQCom.text, txtVUnCom.text, txtVProd.text, _
            txtCEanTrib.text, txtUTrib.text, txtQtrib.text, txtVUnCom.text, txtVFrete.text, txtVSeg.text, txtVDesc.text, _
            txtVOutro.text, Left(cmbxIndTot.text, 1), txtCEST.text, txtInfAdProd.text)
            
            
        If Me.tag <> "" Then
        
              
        
            If NFe.GravaNfe_item_imposto(Me.tag, id_nfe, CStr(CCur(txtVProd.text) * crcTxFederativa_Par / 100), _
                     lblTipoIcms.caption, tipo_pis, tipo_cofins, "") Then
                    
                    If tipo_icms = "ICMSSN102" Or tipo_icms = "102" Then
                
                        If Not NFe.GravaNfe_item_imposto_icmssn102(Me.tag, id_nfe, Left(cmbOrigem.text, 1), TipoICMS.ICMSSN102) Then
                            MsgBox "erro na gravação do item_Imposto_icmssn102"
                        
                        End If
                    ElseIf tipo_icms = "ICMSSN500" Or tipo_icms = "500" Then
                        If Not NFe.GravaNfe_item_imposto_icmssn500(Me.tag, id_nfe, Left(cmbOrigem.text, 1), TipoICMS.ICMSSN500, "0", "0") Then
                            MsgBox "erro na gravação do item_Imposto_icmssn500"
                        
                        End If
                    ElseIf tpo_icms = "ICMS00" Then
                        If Not NFe.GravaNfe_item_imposto_icmsST(Me.tag, id_nfe, Left(cmbOrigem.text, 1), TipoICMS.ICMS00, "00", "0", "0", _
                        "0", "0") Then
                            MsgBox "erro na gravação do item_Imposto_icmssn00"
                            
                        End If
                    End If
                    
                    
                    If tipo_pis = "PISOutro" Or tipo_pis = "" Then
                    
                        If Not NFe.GravaNfe_item_imposto_PISOutr(Me.tag, id_nfe, "49", "0", "0", "0", "0", "0") Then
                            MsgBox "erro na gravação do item_Imposto_pisOutr"
                            
                        End If
                    ElseIf tipo_pis = "PISNT" Then
                        If Not NFe.GravaNfe_item_imposto_PISNT(Me.tag, id_nfe, lblCST_Pis.tag) Then
                            MsgBox "erro na gravação do item_Imposto_pisOutr"
                            
                        End If
                    End If
                    
                    If tipo_cofins = "COFINSOutro" Or tipo_cofins = "" Then
                        If Not NFe.GravaNfe_item_imposto_COFINSOutr(Me.tag, id_nfe, lblCST_Cofins.tag, "0", "0", "0", "0", "0") Then
                            MsgBox "erro na gravação do item_Imposto_pisOutr"
                        
                        End If
                    Else
                        If Not NFe.GravaNfe_item_imposto_COFINSNT(Me.tag, id_nfe, lblCST_Cofins.tag) Then
                            MsgBox "erro na gravação do item_Imposto_pisOutr"
                        
                        End If
                    End If
                                    
                Else
                    MsgBox "erro na gravação do item_Imposto"
                End If
        End If
    End If
        
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    
    bActive = True
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    P_PreencheNFETipoICMS cmbxTipoICMS
    
    strSQL = "select * from nfe_item as i inner join nfe_item_imposto as imp on i.id=imp.id "
    strSQL = strSQL & " and i.id=" & Me.tag
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            txtCean.text = !cEAN
            txtCEanTrib.text = !cEANTrib
            txtCfop.text = !CFOP
            txtCProd.text = !cProd
            
            txtEXTIPI.text = Format(!EXTIPI)
            txtNCM.text = Format(!ncm, "0#######")
            txtNVE.text = Format(!NVE)
            txtQCom.text = FormatVal(!qCom, "###,##0.000")
            txtQtrib.text = FormatVal(!qTrib, "###,##0.000")
            txtUCom.text = !uCom
            txtUTrib.text = !uTrib
            txtVUnCom.text = FormatVal(!vUnCom)
            txtVunTrib.text = FormatVal(!vUnTrib)
            txtVFrete.text = FormatVal(!vFrete)
            txtVDesc.text = FormatVal(!vDesc)
            txtVProd.text = FormatVal(!vProd)
            txtVSeg.text = FormatVal(!vSeg)
            txtVOutro.text = FormatVal(!vOutro)
            txtxProd.text = !xProd
            txtCEST.text = Format(!cEst)
            
            P_LocalizaOperacao !indTot, cmbxIndTot
            strSQL = F_SoNumero(!tipo_icms)
            P_LocalizaOperacao strSQL, cmbxTipoICMS
            txtInfAdProd.text = TiraNull(!infAdProd)
            lblTipoIcms.caption = !tipo_icms
            bActive = True
            buscaICMS
            
        End If
    End With
    bActive = True
    
    
End Sub
Sub buscaICMS()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rType As New ADODB.Recordset
    Dim tabela As String
    Dim i As Integer
    Dim j As Integer
    On Error GoTo trata_Erro
    If bActive = False Then Exit Sub
    tabela = F_BuscaUmCampo("tabela", "nfe_tipo_icms", "codigo='" & Trim(Left(cmbxTipoICMS.text, 3)) & "'")
    
    strSQL = "select * from " & tabela & " where id=" & Me.tag
    abreObjetoAdo rPesq, strSQL
    gridValores.Cols = 3
    gridValores.rows = 1
    gridValores.col = 0
    gridValores.text = "Campo"
    gridValores.col = 1
    gridValores.text = "Valor"
    gridValores.col = 2
    gridValores.text = "Tipo"
    gridValores.ColWidth(2) = 1000
    gridValores.rows = 1
    With rPesq
        If .RecordCount > 0 Then
            Do While Not .EOF
                For i = 0 To .Fields.Count - 1
                    If .Fields(i).Name <> "id" And .Fields(i).Name <> "id_nfe" Then
                        
                        strSQL = .Fields(i).Name & Chr(9) & .Fields(i).Value & Chr(9)
                        
                        'strSQL = strSQL & TypeName(.Fields(i).Type)
                        
                        If .Fields(i).Name = "CST" Or .Fields(i).Name = "CSOSN" Or .Fields(i).Name = "orig" Then
                            strSQL = strSQL & "String"
                        Else
                            strSQL = strSQL & "Double"
                        End If

                        gridValores.AddItem strSQL
                    End If
                Next
            
                .MoveNext
            
            Loop
        Else
            
            strSQL = "show fields from " & tabela
            abreObjetoAdo rType, strSQL
            Do While Not rType.EOF
                If rType!Field <> "id" And rType!Field <> "id_nfe" Then
                    
                    strSQL = rType!Field & Chr(9)
                    If rType!Field = "CST" Or rType!Field = "CSOSN" Or rType!Field = "orig" Then
                        strSQL = strSQL & Chr(9) & "String"
                    Else
                        strSQL = strSQL & Chr(9) & "Double"
                    End If
                    gridValores.AddItem strSQL
                End If
                
                rType.MoveNext
            Loop
        End If
        
        gridValores.row = 1
        gridValores.col = 1
        
            gridValores.text = Left(cmbOrigem.text, 1)
        If gridValores.text = "" Then
            gridValores.text = "0"
        End If
        
    
        
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Private Sub Form_Load()
    
    bActive = False
    sstb.Tab = 0
End Sub

Private Sub gridValores_DblClick()
    
    If gridValores.col = 1 And gridValores.rows > 1 Then
        row = gridValores.row
        txtGrid.Visible = True
        txtGrid.Left = gridValores.CellLeft + gridValores.Left
        txtGrid.top = gridValores.CellTop + gridValores.top
        txtGrid.width = gridValores.CellWidth
        txtGrid.text = gridValores.text
        
        txtGrid.SetFocus
        
    End If
End Sub

Private Sub gridValores_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        gridValores_DblClick
    End If
End Sub

Private Sub lblTipoIcms_Click()
    P_LocalizaOperacao lblTipoIcms.caption, cmbxTipoICMS
End Sub

Private Sub txtCEAN_Change()
    txtCEanTrib.text = txtCean.text
End Sub

Private Sub txtCEANTrib_GotFocus()
    P_SelecTexto txtCEanTrib
End Sub

Private Sub txtCEANTrib_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtCEANTrib_LostFocus()
    P_AjustaTexto txtCEanTrib
End Sub
Private Sub txtCean_GotFocus()
    P_SelecTexto txtCean
End Sub

Private Sub txtCean_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtCean_LostFocus()
    P_AjustaTexto txtCean
End Sub
Private Sub txtCProd_GotFocus()
    P_SelecTexto txtCProd
End Sub

Private Sub txtCProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtCProd_LostFocus()
    P_AjustaTexto txtCProd
End Sub

Private Sub txtGrid_GotFocus()
    P_SelecTexto txtGrid
End Sub

Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtGrid_LostFocus()
    If gridValores.col = 1 Then
        gridValores.row = row 'selecionada
        gridValores.text = txtGrid.text
    End If
    txtGrid.Visible = False
    
    
End Sub

Private Sub txtNCM_GotFocus()
    P_SelecTexto txtNCM
End Sub

Private Sub txtNCM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
        alteraNCM
    End If
    
End Sub
Sub alteraNCM()
    Dim strSQL As String
    
    strSQL = F_BuscaUmCampo("ncm_pro", TAB_PRODUTOS, "cod_pro='" & txtCProd.text & "'")
    If strSQL <> txtNCM.text Then
        If MsgBox("Deseja alterar o NCM do cadastro produto de " & strSQL & " para " & txtNCM.text & "?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update  " & TAB_PRODUTOS & " set ncm_pro='" & txtNCM.text & "' "
            strSQL = strSQL & " where cod_pro='" & txtCProd.text & "' "
            Conn.Execute strSQL
            alteraNCMAntigo
            cmndAlterar_Click
        End If
        
    End If
    
    
End Sub
Sub alteraNCMAntigo()
    
    Dim c As ADODB.Connection
    Dim strSQL As String
    FechaObjeto c
    
        
        
    Set c = New ADODB.Connection
    
    
    c.Open "muitomais"
    c.CursorLocation = adUseClient
    
    strSQL = "update  " & TAB_PRODUTOS & " set ncm_pro='" & txtNCM.text & "' "
    strSQL = strSQL & " where cod_pro='" & txtCProd.text & "' "
    
    c.Execute strSQL
    c.Close
    
    
End Sub
Private Sub txtNCM_LostFocus()
    P_AjustaTexto txtNCM
End Sub
Private Sub txtCFOP_GotFocus()
    P_SelecTexto txtCfop
End Sub

Private Sub txtCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtCFOP_LostFocus()
    P_AjustaTexto txtCfop
End Sub

Private Sub txtNVE_GotFocus()
    P_SelecTexto txtNVE
End Sub

Private Sub txtNVE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtNVE_LostFocus()
    P_AjustaTexto txtNVE
End Sub
Private Sub txtEXTIPI_GotFocus()
    P_SelecTexto txtEXTIPI
End Sub

Private Sub txtEXTIPI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtEXTIPI_LostFocus()
    P_AjustaTexto txtEXTIPI
End Sub
Private Sub txtCEST_GotFocus()
    P_SelecTexto txtCEST
End Sub

Private Sub txtCEST_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtCEST_LostFocus()
    P_AjustaTexto txtCEST
End Sub

Private Sub txtQCom_GotFocus()
    P_SelecTexto txtQCom
End Sub

Private Sub txtQCom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtQCom_LostFocus()
    P_ValorLostFocus txtQCom, "###,##0.000"
End Sub
Private Sub txtqtrib_GotFocus()
    P_SelecTexto txtQtrib
End Sub

Private Sub txtqtrib_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtqtrib_LostFocus()
    P_ValorLostFocus txtQtrib, "###,##0.000"
End Sub

Private Sub txtvDesc_GotFocus()
    P_SelecTexto txtVDesc
End Sub

Private Sub txtvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtvDesc_LostFocus()
    P_ValorLostFocus txtVDesc
End Sub

Private Sub txtvFrete_GotFocus()
    P_SelecTexto txtVFrete
End Sub

Private Sub txtvFrete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtvFrete_LostFocus()
    P_ValorLostFocus txtVFrete
End Sub

Private Sub txtvSeg_GotFocus()
    P_SelecTexto txtVSeg
End Sub

Private Sub txtvSeg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtvSeg_LostFocus()
    P_ValorLostFocus txtVSeg
End Sub
Private Sub txtvOutro_GotFocus()
    P_SelecTexto txtVOutro
End Sub

Private Sub txtvOutro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtvOutro_LostFocus()
    P_ValorLostFocus txtVOutro
End Sub
Private Sub txtvUnCom_GotFocus()
    P_SelecTexto txtVUnCom
End Sub
Private Sub txtvUnTrib_GotFocus()
    P_SelecTexto txtVunTrib
End Sub

Private Sub txtvUnTrib_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtvUnTrib_LostFocus()
    P_ValorLostFocus txtVunTrib
End Sub
Private Sub txtVProd_GotFocus()
    P_SelecTexto txtVProd
End Sub

Private Sub txtVProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtVProd_LostFocus()
    P_ValorLostFocus txtVProd
End Sub
Private Sub txtvUnCom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtvUnCom_LostFocus()
    P_ValorLostFocus txtVUnCom
End Sub
Private Sub txtUCom_GotFocus()
    P_SelecTexto txtUCom
End Sub

Private Sub txtUCom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtUCom_LostFocus()
    P_AjustaTexto txtUCom
End Sub
Private Sub txtUTrib_GotFocus()
    P_SelecTexto txtUTrib
End Sub

Private Sub txtUTrib_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtUTrib_LostFocus()
    P_AjustaTexto txtUTrib
End Sub
Private Sub txtxProd_GotFocus()
    P_SelecTexto txtxProd
End Sub

Private Sub txtxProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtxProd_LostFocus()
    P_AjustaTexto txtxProd
End Sub
