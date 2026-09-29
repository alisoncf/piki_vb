VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAlteraEstInv 
   AutoRedraw      =   -1  'True
   Caption         =   "Alterar Estoque"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab stb 
      Height          =   4215
      Left            =   240
      TabIndex        =   35
      Top             =   2400
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Inventários"
      TabPicture(0)   =   "frmAlteraEstInv.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Compras"
      TabPicture(1)   =   "frmAlteraEstInv.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTotalCompras"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "gridCompras"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "NFe"
      TabPicture(2)   =   "frmAlteraEstInv.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblTotalNFe"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "gridNfe"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin MSDataGridLib.DataGrid grid 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   5953
         _Version        =   393216
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
         Caption         =   "Inventários"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridCompras 
         Height          =   3255
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   5741
         _Version        =   393216
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
         Caption         =   "Compras"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridNfe 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   5530
         _Version        =   393216
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
         Caption         =   "Saídas NFe"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTotalCompras 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   13320
         TabIndex        =   40
         Top             =   3840
         Width           =   315
      End
      Begin VB.Label lblTotalNFe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   -61920
         TabIndex        =   39
         Top             =   3600
         Width           =   315
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<= Corrige estoque"
      Height          =   735
      Left            =   10080
      TabIndex        =   34
      Top             =   1440
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   9705
      TabIndex        =   15
      Top             =   1440
      Width           =   9735
      Begin VB.TextBox txtAcertos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   27
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtVendas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "atualizar dados"
         Height          =   375
         Left            =   6600
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtAte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtCompras 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtV1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acertos"
         Height          =   195
         Left            =   3000
         TabIndex        =   28
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas"
         Height          =   195
         Left            =   2040
         TabIndex        =   26
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   4200
         TabIndex        =   23
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         Height          =   195
         Left            =   5280
         TabIndex        =   21
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entradas:"
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NFe"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar (Esc)"
      Height          =   375
      Left            =   13560
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "<<= Alterar"
      Height          =   375
      Left            =   10080
      TabIndex        =   11
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtT2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtT1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtCusto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtQ2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtQ1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblReserv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   33
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservado"
      Height          =   195
      Left            =   7920
      TabIndex        =   32
      Top             =   720
      Width           =   780
   End
   Begin VB.Label lblEr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   31
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque"
      Height          =   195
      Left            =   6960
      TabIndex        =   30
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      Height          =   435
      Left            =   240
      TabIndex        =   29
      Top             =   120
      Width           =   210
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      Height          =   195
      Left            =   1440
      TabIndex        =   14
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblProduto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   13
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T2"
      Height          =   195
      Left            =   4920
      TabIndex        =   9
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T1"
      Height          =   195
      Left            =   3720
      TabIndex        =   7
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custo"
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q2"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q1"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   210
   End
   Begin VB.Label lblId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "frmAlteraEstInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim WithEvents adocompras As ADODB.Recordset
Attribute adocompras.VB_VarHelpID = -1
Dim WithEvents adoNFe As ADODB.Recordset
Attribute adoNFe.VB_VarHelpID = -1

Private Sub cmdAlterar_Click()
    Dim strSQL As String
    strSQL = "update inventario_it set q1=" & F_GravaDec(txtQ1.text) & ","
    strSQL = strSQL & "q2=" & F_GravaDec(txtQ2.text) & ","
    strSQL = strSQL & "Custo=" & F_GravaDec(txtCusto.text) & " "
    strSQL = strSQL & " where id= " & lblId.caption
    
    
    
    Conn.Execute strSQL
    MsgBox "Dados alterados!", vbInformation
    Unload Me
    
    
End Sub
Sub buscaCompras()
    Dim strSQL As String
    Dim cTotal As Currency
    strSQL = "select tc.seq_not,tc.num_not,nome_for as fornecedor,cod_pro,seq_ite,qtd_ite,unit_ite,custosai_ite from "
    strSQL = strSQL & " tab_compras as tc inner join tab_itcompras as ti on tc.seq_not=ti.seq_not "
    strSQL = strSQL & " inner join tab_fornecedores as tf on tc.cod_for=tf.cod_for "
    strSQL = strSQL & " where tc.emissao_not>=" & F_GravaData(txtDe.text)
    strSQL = strSQL & " and tc.emissao_not<=" & F_GravaData(txtAte.text)
    strSQL = strSQL & " and cod_pro='" & lblProduto.tag & "' "
    
    strSQL = strSQL & " and status_not>1 "
    
    abreObjetoAdo adocompras, strSQL, gridCompras
    With adocompras
        Do While Not .EOF
            cTotal = cTotal + !qtd_ite
            .MoveNext
        Loop
        
    End With
    txtCompras.text = cTotal
    
End Sub
Private Sub cmdBuscar_Click()
    busca
    buscaNFe
    buscaCompras
    total
End Sub
Sub buscaNFe()
    Dim strSQL As String
    Dim cTotal As Currency
    strSQL = "select n.id,n.nnf,n.dhemi,i.qcom,vUnCom,vProd from nfe as n inner join nfe_item as i on n.id=i.id_nfe"
    strSQL = strSQL & " where status=2 "
    strSQL = strSQL & " and left(dhemi,10) >=" & F_GravaData(txtDe.text)
    strSQL = strSQL & " and left(dhemi,10)<=" & F_GravaData(txtAte.text)
    strSQL = strSQL & " and cprod='" & lblProduto.tag & "' "
    
                
                
    abreObjetoAdo adoNFe, strSQL, gridNfe
                
    With adoNFe
        Do While Not .EOF
            cTotal = cTotal + !qCom
            .MoveNext
        Loop
    End With
    txtV1.text = FormatVal(cTotal)
                
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    With rPesq
        
        'If Not IsNull(!total) Then
            
        'Else
        '    txtV1.Text = "0"
        'End If
        
    End With
    strSQL = "select sum(ti.qtd_ite) as total from tab_itcompras as ti,"
    strSQL = strSQL & " tab_compras as ts where ts.seq_not=ti.seq_not "
    
    strSQL = strSQL & " and ti.cod_pro='" & lblProduto.tag & "' "
    strSQL = strSQL & " and ts.emissao_not>=" & F_GravaData(txtDe.text)
    strSQL = strSQL & " and ts.emissao_not<=" & F_GravaData(txtAte.text)
    strSQL = strSQL & " and ts.status_not>=2"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!total) Then
            txtCompras.text = FormatVal(!total)
        Else
            txtCompras.text = "0"
        End If
        .Close
    End With
    
    
     strSQL = "select sum(ti.qtd_ite) as total from tab_itvendas as ti,"
    strSQL = strSQL & " tab_vendas as ts where ts.num_ven=ti.num_ven "
    
    strSQL = strSQL & " and ti.cod_pro='" & lblProduto.tag & "' "
    strSQL = strSQL & " and ts.dataefet_ven>=" & F_GravaData(txtDe.text)
    strSQL = strSQL & " and ts.dataefet_ven<=" & F_GravaData(txtAte.text)
    strSQL = strSQL & " and ts.situacao_ven=4 "
    strSQL = strSQL & " and ts.tipo_operacao=1"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!total) Then
            txtVendas.text = FormatVal(!total)
        Else
            txtVendas.text = "0"
        End If
        .Close
    End With
    
    strSQL = "select sum(ti.qtd_ite) as total from tab_itacertos as ti,"
    strSQL = strSQL & " tab_acertos as ts where ts.seq_ace=ti.seq_ace "
    
    strSQL = strSQL & " and ti.cod_pro='" & lblProduto.tag & "' "
    strSQL = strSQL & " and ts.data_ace>=" & F_GravaData(txtDe.text)
    strSQL = strSQL & " and ts.data_ace<=" & F_GravaData(txtAte.text)
    strSQL = strSQL & " and status_ace=2"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!total) Then
            txtAcertos.text = FormatVal(!total)
        Else
            txtAcertos.text = "0"
        End If
        FechaObjeto rPesq
    End With
    strSQL = " select r_pro,er_pro "
    strSQL = strSQL & " from tab_produtos where cod_pro='" & lblProduto.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
       lblEr.caption = !er_pro
       lblReserv.caption = !r_pro
        .Close
    End With
    buscaInv
End Sub
Sub buscaInv()
    Set ado = New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select i.id,i.data,i.de,i.ate,it.q1, it.q2, it.custo, it.nome from inventario as i,inventario_it as it where it.inventario=i.id "
    strSQL = strSQL & " and codigo='" & lblProduto.tag & "' "
    strSQL = strSQL & " order by i.id desc "
    abreObjetoAdo ado, strSQL, grid
    
End Sub
Private Sub cmdFechar_Click()
    Unload Me
End Sub

Sub total()
    On Error Resume Next
    txtT1.text = FormatVal(CCur(txtQ1.text) * CCur(txtCusto.text))
    txtT2.text = FormatVal(CCur(txtQ2.text) * CCur(txtCusto.text))
        
        On erro GoTo trata_Erro
    Dim cTotal As Currency
    cTotal = 0
    With adocompras
        Do While Not .EOF
            cTotal = cTotal + !qtd_ite
            .MoveNext
        Loop
    End With
    lblTotalCompras.caption = cTotal
    
    
    cTotal = 0
    With adoNFe
        Do While Not .EOF
            cTotal = cTotal + !qCom
            .MoveNext
        Loop
    End With
    lblTotal.caption = cTotal
    
trata_Erro:
    
End Sub

Private Sub Command1_Click()
    strSQL = "update tab_produtos set ef_pro=" & F_GravaDec(txtQ1.text) & ","
    strSQL = strSQL & "er_pro=" & F_GravaDec(txtQ2.text) & ","
    strSQL = strSQL & "cr_pro=" & F_GravaDec(txtCusto.text) & " "
    strSQL = strSQL & " where cod_pro='" & lblProduto.tag & "' "
    
    
    Conn.Execute strSQL
    cmdAlterar_Click
'    MsgBox "Dados alterados!", vbInformation
'    Unload Me
    
End Sub

Private Sub Form_Activate()
    busca
    buscaNFe
    buscaCompras
    total
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Form_Load()
    stb.Tab = 0
    txtDe.text = "01/01/" & Year(Date) - 1
    txtAte.text = "31/12/" & Year(Date) - 1
    Set ado = New ADODB.Recordset
    Set adocompras = New ADODB.Recordset
    Set adoNFe = New ADODB.Recordset
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtCusto_GotFocus()
    P_SelecTexto txtCusto
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub

Private Sub txtQ1_GotFocus()
    P_SelecTexto txtQ1
End Sub

Private Sub txtQ1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQ1_LostFocus()
    P_ValorLostFocus txtQ1
    total
End Sub

Private Sub txtQ2_GotFocus()
    P_SelecTexto txtQ2
End Sub
Private Sub txtCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCusto_LostFocus()
    P_ValorLostFocus txtCusto
    total
End Sub


Private Sub txtQ2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQ2_LostFocus()
    P_ValorLostFocus txtQ2
    total
End Sub
