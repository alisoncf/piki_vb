VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmNFE_Detalhe 
   Caption         =   "Detalhe da Nota Fiscal Eletrônica"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16065
   Icon            =   "frmNFE_Detalhe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   16065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   16065
      TabIndex        =   158
      Top             =   7830
      Width           =   16065
      Begin VB.CommandButton cmdRetransm 
         Caption         =   "Retransmitir"
         Height          =   735
         Left            =   1560
         Picture         =   "frmNFE_Detalhe.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCupom 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4440
         Picture         =   "frmNFE_Detalhe.frx":2594
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar NFe"
         Height          =   735
         Left            =   7320
         Picture         =   "frmNFE_Detalhe.frx":2A42
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar NFe"
         Height          =   735
         Left            =   3000
         Picture         =   "frmNFE_Detalhe.frx":2FD0
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdAtlzRetorno 
         Caption         =   "Atualizar Retorno"
         Height          =   735
         Left            =   10200
         Picture         =   "frmNFE_Detalhe.frx":389A
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   13080
         Picture         =   "frmNFE_Detalhe.frx":3DA2
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   735
         Left            =   8760
         Picture         =   "frmNFE_Detalhe.frx":3FEC
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdInutilizar 
         Caption         =   "Inutilizar"
         Height          =   735
         Left            =   11640
         Picture         =   "frmNFE_Detalhe.frx":44C0
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   735
         Left            =   120
         Picture         =   "frmNFE_Detalhe.frx":54DF
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   120
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab sstb 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "frmNFE_Detalhe.frx":5A1B
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "pcInfAdc"
      Tab(0).Control(1)=   "pcDadosNFe"
      Tab(0).Control(2)=   "pcCont"
      Tab(0).Control(3)=   "Picture7"
      Tab(0).Control(4)=   "cmdSaveDados"
      Tab(0).Control(5)=   "pcNfeRef"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Emitente"
      TabPicture(1)   =   "frmNFE_Detalhe.frx":5A37
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(1)=   "Picture4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Destinatário"
      TabPicture(2)   =   "frmNFE_Detalhe.frx":5A53
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pcDestEnd"
      Tab(2).Control(1)=   "pcDestCab"
      Tab(2).Control(2)=   "cmdSaveDest"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Produtos"
      TabPicture(3)   =   "frmNFE_Detalhe.frx":5A6F
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblTotalProd"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "gridCofins"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "gridPIS"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "gridIPI"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "gridICMS"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "gridProdutos"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Picture5"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "pcCFOP"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Totais"
      TabPicture(4)   =   "frmNFE_Detalhe.frx":5A8B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "pcTotais"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Transporte"
      TabPicture(5)   =   "frmNFE_Detalhe.frx":5AA7
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture2"
      Tab(5).Control(1)=   "cmbxModFrete"
      Tab(5).Control(2)=   "pcTransportadora"
      Tab(5).Control(3)=   "cmdSaveTransp"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Pagamentos"
      TabPicture(6)   =   "frmNFE_Detalhe.frx":5AC3
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblTotalPagCap"
      Tab(6).Control(1)=   "lblTotalPag"
      Tab(6).Control(2)=   "lblVNfe"
      Tab(6).Control(3)=   "Label5"
      Tab(6).Control(4)=   "gridPagamentos"
      Tab(6).Control(5)=   "pcPag"
      Tab(6).Control(6)=   "cmdExcluirPagamento"
      Tab(6).Control(7)=   "cmdAltCPAg"
      Tab(6).ControlCount=   8
      Begin VB.CommandButton cmdAltCPAg 
         Caption         =   "Alterar"
         Height          =   255
         Left            =   -73680
         TabIndex        =   309
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdExcluirPagamento 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   -74760
         TabIndex        =   305
         Top             =   2880
         Width           =   975
      End
      Begin VB.PictureBox pcNfeRef 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   -66600
         ScaleHeight     =   1185
         ScaleWidth      =   5745
         TabIndex        =   301
         Top             =   3000
         Width           =   5775
         Begin VB.TextBox txtNfeRef 
            Appearance      =   0  'Flat
            Height          =   885
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   302
            Top             =   240
            Width           =   5415
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Notas Referenciadas"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   34
            Left            =   2040
            TabIndex        =   303
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.PictureBox pcCFOP 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         ScaleHeight     =   825
         ScaleWidth      =   2745
         TabIndex        =   296
         Top             =   6000
         Width           =   2775
         Begin VB.CommandButton cmdAlterarCFOP 
            Caption         =   "Alterar Todos"
            Height          =   315
            Left            =   840
            TabIndex        =   299
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtCFOP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   298
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CFOP"
            Height          =   195
            Left            =   120
            TabIndex        =   297
            Top             =   120
            Width           =   420
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         ScaleHeight     =   465
         ScaleWidth      =   15345
         TabIndex        =   288
         Top             =   6960
         Width           =   15375
         Begin VB.CommandButton cmdSemGtin 
            Caption         =   "Sem GTIN"
            Height          =   255
            Left            =   4800
            TabIndex        =   323
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdEAN 
            Caption         =   "Testar EAN"
            Height          =   255
            Left            =   3240
            TabIndex        =   304
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdNCM 
            Caption         =   "Atualizar NCMs"
            Height          =   255
            Left            =   1680
            TabIndex        =   295
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdEditProd 
            Caption         =   "Editar Produto"
            Height          =   255
            Left            =   120
            TabIndex        =   290
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdAlterarProd 
            Caption         =   "Alterar Item"
            Height          =   255
            Left            =   13560
            TabIndex        =   289
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdSaveDados 
         Caption         =   "Salvar Alterações"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -62880
         TabIndex        =   286
         Top             =   4320
         Width           =   2055
      End
      Begin VB.PictureBox pcPag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -74880
         ScaleHeight     =   1305
         ScaleWidth      =   13785
         TabIndex        =   276
         Top             =   3240
         Width           =   13815
         Begin VB.CommandButton cmdCancelarPag 
            Caption         =   "Cancelar"
            Height          =   255
            Left            =   12840
            TabIndex        =   317
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cmbxTBand 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":5ADF
            Left            =   6960
            List            =   "frmNFE_Detalhe.frx":5B01
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   313
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtTpIntegra 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12240
            TabIndex        =   307
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmdAddPag 
            Caption         =   "Adicionar"
            Height          =   255
            Left            =   12840
            TabIndex        =   285
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtCAut 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9960
            TabIndex        =   284
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtCNPJPag 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   282
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtValorPag 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   280
            Text            =   "0,00"
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox cmbxTpag 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":5B94
            Left            =   120
            List            =   "frmNFE_Detalhe.frx":5BBC
            Style           =   2  'Dropdown List
            TabIndex        =   278
            Top             =   360
            Width           =   2775
         End
         Begin VB.CommandButton cmdFinanceira 
            Height          =   285
            Left            =   6480
            Picture         =   "frmNFE_Detalhe.frx":5CAB
            Style           =   1  'Graphical
            TabIndex        =   311
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "tBand"
            Height          =   195
            Index           =   6
            Left            =   9240
            TabIndex        =   312
            Top             =   120
            Width           =   420
         End
         Begin VB.Label lblOperadora 
            Caption         =   "Operadora"
            Height          =   495
            Left            =   4320
            TabIndex        =   310
            Top             =   720
            Width           =   5535
         End
         Begin VB.Label lblPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TP Integra"
            Height          =   195
            Index           =   5
            Left            =   12240
            TabIndex        =   308
            Top             =   120
            Width           =   750
         End
         Begin VB.Label lblPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Autorização:"
            Height          =   195
            Index           =   4
            Left            =   9960
            TabIndex        =   283
            Top             =   120
            Width           =   1110
         End
         Begin VB.Label lblPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ Operadora:"
            Height          =   195
            Index           =   2
            Left            =   4320
            TabIndex        =   281
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label lblPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   279
            Top             =   120
            Width           =   405
         End
         Begin VB.Label lblPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalide:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   277
            Top             =   120
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdSaveTransp 
         Caption         =   "Salvar Alterações"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -63120
         TabIndex        =   274
         Top             =   6480
         Width           =   1935
      End
      Begin VB.PictureBox pcTransportadora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   -74760
         ScaleHeight     =   4785
         ScaleWidth      =   13545
         TabIndex        =   209
         Top             =   1560
         Visible         =   0   'False
         Width           =   13575
         Begin VB.CommandButton cmdBuscaTransp 
            Height          =   285
            Left            =   1680
            Picture         =   "frmNFE_Detalhe.frx":5FFB
            Style           =   1  'Graphical
            TabIndex        =   306
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtLacres 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   273
            Top             =   4200
            Width           =   7815
         End
         Begin VB.TextBox txtPesoB 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12000
            TabIndex        =   270
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox txtPesoL 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10560
            TabIndex        =   268
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox txtNVol 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            TabIndex        =   266
            Top             =   3600
            Width           =   2415
         End
         Begin VB.TextBox txtMarcaVol 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4080
            TabIndex        =   264
            Top             =   3600
            Width           =   3855
         End
         Begin VB.TextBox txtEspecieVol 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   262
            Top             =   3600
            Width           =   2535
         End
         Begin VB.TextBox txtQVol 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   260
            Top             =   3600
            Width           =   1215
         End
         Begin VB.TextBox txtBalsa 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10560
            TabIndex        =   257
            Top             =   2760
            Width           =   2775
         End
         Begin VB.TextBox txtVagao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            TabIndex        =   255
            Top             =   2760
            Width           =   2415
         End
         Begin VB.TextBox txtRNTCReboque 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6000
            TabIndex        =   253
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox txtUFReboque 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5400
            TabIndex        =   251
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtPlacaReboque 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4080
            TabIndex        =   249
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtRNTCVeicTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   246
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox txtUfVeicTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   244
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   242
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtCMunFGTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12360
            TabIndex        =   238
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtMuncipioGeradorTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7080
            TabIndex        =   236
            Top             =   1920
            Width           =   4815
         End
         Begin VB.TextBox txtCFOPTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5400
            TabIndex        =   235
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtVIcmsRetTransp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4080
            TabIndex        =   232
            Text            =   "0,00"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtPIcmsRetTransp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            TabIndex        =   230
            Text            =   "0,00"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtvBCRetTransp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   228
            Text            =   "0,00"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtVServTransp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   225
            Text            =   "0,00"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtUFTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12840
            TabIndex        =   223
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtXMunTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7080
            TabIndex        =   221
            Top             =   1080
            Width           =   5655
         End
         Begin VB.TextBox txtXEnderTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   219
            Top             =   1080
            Width           =   6855
         End
         Begin VB.TextBox txtIETransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11160
            TabIndex        =   213
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtCNPJTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   212
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtNomeTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            TabIndex        =   211
            Top             =   480
            Width           =   6855
         End
         Begin VB.TextBox txtCPFTransp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   210
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton cmdMunGerador 
            Height          =   285
            Left            =   11880
            Picture         =   "frmNFE_Detalhe.frx":634B
            Style           =   1  'Graphical
            TabIndex        =   239
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lacres"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   33
            Left            =   120
            TabIndex        =   272
            Top             =   3960
            Width           =   570
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso B."
            Height          =   195
            Index           =   32
            Left            =   12000
            TabIndex        =   271
            Top             =   3360
            Width           =   555
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso L."
            Height          =   195
            Index           =   31
            Left            =   10560
            TabIndex        =   269
            Top             =   3360
            Width           =   540
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numeração:"
            Height          =   195
            Index           =   30
            Left            =   8040
            TabIndex        =   267
            Top             =   3360
            Width           =   870
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca:"
            Height          =   195
            Index           =   29
            Left            =   4080
            TabIndex        =   265
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Espécie"
            Height          =   195
            Index           =   28
            Left            =   1440
            TabIndex        =   263
            Top             =   3360
            Width           =   570
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade"
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   261
            Top             =   3360
            Width           =   825
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Volumes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   26
            Left            =   120
            TabIndex        =   259
            Top             =   3120
            Width           =   750
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balsa"
            Height          =   195
            Index           =   25
            Left            =   10560
            TabIndex        =   258
            Top             =   2520
            Width           =   390
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vagão:"
            Height          =   195
            Index           =   24
            Left            =   8040
            TabIndex        =   256
            Top             =   2520
            Width           =   510
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RNTC"
            Height          =   195
            Index           =   22
            Left            =   6000
            TabIndex        =   254
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
            Height          =   195
            Index           =   21
            Left            =   5400
            TabIndex        =   252
            Top             =   2520
            Width           =   210
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Placa:"
            Height          =   195
            Index           =   20
            Left            =   4080
            TabIndex        =   250
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label Label113 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reboque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4080
            TabIndex        =   248
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RNTC"
            Height          =   195
            Index           =   19
            Left            =   2040
            TabIndex        =   247
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
            Height          =   195
            Index           =   18
            Left            =   1440
            TabIndex        =   245
            Top             =   2520
            Width           =   210
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Placa:"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   243
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações de Veículo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   120
            TabIndex        =   241
            Top             =   2280
            Width           =   2010
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Index           =   15
            Left            =   12360
            TabIndex        =   240
            Top             =   1680
            Width           =   540
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Município Gerador:"
            Height          =   195
            Index           =   14
            Left            =   7080
            TabIndex        =   237
            Top             =   1680
            Width           =   1365
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CFOP"
            Height          =   195
            Index           =   13
            Left            =   5400
            TabIndex        =   234
            Top             =   1680
            Width           =   420
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. ICMS Ret"
            Height          =   195
            Index           =   12
            Left            =   4080
            TabIndex        =   233
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Perc. ICMS Ret."
            Height          =   195
            Index           =   11
            Left            =   2760
            TabIndex        =   231
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr BC. Ret."
            Height          =   195
            Index           =   10
            Left            =   1440
            TabIndex        =   229
            Top             =   1680
            Width           =   825
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações de Retençaõ do ICMS Transporte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   120
            TabIndex        =   227
            Top             =   1440
            Width           =   3810
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr Serviço"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   226
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF:"
            Height          =   195
            Index           =   6
            Left            =   12840
            TabIndex        =   224
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Município:"
            Height          =   195
            Index           =   5
            Left            =   7080
            TabIndex        =   222
            Top             =   840
            Width           =   750
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   220
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IE"
            Height          =   195
            Index           =   3
            Left            =   11160
            TabIndex        =   218
            Top             =   240
            Width           =   150
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transportadora"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   217
            Top             =   0
            Width           =   1290
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   216
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   2
            Left            =   4200
            TabIndex        =   215
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPF:"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   214
            Top             =   240
            Width           =   345
         End
      End
      Begin VB.ComboBox cmbxModFrete 
         Height          =   315
         ItemData        =   "frmNFE_Detalhe.frx":669B
         Left            =   -74520
         List            =   "frmNFE_Detalhe.frx":66AB
         Style           =   2  'Dropdown List
         TabIndex        =   207
         Top             =   960
         Width           =   4695
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -74760
         ScaleHeight     =   825
         ScaleWidth      =   13545
         TabIndex        =   206
         Top             =   600
         Width           =   13575
         Begin VB.Label lblTransp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidade de Frete:"
            Height          =   195
            Index           =   23
            Left            =   240
            TabIndex        =   208
            Top             =   120
            Width           =   1500
         End
      End
      Begin VB.PictureBox pcTotais 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   3345
         ScaleWidth      =   13665
         TabIndex        =   173
         Top             =   600
         Width           =   13695
         Begin VB.CommandButton cmdSalvarTotais 
            Caption         =   "Salvar Alterações"
            Enabled         =   0   'False
            Height          =   615
            Left            =   11520
            TabIndex        =   314
            Top             =   2520
            Width           =   2055
         End
         Begin VB.TextBox txtVTotTrib 
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
            Height          =   285
            Left            =   4200
            TabIndex        =   204
            Text            =   "0,00"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox txtVNF 
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
            Height          =   285
            Left            =   2880
            TabIndex        =   202
            Text            =   "0,00"
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtVOutro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            TabIndex        =   200
            Text            =   "0,00"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtVCofins 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   198
            Text            =   "0,00"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtVPIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   196
            Text            =   "0,00"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtVIPI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   194
            Text            =   "0,00"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtVII 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   192
            Text            =   "0,00"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtVDesc 
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
            Height          =   285
            Left            =   1560
            TabIndex        =   190
            Text            =   "0,00"
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtvSeg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   188
            Text            =   "0,00"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtVFrete 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   186
            Text            =   "0,00"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtVProd 
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
            Height          =   285
            Left            =   240
            TabIndex        =   184
            Text            =   "0,00"
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtVST 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5520
            TabIndex        =   182
            Text            =   "0,00"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtvBCST 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            TabIndex        =   180
            Text            =   "0,00"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtvICMSDeson 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   178
            Text            =   "0,00"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtVICms 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   176
            Text            =   "0,00"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtvBC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   174
            Text            =   "0,00"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label91 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Tributado"
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
            Left            =   4200
            TabIndex        =   205
            Top             =   2520
            Width           =   1320
         End
         Begin VB.Label Label90 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr NF:"
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
            Left            =   2880
            TabIndex        =   203
            Top             =   2520
            Width           =   600
         End
         Begin VB.Label Label89 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr Outro"
            Height          =   195
            Left            =   4200
            TabIndex        =   201
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label88 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr COFINS"
            Height          =   195
            Left            =   2880
            TabIndex        =   199
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr PIS"
            Height          =   195
            Left            =   1560
            TabIndex        =   197
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr IPI"
            Height          =   195
            Left            =   240
            TabIndex        =   195
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr I. Import."
            Height          =   195
            Left            =   2880
            TabIndex        =   193
            Top             =   720
            Width           =   840
         End
         Begin VB.Label Label84 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr Desconto"
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
            Left            =   1560
            TabIndex        =   191
            Top             =   2520
            Width           =   1110
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr Seguro"
            Height          =   195
            Left            =   1560
            TabIndex        =   189
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Frete"
            Height          =   195
            Left            =   240
            TabIndex        =   187
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label81 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr Produtos"
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
            Left            =   240
            TabIndex        =   185
            Top             =   2520
            Width           =   1050
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr ST"
            Height          =   195
            Left            =   5520
            TabIndex        =   183
            Top             =   120
            Width           =   435
         End
         Begin VB.Label Label79 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr BC ST"
            Height          =   195
            Left            =   4200
            TabIndex        =   181
            Top             =   120
            Width           =   690
         End
         Begin VB.Label Label78 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr ICMS Des."
            Height          =   195
            Left            =   2880
            TabIndex        =   179
            Top             =   120
            Width           =   990
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr ICMS"
            Height          =   195
            Left            =   1560
            TabIndex        =   177
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr BC:"
            Height          =   195
            Left            =   240
            TabIndex        =   175
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.CommandButton cmdSaveDest 
         Caption         =   "Salvar Alterações"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -61200
         TabIndex        =   172
         Top             =   5160
         Width           =   1935
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   -74880
         ScaleHeight     =   1185
         ScaleWidth      =   14025
         TabIndex        =   142
         Top             =   5880
         Width           =   14055
         Begin VB.TextBox txtUsuario 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3840
            TabIndex        =   156
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtDigestValue 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            TabIndex        =   153
            Top             =   840
            Width           =   5415
         End
         Begin VB.TextBox txtRecibo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3840
            TabIndex        =   152
            Top             =   840
            Width           =   4095
         End
         Begin VB.TextBox txtProtocolo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   150
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox txtCaminhoXML 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4920
            TabIndex        =   148
            Top             =   240
            Width           =   8175
         End
         Begin VB.ComboBox cmbStatus 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":671E
            Left            =   120
            List            =   "frmNFE_Detalhe.frx":672E
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtVenda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   145
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdAbrirXML 
            Caption         =   "..."
            Height          =   285
            Left            =   13080
            TabIndex        =   155
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   27
            Left            =   3840
            TabIndex        =   157
            Top             =   0
            Width           =   540
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Digest Value"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   31
            Left            =   8040
            TabIndex        =   154
            Top             =   600
            Width           =   900
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Recibo:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   30
            Left            =   3840
            TabIndex        =   151
            Top             =   600
            Width           =   555
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Protocolo:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   149
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho XML:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   28
            Left            =   4920
            TabIndex        =   147
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Venda:"
            Height          =   195
            Index           =   26
            Left            =   2520
            TabIndex        =   144
            Top             =   0
            Width           =   510
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   143
            Top             =   0
            Width           =   495
         End
      End
      Begin MSDataGridLib.DataGrid gridProdutos 
         Height          =   5055
         Left            =   240
         TabIndex        =   141
         Top             =   480
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         FormatLocked    =   -1  'True
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
         Caption         =   "Informações de Produtos da NFe"
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "cprod"
            Caption         =   "Código"
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
            DataField       =   "cean"
            Caption         =   "EAN"
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
         BeginProperty Column02 
            DataField       =   "ncm"
            Caption         =   "NCM"
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
         BeginProperty Column03 
            DataField       =   "xProd"
            Caption         =   "Produto"
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
         BeginProperty Column04 
            DataField       =   "cfop"
            Caption         =   "CFOP"
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
         BeginProperty Column05 
            DataField       =   "qcom"
            Caption         =   "QCom"
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
         BeginProperty Column06 
            DataField       =   "vuncom"
            Caption         =   "Unitário"
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
         BeginProperty Column07 
            DataField       =   "vdesc"
            Caption         =   "Desc."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "tipo_icms"
            Caption         =   "Tipo ICMS"
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
         BeginProperty Column09 
            DataField       =   "tipo_cofins"
            Caption         =   "Tipo COFINS"
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
         BeginProperty Column10 
            DataField       =   "tipo_pis"
            Caption         =   "Tipo PIS"
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
         BeginProperty Column11 
            DataField       =   "tipo_ipi"
            Caption         =   "Tipo IPI"
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
         BeginProperty Column12 
            DataField       =   "cProdAnp"
            Caption         =   "cProdANP"
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
         BeginProperty Column13 
            DataField       =   "descAnp"
            Caption         =   "DescANP"
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
         BeginProperty Column14 
            DataField       =   "UFCons"
            Caption         =   "UFCons"
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   4665,26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pcDestCab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   -74760
         ScaleHeight     =   945
         ScaleWidth      =   15465
         TabIndex        =   119
         Top             =   600
         Width           =   15495
         Begin VB.CommandButton cmdProprioEmitente 
            Caption         =   "Setar Emitente como Destinatário"
            Height          =   495
            Left            =   13560
            TabIndex        =   322
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdPesqCli 
            Caption         =   "Pesquisar Cliente"
            Height          =   315
            Left            =   9600
            TabIndex        =   300
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdCliente 
            Caption         =   "Editar Cliente"
            Height          =   315
            Left            =   8280
            TabIndex        =   287
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtCPFDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   137
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtXNomeDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            TabIndex        =   126
            Top             =   480
            Width           =   6855
         End
         Begin VB.TextBox txtCNPJDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   124
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtIDEstrangeiro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11160
            TabIndex        =   122
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPF:"
            Height          =   195
            Left            =   2160
            TabIndex        =   139
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Left            =   4200
            TabIndex        =   135
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   120
            TabIndex        =   133
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destinatário"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   131
            Top             =   0
            Width           =   990
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID Estrangeiro:"
            Height          =   195
            Left            =   11160
            TabIndex        =   129
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.PictureBox pcDestEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   -74760
         ScaleHeight     =   3345
         ScaleWidth      =   15465
         TabIndex        =   96
         Top             =   1680
         Width           =   15495
         Begin VB.TextBox txtComplDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5160
            TabIndex        =   117
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox txtLogrDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   112
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox txtNumDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            TabIndex        =   114
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtBairroDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            TabIndex        =   120
            Top             =   480
            Width           =   4815
         End
         Begin VB.TextBox txtCodMunDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   123
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtNomeMunDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   127
            Top             =   1080
            Width           =   4695
         End
         Begin VB.TextBox txtUFDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6840
            TabIndex        =   132
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCEPDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7440
            TabIndex        =   136
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtCOdPaisDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            TabIndex        =   140
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtPaisDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9840
            TabIndex        =   103
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox txtFoneDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   102
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox txtIEDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            TabIndex        =   101
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtISUF 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            TabIndex        =   100
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtIMDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5760
            TabIndex        =   99
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtEmailDest 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   98
            Top             =   2280
            Width           =   6495
         End
         Begin VB.ComboBox cmbxIndIEDest 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":6769
            Left            =   7200
            List            =   "frmNFE_Detalhe.frx":6776
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   1680
            Width           =   6150
         End
         Begin VB.CommandButton cmdCidade 
            Height          =   285
            Left            =   6240
            Picture         =   "frmNFE_Detalhe.frx":6864
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
            Height          =   195
            Left            =   5160
            TabIndex        =   115
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   116
            Top             =   0
            Width           =   825
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logradouro:"
            Height          =   195
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
            Height          =   195
            Left            =   4200
            TabIndex        =   113
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   8520
            TabIndex        =   118
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código Município"
            Height          =   195
            Left            =   120
            TabIndex        =   121
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do Município:"
            Height          =   195
            Left            =   1560
            TabIndex        =   125
            Top             =   840
            Width           =   1440
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF:"
            Height          =   195
            Left            =   6840
            TabIndex        =   130
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CEP:"
            Height          =   195
            Left            =   7440
            TabIndex        =   134
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código País:"
            Height          =   195
            Left            =   8520
            TabIndex        =   138
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "País:"
            Height          =   195
            Left            =   9840
            TabIndex        =   110
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fone:"
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inscrição Estadual"
            Height          =   195
            Left            =   2400
            TabIndex        =   108
            Top             =   1440
            Width           =   1305
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ISUF"
            Height          =   195
            Left            =   4560
            TabIndex        =   107
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IM:"
            Height          =   195
            Left            =   5760
            TabIndex        =   106
            Top             =   1440
            Width           =   225
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "email:"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador da IE:"
            Height          =   195
            Left            =   7200
            TabIndex        =   104
            Top             =   1440
            Width           =   1125
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   -74760
         ScaleHeight     =   2385
         ScaleWidth      =   15105
         TabIndex        =   62
         Top             =   1680
         Width           =   15135
         Begin VB.ComboBox cmbxCRTEmit 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":6BB4
            Left            =   10680
            List            =   "frmNFE_Detalhe.frx":6BC1
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   1680
            Width           =   2790
         End
         Begin VB.TextBox txtCNAEEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   92
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtIMEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   90
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtIESTEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtIEEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtFoneEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox txtNomePaisEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   1080
            Width           =   4815
         End
         Begin VB.TextBox txtCodPaisEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCEP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtUFEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtNomeMunEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox txtCodMunEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtBairroEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   480
            Width           =   4815
         End
         Begin VB.TextBox txtNumEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtXLogrEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox txtComplEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código do Regime Tributário:"
            Height          =   195
            Left            =   10680
            TabIndex        =   94
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNAE Fiscal"
            Height          =   195
            Left            =   8520
            TabIndex        =   93
            Top             =   1440
            Width           =   885
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inscrição Municipal:"
            Height          =   195
            Left            =   6720
            TabIndex        =   91
            Top             =   1440
            Width           =   1410
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I.E. Subst. Tributário:"
            Height          =   195
            Left            =   4560
            TabIndex        =   89
            Top             =   1440
            Width           =   1485
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inscrição Estadual"
            Height          =   195
            Left            =   2400
            TabIndex        =   87
            Top             =   1440
            Width           =   1305
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fone:"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "País:"
            Height          =   195
            Left            =   8520
            TabIndex        =   83
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código País:"
            Height          =   195
            Left            =   7200
            TabIndex        =   81
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CEP:"
            Height          =   195
            Left            =   5760
            TabIndex        =   79
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF:"
            Height          =   195
            Left            =   5160
            TabIndex        =   77
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do Município:"
            Height          =   195
            Left            =   1560
            TabIndex        =   75
            Top             =   840
            Width           =   1440
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código Município"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   8520
            TabIndex        =   71
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
            Height          =   195
            Left            =   4200
            TabIndex        =   69
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logradouro:"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   67
            Top             =   0
            Width           =   825
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
            Height          =   195
            Left            =   5160
            TabIndex        =   66
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   -74760
         ScaleHeight     =   945
         ScaleWidth      =   15225
         TabIndex        =   54
         Top             =   600
         Width           =   15255
         Begin VB.TextBox txtXFantEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   480
            Width           =   4815
         End
         Begin VB.TextBox txtCNPJEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtxNomeEmit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   480
            Width           =   6255
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fantasia:"
            Height          =   195
            Left            =   8520
            TabIndex        =   61
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Emitente:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   59
            Top             =   0
            Width           =   780
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Left            =   2160
            TabIndex        =   57
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.PictureBox pcCont 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   -74880
         ScaleHeight     =   1185
         ScaleWidth      =   8145
         TabIndex        =   48
         Top             =   3000
         Width           =   8175
         Begin VB.TextBox txtJust 
            Appearance      =   0  'Flat
            Height          =   885
            Left            =   2040
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   6015
         End
         Begin VB.TextBox txtDHCont 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contingência:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   22
            Left            =   120
            TabIndex        =   53
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Justificativa:"
            Height          =   195
            Index           =   24
            Left            =   4680
            TabIndex        =   52
            Top             =   0
            Width           =   870
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Hora"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox pcDadosNFe 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   -74880
         ScaleHeight     =   2385
         ScaleWidth      =   14025
         TabIndex        =   1
         Top             =   480
         Width           =   14055
         Begin VB.TextBox txtNum_Ven 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12480
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox txtVerProc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11400
            TabIndex        =   44
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox cmbxProcEmi 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":6C30
            Left            =   6000
            List            =   "frmNFE_Detalhe.frx":6C40
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   2040
            Width           =   5295
         End
         Begin VB.ComboBox cmbxIndPres 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":6D3A
            Left            =   2160
            List            =   "frmNFE_Detalhe.frx":6D50
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   2040
            Width           =   3735
         End
         Begin VB.ComboBox cmbxIndFinal 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":6E2E
            Left            =   120
            List            =   "frmNFE_Detalhe.frx":6E38
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   2040
            Width           =   1935
         End
         Begin VB.ComboBox cmbxFinNfe 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":6E5E
            Left            =   11400
            List            =   "frmNFE_Detalhe.frx":6E6E
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox txtDV 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   13080
            TabIndex        =   42
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cmbxTpAmb 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":6ECB
            Left            =   9240
            List            =   "frmNFE_Detalhe.frx":6ED5
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1440
            Width           =   2055
         End
         Begin VB.ComboBox cmbxTpEmis 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":6EF8
            Left            =   3360
            List            =   "frmNFE_Detalhe.frx":6F14
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1440
            Width           =   5775
         End
         Begin VB.ComboBox cmbxTpImp 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":7106
            Left            =   120
            List            =   "frmNFE_Detalhe.frx":711C
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1440
            Width           =   3135
         End
         Begin VB.ComboBox cmbxIdDest 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":71C2
            Left            =   11400
            List            =   "frmNFE_Detalhe.frx":71CF
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   840
            Width           =   2055
         End
         Begin VB.ComboBox cmbxTPNf 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":7222
            Left            =   9240
            List            =   "frmNFE_Detalhe.frx":722C
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtDHSaiEnt 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6840
            TabIndex        =   24
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtDHEmi 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   22
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtChave 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   4215
         End
         Begin VB.TextBox txtcNF 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   12000
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtSerie 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11400
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cmbxModelo 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":7248
            Left            =   9240
            List            =   "frmNFE_Detalhe.frx":7252
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox cmbxIndPag 
            Height          =   315
            ItemData        =   "frmNFE_Detalhe.frx":727E
            Left            =   6840
            List            =   "frmNFE_Detalhe.frx":728B
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtNatOp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            TabIndex        =   9
            Top             =   240
            Width           =   4095
         End
         Begin VB.TextBox txtnNF 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtUF 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Venda:"
            Height          =   195
            Index           =   32
            Left            =   12480
            TabIndex        =   45
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Versão Emiss."
            Height          =   195
            Index           =   21
            Left            =   11400
            TabIndex        =   43
            Top             =   1800
            Width           =   990
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Processo da Emissão"
            Height          =   195
            Index           =   20
            Left            =   6000
            TabIndex        =   41
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Presença:"
            Height          =   195
            Index           =   19
            Left            =   2160
            TabIndex        =   39
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   37
            Top             =   1800
            Width           =   705
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Finalidade"
            Height          =   195
            Index           =   17
            Left            =   11400
            TabIndex        =   35
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DV"
            Height          =   195
            Index           =   8
            Left            =   13080
            TabIndex        =   18
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ambiente"
            Height          =   195
            Index           =   16
            Left            =   9240
            TabIndex        =   33
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Emissão:"
            Height          =   195
            Index           =   15
            Left            =   3360
            TabIndex        =   31
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Impressão"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   1080
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Identif. Destino"
            Height          =   195
            Index           =   13
            Left            =   11400
            TabIndex        =   27
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Operação:"
            Height          =   195
            Index           =   12
            Left            =   9240
            TabIndex        =   25
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Movimento (E/S)"
            Height          =   195
            Index           =   11
            Left            =   6840
            TabIndex        =   23
            Top             =   600
            Width           =   1590
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Emissão:"
            Height          =   195
            Index           =   10
            Left            =   4440
            TabIndex        =   21
            Top             =   600
            Width           =   630
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Chave de Acesso:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Index           =   7
            Left            =   12000
            TabIndex        =   16
            Top             =   0
            Width           =   540
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Série"
            Height          =   195
            Index           =   6
            Left            =   11400
            TabIndex        =   14
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo:"
            Height          =   195
            Index           =   5
            Left            =   9240
            TabIndex        =   12
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lblID 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ind. Pag:"
            Height          =   195
            Index           =   4
            Left            =   6840
            TabIndex        =   10
            Top             =   0
            Width           =   645
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Natureza da Operação"
            Height          =   195
            Index           =   3
            Left            =   2640
            TabIndex        =   8
            Top             =   0
            Width           =   1620
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
            Height          =   195
            Index           =   2
            Left            =   1560
            TabIndex        =   6
            Top             =   0
            Width           =   555
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF:"
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   4
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblDados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "ID:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   210
         End
      End
      Begin MSDataGridLib.DataGrid gridICMS 
         Height          =   975
         Left            =   3240
         TabIndex        =   168
         Top             =   5880
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1720
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         Caption         =   "ICMS"
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
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridIPI 
         Height          =   975
         Left            =   13200
         TabIndex        =   169
         Top             =   5880
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1720
         _Version        =   393216
         Appearance      =   0
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
         Caption         =   "IPI"
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
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridPIS 
         Height          =   975
         Left            =   10560
         TabIndex        =   170
         Top             =   5880
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1720
         _Version        =   393216
         Appearance      =   0
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
         Caption         =   "PIS"
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
               ColumnWidth     =   494,929
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridCofins 
         Height          =   975
         Left            =   7440
         TabIndex        =   171
         Top             =   5880
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1720
         _Version        =   393216
         Appearance      =   0
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
         Caption         =   "Cofins"
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
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridPagamentos 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   275
         Top             =   720
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         Caption         =   "Pagamentos"
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pcInfAdc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   -74880
         ScaleHeight     =   1425
         ScaleWidth      =   11865
         TabIndex        =   291
         Top             =   4320
         Width           =   11895
         Begin VB.TextBox txtInfCpl 
            Appearance      =   0  'Flat
            Height          =   1125
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   293
            Top             =   240
            Width           =   11655
         End
         Begin VB.Label lblInfAdFisco 
            AutoSize        =   -1  'True
            Caption         =   "Inf. Adc. Fisco"
            Height          =   195
            Left            =   120
            TabIndex        =   294
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblDados 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações Adicionais"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   33
            Left            =   120
            TabIndex        =   292
            Top             =   0
            Width           =   1950
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soma dos produtos:"
         Height          =   195
         Left            =   13200
         TabIndex        =   321
         Top             =   5640
         Width           =   1410
      End
      Begin VB.Label lblTotalProd 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   14880
         TabIndex        =   320
         Top             =   5640
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da NFe:"
         Height          =   195
         Left            =   -67200
         TabIndex        =   319
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblVNfe 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   -66120
         TabIndex        =   318
         Top             =   2880
         Width           =   390
      End
      Begin VB.Label lblTotalPag 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   -62640
         TabIndex        =   316
         Top             =   2880
         Width           =   390
      End
      Begin VB.Label lblTotalPagCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de pagamentos:"
         Height          =   195
         Left            =   -64320
         TabIndex        =   315
         Top             =   2880
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmNFE_Detalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Public bNotaFechada As Boolean
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public WithEvents adoPIS As ADODB.Recordset
Attribute adoPIS.VB_VarHelpID = -1
Public WithEvents adoIPI As ADODB.Recordset
Attribute adoIPI.VB_VarHelpID = -1
Public WithEvents adoCOFINS As ADODB.Recordset
Attribute adoCOFINS.VB_VarHelpID = -1
Public WithEvents adoICMS As ADODB.Recordset
Attribute adoICMS.VB_VarHelpID = -1
Public WithEvents adoPag As ADODB.Recordset
Attribute adoPag.VB_VarHelpID = -1
Dim editando As Boolean

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    CarregaImpostos
End Sub

Private Sub ado_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'On Error Resume Next
    
    
    
End Sub

Private Sub cmbxModFrete_Click()
    If Left(cmbxModFrete.Text, 1) = "9" Then
        pcTransportadora.Visible = False
    Else
        pcTransportadora.Visible = True
            
        pcTransportadora.Enabled = True
        cmdSaveTransp.Enabled = True
    End If
End Sub

Private Sub cmbxTpEmis_Click()
    Dim strSQL As String
    If bActive = False Then Exit Sub
     
    
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    alteraChave
    
    
    
End Sub
Sub alteraChave()
    Dim strSQL As String
    strSQL = GeraChaveNfE(Left(cmbxModelo.Text, 2), txtSerie.Text, txtnNF.Text, txtcNF.Text, Left(cmbxTpEmis.Text, 1))
    If strSQL = txtChave.Text Then
        Exit Sub
    End If
    
    If MsgBox("Modificar o tipo de emissão modificará também a chave de acesso. Confirma modificação?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
     NFe.AlteraChaveNfe Me.tag, strSQL
    MsgBox "Dados alterados!", vbInformation
End Sub
Private Sub cmdAbrirXML_Click()
    'On Error Resume Next
    
    Shell "explorer " & mudarCaractere(txtCaminhoXML.Text, "/", "\")
End Sub

Private Sub cmdAddNfeRef_Click()

End Sub

Private Sub cmdAddPag_Click()
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    NFe.GravaNFe_Pag Me.tag, Left(cmbxTpag.Text, 2), txtValorPag.Text, cmbxTBand.tag, txtCNPJPag.Text, Left(cmbxTBand.Text, 2), _
    txtCAut.Text, txtTpIntegra.Text, "0", cmdAddPag.tag
    CarregaPagamentos
    cmdCancelarPag_Click
End Sub

Private Sub cmdAltCPAg_Click()
    On Error GoTo trata_Erro:
    If Not editando Then
        MsgBox "A NFe não está em modo de edição. Clique no botão editar!", vbExclamation
        
        Exit Sub
    End If
    With adoPag
        P_LocalizaOperacao TiraNull(!tPag), cmbxTpag
        txtValorPag.Text = FormatVal(!vPag)
        txtCNPJPag.Text = TiraNull(!CNPJ)
        P_LocalizaOperacao TiraNull(!tBand), cmbxTBand
        
        txtCAut.Text = TiraNull(!cAut)
        txtTpIntegra.Text = TiraNull(!tpIntegra)
        cmdAddPag.caption = "Alterar"
        cmdAddPag.tag = !id
        cmdCancelar.Visible = True
    End With
    
    Exit Sub
trata_Erro:
    
    
End Sub

Private Sub cmdAlterarCFOP_Click()
    Dim strSQL As String
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    If txtCfop.Text = "" Then
        MsgBox "Informe o CFOP que deseja alterar!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja alterar o CFOP de todos os produtos desta nota para " & txtCfop.Text & "?", vbYesNo + vbQuestion) = vbYes Then
    
        strSQL = "update nfe_item set cfop='" & txtCfop.Text & "' "
        strSQL = strSQL & " where id_nfe=" & Me.tag
        Conn.Execute strSQL
        
        CarregaProdutos
    
    End If
    
    
    
End Sub

Private Sub cmdAlterarProd_Click()
    frmNFe_Item.tag = ado!id
    frmNFe_Item.id_nfe = Me.tag
    frmNFe_Item.caption = "NFe id: " & Me.tag & " Destalhe do produto " & ado!id
    frmNFe_Item.Show 1
    ado.Requery
    Unload frmNFe_Item
    
End Sub

Private Sub cmdBuscaTransp_Click()
    buscaTransportadora
End Sub
Sub buscaTransportadora()
    Dim f As SpedParticipante.participante
    Dim c As Cidades.Cidade
    f = SpedParticipante.F_BuscaParticipante(txtCNPJTransp.Text)
    If f.codigo <> "" Then
        c = F_BuscaCidadeCodigo(f.cod_mun)
        txtCNPJTransp.Text = f.CNPJ
        txtCNPJTransp.tag = f.codigo
        txtNomeTransp.Text = f.nome
        txtIETransp.Text = f.IE
        txtCPFTransp.Text = f.cpf
        
        txtXEnderTransp.Text = f.Endereco
        txtXMunTransp.Text = c.nome_cid
        txtUFTransp.Text = c.uf_cid
        txtXMunTransp.tag = c.cod_cid
        
    End If
    
    
End Sub


Private Sub cmdCancelarPag_Click()
    cmdAddPag.tag = ""
    cmdAddPag.caption = "Adicionar"
    cmdCancelar.Visible = False
End Sub

Private Sub cmdCidade_Click()
    buscaCidade
End Sub
Sub buscaCidade()
    Dim c As Cidade
    
    c = Cidades.F_BuscaCidades(txtNomeMunDest.Text)
    If c.cod_cid <> "" Then
        txtNomeMunDest.Text = c.nome_cid
        txtCodMunDest.Text = c.cod_cid
        txtUFDest.Text = c.uf_cid
        
        
    End If
End Sub

Private Sub cmdCliente_Click()
    frmEdiClientes.txtbxExpr.Text = txtXNomeDest.Text
    frmEdiClientes.Show 1
End Sub

Sub HabilitaFrame(obj As Object, opcao As Boolean)
    If opcao = True Then
        obj.Enabled = True
        obj.BackColor = vbWhite
    Else
        obj.Enabled = False
        obj.BackColor = &HE0E0E0
    End If
    
    
End Sub



Private Sub cmdCupom_Click()
    Dim s As String
     s = RetornoConteudoNFeXML(txtChave, Format(Left(txtDHEmi.Text, 10), "YYYYMM"))
        p = ApontarPrinterToPDF(txtChave.Text & ".pdf")
        i = ImprimirXMLNFeDANFE(s, Printer)
        Printer.EndDoc
        Exit Sub
End Sub

Private Sub cmdEAN_Click()
     Dim strSQL As String
    Dim iLen As Integer
    With ado
        .MoveFirst
        Do While Not .EOF
            iLen = Len(TiraNull(!cEAN))
            If iLen = 8 Or iLen = 13 Or iLen = 14 Then
                strSQL = base_10(Left(!cEAN, 12))
                If Left(strSQL, 13) <> !cEAN Then
                    MsgBox "Erro de dígito verificador: " & strSQL
                End If
            End If
            
            iLen = Len(TiraNull(!cEANTrib))
            If iLen = 8 Or iLen = 13 Or iLen = 14 Then
                strSQL = base_10(Left(!cEANTrib, 12))
                If Left(strSQL, 13) <> !cEANTrib Then
                    MsgBox "Erro de dígito verificador: " & strSQL
                End If
            End If
            
            .MoveNext
        Loop
        .MoveFirst
    End With
    
End Sub

Private Sub cmdEdit_Click()
    If sstb.Tab = 0 Then
        HabilitaFrame pcDadosNFe, True
        HabilitaFrame pcCont, True
        HabilitaFrame pcInfAdc, True
        HabilitaFrame pcNfeRef, True
        cmdSaveDados.Enabled = True
        
        txtNatOp.SetFocus
        
    ElseIf sstb.Tab = 1 Then
    ElseIf sstb.Tab = 2 Then
        pcDestCab.BackColor = vbWhite
        pcDestEnd.BackColor = vbWhite
        pcDestCab.Enabled = True
        pcDestEnd.Enabled = True
        cmdSaveDest.Enabled = True
        txtCNPJDest.SetFocus
    ElseIf sstb.Tab = 3 Then
    ElseIf sstb.Tab = 4 Then
        cmdSalvarTotais.Enabled = True
    ElseIf sstb.Tab = 5 Then
        cmdSaveTransp.Enabled = True
    ElseIf sstb.Tab = 6 Then
        pcPag.Visible = True
    End If
    If buscaStatusNFe(Me.tag) = "1" Then
        editando = True
    End If
    
End Sub


Private Sub cmdEditProd_Click()
    If strAltprod_usu = "S" Then
        frmCadPro.tag = ado!cProd
        frmCadPro.Show 1
    End If
End Sub

Private Sub cmdExcluirPagamento_Click()
    Dim strSQL As String
    
    If Not editando Then Exit Sub
    
    
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma exclusão do pagamento?", vbYesNo + vbQuestion) = vbNo Then
        
        Exit Sub
        
        
    End If
    strSQL = "delete from nfe_pag where id=" & adoPag!id
    Conn.Execute strSQL
    CarregaPagamentos
    
        
    
End Sub

Private Sub cmdFinanceira_Click()
    Dim f As financeira.financeira
    f = financeira.F_BuscaFinanceira(txtCNPJPag.Text)
    If f.cod_fin <> "" Then
        txtCNPJPag.Text = f.CNPJ
        lblOperadora.caption = f.cod_fin & "-" & f.nome_fin
        lblOperadora.caption = f.cod_fin
        
    End If
End Sub

Private Sub cmdMunGerador_Click()
    buscaMunGerador
End Sub
Sub buscaMunGerador()
    Dim c As Cidade
    c = Cidades.F_BuscaCidades(txtMuncipioGeradorTransp.Text)
    If c.cod_cid <> "" Then
        txtCMunFGTransp.Text = c.cod_cid
        txtCMunFGTransp.tag = c.cod_cid
        txtMuncipioGeradorTransp.Text = c.nome_cid
        txtMuncipioGeradorTransp.tag = c.cod_cid
        
    End If
End Sub
Private Sub cmdNCM_Click()
    Dim ncm As String


 Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    With ado
        .MoveFirst
        Do While Not .EOF
            ncm = F_BuscaUmCampo("ncm_pro", "tab_produtos", "cod_Pro='" & ado!cProd & "'")
            If Len(ncm) = 8 Then
                strSQL = "update nfe_item set ncm='" & ncm & "' where id=" & ado!id & ""
                Conn.Execute strSQL
            End If
            
            
            .MoveNext
        Loop
        ado.Requery
    End With
    
End Sub

Private Sub cmdPesqCli_Click()
    
    buscaCli
End Sub
Sub buscaCli()
    Dim c As cliente
    
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    c = F_BuscaCliente(txtXNomeDest.Text)
    If c.cod_cli <> "" Then
        txtXNomeDest.Text = c.nome_Cli
        txtBairroDest.Text = c.bairro_cli
        txtNomeMunDest.Text = c.c.nome_cid
        txtNomeMunDest.tag = c.c.cod_mun
        
        txtLogrDest.Text = c.Logradouro
        txtNumDest.Text = c.numero
        txtComplDest.Text = c.complemento
        txtCodMunDest.Text = c.c.cod_cid
        txtCEPDest.Text = c.cep_cli
        txtUFDest.Text = c.c.uf_cid
        txtIEDest.Text = c.ie_Cli
        txtCNPJDest.Text = c.cnpj_Cli
        txtCPFDest.Text = c.cpf_cli
        txtIEDest.Text = c.ie_Cli
        
        If buscaStatusNFe(Me.tag) <> "1" Then
            MsgBox "Esta NFe não pode ser alterada!", vbExclamation
            Exit Sub
        Else
            cmdSaveDest_Click
        End If
        
    End If
End Sub

Private Sub cmdProprioEmitente_Click()
    If editando = True Then
        txtCNPJDest.Text = txtCNPJEmit.Text
        txtXNomeDest.Text = txtxNomeEmit.Text
        txtLogrDest.Text = txtXLogrEmit.Text
        txtNumDest.Text = txtNumEmit.Text
        txtNomeMunDest.Text = txtNomeMunEmit.Text
        txtNomeMunDest.tag = txtNomeMunEmit.tag
        txtCodMunDest.Text = txtCodMunEmit.Text
        txtComplDest.Text = txtComplEmit.Text
        txtBairroDest.Text = txtBairroEmit.Text
        txtCEPDest.Text = txtcep.Text
        txtUFDest.Text = txtUFEmit.Text
        txtIEDest.Text = txtIEEmit.Text
        txtFoneDest.Text = txtFoneEmit.Text
        cmbxCRTEmit.ListIndex = 0
        
    End If
    
End Sub

Private Sub cmdRetransm_Click()
    On Error GoTo trata_Erro
    If NFe.buscaStatusNFe(Me.tag) = "1" Then
        TransmitirNFcE Me.tag, ""
        If buscaStatusNFe(Me.tag) = "2" Then
            MsgBox "Nota Fiscal transmitida e autorizada!", vbInformation
            
        End If
        
    Else
        MsgBox "O status dessa NFe não permite retransmissão!", vbExclamation
    End If
    Exit Sub
trata_Erro:
P_MostraErro
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub cmdSalvarTotais_Click()
    If Not NFe.GravaNfe_IcmsTot(Me.tag, txtvBC.Text, txtVICms.Text, txtvICMSDeson.Text, txtvBCST.Text, _
         txtVST.Text, txtVProd.Text, txtVFrete.Text, txtVSeg.Text, txtVDesc.Text, txtVII.Text, _
         txtVIPI.Text, txtVPIS.Text, txtVCofins.Text, txtVOutro.Text, txtVNF.Text, txtVTotTrib.Text) Then
        MsgBox "Erro na gravação dos totais de impostos", vbExclamation
    Else
        MsgBox "Dados gravados!", vbInformation
    End If
End Sub

Private Sub cmdSaveDados_Click()
    
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    
    If NFe.NFE_AlteraDadosNfe(Me.tag, txtUF.Text, txtNatOp.Text, Left(cmbxIndPag.Text, 1), Left(cmbxModelo, 2), txtDHEmi.Text, txtDHSaiEnt.Text, _
       Left(cmbxTPNf.Text, 1), Left(cmbxIdDest.Text, 1), Left(cmbxTpImp.Text, 1), Left(cmbxTpEmis.Text, 1), Left(cmbxTpAmb.Text, 1), _
       Left(cmbxFinNfe.Text, 1), Left(cmbxIndFinal.Text, 1), Left(cmbxIndPres.Text, 1), Left(cmbxProcEmi.Text, 1), _
       txtVerProc.Text, txtDHCont.Text, txtJust.Text) Then
       
       NFe.AlteraInfAdicNFe Me.tag, txtInfCpl.Text
        
        GravaNFe_Ref Me.tag, txtNfeRef
        
        
        
        MsgBox "Dados alterados", vbInformation
        
        
        
        
        Exit Sub
    End If
    
End Sub

Private Sub cmdSaveDest_Click()
    Dim status As String
    
    
    
    status = F_BuscaUmCampo("status", "nfe", "id=" & Me.tag)
    If status <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    If NFe.AlteraNFe_Dest(Me.tag, txtCNPJDest.Text, txtCPFDest.Text, txtIDEstrangeiro.Text, txtXNomeDest.Text) Then
        If NFe.AlteraNfe_enderdest(Me.tag, txtLogrDest.Text, txtNumDest.Text, txtComplDest.Text, txtBairroDest.Text, txtCodMunDest.Text, _
         txtNomeMunDest.Text, txtUFDest.Text, txtCEPDest.Text, txtCOdPaisDest.Text, txtPaisDest.Text, txtFoneDest.Text, _
         txtIEDest.Text, Left(cmbxIndIEDest.Text, 1), txtISUF.Text, txtIMDest.Text, txtEmailDest.Text) Then
         
            MsgBox "Dados do Destinatário Alterados!", vbInformation
        End If
    End If
    
    Exit Sub
End Sub

Private Sub cmdSaveTransp_Click()
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    P_Status "salvando dados..."
    
    If Not NFe.GravaNFe_Transp(Me.tag, Left(cmbxModFrete.Text, 1)) Then
    
        MsgBox "Erro na gravação NfeTransp", vbExclamation
    End If
    
    NFe.GravaNFe_Transp_Transportadora Me.tag, txtCNPJTransp.Text, txtCPFTransp.Text, txtNomeTransp.Text, txtIETransp.Text, txtXEnderTransp.Text, txtXMunTransp.Text, txtUFTransp.Text
    NFe.GravaNFe_Transp_Reboque Me.tag, txtPlacaReboque.Text, txtUFReboque.Text, txtRNTCReboque.Text, txtVagao.Text, txtBalsa.Text
    NFe.GravaNFe_Transp_RetTransp Me.tag, txtVServTransp.Text, txtvBCRetTransp.Text, txtPIcmsRetTransp.Text, txtVIcmsRetTransp.Text, txtCFOPTransp.Text, txtCMunFGTransp.Text
    NFe.GravaNFe_Transp_VeicTransp Me.tag, txtPlacaReboque.Text, txtUfVeicTransp.Text, txtRNTCVeicTransp.Text
    P_Status
End Sub



Private Sub cmdSemGtin_Click()
    Dim strSQL As String
    strSQL = "update nfe_item set cean='SEM GTIN',ceantrib='SEM GTIN' "
        strSQL = strSQL & " where id_nfe=" & Me.tag & " and (cean='' or cean is null)"
        Conn.Execute strSQL
        CarregaProdutos
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub



Private Sub Form_Load()
    bActive = False
    Set ado = New ADODB.Recordset
    Set adoPIS = New ADODB.Recordset
    Set adoIPI = New ADODB.Recordset
    Set adoCOFINS = New ADODB.Recordset
    Set adoICMS = New ADODB.Recordset
    Set adoPag = New ADODB.Recordset
    sstb.Tab = 0
End Sub
Sub carrega()
    CarregaNfe
    CarregaEmit
    CarregaDest
    CarregaProdutos
    CarregaTot
    CarregaTransp
    CarregaPagamentos
    CarregaInfAdic
    CarregaTransp
End Sub

Sub CarregaNfe()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from nfe"
    
    strSQL = strSQL & " where nfe.id=" & Me.tag
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Me.caption = "NFe " & !id & " - " & !Chave
        lblId.caption = !id
        txtUF.Text = !cUF
        txtcNF.Text = !cNF
        txtnNF.Text = !nNF
        txtNatOp.Text = !natOp
        P_LocalizaOperacao !indPag, cmbxIndPag
        P_LocalizaOperacao !modelo, cmbxModelo
        txtSerie.Text = !serie
        txtDV.Text = !cDV
        txtChave.Text = !Chave
        txtDHEmi.Text = !dhEmi
        txtNum_Ven.Text = !num_ven
        txtDHSaiEnt.Text = !dhSaiEnt
        P_LocalizaOperacao !tpNf, cmbxTPNf
        P_LocalizaOperacao !idDest, cmbxIdDest
        P_LocalizaOperacao !tpImp, cmbxTpImp
        P_LocalizaOperacao !tpEmis, cmbxTpEmis
        P_LocalizaOperacao !tpAmb, cmbxTpAmb
        P_LocalizaOperacao !finNFe, cmbxFinNfe
        P_LocalizaOperacao !indFinal, cmbxIndFinal
        P_LocalizaOperacao !indPres, cmbxIndPres
        P_LocalizaOperacao !procEmi, cmbxProcEmi
        
        txtVerProc.Text = !verProc
        txtDHCont.Text = !dhCont
        txtJust.Text = !xJust
        
        txtNfeRef.Text = BuscaNFe_Ref(!id)
             
        P_LocalizaOperacao Format(!status), cmbStatus
        txtVenda.Text = Format(!num_ven, "000000#")
        txtCaminhoXML.Text = Format(!caminhoXML)
        txtProtocolo.Text = Format(!protocolo)
        txtRecibo.Text = Format(!recibo)
        txtUsuario.Text = Format(!cod_usu)
        txtDigestValue.Text = Format(!DigestValue)
        
        If !status = "2" Then
            cmdCupom.Enabled = True
        End If
        
        
    End With
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
    
    
End Sub

Sub CarregaEmit()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from nfe_emit inner join nfe_enderemit on nfe_emit.id_nfe=nfe_enderemit.id_nfe "
    
    strSQL = strSQL & " where nfe_emit.id_nfe=" & Me.tag
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        txtxNomeEmit.Text = !xNome
        txtCNPJEmit.Text = !CNPJ
        txtXFantEmit.Text = Format(!xFant)
        txtXLogrEmit.Text = Format(!xLgr)
        txtNumEmit.Text = Format(!nro)
        txtComplEmit.Text = Format(!xCpl)
        txtBairroEmit.Text = Format(!xBairro)
        txtCodMunEmit.Text = !cMun
        txtNomeMunEmit.Text = !xMun
        txtUFEmit.Text = !UF
        txtcep.Text = !cep
        txtCodPaisEmit.Text = !cPais
        txtNomePaisEmit.Text = !xPais
        txtFoneEmit.Text = !fone
        txtIEEmit.Text = !IE
        txtIESTEmit.Text = Format(!iEst)
        txtIMEmit.Text = Format(!IM)
        txtCNAEEmit.Text = Format(!CNAE)
        P_LocalizaOperacao !CRT, cmbxCRTEmit
        
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
    
    
End Sub



Sub CarregaDest()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from nfe_dest left join nfe_enderdest on nfe_dest.id_nfe=nfe_enderdest.id_nfe "
    
    strSQL = strSQL & " where nfe_dest.id_nfe=" & Me.tag
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        txtXNomeDest.Text = !xNome
        txtCNPJDest.Text = !CNPJ
        txtCPFDest.Text = !cpf
        txtIDEstrangeiro.Text = Format(!idEstrangeiro)
        txtLogrDest.Text = Format(!xLgr)
        txtNumDest.Text = Format(!nro)
        txtComplDest.Text = Format(!xCpl)
        txtBairroDest.Text = Format(!xBairro)
        txtCodMunDest.Text = !cMun
        txtNomeMunDest.Text = !xMun
        txtUFDest.Text = !UF
        txtCEPDest.Text = !cep
        txtCOdPaisDest.Text = !cPais
        txtPaisDest.Text = !xPais
        txtFoneDest.Text = !fone
        txtIEDest.Text = !IE
        
        txtIMDest.Text = Format(!IM)
        
        
        P_LocalizaOperacao !IndIEDest, cmbxIndIEDest
        
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
    
    
End Sub
Sub CarregaTot()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from nfe_icmstot "
    
    strSQL = strSQL & " where id_nfe=" & Me.tag
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        txtvBC.Text = FormatVal(!vBC)
        txtVICms.Text = FormatVal(!vICMS)
        txtvICMSDeson.Text = FormatVal(!vICMSDeson)
        txtvBCST.Text = FormatVal(!vBCST)
        txtVST.Text = FormatVal(!vST)
        txtVProd.Text = FormatVal(!vProd)
        txtVFrete.Text = FormatVal(!vFrete)
        txtVSeg.Text = FormatVal(!vSeg)
        txtVDesc.Text = FormatVal(!vDesc)
        txtVII.Text = FormatVal(!vII)
        txtVIPI.Text = FormatVal(!vIPI)
        txtVPIS.Text = FormatVal(!vPIS)
        txtVCofins.Text = FormatVal(!vCOFINS)
        txtVOutro.Text = FormatVal(!vOutro)
        txtVNF.Text = FormatVal(!vNF)
        txtVTotTrib.Text = FormatVal(!vTotTrib)
        
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
    
    
End Sub
Sub CarregaTransp()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from nfe_transp as t left join nfe_transp_reboque as r on t.id_nfe=r.id_nfe "
    strSQL = strSQL & " left join nfe_transp_transportadora as tt on tt.id_nfe=r.id_nfe  "
    strSQL = strSQL & " left join nfe_transp_veictransp as tvt on tvt.id_nfe=r.id_nfe  "
    strSQL = strSQL & " left join nfe_transp_vol as v on v.id_nfe=r.id_nfe  "
    
    
    strSQL = strSQL & " where t.id_nfe=" & Me.tag
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        P_LocalizaOperacao !modFrete, cmbxModFrete
        txtCNPJTransp.Text = TiraNull(!CNPJ)
        'txtCNPJTransp.tag =  f_Bu !codigo
        txtNomeTransp.Text = TiraNull(!xNome)
        txtIETransp.Text = TiraNull(!IE)
        txtCPFTransp.Text = TiraNull(!cpf)
        
        txtXEnderTransp.Text = TiraNull(!xEnder)
        txtXMunTransp.Text = TiraNull(!xMun)
        txtUFTransp.Text = TiraNull(!UF)
                
        
    End With
    
    strSQL = "select * fr"
    
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
    
    
End Sub
Sub CarregaPagamentos()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from nfe_pag"
    
    strSQL = strSQL & " where id_nfe=" & Me.tag
    
    
    abreObjetoAdo adoPag, strSQL, gridPagamentos
    
    Dim c As Currency
    With adoPag
         Do While Not .EOF
            c = c + !vPag
            .MoveNext
         Loop
    End With
    lblVNfe.caption = txtVNF.Text
    lblTotalPag.caption = FormatVal(c)
    If adoPag.RecordCount > 0 Then
        adoPag.MoveFirst
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
    
    
    
End Sub
Sub CarregaInfAdic()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from nfe_infAdic"
    
    strSQL = strSQL & " where id_nfe=" & Me.tag
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            txtInfCpl.Text = !infCpl
            lblInfAdFisco.caption = "Retorno Fisco: " & !infAdFisco
        
        End If
        .Close
    End With
    
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
    
    
End Sub

Sub CarregaProdutos()
    Dim strSQL As String
    P_Status "carregando produtos"
    strSQL = "select i.*,ii.*,comb.cProdAnp,comb.DescAnp,comb.UFCons from nfe_item as i "
    strSQL = strSQL & " left join nfe_item_imposto as ii on i.id=ii.id  "
    strSQL = strSQL & " left join nfe_item_comb as comb on i.id=comb.id_item"
    
    strSQL = strSQL & " where i.id_nfe=" & Me.tag
    
    abreObjetoAdo ado, strSQL, gridProdutos
    Dim c As Currency
    With ado
        gridProdutos.Visible = False
         Do While Not .EOF
            c = c + !vProd
            .MoveNext
         Loop
         lblTotalProd.caption = FormatVal(c)
        If .RecordCount > 0 Then .MoveFirst
    End With
    gridProdutos.Visible = True
    P_Status
End Sub
Sub CarregaImpostos()
    
    
    Dim strSQL As String
    Dim NomeTabelaICMS As String
    Dim compl As String
    'If ado!Tipo_Icms = "500" Then
   '
   ' End If
    On Error GoTo trata_Erro
    If Left(UCase(ado!tipo_icms), 4) <> "ICMS" Then
        compl = "icms"
    End If
    
    strSQL = "select * from nfe_item_imposto_" & compl & ado!tipo_icms
    strSQL = strSQL & " where id=" & ado!id
    
    abreObjetoAdo adoICMS, strSQL, gridICMS
    
    strSQL = "select * from nfe_item_imposto_" & ado!tipo_ipi
    strSQL = strSQL & " where id=" & ado!id
    abreObjetoAdo adoIPI, strSQL, gridIPI
    
    strSQL = "select * from nfe_item_imposto_" & ado!tipo_pis
    strSQL = strSQL & " where id=" & ado!id
    abreObjetoAdo adoPIS, strSQL, gridPIS
    
    strSQL = "select * from nfe_item_imposto_" & ado!tipo_cofins
    strSQL = strSQL & " where id=" & ado!id
    abreObjetoAdo adoCOFINS, strSQL, gridCofins
    
    
trata_Erro:
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGradeWidth sstb, Me, 400
    PosicionaGradeHeigth sstb, pcBot, 200
    posicionaADireita cmdSair, pcBot
    If bActive = True Then
    
        PosicionaGradeWidth gridProdutos, sstb, 200
    End If
End Sub

Private Sub gridPagamentos_DblClick()
    cmdAltCPAg_Click
End Sub

Private Sub gridPagamentos_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        excluiPag
    End If

End Sub
Sub excluiPag()
    On Error GoTo trata_Erro
    excluir_nfe_pag (adoPag!id)
trata_Erro:
    
End Sub
Private Sub gridProdutos_DblClick()
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    cmdAlterarProd_Click
End Sub

Private Sub gridProdutos_HeadClick(ByVal ColIndex As Integer)
    OrdenaGrid gridProdutos, ado
End Sub

Private Sub sstb_Click(PreviousTab As Integer)
    If sstb.Tab = 1 Then
        cmdEdit.Enabled = False
    Else
        cmdEdit.Enabled = True
    End If
End Sub

Private Sub txtCFOP_GotFocus()
    P_SelecTexto txtCfop
End Sub

Private Sub txtCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCFOP
    End If
    
End Sub
Sub buscaCFOP()
    Dim p As Padrao1
    p = Sped_Lista_CFOP(txtCfop.Text)
    If p.codigo <> "" Then
        txtCfop.Text = p.codigo
        txtCfop.tag = p.codigo
        
    End If
End Sub
Private Sub txtCNPJDest_LostFocus()
    If txtCNPJDest.Text = "" Then Exit Sub
    If Not F_CGC(txtCNPJDest.Text) Then
        MsgBox "CNPJ INválido!", vbExclamation
    End If
End Sub

Private Sub txtCodMunDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCidade
    End If
End Sub

Private Sub txtDHCont_GotFocus()
    If txtDHCont.Text = "" Then
        txtDHCont.Text = FormataDataNfe(Now)
    End If
End Sub

Private Sub txtnNF_KeyDown(KeyCode As Integer, Shift As Integer)
    If buscaStatusNFe(Me.tag) <> "1" Then
        MsgBox "Esta NFe não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    
    alteraChave
End Sub
