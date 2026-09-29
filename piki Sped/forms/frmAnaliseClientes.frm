VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAnaliseClientes 
   Caption         =   "Análise de Movimentação de Clientes"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   3495
      Left            =   5280
      ScaleHeight     =   3435
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3465
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4260
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Informe os clientes"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmAnaliseClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
