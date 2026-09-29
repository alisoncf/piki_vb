VERSION 5.00
Begin VB.Form frmCorrecaoRep 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Correção de Lançamentos de Representantes"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11685
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10455
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   495
         Left            =   10440
         Picture         =   "frmCorrecaoRep.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCorrecaoRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
