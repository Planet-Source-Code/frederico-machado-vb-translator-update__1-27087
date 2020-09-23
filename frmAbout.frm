VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VB Translator"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "fredisoft@bol.com.br"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Copyright (c) 2001 Fredisoft Corp."
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   2490
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "developed by Frederico Machado"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "VB Translator Lite"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   600
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
  Unload Me
End Sub
