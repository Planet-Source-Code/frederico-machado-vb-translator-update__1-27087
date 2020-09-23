VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Words"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtEng 
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Portuguese:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "English:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  If txtEng = "" Or txtPort = "" Then Exit Sub
  dictionary.FindFirst "english='" & txtEng & "'" ' Verify if the word alread exist in Database
  If dictionary.NoMatch Then ' If it doesn't exist then add a new
    dictionary.AddNew
  Else
    If MsgBox("This word already exist in the Database, do you want to substitute it?", vbYesNo) = vbYes Then ' If it already exists and the user clicked on yes lets edit it
      dictionary.Edit
    Else ' If the user clicked on No then lets exit sub
      txtEng = ""
      txtPort = ""
      Exit Sub
    End If
  End If
  dictionary!english = txtEng ' Insert the English word to the DB
  dictionary!portuguese = txtPort ' Insert the Portuguese word to the DB
  dictionary.Update ' Update the DB
  txtEng = ""
  txtPort = ""
  txtEng.SetFocus
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
