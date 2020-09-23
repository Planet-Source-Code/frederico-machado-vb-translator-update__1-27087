VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Translator Lite"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Add Words"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.OptionButton optEng 
      Caption         =   "English/Portuguese"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Portuguese/English"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2115
      Width           =   3735
   End
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "Translate"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1665
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "frmMain.frx":058A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":0B14
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":0F56
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Frederico Machado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   13
      Top             =   720
      Width           =   1995
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lite"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Translator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   585
      Left            =   2880
      TabIndex        =   11
      Top             =   120
      Width           =   2160
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Translated Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   9
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Original Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   -120
      Top             =   -120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdTranslate_Click()
  If Text1 = "" Then Exit Sub
  If optPort.Value = True Then ' If optPort is selected then it translate from Portuguese to English
    dictionary.FindFirst "portuguese='" & Text1 & "'" ' Verify if the word exists
    If dictionary.NoMatch Then ' If it doesn't exists then put text1 into text2
      Text2 = Text1
    Else
      Text2 = dictionary!english ' If it exists then put the value into text2
    End If
  ElseIf optEng.Value = True Then ' If optEng is selected then it translate from English to Portuguese
    dictionary.FindFirst "english='" & Text1 & "'" ' Verify if the word exists
    If dictionary.NoMatch Then ' If it doesn't exists then put text1 into text2
      Text2 = Text1
    Else
      Text2 = dictionary!portuguese ' If it exists then put the value into text2
    End If
  End If
End Sub

Private Sub Command1_Click()
  frmHelp.Show 1
End Sub

Private Sub Command2_Click()
  frmAdd.Show 1
End Sub

Private Sub Form_Load()
  mypath = App.Path ' Sets the variable to app path
  If Right$(mypath, 1) <> "\" Then mypath = mypath & "\" ' If is a folder, it needs a \
  
  Set dbase = OpenDatabase(mypath & "data.mdb") ' Open the Database into the variable
  Set dictionary = dbase.OpenRecordset("select * from dictionary") ' Open the recordset into a variable
End Sub

Private Sub optEng_Click()
  Text1.SetFocus
End Sub

Private Sub optPort_Click()
  Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 32 Then KeyAscii = 0 ' It doesn't let the space key to work
End Sub
