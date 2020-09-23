VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter username and password :"
      Height          =   2055
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "OK!"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Password :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Username :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Enter username and password (must use caps where needed)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label Label2 
      Caption         =   "Password : 1234"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Username : Rod"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As String

Private Sub Command1_Click()
If Text1.Text & Text2.Text = i Then 'this means : if text1 and text2's text is Rod1234 when put together then...
MsgBox "correct password!", vbInformation, "correct"
Else 'if its not Rod1234
MsgBox "Incorrect password, please try again", vbCritical, "Error"
End If
End Sub

Private Sub Command2_Click()
Load Form1
Form1.Show
Me.Hide
End Sub

Private Sub Form_Load()
i = "Rod1234"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "thank for looking at my code, i hope it helped you, you can use this stuff for lots of things!!!, please vote, it be less than a minute of your time, it took me hours to construct this so please vote, or give feedback, thank you.", vbExclamation, "Goodbye"
End Sub
