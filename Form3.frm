VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   1920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Say default!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Type it!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "This is my ""typewriter"" text....... type in what you want and click Ok, then watch as your text is typed out!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As String 'i is a text
Dim a As String 'a is also a text

Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = False
i = Text1.Text ' the text of a is whatever the person types in textbox1
End Sub

Private Sub Command2_Click()
Timer2.Enabled = True 'turn on timer2 but then turn off timer1 so that they are not both on at the same time!
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
Load Form1
Form1.Show
Me.Hide
End Sub

Private Sub Form_Load()
a = "Hello, this is the typewriter text. Well, what ya think....if you think my code is anygood and helped ya...please take the time to vote, it will take over a minute, thats nothing compared to the time it took to make this!!!!!" 'this is what a is
End Sub

Private Sub Timer1_Timer()
  Label2.Caption = Mid(i, 1, Len(Label2.Caption) + 1) 'add one letter from i to the caption of label2 when each interval passes on the timer
End Sub

Private Sub Timer2_Timer()
Label2.Caption = Mid(a, 1, Len(Label2.Caption) + 1) 'add one letter from a
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "thank for looking at my code, i hope it helped you, you can use this stuff for lots of things!!!, please vote, it be less than a minute of your time, it took me hours to construct this so please vote, or give feedback, thank you.", vbExclamation, "Goodbye"
End Sub
