VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Back to main"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start!"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   7935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Dim leng As Integer

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub


Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
Form1.Show
Load Form1
Me.Hide
End Sub

Private Sub Form_Load()
s = "LOOK IM MOVING!!" + Space(40)
leng = Len(s)
End Sub


Private Sub Timer1_Timer()
s = Left(Right(s, 1), leng) + Left(s, leng)
Label1.Caption = s
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "thank for looking at my code, i hope it helped you, you can use this stuff for lots of things!!!, please vote, it be less than a minute of your time, it took me hours to construct this so please vote, or give feedback, thank you.", vbExclamation, "Goodbye"
End Sub
