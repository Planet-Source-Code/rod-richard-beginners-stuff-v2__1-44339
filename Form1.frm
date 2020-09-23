VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Moving text!"
      Height          =   375
      Left            =   3480
      TabIndex        =   28
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Usrname + Pswd"
      Height          =   375
      Left            =   2040
      TabIndex        =   27
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Typerwriter"
      Height          =   375
      Left            =   960
      TabIndex        =   26
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Msgboxs"
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Progress bar! stuff"
      Height          =   3135
      Left            =   4320
      TabIndex        =   17
      Top             =   1920
      Width           =   3375
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
         Height          =   375
         Left            =   1440
         TabIndex        =   24
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "0"
         Top             =   2040
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2760
         Top             =   360
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Start"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Enter a number between 1 - 99......"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Click!!!"
      Height          =   1455
      Left            =   4440
      TabIndex        =   15
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "Click me!!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mouse over"
      Height          =   2775
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   2655
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "msgbox"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Normal"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cross"
         Height          =   375
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   12
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hour glass"
         Height          =   375
         Left            =   120
         MousePointer    =   11  'Hourglass
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "I beam"
         Height          =   375
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
      Begin VB.OptionButton Option4 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Blue"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Red"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Green"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "Add to list"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim tu As Boolean
Private Sub Command1_Click()
List1.AddItem Text1.Text
End Sub

Private Sub Command2_Click()
i = i + 1
If i = 1 Then
MsgBox "weldone! but that hurt, please stop!", vbInformation
End If
If i = 2 Then
MsgBox "Ouch!!!! Stop!!!", vbCritical
End If
If i = 3 Then
MsgBox "im warning you, u better stop! last chance!!!! i will take the button away!!"
End If
If i = 4 Then
MsgBox "you asked for it wise guy! the button goes!!", vbCritical
Command2.Enabled = False
End If
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
End Sub

Private Sub Command5_Click()
ProgressBar1.Value = ProgressBar1.Value + Val(Text2.Text)
Label6.Caption = ProgressBar1.Value & "%"
End Sub

Private Sub Command6_Click()
Load Form2
Form2.Show
Form1.Hide
End Sub

Private Sub Command7_Click()
Load Form3
Form3.Show
Form1.Hide
End Sub

Private Sub Command8_Click()
Load Form4
Form4.Show
Me.Hide

End Sub

Private Sub Command9_Click()
Load Form5
Form5.Show
Me.Hide
End Sub

Private Sub Form_Load()
Label6.Caption = 0 & "%"
i = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "thank for looking at my code, i hope it helped you, you can use this stuff for lots of things!!!, please vote, it be less than a minute of your time, it took me hours to construct this so please vote, or give feedback, thank you.", vbExclamation, "Goodbye"
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "you have moved the mouse over here!!! weldone!", vbInformation, "Good"
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Me.BackColor = vbGreen
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Me.BackColor = vbRed
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Me.BackColor = vbBlue
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Me.BackColor = &H8000000F
End If
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label6.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value > 99 Then
Timer1.Enabled = False
End If
End Sub




