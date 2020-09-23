VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show custom msgbox!!!!"
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Msgbox heading"
      Height          =   1215
      Left            =   3840
      TabIndex        =   7
      Top             =   1920
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Enter the msgbox heading"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select icon"
      Height          =   1455
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   3495
      Begin VB.OptionButton Option3 
         Caption         =   "Exclamation"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Information"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Critical"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text of msgbox"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Please enter what you wish to have the msgbox say."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'just to let you know incase you dont : option.value
'just means weather on not,the option button is selected
'if its eg-option1.value = true
'that means that option1 is selected

Private Sub Command1_Click()
If Option1.Value = True Then
MsgBox Text1.Text, vbCritical, Text2.Text
End If
If Option2.Value = True Then
MsgBox Text1.Text, vbInformation, Text2.Text
End If
If Option3.Value = True Then
MsgBox Text1.Text, vbExclamation, Text2.Text
End If
End Sub

Private Sub Command2_Click()
Load Form1
Form1.Show
Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "thank for looking at my code, i hope it helped you, you can use this stuff for lots of things!!!, please vote, it be less than a minute of your time, it took me hours to construct this so please vote, or give feedback, thank you.", vbExclamation, "Goodbye"
End Sub
