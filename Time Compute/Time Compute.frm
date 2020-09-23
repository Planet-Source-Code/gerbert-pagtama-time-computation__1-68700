VERSION 5.00
Begin VB.Form TimeCompute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Compute (by : Gerbert Pagtama >>> gerbert_p@yahoo.com)"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame TimeCompute 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   7
         Text            =   " "
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   6
         Text            =   " "
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Compute"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   3
         Text            =   " "
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   2160
         Y1              =   3600
         Y2              =   4800
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2160
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(8 hr. 14 min.)"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   13
         Top             =   4440
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "End Time"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   12
         Top             =   3720
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Start Time"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "9:34"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   3960
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "1:20"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   3960
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "example "
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   8
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "End Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   2
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1425
      End
   End
End
Attribute VB_Name = "TimeCompute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'a simple time compute (by : GERBERT PAGTAMA)
'e-mail add : gerbert_p@yahoo.com



Private Sub Command1_Click()
 End
End Sub

Private Sub Command2_Click()
  xVal = DateDiff("n", Text1.Text, Text2.Text)
  Text3.Text = "(" & Val(Int(xVal / 60)) & " hr. " & Val(Val(xVal) Mod 60) & " min." & ")"
End Sub
