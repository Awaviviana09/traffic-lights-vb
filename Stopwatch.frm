VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Stopwatch"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11115
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Next  Form 2"
      Height          =   615
      Left            =   7680
      TabIndex        =   8
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1080
      Top             =   2280
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset"
      Height          =   585
      Left            =   8880
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   615
      Left            =   6240
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Resume"
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label detik 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label menit 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label jam 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "STOPWATCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
detik.Caption = 0
menit.Caption = 0
jam.Caption = 0
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
End Sub

Private Sub Command4_Click()
detik.Caption = 0
menit.Caption = 0
jam.Caption = 0
Timer1.Enabled = False
End Sub

Private Sub Command5_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Timer1_Timer()
detik.Caption = detik.Caption + 1
If detik.Caption = "60" Then
menit.Caption = menit.Caption + 1
detik.Caption = 0

End If


If menit.Caption = "60" Then
jam.Caption = jam.Caption + 1
menit.Caption = 0
detik.Caption = 0

End If

End Sub
