VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00808000&
   Caption         =   "Lampu Lalu Lintas"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   7020
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7560
      Top             =   5040
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6000
      Top             =   5040
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4440
      Top             =   5040
   End
   Begin VB.CommandButton stop 
      Caption         =   "Stop"
      Height          =   735
      Left            =   5400
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton start 
      Caption         =   "Start"
      Height          =   735
      Left            =   5400
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label waktu 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Shape hijau 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape kuning 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   735
   End
   Begin VB.Shape merah 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   1800
      Top             =   4440
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   1320
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub start_Click()
Timer1.Enabled = True
waktu.Caption = 10
waktu.ForeColor = vbRed
merah.BackColor = vbRed
End Sub

Private Sub stop_Click()
waktu.Caption = 0
waktu.ForeColor = vbBlack
merah.BackColor = vbWhite
kuning.BackColor = vbWhite
hijau.BackColor = vbWhite
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False

End Sub

Private Sub Timer1_Timer()
waktu.Caption = waktu.Caption - 1
merah.BackColor = vbRed
kuning.BackColor = vbWhite
hijau.BackColor = vbWhite

If waktu.Caption = "0" Then
waktu.Caption = 3
waktu.ForeColor = vbYellow
merah.BackColor = vbWhite
kuning.BackColor = vbYellow
hijau.BackColor = vbWhite
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
waktu.Caption = waktu.Caption - 1

If waktu.Caption = "0" Then
waktu.Caption = 15
waktu.ForeColor = vbGreen
merah.BackColor = vbWhite
kuning.BackColor = vbWhite
hijau.BackColor = vbGreen
Timer2.Enabled = False
Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
waktu.Caption = waktu.Caption - 1

If waktu.Caption = "0" Then
waktu.Caption = 10
waktu.ForeColor = vbRed
merah.BackColor = vbRed
kuning.BackColor = vbWhite
hijau.BackColor = vbWhite
Timer3.Enabled = False
Timer1.Enabled = True
End If
End Sub
