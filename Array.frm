VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808000&
   Caption         =   "Program Array"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7155
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Next Form 3"
      Height          =   615
      Left            =   7680
      TabIndex        =   5
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   4335
      Left            =   3480
      TabIndex        =   3
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "DATA YANG DI INPUT"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "JUMLAH DATA"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim Larik(9) As String
Dim i As Integer
Dim data As Integer
data = CInt(Text1.Text)
If data > 9 Then
    MsgBox "Jumlah data tidak boleh lebih dari 9", vbCritical
Else
If data <= 0 Then
    MsgBox "Jumlah data tidak boleh kurang dari 1", vbCritical
Else
For i = 0 To data - 1
    Prompt$ = "Input data untuk Array"
    Nilai$ = InputBox(Prompt$, "Array Dimensi Satu")
    Larik(i) = Nilai$
    List1.AddItem Larik(i), i
    Next i
End If
End If


End Sub

Private Sub Command2_Click()
Form2.Hide
Form3.Show
End Sub
