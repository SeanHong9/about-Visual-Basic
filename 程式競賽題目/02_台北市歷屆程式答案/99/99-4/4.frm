VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "標楷體"
      Size            =   18
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   11895
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "排序"
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   480
      Left            =   2040
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   14
      Left            =   720
      TabIndex        =   16
      Top             =   9120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   13
      Left            =   720
      TabIndex        =   15
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   12
      Left            =   720
      TabIndex        =   14
      Top             =   7920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   11
      Left            =   720
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   10
      Left            =   720
      TabIndex        =   12
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   9
      Left            =   720
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   8
      Left            =   720
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   7
      Left            =   720
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   6
      Left            =   720
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   5
      Left            =   720
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   3
      Left            =   720
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   480
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "K="
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "N="
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim o As Boolean

Private Sub Command1_Click()
Dim k As Integer
Dim p As Boolean
Dim a As Integer
k = Val(Text1)
x = Text3
Do
  p = True
  For i = 0 To k - 2
    If Val(Text2(i)) > Val(Text2(i + 1)) Then p = False
    If Val(Text2(i)) = k Then a = i
  Next i
  If p = False Then
    Text3 = a + 1
    Text3 = k
    k = k - 1
  End If
Loop Until p = True
Text3 = x
End Sub

Private Sub Text1_Change()
For i = 0 To 14
  Text2(i) = ""
  Text2(i).Visible = flase
Next i
If Val(Text1) < 1 Or Val(Text1) > 15 Then
  Text1 = ""
  
Else
  For i = 0 To Val(Text1) - 1
    Text2(i).Visible = True
  Next i
End If
End Sub

Private Sub Text2_Change(Index As Integer)
If o = False Then
  If Val(Text2(Index)) > Val(Text1) Then Text2(Index) = ""
    For i = 0 To Val(Text1) - 1
      If Text2(Index) = Text2(i) And i <> Index Then Text2(Index) = ""
    Next i
End If
End Sub

Private Sub Text3_Change()
If Val(Text3) <= 1 Or Val(Text3) > Val(Text1) Then
  Text3 = ""
Else
  o = True
  For i = 1 To Val(Text3) \ 2
    x = Text2(i - 1): Text2(i - 1) = Text2(Val(Text3) - i): Text2(Val(Text3) - i) = x
  Next i
  o = False
End If
End Sub
