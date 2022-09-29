VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form4"
   ScaleHeight     =   5625
   ScaleWidth      =   10800
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   120
      Top             =   4200
   End
   Begin VB.Timer Timer6 
      Interval        =   1000
      Left            =   120
      Top             =   3600
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   120
      Top             =   2880
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   120
      Top             =   2160
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   120
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   840
   End
   Begin VB.CommandButton Command7 
      Caption         =   "seven"
      Height          =   495
      Left            =   5280
      TabIndex        =   53
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "six"
      Height          =   615
      Left            =   5280
      TabIndex        =   52
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "five"
      Height          =   495
      Left            =   5280
      TabIndex        =   51
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "four"
      Height          =   615
      Left            =   5160
      TabIndex        =   50
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "three"
      Height          =   615
      Left            =   5160
      TabIndex        =   49
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "two"
      Height          =   495
      Left            =   5160
      TabIndex        =   48
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "one"
      Height          =   615
      Left            =   5160
      TabIndex        =   47
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   44
      Left            =   3960
      TabIndex        =   46
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   43
      Left            =   3600
      TabIndex        =   45
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   42
      Left            =   3240
      TabIndex        =   44
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   41
      Left            =   2880
      TabIndex        =   43
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   40
      Left            =   2520
      TabIndex        =   42
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   39
      Left            =   2160
      TabIndex        =   41
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   38
      Left            =   1800
      TabIndex        =   40
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   37
      Left            =   1440
      TabIndex        =   39
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   36
      Left            =   1080
      TabIndex        =   38
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   35
      Left            =   3960
      TabIndex        =   37
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   34
      Left            =   3600
      TabIndex        =   36
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   33
      Left            =   3240
      TabIndex        =   35
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   32
      Left            =   2880
      TabIndex        =   34
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   31
      Left            =   2520
      TabIndex        =   33
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   30
      Left            =   2160
      TabIndex        =   32
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   29
      Left            =   1800
      TabIndex        =   31
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   28
      Left            =   1440
      TabIndex        =   30
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   27
      Left            =   1080
      TabIndex        =   29
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   26
      Left            =   3960
      TabIndex        =   28
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   25
      Left            =   3600
      TabIndex        =   27
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   24
      Left            =   3240
      TabIndex        =   26
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   23
      Left            =   2880
      TabIndex        =   25
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   22
      Left            =   2520
      TabIndex        =   24
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   21
      Left            =   2160
      TabIndex        =   23
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   20
      Left            =   1800
      TabIndex        =   22
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   19
      Left            =   1440
      TabIndex        =   21
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   18
      Left            =   1080
      TabIndex        =   20
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   17
      Left            =   3960
      TabIndex        =   19
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   16
      Left            =   3600
      TabIndex        =   18
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   15
      Left            =   3240
      TabIndex        =   17
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   14
      Left            =   2880
      TabIndex        =   16
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   13
      Left            =   2520
      TabIndex        =   15
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   12
      Left            =   2160
      TabIndex        =   14
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   11
      Left            =   1800
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   10
      Left            =   1440
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   8
      Left            =   3960
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   7
      Left            =   3600
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   6
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "O"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "time"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, z As Integer
Private Sub Command1_Click()
Timer2.Enabled = True
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
a = 0: z = 0
For i = 0 To 44
Label3(i) = "○"
Next
End Sub

Private Sub Command2_Click()
Timer2.Enabled = False
Timer3.Enabled = True
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
a = 0: z = 0
For i = 0 To 44
Label3(i) = "○"
Next
End Sub

Private Sub Command3_Click()
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = True
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
a = 0: z = 0
For i = 0 To 44
Label3(i) = "○"
Next
End Sub

Private Sub Command4_Click()
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = True
Timer6.Enabled = False
Timer7.Enabled = False
a = 0: z = 0
For i = 0 To 44
Label3(i) = "○"
Next
End Sub

Private Sub Command5_Click()
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = True
Timer7.Enabled = False
a = 0: z = 0
For i = 0 To 44
Label3(i) = "○"
Next
End Sub

Private Sub Command6_Click()
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = True
a = 0: z = 0
For i = 0 To 44
Label3(i) = "○"
Next
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Form_Activate()
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
End Sub

Private Sub Timer1_Timer()
Label1 = Format(Time, "hh:mm:ss")
End Sub

Private Sub Timer2_Timer()

If z = 1 Then
For i = 0 To 44
Label3(i) = "○"
Next
Timer2.Enabled = False
End If

If z = 0 Then
If a Mod 9 <> 0 Then Label3(a Mod 9 - 1) = "○"
Label3(a Mod 9) = "●"
a = a + 1
If a Mod 9 = 0 Then
Label3(8) = "●"
z = 1
End If
End If

End Sub

Private Sub Timer3_Timer()
b = Array(0, 1, 2, 3, 4, 3, 2, 1, 0)
If z = 1 Then
For i = 0 To 44
Label3(i) = "○"
Next
Timer3.Enabled = False
End If

If z = 0 Then
If a Mod 9 <> 0 Then Label3(Val(b(a Mod 9 - 1))) = "○"
Label3(Val(b(a Mod 9))) = "●"
a = a + 1

If a Mod 9 = 0 Then
Label3(0) = "●"
z = 1
End If
End If
End Sub

Private Sub Timer4_Timer()
b = Array(0, 2, 4, 6, 1, 3, 5, 7)

If z = 1 Then
For i = 0 To 44
Label3(i) = "○"
Next
Timer4.Enabled = False
End If

If z = 0 Then
If a Mod 8 <> 0 Then Label3(Val(b(a Mod 8 - 1))) = "○"
Label3(Val(b(a Mod 8))) = "●"
a = a + 1
If a Mod 8 = 0 Then
Label3(7) = "●"
z = 1
End If
End If
End Sub

Private Sub Timer5_Timer()
b = Array(0, 10, 20, 30, 40, 32, 24, 16, 8)


If z = 1 Then
For i = 0 To 44
Label3(i) = "○"
Next
Timer5.Enabled = False
End If

If z = 0 Then
If a Mod 9 <> 0 Then Label3(Val(b(a Mod 9 - 1))) = "○"
Label3(Val(b(a Mod 9))) = "●"
a = a + 1

If a Mod 9 = 0 Then
Label3(8) = "●"
z = 1
End If
End If
End Sub

Private Sub Timer6_Timer()
If a Mod 5 = 0 And z = 1 Then
For i = 0 To 44
Label3(i) = "O"
Next
Timer6.Enabled = False
z = 2
End If

If z < 2 Then
If a Mod 5 = 0 Then Label3(40) = "●": Label3(0) = "○": Label3(8) = "○": z = 1
If a Mod 5 = 1 Then Label3(30) = "●": Label3(32) = "●": Label3(40) = "○"
If a Mod 5 = 2 Then Label3(20) = "●": Label3(24) = "●": Label3(30) = "○": Label3(32) = "○"
If a Mod 5 = 3 Then Label3(10) = "●": Label3(16) = "●": Label3(20) = "○": Label3(24) = "○"
If a Mod 5 = 4 Then Label3(0) = "●": Label3(8) = "●": Label3(10) = "○": Label3(16) = "○"
a = a + 1
End If
End Sub

Private Sub Timer7_Timer()
b = Array(0, 1, 10, 2, 11, 20, 3, 12, 4, 5, 14, 6, 15, 24, 7, 16, 8)


If a Mod 9 = 0 And z = 1 Then
For i = 0 To 44
Label3(i) = "O"
Next
Timer7.Enabled = False
z = 2
End If


If z < 2 Then
If a Mod 9 = 0 Then
For i = 0 To 0
Label3(b(i)) = "●"
Next
z = 1
End If

If a Mod 9 = 1 Then
For i = 0 To 2
Label3(b(i)) = "●"
Next
End If

If a Mod 9 = 2 Then
For i = 0 To 5
Label3(b(i)) = "●"
Next
End If

If a Mod 9 = 3 Then
For i = 0 To 7
Label3(b(i)) = "●"
Next
End If

If a Mod 9 = 4 Then
For i = 0 To 8
Label3(b(i)) = "●"
Next
End If

If a Mod 9 = 5 Then
For i = 0 To 10
Label3(b(i)) = "●"
Next
End If

If a Mod 9 = 6 Then
For i = 0 To 13
Label3(b(i)) = "●"
Next
End If

If a Mod 9 = 7 Then
For i = 0 To 15
Label3(b(i)) = "●"
Next
End If


If a Mod 9 = 8 Then
For i = 0 To 16
Label3(b(i)) = "●"
Next
End If
a = a + 1
End If
End Sub
