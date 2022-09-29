VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10740
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command6 
      Caption         =   "end "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   21
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "star"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   20
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "↑"
      Height          =   735
      Left            =   8400
      TabIndex        =   19
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "→"
      Height          =   735
      Left            =   9360
      TabIndex        =   18
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "↓"
      Height          =   735
      Left            =   8400
      TabIndex        =   17
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "←"
      Height          =   735
      Left            =   7320
      TabIndex        =   16
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   15
      Left            =   5640
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   14
      Left            =   4080
      TabIndex        =   14
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   13
      Left            =   2520
      TabIndex        =   13
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   12
      Left            =   960
      TabIndex        =   12
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   11
      Left            =   5640
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   10
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   9
      Left            =   2520
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   5640
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   4080
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   5640
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = 0
For j = 0 To 4

For i = 3 To 1 Step -1
If Label1(i) <> "" Then
If Label1(i) = Label1(i - 1) Then
Label1(i - 1) = Label1(i - 1) * 2
Label1(i) = ""
Else
If Label1(i - 1) = "" Then
Label1(i - 1) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 7 To 5 Step -1
If Label1(i) <> "" Then
If Label1(i) = Label1(i - 1) Then
Label1(i - 1) = Label1(i - 1) * 2
Label1(i) = ""
Else
If Label1(i - 1) = "" Then
Label1(i - 1) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 11 To 9 Step -1
If Label1(i) <> "" Then
If Label1(i) = Label1(i - 1) Then
Label1(i - 1) = Label1(i - 1) * 2
Label1(i) = ""
Else
If Label1(i - 1) = "" Then
Label1(i - 1) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 15 To 13 Step -1
If Label1(i) <> "" Then
If Label1(i) = Label1(i - 1) Then
Label1(i - 1) = Label1(i - 1) * 2
Label1(i) = ""
Else
If Label1(i - 1) = "" Then
Label1(i - 1) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next
Next

For i = 0 To 15
If Label1(i) <> "" Then a = a - (-1)
If Label1(i) = "2048" Then
MsgBox "end", vbOK, "end"
End If
Next

If a = 16 Then MsgBox "you lost", vbOK, "you lost"

For i = 1 To 1
Randomize
a = Int(Rnd * 16)
If Label1(a) = "" Then
b = Int(Rnd() * 2)
If b = 1 Then Label1(a) = 2
If b = 0 Then Label1(a) = 4
End If
Next



End Sub

Private Sub Command2_Click()
a = 0
For j = 0 To 4
For i = 0 To 8 Step 4
If Label1(i) <> "" Then
If Label1(i) = Label1(i - (-4)) Then
Label1(i - (-4)) = Label1(i - (-4)) * 2
Label1(i) = ""
Else
If Label1(i - (-4)) = "" Then
Label1(i - (-4)) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 1 To 9 Step 4
If Label1(i) <> "" Then
If Label1(i) = Label1(i - (-4)) Then
Label1(i - (-4)) = Label1(i - (-4)) * 2
Label1(i) = ""
Else
If Label1(i - (-4)) = "" Then
Label1(i - (-4)) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 2 To 10 Step 4
If Label1(i) <> "" Then
If Label1(i) = Label1(i - (-4)) Then
Label1(i - (-4)) = Label1(i - (-4)) * 2
Label1(i) = ""
Else
If Label1(i - (-4)) = "" Then
Label1(i - (-4)) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 3 To 11 Step 4
If Label1(i) <> "" Then
If Label1(i) = Label1(i - (-4)) Then
Label1(i - (-4)) = Label1(i - (-4)) * 2
Label1(i) = ""
Else
If Label1(i - (-4)) = "" Then
Label1(i - (-4)) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next
Next

For i = 0 To 15
If Label1(i) <> "" Then a = a - (-1)
If Label1(i) = "2048" Then
MsgBox "end", vbOK, "end"
End If
Next

If a = 16 Then MsgBox "you lost", vbOK, "you lost"

For i = 1 To 1
Randomize
a = Int(Rnd * 16)
If Label1(a) = "" Then
b = Int(Rnd() * 2)
If b = 1 Then Label1(a) = 2
If b = 0 Then Label1(a) = 4
End If
Next
End Sub

Private Sub Command3_Click()
a = 0
For j = 0 To 4
For i = 0 To 2
If Label1(i) <> "" Then
If Label1(i) = Label1(i - (-1)) Then
Label1(i - (-1)) = Label1(i - (-1)) * 2
Label1(i) = ""
Else
If Label1(i - (-1)) = "" Then
Label1(i - (-1)) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 4 To 6
If Label1(i) <> "" Then
If Label1(i) = Label1(i - (-1)) Then
Label1(i - (-1)) = Label1(i - (-1)) * 2
Label1(i) = ""
Else
If Label1(i - (-1)) = "" Then
Label1(i - (-1)) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 8 To 10
If Label1(i) <> "" Then
If Label1(i) = Label1(i - (-1)) Then
Label1(i - (-1)) = Label1(i - (-1)) * 2
Label1(i) = ""
Else
If Label1(i - (-1)) = "" Then
Label1(ii - (-1)) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 12 To 14
If Label1(i) <> "" Then
If Label1(i) = Label1(i - (-1)) Then
Label1(i - (-1)) = Label1(i - (-1)) * 2
Label1(i) = ""
Else
If Label1(i - (-1)) = "" Then
Label1(i - (-1)) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next
Next


For i = 0 To 15
If Label1(i) <> "" Then a = a - (-1)
If Label1(i) = "2048" Then
MsgBox "end", vbOK, "end"
End If
Next

If a = 16 Then MsgBox "you lost", vbOK, "you lost"

For i = 1 To 1
Randomize
a = Int(Rnd * 16)
If Label1(a) = "" Then
b = Int(Rnd() * 2)
If b = 1 Then Label1(a) = 2
If b = 0 Then Label1(a) = 4
End If
Next
End Sub

Private Sub Command4_Click()
a = 0

For j = 0 To 4
For i = 12 To 4 Step -4
If Label1(i) <> "" Then
If Label1(i) = Label1(i - 4) Then
Label1(i - 4) = Label1(i - 4) * 2
Label1(i) = ""
Else
If Label1(i - 4) = "" Then
Label1(i - 4) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 13 To 5 Step -4
If Label1(i) <> "" Then
If Label1(i) = Label1(i - 4) Then
Label1(i - 4) = Label1(i - 4) * 2
Label1(i) = ""
Else
If Label1(i - 4) = "" Then
Label1(i - 4) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 14 To 6 Step -4
If Label1(i) <> "" Then
If Label1(i) = Label1(i - 4) Then
Label1(i - 4) = Label1(i - 4) * 2
Label1(i) = ""
Else
If Label1(i - 4) = "" Then
Label1(i - 4) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next

For i = 15 To 7 Step -4
If Label1(i) <> "" Then
If Label1(i) = Label1(i - 4) Then
Label1(i - 4) = Label1(i - 4) * 2
Label1(i) = ""
Else
If Label1(i - 4) = "" Then
Label1(i - 4) = Label1(i)
Label1(i) = ""
End If
End If
End If
Next
Next


For i = 0 To 15
If Label1(i) <> "" Then a = a - (-1)
If Label1(i) = "2048" Then
MsgBox "end", vbOK, "end"
End If
Next


If a = 16 Then MsgBox "you lost", vbOK, "you lost"


For i = 1 To 1
Randomize
a = Int(Rnd * 16)
If Label1(a) = "" Then
b = Int(Rnd() * 2)
If b = 1 Then Label1(a) = 2
If b = 0 Then Label1(a) = 4
End If
Next
End Sub

Private Sub Command5_Click()
For i = 0 To 15
Label1(i) = ""
Next



For i = 0 To 1
Randomize
a = Int(Rnd * 16)
If Label1(a) = "" Then
Label1(a) = 2
End If
Next

Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Activate()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
End Sub

