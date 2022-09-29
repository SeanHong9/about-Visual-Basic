VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14340
   LinkTopic       =   "Form2"
   ScaleHeight     =   7560
   ScaleWidth      =   14340
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command6 
      Caption         =   "新增"
      Height          =   855
      Left            =   10080
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   10560
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "選購結帳完成"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   3840
      ItemData        =   "第三題.frx":0000
      Left            =   5400
      List            =   "第三題.frx":0002
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   3840
      ItemData        =   "第三題.frx":0004
      Left            =   480
      List            =   "第三題.frx":0006
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   26.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      TabIndex        =   14
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "價錢"
      Height          =   375
      Index           =   3
      Left            =   10560
      TabIndex        =   12
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "商品"
      Height          =   375
      Index           =   2
      Left            =   8760
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "顧客選購物品區"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "商品陳列區"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = List1.ListIndex
List2.AddItem List1.Text
List1.List(a) = ""

For i = 0 To List1.ListCount
If List1.List(i) = "" Then List1.List(i) = List1.List(i - (-1)): List1.List(i - (-1)) = ""
Next


End Sub

Private Sub Command2_Click()
For i = 0 To List1.ListCount
a = List1.List(i)
List2.AddItem a
Next

List1.Clear
End Sub

Private Sub Command3_Click()
a = List2.ListIndex
List1.AddItem List2.Text
List2.List(a) = ""


For i = 0 To List2.ListCount
If List2.List(i) = "" Then List2.List(i) = List2.List(i - (-1)): List2.List(i - (-1)) = ""
Next
End Sub

Private Sub Command4_Click()
For i = 0 To List2.ListCount
a = List2.List(i)
List1.AddItem a
Next

List2.Clear
End Sub

Private Sub Command5_Click()
Dim d(20) As String
x = 0
For i = 0 To List2.ListCount
If List2.List(i) <> "" Then
      For j = 1 To Len(List2.List(i))
      If Mid(List2.List(i), j, 1) = "$" Then
      d(x) = Mid(List2.List(i), j - (-1))
      x = x - (-1)
      Exit For
      End If
      Next
End If
Next

Sum = 0
For i = 0 To UBound(d)
Sum = Sum - (-Val(d(i)))
Next

Label2 = Sum
End Sub

Private Sub Command6_Click()
If Text1(0) <> "" And Text1(1) <> "" Then
List1.AddItem Text1(0) & "---$" & Text1(1)
End If
Text1(0) = "": Text1(1) = ""
End Sub

