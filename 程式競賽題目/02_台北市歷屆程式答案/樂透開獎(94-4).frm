VERSION 5.00
Begin VB.Form frmHappy 
   Caption         =   "�ֳz�}��"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5955
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton over 
      Caption         =   "���}"
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   3960
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�}��"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "���z����"
      BeginProperty Font 
         Name            =   "�رd�ɫF�y"
         Size            =   26.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "�S�O��"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.Label number 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   6
      Left            =   4440
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label number 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   5
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label number 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   4
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.Label number 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Label number 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.Label number 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label number 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmHappy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num(1 To 42) As Boolean
Dim i As Byte
Private Sub Command1_Click()
  Erase num
  For i = 0 To 6
    number(i) = ""
  Next i
  i = 0
  Timer1.Enabled = True
  Timer2.Enabled = True
End Sub

Private Sub over_Click()
  End
End Sub

Private Sub Timer1_Timer()
  
    Do
      n = Int(Rnd * 42) + 1
    Loop While num(n) = True
    
    number(i) = n
    num(n) = True
    i = i + 1
    If i = 7 Then
      Timer1.Enabled = False
      Timer2.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
   number(i) = Int(Rnd * 42) + 1
End Sub
