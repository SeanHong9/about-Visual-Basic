VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox txt_b 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   "53284716"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txt_a 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "46317825"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmd_war 
      Caption         =   "�@��"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "�A�X����"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�ҥX����"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_war_Click()
 Dim na As Integer
 Dim nb As Integer
 Cls
 
 na = 1: nb = 1
  
 Do
   a = Mid(txt_a, na, 1)
   b = Mid(txt_b, nb, 1)
   Print a; b
   If a >= b Then
     If a = 8 And b = 1 Then
       na = na + 1
     Else
       nb = nb + 1
     End If
   End If
   If a <= b Then
     If a = 1 And b = 8 Then
       nb = nb + 1
     Else
       na = na + 1
     End If
   End If
 Loop Until na > 8 Or nb > 8
 
 If na = nb Then
    Print "����"
 Else
   If na > 8 Then Print "�Awin" Else Print "��win"
 End If
 
End Sub
