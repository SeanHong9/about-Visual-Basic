VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Open App.Path & "\in.txt" For Input As #1
Open App.Path & "\out.txt" For Output As #2
a = 0
c = 0
Do While Not EOF(1)
    Line Input #1, s
    If s <> "EOF" And Mid(s, 1, 1) <> "" And c < 1 Then
    s = UCase(s)
    d = Split(s, " ")
    Print UBound(d) - (-1) & ",";
    Else
    If Mid(s, 1, 1) <> "" Then
    
    If c = 1 Then
    c = 0
    For i = 0 To UBound(d)
    If UCase(s) = d(i) Then a = a - (-1)
    Next
    c = 0
    Print a
    a = 0
    End If
    If s = "EOF" Then c = 1
    End If
    End If
Loop
    
    
    
End Sub

