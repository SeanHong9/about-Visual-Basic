VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "�ө���"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c$(), goodx$, goody$, goodstep% '�g�c�B�̵u���|�B�̤֨B��

Private Sub Form_Load()
Dim tx As Variant, ty As Variant, Sx%, Sy% '�Ȧs���|�B�_�I�y��
tx = 0
Open "test1.txt" For Input As 1
Do Until EOF(1)
  Line Input #1, ty
  tx = tx + 1
Loop
Close
ty = Split(Trim(ty), " ")
ReDim c(tx, UBound(ty) + 1) '�g�c(����,�e��)
Open "test1.txt" For Input As 1
Open "result1.txt" For Output As 2
For x = 1 To UBound(c, 1)
  Line Input #1, ty
  ty = Split(Trim(ty), " ")
  For y = 1 To UBound(c, 2)
    c(x, y) = ty(y - 1)
    If ty(y - 1) = "S" Then Sx = x: Sy = y '�����_�I
  Next y
Next x
goodstep = UBound(c, 1) * UBound(c, 2) '�w�]�B�Ƴ̦h��[����*�e��]
gogo Sx, Sy, "", "", 0 '��̵u���|
tx = Split(goodx, " ") '���Ѹ��|
ty = Split(goody, " ")
For i = 1 To UBound(tx)
  c(tx(i), ty(i)) = "H" '�ФW�̨θ��|
Next i
For i = 1 To UBound(c, 1) '��X
  For j = 1 To UBound(c, 2)
    Print #2, c(i, j);
    If j < UBound(c, 2) Then Print #2, " ";
  Next j
  Print #2,
Next i
Close
End
End Sub


Sub gogo(ByVal Nx%, ByVal Ny%, ByVal wayx$, ByVal wayy$, ByVal step%)
Dim tx As Variant, ty As Variant '�Ȧs���|
If Nx > UBound(c, 1) Or Nx < 1 Or Ny > UBound(c, 2) Or Ny < 1 Then Exit Sub '�ˬd�d��
If c(Nx, Ny) = "0" Then Exit Sub '�ˬd����
If c(Nx, Ny) = "X" And step < goodstep Then '�ˬd�O�_���I�A�B�ƬO�_����
  goodx = Trim(wayx) '�����̵u���|
  goody = Trim(wayy)
  goodstep = step '�����̤֨B��
  Exit Sub
End If
tx = Split(wayx, " ") '���Ѹ��|
ty = Split(wayy, " ")
For i = 0 To UBound(tx) '�ˬd�y�ЬO�_����
  If Nx = tx(i) And Ny = ty(i) Then Exit Sub
Next i
wayx = wayx & " " & Nx '�������|
wayy = wayy & " " & Ny '
step = step + 1 '�W�[�B��
gogo Nx - 1, Ny, wayx, wayy, step '���W��
gogo Nx + 1, Ny, wayx, wayy, step '���U��
gogo Nx, Ny - 1, wayx, wayy, step '������
gogo Nx, Ny + 1, wayx, wayy, step '���k��
End Sub
