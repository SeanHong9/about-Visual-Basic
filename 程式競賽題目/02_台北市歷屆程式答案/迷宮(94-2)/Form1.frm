VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "細明體"
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
   StartUpPosition =   3  '系統預設值
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c$(), goodx$, goody$, goodstep% '迷宮、最短路徑、最少步數

Private Sub Form_Load()
Dim tx As Variant, ty As Variant, Sx%, Sy% '暫存路徑、起點座標
tx = 0
Open "test1.txt" For Input As 1
Do Until EOF(1)
  Line Input #1, ty
  tx = tx + 1
Loop
Close
ty = Split(Trim(ty), " ")
ReDim c(tx, UBound(ty) + 1) '迷宮(高度,寬度)
Open "test1.txt" For Input As 1
Open "result1.txt" For Output As 2
For x = 1 To UBound(c, 1)
  Line Input #1, ty
  ty = Split(Trim(ty), " ")
  For y = 1 To UBound(c, 2)
    c(x, y) = ty(y - 1)
    If ty(y - 1) = "S" Then Sx = x: Sy = y '紀錄起點
  Next y
Next x
goodstep = UBound(c, 1) * UBound(c, 2) '預設步數最多為[高度*寬度]
gogo Sx, Sy, "", "", 0 '找最短路徑
tx = Split(goodx, " ") '分解路徑
ty = Split(goody, " ")
For i = 1 To UBound(tx)
  c(tx(i), ty(i)) = "H" '標上最佳路徑
Next i
For i = 1 To UBound(c, 1) '輸出
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
Dim tx As Variant, ty As Variant '暫存路徑
If Nx > UBound(c, 1) Or Nx < 1 Or Ny > UBound(c, 2) Or Ny < 1 Then Exit Sub '檢查範圍
If c(Nx, Ny) = "0" Then Exit Sub '檢查死路
If c(Nx, Ny) = "X" And step < goodstep Then '檢查是否終點，步數是否較少
  goodx = Trim(wayx) '紀錄最短路徑
  goody = Trim(wayy)
  goodstep = step '紀錄最少步數
  Exit Sub
End If
tx = Split(wayx, " ") '分解路徑
ty = Split(wayy, " ")
For i = 0 To UBound(tx) '檢查座標是否重複
  If Nx = tx(i) And Ny = ty(i) Then Exit Sub
Next i
wayx = wayx & " " & Nx '紀錄路徑
wayy = wayy & " " & Ny '
step = step + 1 '增加步數
gogo Nx - 1, Ny, wayx, wayy, step '往上走
gogo Nx + 1, Ny, wayx, wayy, step '往下走
gogo Nx, Ny - 1, wayx, wayy, step '往左走
gogo Nx, Ny + 1, wayx, wayy, step '往右走
End Sub
