VERSION 5.00
Begin VB.Form main 
   AutoRedraw      =   -1  'True
   Caption         =   "§�~�I��"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   6735
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton cmd_print 
      Caption         =   "�C�L�H�e�M��"
      Height          =   495
      Left            =   5040
      TabIndex        =   25
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txt_date 
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Data data_ok 
      Caption         =   "ok"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\bin\�ୱ\§�~�I��\gift.mdb"
      DefaultCursorType=   0  '�w�]����ƫ���
      DefaultType     =   2  '�ϥ� ODBCDirect
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '�ʺA��(Dynaset)
      RecordSource    =   "�I�����"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   10
      Left            =   2640
      TabIndex        =   20
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   19
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   8
      Left            =   2640
      TabIndex        =   18
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   15
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk 
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Data data_gift 
      Caption         =   "gift"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\bin\�ୱ\§�~�I��\gift.mdb"
      DefaultCursorType=   0  '�w�]����ƫ���
      DefaultType     =   2  '�ϥ� ODBCDirect
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '�ʺA��(Dynaset)
      RecordSource    =   "§�~�d��"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cmb_part 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "gift.frx":0000
      Left            =   1680
      List            =   "gift.frx":0010
      TabIndex        =   9
      Text            =   "�̺����d��"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox cmb_dot 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "gift.frx":0032
      Left            =   240
      List            =   "gift.frx":0045
      TabIndex        =   8
      Text            =   "���I�Ƭd��"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmd_work 
      Caption         =   "�T�w�I��"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txt_pass 
      Alignment       =   2  '�m�����
      Height          =   375
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "08140726"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "�T�w��J"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txt_ID 
      Alignment       =   2  '�m�����
      DataSource      =   "data_mem"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "00003"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Data data_mem 
      Caption         =   "member"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\bin\�ୱ\§�~�I��\gift.mdb"
      DefaultCursorType=   0  '�w�]����ƫ���
      DefaultType     =   2  '�ϥ� ODBCDirect
      Exclusive       =   0   'False
      Height          =   405
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '�ʺA��(Dynaset)
      RecordSource    =   "�|�����"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   6240
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label5 
      Caption         =   "�I�����"
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lbl_else 
      Alignment       =   2  '�m�����
      BorderStyle     =   1  '��u�T�w
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "�Ѿl�I��"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lbl_dot 
      Alignment       =   2  '�m�����
      BorderStyle     =   1  '��u�T�w
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   6000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "���O���Q�I��"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "��J�K�X"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�|���s��"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim g_name(20) As String
Dim g_dot(20) As Long
Dim c As Integer

Private Sub chk_Click(Index As Integer)
  If chk(Index).Value = 1 Then
    lbl_else = Val(lbl_else) - g_dot(Index)
    If Val(lbl_else) < 0 Then
      chk(Index).Value = 0
      MsgBox "�I�Ƥ���", vbCritical + vbOKOnly, "���~"
    End If
  Else
    lbl_else = Val(lbl_else) + g_dot(Index)
  End If
End Sub

Private Sub cmb_dot_Click()
  Cls
  
  For i = 0 To 10
    chk(i).Value = 0
  Next i
   For i = 0 To c
    chk(i).Visible = False
  Next i
  Erase g_name
  Erase g_dot
  CurrentY = 4000
  Print Tab(5); "§�~�W��", "�I���I��", "�I���Ŀ�"
  Print String(55, "-")
  c = 0
  data_gift.Recordset.MoveFirst
  Do While Not data_gift.Recordset.EOF
    Select Case cmb_dot.ListIndex
      Case 0
       If data_gift.Recordset![�I���I��] <= 1000 Then
         Call printfield
       End If
      Case 1
       If data_gift.Recordset![�I���I��] > 1000 And data_gift.Recordset![�I���I��] <= 2000 Then
         Call printfield
       End If
      Case 2
       If data_gift.Recordset![�I���I��] > 2000 And data_gift.Recordset![�I���I��] <= 3000 Then
         Call printfield
       End If
      Case 3
       If data_gift.Recordset![�I���I��] > 3000 And data_gift.Recordset![�I���I��] <= 4000 Then
         Call printfield
       End If
      Case Else
       If data_gift.Recordset![�I���I��] > 4000 Then
         Call printfield
       End If
    End Select
    data_gift.Recordset.MoveNext
    
  Loop
  data_gift.Refresh
End Sub
Sub printfield()
  
  g_name(c) = data_gift.Recordset![§�~�W��]
  g_dot(c) = data_gift.Recordset![�I���I��]
  chk(c).Visible = True
  Print Tab(5); data_gift.Recordset![§�~�W��],
  Print data_gift.Recordset![�I���I��]
  Print
  c = c + 1
End Sub

Private Sub cmb_part_Click()
  Cls
  
  For i = 0 To 10
    chk(i).Value = 0
  Next i
   For i = 0 To c
    chk(i).Visible = False
  Next i
  Erase g_name
  Erase g_dot
  CurrentY = 4000
  Print Tab(5); "§�~�W��", "�I���I��", "�I���Ŀ�"
  Print String(55, "-")
  c = 0
  data_gift.Recordset.MoveFirst
  Do While Not data_gift.Recordset.EOF
    Select Case cmb_part.ListIndex
      Case 0
       If data_gift.Recordset![§�~����] = 0 Then
         Call printfield
       End If
      Case 1
       If data_gift.Recordset![§�~����] = 1 Then
         Call printfield
       End If
      Case 2
       If data_gift.Recordset![§�~����] = 2 Then
         Call printfield
       End If
      Case 3
       If data_gift.Recordset![§�~����] = 3 Then
         Call printfield
       End If
      Case Else
        Call printfield
    End Select
    data_gift.Recordset.MoveNext
  Loop
  data_gift.Refresh

End Sub

Private Sub Cmd_ok_Click()
  Dim flag As Boolean
  Dim pass As String, t As String
  
  data_mem.Recordset.MoveFirst
  
  flag = False
  Do While Not data_mem.Recordset.EOF
     If data_mem.Recordset.Fields(0) = txt_ID Then
        flag = True
        pass = ""
        For i = 1 To Len(txt_pass)
          t = Right(Str((Val(Mid(txt_pass, i, 1)) * 3 + i) Mod 10), 1)
          pass = pass + t
        Next i
        If data_mem.Recordset.Fields(2) <> pass Then
           MsgBox "�K�X���~", vbCritical + vbOKOnly, "���~"
           Exit Do
        Else
           lbl_dot = data_mem.Recordset![�Ѿl���Q�I��]
           lbl_else = data_mem.Recordset![�Ѿl���Q�I��]
           cmd_work.Enabled = True
           cmb_dot.Enabled = True
           cmb_part.Enabled = True
           txt_date = Date
        End If
        Exit Do
     Else
        data_mem.Recordset.MoveNext
     End If
  Loop
  If flag = False Then
    MsgBox "�|���s�����s�b", vbCritical + vbOKOnly, "���~"
    data_mem.Refresh
    
  End If
     
End Sub

Private Sub cmd_print_Click()
  g_print.Show
End Sub

Private Sub cmd_work_Click()
  data_mem.Recordset.Edit
  data_mem.Recordset![�Ѿl���Q�I��] = Val(lbl_else)
  data_mem.Recordset.Update
  
  For i = 0 To c
     If chk(i).Value = 1 Then
        data_ok.Recordset.AddNew
        data_ok.Recordset![�|���s��] = txt_ID
        data_ok.Recordset![§�~�W��] = g_name(i)
        data_ok.Recordset![�I�����] = Format(txt_date, "yyyy/mm/dd")
        data_ok.Recordset![�H�e���O] = False
        data_ok.Recordset.Update
        
     End If
  Next i
  MsgBox "§�~�I���n�O����!", vbExclamation + vbOKOnly, "���\"
  For i = 0 To c
    chk(i).Value = 0
  Next i
  For i = 0 To c
    chk(i).Visible = False
  Next i
  Cls
  cmd_work.Enabled = False
  cmb_dot.Enabled = False
  cmb_part.Enabled = False
  lbl_else = ""
  lbl_dot = ""
  txt_date = ""
  txt_aa = ""
End Sub

Private Sub Form_Load()
  data_gift.DatabaseName = App.Path + "\gift.mdb"
  data_mem.DatabaseName = App.Path + "\gift.mdb"
  data_ok.DatabaseName = App.Path + "\gift.mdb"
End Sub
