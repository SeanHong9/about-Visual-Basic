VERSION 5.00
Begin VB.Form g_print 
   AutoRedraw      =   -1  'True
   Caption         =   "�C�L§�~�M��"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   7620
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Data data_print 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\bin\�ୱ\§�~�I��\gift.mdb"
      DefaultCursorType=   0  '�w�]����ƫ���
      DefaultType     =   2  '�ϥ� ODBCDirect
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '�ʺA��(Dynaset)
      RecordSource    =   "�C�L�M��"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "g_print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload g_print
  main.Show
  
End Sub

Private Sub Form_Activate()
  Dim dat1 As Date, dat2 As Date
  Cls
  dat1 = Format(InputBox("��J�_�l���", "���1", Date - 10), "yyyy/mm/dd")
  dat2 = Format(InputBox("��J�פ���", "���2", Date), "yyyy/mm/dd")

  Print
  Print "�|���s��", "�|���m�W", "�H�e�a�}", , "�I�����", "§�~�W��"
  Print String(120, "-")
  data_print.Recordset.MoveFirst
  c = 0
  Do While Not data_print.Recordset.EOF
     If data_print.Recordset![�I�����] >= dat1 And data_print.Recordset![�I�����] <= dat2 Then
       data_print.Recordset.Edit
       data_print.Recordset![�H�e���O] = True
       data_print.Recordset.Update
       
       If data_print.Recordset![�|���s��] = s_ID Then c = c + 1 Else c = 0
       If c > 0 Then
         Print , , , , data_print.Recordset![�I�����], data_print.Recordset![§�~�W��]
         data_print.Recordset.MoveNext
       Else
         s_ID = data_print.Recordset![�|���s��]
         Print data_print.Recordset![�|���s��],
         Print data_print.Recordset![�|���m�W],
         Print data_print.Recordset![�H�e�a�}],
         Print data_print.Recordset![�I�����],
         Print data_print.Recordset![§�~�W��]
         data_print.Recordset.MoveNext
       End If
     Else
       data_print.Recordset.MoveNext
     End If
  Loop
End Sub

Private Sub Form_Load()
    data_print.DatabaseName = App.Path + "\gift.mdb"

End Sub
