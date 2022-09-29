VERSION 5.00
Begin VB.Form g_print 
   AutoRedraw      =   -1  'True
   Caption         =   "列印禮品清單"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   7620
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "關閉"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Data data_print 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\bin\桌面\禮品兌換\gift.mdb"
      DefaultCursorType=   0  '預設的資料指標
      DefaultType     =   2  '使用 ODBCDirect
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '動態集(Dynaset)
      RecordSource    =   "列印清單"
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
  dat1 = Format(InputBox("輸入起始日期", "日期1", Date - 10), "yyyy/mm/dd")
  dat2 = Format(InputBox("輸入終止日期", "日期2", Date), "yyyy/mm/dd")

  Print
  Print "會員編號", "會員姓名", "寄送地址", , "兌換日期", "禮品名稱"
  Print String(120, "-")
  data_print.Recordset.MoveFirst
  c = 0
  Do While Not data_print.Recordset.EOF
     If data_print.Recordset![兌換日期] >= dat1 And data_print.Recordset![兌換日期] <= dat2 Then
       data_print.Recordset.Edit
       data_print.Recordset![寄送註記] = True
       data_print.Recordset.Update
       
       If data_print.Recordset![會員編號] = s_ID Then c = c + 1 Else c = 0
       If c > 0 Then
         Print , , , , data_print.Recordset![兌換日期], data_print.Recordset![禮品名稱]
         data_print.Recordset.MoveNext
       Else
         s_ID = data_print.Recordset![會員編號]
         Print data_print.Recordset![會員編號],
         Print data_print.Recordset![會員姓名],
         Print data_print.Recordset![寄送地址],
         Print data_print.Recordset![兌換日期],
         Print data_print.Recordset![禮品名稱]
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
