VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "�I�\��"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   5415
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   8
      Left            =   3840
      TabIndex        =   18
      Text            =   "0"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   7
      Left            =   3840
      TabIndex        =   17
      Text            =   "0"
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   6
      Left            =   3840
      TabIndex        =   16
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   5
      Left            =   3840
      TabIndex        =   15
      Text            =   "0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   4
      Left            =   3840
      TabIndex        =   14
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   3
      Left            =   3840
      TabIndex        =   13
      Text            =   "0"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   2
      Left            =   3840
      TabIndex        =   12
      Text            =   "0"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   1
      Left            =   3840
      TabIndex        =   11
      Text            =   "0"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txt_num 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      Height          =   270
      Index           =   0
      Left            =   3840
      TabIndex        =   10
      Text            =   "0"
      Top             =   1920
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "�\���޲z.frx":0000
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '�m�����
      DataField       =   "�ู"
      DataSource      =   "data_no"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Data data_no 
      Caption         =   "�\��"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\work_new\red\VB�{���]�p\�x�_���{���D��\�\���޲z(97-1).mdb"
      DefaultCursorType=   0  '�w�]����ƫ���
      DefaultType     =   2  '�ϥ� ODBCDirect
      Exclusive       =   0   'False
      Height          =   405
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '�ʺA��(Dynaset)
      RecordSource    =   "�\��"
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Data data_menu 
      Caption         =   "menu"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\work_new\red\VB�{���]�p\�x�_���{���D��\�\���޲z(97-1).mdb"
      DefaultCursorType=   0  '�w�]����ƫ���
      DefaultType     =   2  '�ϥ� ODBCDirect
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '�ʺA��(Dynaset)
      RecordSource    =   "�\��"
      Top             =   5520
      Width           =   1860
   End
   Begin VB.Data data_order 
      Caption         =   "�\�椺�e"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\work_new\red\VB�{���]�p\�x�_���{���D��\�\���޲z(97-1).mdb"
      DefaultCursorType=   0  '�w�]����ƫ���
      DefaultType     =   2  '�ϥ� ODBCDirect
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '�ʺA��(Dynaset)
      RecordSource    =   "�\�椺�e"
      Top             =   6480
      Width           =   2100
   End
   Begin VB.Label Label11 
      Caption         =   "�ƶq"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label7 
      Caption         =   "�ɶ�"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "���"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      DataField       =   "�I�\�ɶ�"
      DataSource      =   "data_no"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      DataField       =   "�I�\���"
      DataSource      =   "data_no"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "�ู"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   1  '��u�T�w
      DataField       =   "�\��s��"
      DataSource      =   "data_no"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�\��s��"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
