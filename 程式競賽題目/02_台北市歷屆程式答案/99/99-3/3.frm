VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "標楷體"
      Size            =   18
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   10575
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "搜尋字串"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "迴文"
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   224
      Left            =   6840
      TabIndex        =   226
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   223
      Left            =   6360
      TabIndex        =   225
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   222
      Left            =   5880
      TabIndex        =   224
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   221
      Left            =   5400
      TabIndex        =   223
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   220
      Left            =   4920
      TabIndex        =   222
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   219
      Left            =   4440
      TabIndex        =   221
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   218
      Left            =   3960
      TabIndex        =   220
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   217
      Left            =   3480
      TabIndex        =   219
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   216
      Left            =   3000
      TabIndex        =   218
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   215
      Left            =   2520
      TabIndex        =   217
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   214
      Left            =   2040
      TabIndex        =   216
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   213
      Left            =   1560
      TabIndex        =   215
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   212
      Left            =   1080
      TabIndex        =   214
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   211
      Left            =   600
      TabIndex        =   213
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   210
      Left            =   120
      TabIndex        =   212
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   209
      Left            =   6840
      TabIndex        =   211
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   208
      Left            =   6360
      TabIndex        =   210
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   207
      Left            =   5880
      TabIndex        =   209
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   206
      Left            =   5400
      TabIndex        =   208
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   205
      Left            =   4920
      TabIndex        =   207
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   204
      Left            =   4440
      TabIndex        =   206
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   203
      Left            =   3960
      TabIndex        =   205
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   202
      Left            =   3480
      TabIndex        =   204
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   201
      Left            =   3000
      TabIndex        =   203
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   200
      Left            =   2520
      TabIndex        =   202
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   199
      Left            =   2040
      TabIndex        =   201
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   198
      Left            =   1560
      TabIndex        =   200
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   197
      Left            =   1080
      TabIndex        =   199
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   196
      Left            =   600
      TabIndex        =   198
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   195
      Left            =   120
      TabIndex        =   197
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   194
      Left            =   6840
      TabIndex        =   196
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   193
      Left            =   6360
      TabIndex        =   195
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   192
      Left            =   5880
      TabIndex        =   194
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   191
      Left            =   5400
      TabIndex        =   193
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   190
      Left            =   4920
      TabIndex        =   192
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   189
      Left            =   4440
      TabIndex        =   191
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   188
      Left            =   3960
      TabIndex        =   190
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   187
      Left            =   3480
      TabIndex        =   189
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   186
      Left            =   3000
      TabIndex        =   188
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   185
      Left            =   2520
      TabIndex        =   187
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   184
      Left            =   2040
      TabIndex        =   186
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   183
      Left            =   1560
      TabIndex        =   185
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   182
      Left            =   1080
      TabIndex        =   184
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   181
      Left            =   600
      TabIndex        =   183
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   180
      Left            =   120
      TabIndex        =   182
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   179
      Left            =   6840
      TabIndex        =   181
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   178
      Left            =   6360
      TabIndex        =   180
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   177
      Left            =   5880
      TabIndex        =   179
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   176
      Left            =   5400
      TabIndex        =   178
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   175
      Left            =   4920
      TabIndex        =   177
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   174
      Left            =   4440
      TabIndex        =   176
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   173
      Left            =   3960
      TabIndex        =   175
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   172
      Left            =   3480
      TabIndex        =   174
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   171
      Left            =   3000
      TabIndex        =   173
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   170
      Left            =   2520
      TabIndex        =   172
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   169
      Left            =   2040
      TabIndex        =   171
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   168
      Left            =   1560
      TabIndex        =   170
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   167
      Left            =   1080
      TabIndex        =   169
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   166
      Left            =   600
      TabIndex        =   168
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   165
      Left            =   120
      TabIndex        =   167
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   164
      Left            =   6840
      TabIndex        =   166
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   163
      Left            =   6360
      TabIndex        =   165
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   162
      Left            =   5880
      TabIndex        =   164
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   161
      Left            =   5400
      TabIndex        =   163
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   160
      Left            =   4920
      TabIndex        =   162
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   159
      Left            =   4440
      TabIndex        =   161
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   158
      Left            =   3960
      TabIndex        =   160
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   157
      Left            =   3480
      TabIndex        =   159
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   156
      Left            =   3000
      TabIndex        =   158
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   155
      Left            =   2520
      TabIndex        =   157
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   154
      Left            =   2040
      TabIndex        =   156
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   153
      Left            =   1560
      TabIndex        =   155
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   152
      Left            =   1080
      TabIndex        =   154
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   151
      Left            =   600
      TabIndex        =   153
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   150
      Left            =   120
      TabIndex        =   152
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   149
      Left            =   6840
      TabIndex        =   151
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   148
      Left            =   6360
      TabIndex        =   150
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   147
      Left            =   5880
      TabIndex        =   149
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   146
      Left            =   5400
      TabIndex        =   148
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   145
      Left            =   4920
      TabIndex        =   147
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   144
      Left            =   4440
      TabIndex        =   146
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   143
      Left            =   3960
      TabIndex        =   145
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   142
      Left            =   3480
      TabIndex        =   144
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   141
      Left            =   3000
      TabIndex        =   143
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   140
      Left            =   2520
      TabIndex        =   142
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   139
      Left            =   2040
      TabIndex        =   141
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   138
      Left            =   1560
      TabIndex        =   140
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   137
      Left            =   1080
      TabIndex        =   139
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   136
      Left            =   600
      TabIndex        =   138
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   135
      Left            =   120
      TabIndex        =   137
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   134
      Left            =   6840
      TabIndex        =   136
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   133
      Left            =   6360
      TabIndex        =   135
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   132
      Left            =   5880
      TabIndex        =   134
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   131
      Left            =   5400
      TabIndex        =   133
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   130
      Left            =   4920
      TabIndex        =   132
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   129
      Left            =   4440
      TabIndex        =   131
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   128
      Left            =   3960
      TabIndex        =   130
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   127
      Left            =   3480
      TabIndex        =   129
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   126
      Left            =   3000
      TabIndex        =   128
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   125
      Left            =   2520
      TabIndex        =   127
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   124
      Left            =   2040
      TabIndex        =   126
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   123
      Left            =   1560
      TabIndex        =   125
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   122
      Left            =   1080
      TabIndex        =   124
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   121
      Left            =   600
      TabIndex        =   123
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   120
      Left            =   120
      TabIndex        =   122
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   119
      Left            =   6840
      TabIndex        =   121
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   118
      Left            =   6360
      TabIndex        =   120
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   117
      Left            =   5880
      TabIndex        =   119
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   116
      Left            =   5400
      TabIndex        =   118
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   115
      Left            =   4920
      TabIndex        =   117
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   114
      Left            =   4440
      TabIndex        =   116
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   113
      Left            =   3960
      TabIndex        =   115
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   112
      Left            =   3480
      TabIndex        =   114
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   111
      Left            =   3000
      TabIndex        =   113
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   110
      Left            =   2520
      TabIndex        =   112
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   109
      Left            =   2040
      TabIndex        =   111
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   108
      Left            =   1560
      TabIndex        =   110
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   107
      Left            =   1080
      TabIndex        =   109
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   106
      Left            =   600
      TabIndex        =   108
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   105
      Left            =   120
      TabIndex        =   107
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   104
      Left            =   6840
      TabIndex        =   106
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   103
      Left            =   6360
      TabIndex        =   105
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   102
      Left            =   5880
      TabIndex        =   104
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   101
      Left            =   5400
      TabIndex        =   103
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   100
      Left            =   4920
      TabIndex        =   102
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   99
      Left            =   4440
      TabIndex        =   101
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   98
      Left            =   3960
      TabIndex        =   100
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   97
      Left            =   3480
      TabIndex        =   99
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   96
      Left            =   3000
      TabIndex        =   98
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   95
      Left            =   2520
      TabIndex        =   97
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   94
      Left            =   2040
      TabIndex        =   96
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   93
      Left            =   1560
      TabIndex        =   95
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   92
      Left            =   1080
      TabIndex        =   94
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   91
      Left            =   600
      TabIndex        =   93
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   90
      Left            =   120
      TabIndex        =   92
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   89
      Left            =   6840
      TabIndex        =   91
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   88
      Left            =   6360
      TabIndex        =   90
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   87
      Left            =   5880
      TabIndex        =   89
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   86
      Left            =   5400
      TabIndex        =   88
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   85
      Left            =   4920
      TabIndex        =   87
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   84
      Left            =   4440
      TabIndex        =   86
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   83
      Left            =   3960
      TabIndex        =   85
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   82
      Left            =   3480
      TabIndex        =   84
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   81
      Left            =   3000
      TabIndex        =   83
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   80
      Left            =   2520
      TabIndex        =   82
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   79
      Left            =   2040
      TabIndex        =   81
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   78
      Left            =   1560
      TabIndex        =   80
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   77
      Left            =   1080
      TabIndex        =   79
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   76
      Left            =   600
      TabIndex        =   78
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   75
      Left            =   120
      TabIndex        =   77
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   74
      Left            =   6840
      TabIndex        =   76
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   73
      Left            =   6360
      TabIndex        =   75
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   72
      Left            =   5880
      TabIndex        =   74
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   71
      Left            =   5400
      TabIndex        =   73
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   70
      Left            =   4920
      TabIndex        =   72
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   69
      Left            =   4440
      TabIndex        =   71
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   68
      Left            =   3960
      TabIndex        =   70
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   67
      Left            =   3480
      TabIndex        =   69
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   66
      Left            =   3000
      TabIndex        =   68
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   65
      Left            =   2520
      TabIndex        =   67
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   64
      Left            =   2040
      TabIndex        =   66
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   63
      Left            =   1560
      TabIndex        =   65
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   62
      Left            =   1080
      TabIndex        =   64
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   61
      Left            =   600
      TabIndex        =   63
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   60
      Left            =   120
      TabIndex        =   62
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   59
      Left            =   6840
      TabIndex        =   61
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   58
      Left            =   6360
      TabIndex        =   60
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   57
      Left            =   5880
      TabIndex        =   59
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   56
      Left            =   5400
      TabIndex        =   58
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   55
      Left            =   4920
      TabIndex        =   57
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   54
      Left            =   4440
      TabIndex        =   56
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   53
      Left            =   3960
      TabIndex        =   55
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   52
      Left            =   3480
      TabIndex        =   54
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   51
      Left            =   3000
      TabIndex        =   53
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   50
      Left            =   2520
      TabIndex        =   52
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   49
      Left            =   2040
      TabIndex        =   51
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   48
      Left            =   1560
      TabIndex        =   50
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   47
      Left            =   1080
      TabIndex        =   49
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   46
      Left            =   600
      TabIndex        =   48
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   45
      Left            =   120
      TabIndex        =   47
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   44
      Left            =   6840
      TabIndex        =   46
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   43
      Left            =   6360
      TabIndex        =   45
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   42
      Left            =   5880
      TabIndex        =   44
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   41
      Left            =   5400
      TabIndex        =   43
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   40
      Left            =   4920
      TabIndex        =   42
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   39
      Left            =   4440
      TabIndex        =   41
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   38
      Left            =   3960
      TabIndex        =   40
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   37
      Left            =   3480
      TabIndex        =   39
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   36
      Left            =   3000
      TabIndex        =   38
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   35
      Left            =   2520
      TabIndex        =   37
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   34
      Left            =   2040
      TabIndex        =   36
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   33
      Left            =   1560
      TabIndex        =   35
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   32
      Left            =   1080
      TabIndex        =   34
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   31
      Left            =   600
      TabIndex        =   33
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   30
      Left            =   120
      TabIndex        =   32
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   29
      Left            =   6840
      TabIndex        =   31
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   28
      Left            =   6360
      TabIndex        =   30
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   27
      Left            =   5880
      TabIndex        =   29
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   26
      Left            =   5400
      TabIndex        =   28
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   25
      Left            =   4920
      TabIndex        =   27
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   24
      Left            =   4440
      TabIndex        =   26
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   23
      Left            =   3960
      TabIndex        =   25
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   22
      Left            =   3480
      TabIndex        =   24
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   21
      Left            =   3000
      TabIndex        =   23
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   20
      Left            =   2520
      TabIndex        =   22
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   19
      Left            =   2040
      TabIndex        =   21
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   18
      Left            =   1560
      TabIndex        =   20
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   17
      Left            =   1080
      TabIndex        =   19
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   16
      Left            =   600
      TabIndex        =   18
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   14
      Left            =   6840
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   13
      Left            =   6360
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   12
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   11
      Left            =   5400
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   10
      Left            =   4920
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   9
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   8
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   7
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   6
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   4
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  '置中對齊
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 15, 1 To 15) As String
Dim o As Boolean
Dim xx(1 To 15) As Integer
Dim yy(1 To 15) As Integer
Dim zz As Integer

Private Sub Command1_Click()
Dim s As String
Dim p As Boolean
For i = 0 To 224
  Label(i).BackColor = RGB(255, 255, 255)
Next i
For i = 1 To 15
  For j = 1 To 15
    '-----------------
    s = a(j, i)
    p = True
    For k = j + 1 To 15
      s = s & a(k, i)
      If a(j, i) = a(k, i) Then
        Call qqq(s, p)
        If p = True Then
          For l = 1 To Len(s)
            Label(j + l - 2 + (i - 1) * 15).BackColor = RGB(0, 0, 255)
          Next l
        End If
      End If
    Next k
    s = a(j, i)
    p = True
    For k = i + 1 To 15
      s = s & a(j, k)
      If a(j, i) = a(j, k) Then
        Call qqq(s, p)
        If p = True Then
          For l = 1 To Len(s)
            Label(j - 1 + (i + l - 2) * 15).BackColor = RGB(0, 0, 255)
          Next l
        End If
      End If
    Next k
    s = a(j, i)
    p = True
    c = 15 - j
    If 15 - i < c Then c = 15 - i
    For k = 1 To c
      s = s & a(j + k, i + k)
      If a(j, i) = a(j + k, i + k) Then
        Call qqq(s, p)
        If p = True Then
          For l = 0 To Len(s) - 1
            Label(j + l - 1 + (i + l - 1) * 15).BackColor = RGB(0, 0, 255)
          Next l
        End If
      End If
    Next k
    s = a(j, i)
    p = True
    c = j - 1
    If 15 - i < c Then c = 15 - i
    For k = 1 To c
      s = s & a(j - k, i + k)
      If a(j, i) = a(j - k, i + k) Then
        Call qqq(s, p)
        If p = True Then
          For l = 0 To Len(s) - 1
            Label(j - l - 1 + (i + l - 1) * 15).BackColor = RGB(0, 0, 255)
          Next l
        End If
      End If
    Next k
  Next j
Next i
End Sub

Private Sub Command2_Click()
Dim b As String
b = InputBox("搜尋字串", "搜尋")
o = False
zz = 0
Erase xx, yy
For i = 0 To 224
  Label(i).BackColor = RGB(255, 255, 255)
Next i
For i = 1 To 15
  For j = 1 To 15
    If o = False Then Call ppp(j, i, b, 0)
  Next j
Next i
If o = True Then
  For i = 1 To 15
    For j = 1 To 15
      For k = 1 To zz
        If xx(k) = j And yy(k) = i Then
          Label(j + (i - 1) * 15 - 1).BackColor = RGB(255, 0, 0)
          Exit For
        End If
      Next k
    Next j
  Next i
End If
End Sub

Private Sub Form_Activate()
Dim b As String
Open "data.txt" For Input As #1
Do Until EOF(1)
  c = c + 1
  Line Input #1, b
  For i = 1 To Len(b)
    a(i, c) = Mid(b, i, 1)
    Label(i + (c - 1) * 15 - 1) = a(i, c)
  Next i
Loop
For i = 0 To 224
  Label(i).BackColor = RGB(255, 255, 255)
Next i
End Sub

Private Sub ppp(ByVal x, ByVal y, ByVal ans, ByVal z)
If x >= 1 And y >= 1 And x <= 15 And y <= 15 Then
  If Len(ans) > 0 Then
    If a(x, y) = Left(ans, 1) Then
      zz = zz + 1
      xx(zz) = x
      yy(zz) = y
      If o = False And (z = 0 Or z = 1) Then Call ppp(x + 1, y, Right(ans, Len(ans) - 1), 1)
      If o = False And (z = 0 Or z = 2) Then Call ppp(x - 1, y, Right(ans, Len(ans) - 1), 2)
      If o = False And (z = 0 Or z = 3) Then Call ppp(x, y + 1, Right(ans, Len(ans) - 1), 3)
      If o = False And (z = 0 Or z = 4) Then Call ppp(x, y - 1, Right(ans, Len(ans) - 1), 4)
      If o = False And (z = 0 Or z = 5) Then Call ppp(x + 1, y + 1, Right(ans, Len(ans) - 1), 5)
      If o = False And (z = 0 Or z = 6) Then Call ppp(x + 1, y - 1, Right(ans, Len(ans) - 1), 6)
      If o = False And (z = 0 Or z = 7) Then Call ppp(x - 1, y + 1, Right(ans, Len(ans) - 1), 7)
      If o = False And (z = 0 Or z = 8) Then Call ppp(x - 1, y - 1, Right(ans, Len(ans) - 1), 8)
    End If
    If o = False And z > 0 Then
      zz = 1
    ElseIf o = False And z = 0 Then
      zz = 0
      Erase xx, yy
    End If
  Else
    o = True
  End If
Else
  If o = False And z > 0 Then
    zz = 1
  ElseIf o = False And z = 0 Then
    zz = 0
    Erase xx, yy
  End If
End If
End Sub

Private Sub qqq(s, p As Boolean)
For l = 1 To Len(s) \ 2
  If Mid(s, l, 1) <> Mid(s, Len(s) - l + 1, 1) Or Len(s) < 4 Then
    p = False
    Exit For
  End If
Next l
End Sub
