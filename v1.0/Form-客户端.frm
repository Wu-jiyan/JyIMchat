VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "即时通讯-客户端"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11760
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000018&
      Caption         =   "连接"
      Height          =   375
      Left            =   11280
      MaskColor       =   &H80000018&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Text            =   "1000"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000018&
      Caption         =   "设置"
      Height          =   375
      Left            =   10680
      MaskColor       =   &H80000018&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      ToolTipText     =   "输入服务端(对方)ip地址"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "发送"
      Height          =   495
      Left            =   10200
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "发送当前的文本"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "等线"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   9975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   ""
      RemotePort      =   1000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "端口号(和服务端一致)"
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "服务端(对方)ip地址"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   0
      Top             =   6720
      Width           =   11775
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   11775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000018&
      FillStyle       =   0  'Solid
      Height          =   6855
      Left            =   0
      Top             =   600
      Width           =   11775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "基岩即时通讯"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call Winsock1.SendData(Text2.Text)
Text2.SelStart = Len(Text2.Text)
Text1.Caption = Text1.Caption & "我：" & Text2.Text + vbCrLf + " "
Text2.Text = ""
End Sub

Private Sub Command2_Click()
Winsock1.RemoteHost = Text3.Text
Winsock1.RemotePort = Text4.Text
End Sub


Private Sub Command3_Click()
Winsock1.Connect
End Sub

Private Sub Form_Terminate()
Winsock1.Close
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim message As String
Call Winsock1.GetData(message)
Text1.Caption = Text1.Caption & "对方：" & message + vbCrLf + " "
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Connect
End Sub
