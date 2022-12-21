VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "即时通讯-服务端"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14175
   Icon            =   "聊吧-服务端.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   14175
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox LineIn 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   9360
      Width           =   12615
   End
   Begin VB.CommandButton Send 
      BackColor       =   &H0000FF00&
      Caption         =   "发送(Enter)"
      Height          =   495
      Left            =   12840
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "设置"
      Height          =   495
      Left            =   12960
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "配对"
      Height          =   495
      Left            =   13560
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Port 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   1
      Text            =   "1000"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Chat 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   13935
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   13560
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   25272
      LocalPort       =   25272
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "基岩即时通讯"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   13935
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "端口号(与客户端一致)："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock.LocalPort = Port.Text
End Sub

Private Sub Command2_Click()
Winsock.Listen
End Sub

Private Sub Form_Load()
Title.Caption = "基岩即时通讯"
End Sub
Private Sub LineIn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call Winsock.SendData(LineIn.Text)
Chat.Text = Chat.Text & "我：" & LineIn.Text + vbCrLf + ""
LineIn.Text = ""
End If
End Sub
Private Sub Send_Click()
Call Winsock.SendData(LineIn.Text)
LineIn.SelStart = Len(LineIn.Text)
Chat.Text = Chat.Text & "我：" & LineIn.Text + vbCrLf + ""
LineIn.Text = ""
End Sub


Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
If Winsock.State <> sckClosed Then
    Winsock.Close
End If
Call Winsock.Accept(requestID)
End Sub
Private Sub Winsock_Close()
Winsock.Close
Winsock.Listen
End Sub
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim message As String
Call Winsock.GetData(message)
Chat.Text = Chat.Text & "对方：" & message + vbCrLf + ""
End Sub
Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock.Close
Winsock.Listen
End Sub

