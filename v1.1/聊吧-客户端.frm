VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʱͨѶ-�ͻ���"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13590
   BeginProperty Font 
      Name            =   "����"
      Size            =   7.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "�İ�-�ͻ���.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13590
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox IP 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   9
      Text            =   "219.150.218.20"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Port 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Left            =   11400
      TabIndex        =   4
      Text            =   "25272"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Chat 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   2
      Top             =   720
      Width           =   13335
   End
   Begin VB.CommandButton Send 
      BackColor       =   &H0080FF80&
      Caption         =   "����(Enter)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox LineIn 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   9360
      Width           =   12015
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   12600
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�Է�ip��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�Է��˿ڣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���Ҽ�ʱͨѶ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Winsock.RemotePort = Port.Text
Winsock.RemoteHost = IP.Text
End Sub

Private Sub Command2_Click()
Winsock.Connect
End Sub

Private Sub Form_Load()
Title.Caption = "���Ҽ�ʱͨѶ"
End Sub

Private Sub LineIn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call Winsock.SendData(LineIn.Text)
LineIn.SelStart = Len(LineIn.Text)
Chat.Text = Chat.Text & "�ң�" & LineIn.Text + vbCrLf + ""
LineIn.Text = ""
End If
End Sub

Private Sub Send_Click()
Call Winsock.SendData(LineIn.Text)
LineIn.SelStart = Len(LineIn.Text)
Chat.Text = Chat.Text & "�ң�" & LineIn.Text + vbCrLf + ""
LineIn.Text = ""
End Sub

Private Sub Winsock_Close()
Winsock.Close
End Sub
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim message As String
Call Winsock.GetData(message)
Chat.Text = Chat.Text & "�Է���" & message + vbCrLf + ""
LineIn.SelStart = Len(LineIn.Text)
End Sub
Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock.Close
End Sub
