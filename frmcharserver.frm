VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmchatserver 
   Caption         =   "Server, Lizard Soft Programming CHAT"
   ClientHeight    =   4275
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmcharserver.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   360
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2640
      Top             =   1920
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   0
      Width           =   6495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SEND"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Text            =   "Type Message Here!"
      Top             =   3720
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Menu mnumin 
      Caption         =   "&Minimize"
   End
   Begin VB.Menu mnuclose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "frmchatserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dan Koestler
'Lizard Soft Programming
'Free for PlanetSourceCode
'Any comments email DKoestler@lsoft.8m.com

'Alrm is used when a message is recived and the form is minimized
Dim alrm As Integer
'Header is what appears before your message
Dim header As String
'msg is the actuall message text
Dim msg As String

Private Sub Command1_Click()
'The following line gets the localport that the remote computer will connect to
Winsock1.LocalPort = InputBox("Please enter the local port.", "LocalPort")
'Winsock1 needs to listen, when a connection attempt is recived winsock1.
Winsock1.Listen
End Sub

Private Sub Command2_Click()
'Set the focus on the text
text1.SetFocus
'Set the msg format
msg = header + text1.Text + vbCrLf
'In a server you need two winsock controls, one to recive the connection attempt and the other
'to send messages, recive messages and accept the connection attempt.
'This command sends msg to the connected computer
Winsock2.SendData msg
'Display the message you just sent to the remote computer on this computer
Text2.Text = Text2.Text + msg + vbCrLf
'Clear the message textbox
text1.Text = ""
End Sub

Private Sub Form_Load()
'Initilize alrm, which is used for poping up the form when a message is recived
alrm = 0
'Get the header, which appears before the text on both computers.
header = InputBox("Please enter your 'header'", "Header")
End Sub

Private Sub mnuclose_Click()
'Quit
Unload Me
End Sub

Private Sub mnumin_Click()
'When you minimize the form enable the timer to see if a alrm = 1
Timer1.Enabled = True
'Minimze the form
frmchatserver.WindowState = 1
End Sub

Private Sub Text2_DblClick()
'When you double click on the main text box clear it
Text2.Text = ""
End Sub

Private Sub Timer1_Timer()
'If a message is recived
If alrm = 1 Then
    'Disable the timer
    Timer1.Enabled = False
    'Make the window go to normal size
    frmchatserver.WindowState = 0
    'Reset alrm to 0
    alrm = 0
End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'If there is a connection request make WINSOCK 2 (very important that it is 2) accept the request
'therefore connecting the two computers
Winsock2.Accept requestID
End Sub


Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
'If data is recived:
'Make alrm = 1
alrm = 1
'Make winsock2 (the winsock that is connected to the remote computer) get the data and put it under MSG
Winsock2.GetData msg
'Make the textbox equal MSG
Text2.Text = Text2.Text + msg
End Sub

