VERSION 5.00
Begin VB.Form frmclient 
   Caption         =   "Client - Lizard Soft Programming CHAT 0.1"
   ClientHeight    =   3315
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6870
   ControlBox      =   0   'False
   Icon            =   "frmclient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1560
      Top             =   960
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Text            =   "Message Here"
      Top             =   2640
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtmsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Text            =   "Address Here"
      Top             =   3000
      Width           =   5895
   End
   Begin VB.PictureBox Winsock1 
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Menu mnumin 
      Caption         =   "&Minimize"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmclient"
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
'Set winsock's remote host to text1.text
Winsock1.RemoteHost = Text1.Text
'Attempt a connection
Winsock1.Connect
End Sub

Private Sub Command2_Click()
'Set text2 to current focus, so you don't have to click on it
Text2.SetFocus
'Make the message format, which is the header, the message, and then a blank line
msg = header + Text2.Text + vbCrLf
'Tell winsock1 to send the message to the remote computer
Winsock1.SendData msg
'Put the text of the message you just sent to the text window
txtmsg.Text = txtmsg.Text + msg + vbCrLf
'Clear the message textbox
Text2.Text = ""
End Sub

Private Sub Form_Load()
'Initilize and find the header and the remote port
alrm = 0
header = InputBox("Please enter your 'header'", "Enter")
Winsock1.RemotePort = InputBox("Enter the remote port.", "Remote Port")
End Sub

Private Sub mnuexit_Click()
'Quit
Unload frmclient
End Sub

Private Sub mnumin_Click()
'Minimize the window and popup if there is an incoming message
'See timer1 code
Timer1.Enabled = True
frmclient.WindowState = 1
End Sub

Private Sub Timer1_Timer()
'If a message is recived timer
If alrm = 1 Then
    'Maximize window
    frmclient.WindowState = 0
    'Reset alrm's value
    alrm = 0
    'Disable the timer
    Timer1.Enabled = False
End If
End Sub

Private Sub txtmsg_DblClick()
'When you double click on the main message text box it clears
txtmsg.Text = ""
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'When data is recived make alrm = 1
alrm = 1
'The follwing puts the incoming data into MSG
Winsock1.GetData msg
'Post the msg text to the textbox
txtmsg.Text = txtmsg.Text + msg
End Sub

