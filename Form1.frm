VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "send"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Type something to send!"
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   1440
      Width           =   6495
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   5640
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Login"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   4080
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listen"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "21"
      Top             =   3000
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":0010
      Left            =   0
      List            =   "Form1.frx":0012
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "Test to see if it works by login in to yourself! Also, send a message or something and you will see"
      Height          =   435
      Left            =   1920
      TabIndex        =   9
      Top             =   3840
      Width           =   3885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Listen on port"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Server v1.0
'Author: Dustin Davis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'
'This is to show how to accept multiple logins from on socket!
'Please do not steal this code.
'Enjoy!
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public i As Long 'counter for the winsock control

Private Sub Command1_Click()
'Listens on a port
Winsock1.LocalPort = Text1.Text 'set the port
Winsock1.Listen 'tell it to listen
Command1.Enabled = False
Command2.Enabled = True
List1.AddItem "Listening on port: " & Text1.Text
End Sub

Private Sub Command2_Click()
'disconnect
Winsock1.Close 'close port
Command2.Enabled = False
Command1.Enabled = True
List1.AddItem "Disconnected"
End Sub

Private Sub Command3_Click()
'this is the button that will login to yourself
Winsock3.Close 'make sure it isnt open cuz of errors
Winsock3.RemoteHost = Winsock3.LocalIP 'this sets the remotehost to you
Winsock3.RemotePort = 21 'port to login to, the port that winsock1 is watching
Winsock3.Connect 'connect
End Sub

Private Sub Command4_Click()
Winsock3.SendData Text3.Text 'this will send data
End Sub

Private Sub Form_Load()
Winsock1.Close 'makes sure it isnt already open cuz of errors
i = 0 'set this to zero
End Sub

Private Sub Winsock1_Close()
'what to do when the socket closes
List1.AddItem "Winsock 1 Closed"
Winsock1.Close
Command2.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Winsock1_Connect()
'it should not be connected! only to listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'this is the big cheese! When someone tries to login, it will open a diffrent
'winsock and accept the connection! That way winsock1 keeps watching port 21
i = i + 1 'adds one to make sure we dont get errors
Load Winsock2(i) 'load a new winsock control
Winsock2(i).Close 'close it cuz of errors
Winsock2(i).Accept requestID 'accept the connection
List1.AddItem "Winsock " & i & " Accepting Connection request: " & requestID
'set up winsock1 for listening again
Winsock1.Close
Winsock1.LocalPort = Text1.Text
Winsock1.Listen
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'There really should not be any errors since winsock1 does not do anything
List1.AddItem "Winsock 1 Err: " & Number & " " & Description
List1.AddItem "Winsock 1 State: " & Winsock1.State
End Sub

Private Sub Winsock2_Close(Index As Integer)
'what to do when a socket closes
List1.AddItem "Winsock " & Index & " Closed"
Winsock2(Index).Close
End Sub

Private Sub Winsock2_Connect(Index As Integer)
'will need to add new stuff if you want, u can connect to them. Just
'creat a new winsock control, load a new object from that control and
'connect to them
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'data arrival, you can change this to do what you want.
Dim Data As String
Winsock2(Index).GetData Data 'gets the data
Text2.Text = Text2.Text + vbCrLf & Data
End Sub

Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'what to do if an error occurs
List1.AddItem "Winsock " & Index & " Err: " & Number & " " & Description
List1.AddItem "Winsock " & Index & " State: " & Winsock2(Index).State
End Sub

Private Sub Winsock3_Connect()
'tells u it has connected!
List1.AddItem "Winsock 3 connected"
End Sub

