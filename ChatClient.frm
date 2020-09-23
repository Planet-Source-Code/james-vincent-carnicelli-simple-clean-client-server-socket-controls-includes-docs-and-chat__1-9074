VERSION 5.00
Object = "{27BB2290-4631-11D4-8F1E-4000500C1033}#25.0#0"; "Sockets.ocx"
Begin VB.Form ChatClient 
   Caption         =   "Sockets Demo - Chat Client"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton AnotherClient 
      Caption         =   "Launch Another Client"
      Height          =   315
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   2475
   End
   Begin VB.CommandButton StartServer 
      Caption         =   "Start Server"
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton Send 
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   5700
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox ToSend 
      Enabled         =   0   'False
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Text            =   "Hi, there!"
      Top             =   3600
      Width           =   5535
   End
   Begin VB.CommandButton Connect 
      Caption         =   "Connect"
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox Port 
      Height          =   285
      Left            =   4860
      TabIndex        =   4
      Text            =   "100"
      Top             =   540
      Width           =   435
   End
   Begin VB.TextBox Host 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   540
      Width           =   3735
   End
   Begin Sockets.ClientSocket Socket 
      Left            =   120
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.TextBox History 
      Enabled         =   0   'False
      Height          =   2655
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   900
      Width           =   6795
   End
   Begin VB.Label Label 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4740
      TabIndex        =   3
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "Host/Port"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "ChatClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AnotherClient_Click()
    Dim NewClient As ChatClient
    Set NewClient = New ChatClient
    NewClient.Show
End Sub

Private Sub Connect_Click()
    If Socket.Connected Then
        Socket.Disconnect
        Connect.Caption = "Connect"
        History.Enabled = False
        ToSend.Enabled = False
        Send.Enabled = False
    Else
        On Error Resume Next
        Socket.Connect Host.Text, Port.Text
        If Err.Number <> 0 Then
            MsgBox Err.Description
        Else
            Connect.Caption = "Disconnect"
            History.Enabled = True
            ToSend.Enabled = True
            Send.Enabled = True
            ToSend.SetFocus
        End If
    End If
End Sub

Private Sub Send_Click()
    Socket.Send ToSend.Text
    ToSend.Text = ""
    ToSend.SetFocus
End Sub

Private Sub Socket_DataArrival(Bytes As Long)
    History.SelText = Socket.Receive & vbCrLf
End Sub

Private Sub Socket_Disconnect()
    Connect.Caption = "Connect"
    History.Enabled = False
    ToSend.Enabled = False
    Send.Enabled = False
End Sub

Private Sub StartServer_Click()
    ChatServer.Show
End Sub
