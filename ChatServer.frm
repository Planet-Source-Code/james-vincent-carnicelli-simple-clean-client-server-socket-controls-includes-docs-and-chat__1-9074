VERSION 5.00
Object = "{27BB2290-4631-11D4-8F1E-4000500C1033}#25.0#0"; "Sockets.ocx"
Begin VB.Form ChatServer 
   Caption         =   "Sockets Demo - Chat Server"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DisconnectAll 
      Caption         =   "Disconnect All"
      Height          =   375
      Left            =   4260
      TabIndex        =   2
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton Listen 
      Caption         =   "Stop Listening"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2595
   End
   Begin Sockets.ServerSocketBank Sockets 
      Left            =   0
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.TextBox History 
      Height          =   3435
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   5835
   End
End
Attribute VB_Name = "ChatServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DisconnectAll_Click()
    Sockets.DisconnectAll
End Sub

Private Sub Form_Load()
    Dim Port As Integer
    Port = 100
    Port = InputBox("Port to listen on", , Port)
    Sockets.Listen Port
    History.SelText = "Listening for connections on port " & Port & vbCrLf
End Sub

Private Sub Listen_Click()
    If Sockets.Listening Then
        Sockets.StopListening
        Listen.Caption = "Listen"
        History.SelText = "Stopped listening" & vbCrLf
    Else
        Sockets.Listen
        Listen.Caption = "Stop Listening"
        History.SelText = "Started listening" & vbCrLf
    End If
End Sub

Private Sub Sockets_Connected(Index As Integer, Socket As Sockets.ServerSocket)
    Dim Tag As Object
    History.SelText = "Connected to " & Index & " from " & Socket.Host & _
      " at " & Socket.ExtraTag.Item("ConnectTime") & vbCrLf
    Sockets.Broadcast Index & " has connected from " & Socket.Host & "."
End Sub

Private Sub Sockets_ConnectionRequest(requestID As Long, FromHost As String, Cancel As Boolean, NewTag As Variant)
    Set NewTag = New Collection
    NewTag.Add Now, "ConnectTime"
End Sub

Private Sub Sockets_DataArrival(Index As Integer, Socket As Sockets.ServerSocket, Bytes As Long)
    Dim Text As String
    Text = Socket.Receive
    History.SelText = Index & "> " & Text & vbCrLf
    Sockets.Broadcast Index & "> " & Text
End Sub

Private Sub Sockets_Disconnect(Index As Integer, Socket As Sockets.ServerSocket)
    Dim Tag As Object
    History.SelText = "Disconnected from " & Index & _
      ", connected for " & DateDiff("s", Socket.ExtraTag.Item("ConnectTime"), Now) & " seconds" & vbCrLf
    Sockets.Broadcast Index & " has disconnected."
End Sub
