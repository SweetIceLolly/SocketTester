VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Socket Tester"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   13305
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuControl 
      Caption         =   "&Control"
      Begin VB.Menu mnuNewClient 
         Caption         =   "New Event (&Client Mode)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuNewServer 
         Caption         =   "New Event (&Server Mode)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuNewUDP 
         Caption         =   "New Event (&UDP Mode)"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuTerminateAllEvents 
         Caption         =   "&Terminate All Events"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEvents As Integer

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim a As VbMsgBoxResult
    a = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Exit")
    If a = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNewClient_Click()
    Dim Client As New frmClient
    Load Client
    iEvents = iEvents + 1
    Client.Caption = "Event " & iEvents & " - Client Mode"
End Sub

Private Sub mnuNewServer_Click()
    Dim Server As New frmServer
    Load Server
    iEvents = iEvents + 1
    Server.Caption = "Event " & iEvents & " - Server Mode"
End Sub

Private Sub mnuNewUDP_Click()
    Dim UDP As New frmUDP
    Load UDP
    iEvents = iEvents + 1
    UDP.Caption = "Event " & iEvents & " - UDP Mode"
End Sub

Private Sub mnuTerminateAllEvents_Click()
    Dim Window As Form
    For Each Window In Forms
        If Not Window Is frmMain Then
            Unload Window
        End If
    Next Window
End Sub
