VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "CODEJO~4.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event 1 - Server Mode"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   7605
   Begin MSWinsockLib.Winsock wsMain 
      Left            =   840
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame fraDataReceived 
      Caption         =   "Data Received (Total: 0, 0B)"
      Height          =   3375
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   7335
      Begin VB.ListBox lstDataReceived 
         Height          =   2400
         ItemData        =   "frmServer.frx":0000
         Left            =   120
         List            =   "frmServer.frx":0002
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkAutoScroll 
         Caption         =   "Auto Scroll"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy To Send Data"
         Height          =   375
         Left            =   4320
         TabIndex        =   9
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save File"
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   2880
         Width           =   1095
      End
      Begin XtremeSuiteControls.TabControl tabViewData 
         Height          =   2415
         Left            =   2040
         TabIndex        =   16
         Top             =   360
         Width           =   5175
         _Version        =   983043
         _ExtentX        =   9128
         _ExtentY        =   4260
         _StockProps     =   68
         ItemCount       =   2
         Item(0).Caption =   "Text View"
         Item(0).Tooltip =   "Text View"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "edTextView"
         Item(1).Caption =   "Binary View"
         Item(1).Tooltip =   "Binary View"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "edBinaryView"
         Begin VB.TextBox edTextView 
            Height          =   1935
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   360
            Width           =   4935
         End
         Begin XtremeSuiteControls.HexEdit edBinaryView 
            Height          =   1935
            Left            =   -69880
            TabIndex        =   17
            Top             =   360
            Visible         =   0   'False
            Width           =   4935
            _Version        =   983043
            _ExtentX        =   8705
            _ExtentY        =   3413
            _StockProps     =   73
            WideAddress     =   0   'False
            ReadOnly        =   -1  'True
         End
      End
      Begin VB.Label labPacketSize 
         AutoSize        =   -1  'True
         Caption         =   "Packet No. 0: Size = 0B"
         Height          =   195
         Left            =   2040
         TabIndex        =   18
         Top             =   2880
         Width           =   1710
      End
   End
   Begin VB.Frame fraSendData 
      Caption         =   "Send Data"
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   7335
      Begin VB.CommandButton cmdSendData 
         Caption         =   "Send"
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin XtremeSuiteControls.TabControl tabSendData 
         Height          =   1215
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6135
         _Version        =   983043
         _ExtentX        =   10821
         _ExtentY        =   2143
         _StockProps     =   68
         ItemCount       =   2
         Item(0).Caption =   "Text Mode"
         Item(0).Tooltip =   "Text Mode"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "edSendText"
         Item(1).Caption =   "Binary Mode"
         Item(1).Tooltip =   "Binary Mode"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "edSendBinary"
         Begin VB.TextBox edSendText 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   360
            Width           =   5895
         End
         Begin XtremeSuiteControls.HexEdit edSendBinary 
            Height          =   735
            Left            =   -69880
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   5895
            _Version        =   983043
            _ExtentX        =   10398
            _ExtentY        =   1296
            _StockProps     =   73
            WideAddress     =   0   'False
         End
      End
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   1000
      Left            =   1440
      Top             =   480
   End
   Begin VB.CommandButton cmdStartServer 
      Caption         =   "Start"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox edServerPort 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CDL 
      Left            =   240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label labTraffic 
      Caption         =   "Traffic: 0B/s, 0B in total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      TabIndex        =   19
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label labState 
      AutoSize        =   -1  'True
      Caption         =   "Current State: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   20
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      Caption         =   "Server Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrevDataSize    As Long
Dim DataSize        As Long
Dim RecvDataSize    As Long
Dim DataPos()       As Long
Dim RecvBuffer()    As Byte
Dim PacketSize()    As Long
Dim DataIndex       As Long

Private Function SizeWithFormat(lBytes) As String
    Select Case lBytes
        Case Is < 1024
           SizeWithFormat = lBytes & "B"
        
        Case Is < 1024 ^ 2
            SizeWithFormat = Format(lBytes / 1024, "0.00") & "KB"
        
        Case Is < 1024 ^ 3
            SizeWithFormat = Format(lBytes / (1024 ^ 2), "0.00") & "MB"
        
        Case Is < 1024 ^ 4
            SizeWithFormat = Format(lBytes / (1024 ^ 3), "0.00") & "GB"
        
    End Select
End Function

Private Sub cmdClose_Click()
    Me.wsMain.Close
End Sub

Private Sub cmdStartServer_Click()
    On Error Resume Next
    Me.wsMain.Close
    Me.wsMain.Bind Me.edServerPort.Text
    Me.wsMain.Listen
End Sub

Private Sub edSendText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdSendData_Click
    End If
End Sub

Private Sub edServerPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdStartServer_Click
    End If
End Sub

Private Sub wsMain_Close()
    Me.labState.Caption = "Current State: " & Me.wsMain.state
End Sub

Private Sub tmrRefresh_Timer()
    Me.labState.Caption = "Current State: " & Me.wsMain.state
    Me.labTraffic.Caption = "Traffic: " & SizeWithFormat(DataSize - PrevDataSize) & _
        "/s, " & SizeWithFormat(DataSize) & " in total"
    PrevDataSize = DataSize
End Sub

Private Sub wsMain_ConnectionRequest(ByVal requestID As Long)
    Me.wsMain.Close
    Me.wsMain.Accept requestID
    Me.labState.Caption = "Current State: " & Me.wsMain.state
End Sub

Private Sub wsMain_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim temp() As Byte
    
    Me.wsMain.GetData temp
    
    RecvDataSize = RecvDataSize + UBound(temp) + 1
    DataSize = DataSize + UBound(temp) + 1
    
    DataPos(DataIndex) = UBound(RecvBuffer)
    PacketSize(DataIndex) = UBound(temp)
    ReDim Preserve RecvBuffer(UBound(RecvBuffer) + UBound(temp) + 1)
    ReDim Preserve DataPos(UBound(DataPos) + 1)
    ReDim Preserve PacketSize(UBound(PacketSize) + 1)
    CopyMemory RecvBuffer(DataPos(DataIndex)), temp(0), ByVal (UBound(temp) + 1)
    DataIndex = DataIndex + 1
    
    If Me.chkAutoScroll.Value = 1 Then
        Me.lstDataReceived.AddItem StrConv(temp, vbUnicode)
        Me.edBinaryView.SetData temp
        Me.lstDataReceived.ListIndex = DataIndex - 1
    End If
    Me.fraDataReceived.Caption = "Data Received (Total: " & DataIndex & _
        ", " & SizeWithFormat(RecvDataSize) & ")"
End Sub

Private Sub Form_Load()
    ReDim RecvBuffer(0)
    ReDim PacketSize(0)
    ReDim DataPos(0)
    DataIndex = 0
End Sub

Private Sub cmdCopy_Click()
    Me.edSendText.Text = Me.edTextView.Text
    Me.edSendBinary.SetData Me.edBinaryView.GetData
End Sub

Private Sub cmdLoad_Click()
    On Error Resume Next
    Dim tmp As String, fString As String
    Dim temp() As Byte
    
    Me.CDL.Filter = "All Files(*.*)|*"
    If Me.tabSendData.Selected.Index = 0 Then
        Me.CDL.DialogTitle = "Load text file"
        Me.CDL.ShowOpen
        If Me.CDL.FileName = "" Or Err.Number <> 0 Then
            Exit Sub
        End If
        Open Me.CDL.FileName For Input As #1
            Do While Not EOF(1)
                Input #1, tmp
                fString = fString & tmp
            Loop
        Close #1
        Me.edSendText.Text = fString
    Else
        Me.CDL.DialogTitle = "Load binary file"
        Me.CDL.ShowOpen
        If Me.CDL.FileName = "" Or Err.Number <> 0 Then
            Exit Sub
        End If
        Open Me.CDL.FileName For Binary As #1
            ReDim temp(LOF(1))
            Get #1, , temp
        Close #1
        Me.edSendBinary.SetData temp
    End If
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    
    Me.CDL.Filter = "All Files(*.*)|*"
    If Me.tabViewData.Selected.Index = 0 Then
        Me.CDL.DialogTitle = "Save file as text"
        Me.CDL.ShowSave
        If Me.CDL.FileName = "" Or Err.Number <> 0 Then
            Exit Sub
        End If
        Open Me.CDL.FileName For Output As #1
            Print #1, Me.edTextView.Text
        Close #1
    Else
        Me.CDL.DialogTitle = "Save file as binary"
        Me.CDL.ShowSave
        If Me.CDL.FileName = "" Or Err.Number <> 0 Then
            Exit Sub
        End If
        
        Dim temp() As Byte
        temp = Me.edBinaryView.GetData
        Open Me.CDL.FileName For Binary As #1
            Put #1, , temp
        Close #1
    End If
End Sub

Private Sub cmdSendData_Click()
    On Error Resume Next
    If Me.tabSendData.Selected.Index = 0 Then
        Me.wsMain.SendData Me.edSendText.Text
        If Err.Number = 0 Then
            Me.edSendText.Text = ""
            Me.edSendText.SetFocus
        End If
    Else
        Me.wsMain.SendData Me.edSendBinary.GetData
        If Err.Number = 0 Then
            Dim t(0) As Byte
            Me.edSendBinary.SetData t
            Me.edSendBinary.SetFocus
        End If
    End If
End Sub

Private Sub lstDataReceived_Click()
    Dim temp()  As Byte
    Dim iData   As Long
    
    iData = Me.lstDataReceived.ListIndex
    ReDim temp(PacketSize(iData))
    CopyMemory temp(0), RecvBuffer(DataPos(iData)), ByVal (PacketSize(iData) + 1)
    Me.edTextView.Text = StrConv(temp, vbUnicode)
    Me.edBinaryView.SetData temp
    Me.labPacketSize.Caption = "Packet No. " & iData + 1 & ": Size = " & SizeWithFormat(PacketSize(iData) + 1)
End Sub

Private Sub wsMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Me.labState.Caption = "Current State: " & Me.wsMain.state
End Sub

Private Sub wsMain_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    DataSize = DataSize + bytesSent
End Sub
