VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmMainServer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "NeoChat V.1.0.0 (TCP Client)"
   ClientHeight    =   4710
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5205
   Icon            =   "frmMainServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainServer.frx":030A
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSend 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2760
      Width           =   4695
   End
   Begin VB.CommandButton cmdOption 
      DownPicture     =   "frmMainServer.frx":15B0
      Height          =   255
      Left            =   3000
      Picture         =   "frmMainServer.frx":1853
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   180
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Index           =   0
      Left            =   3000
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   2520
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Default         =   -1  'True
      DisabledPicture =   "frmMainServer.frx":1AFC
      DownPicture     =   "frmMainServer.frx":1DCC
      Height          =   495
      Left            =   240
      Picture         =   "frmMainServer.frx":20A1
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "frmMainServer.frx":237C
      Height          =   255
      Left            =   4410
      Picture         =   "frmMainServer.frx":24D9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      DownPicture     =   "frmMainServer.frx":263B
      Height          =   255
      Left            =   4680
      Picture         =   "frmMainServer.frx":27CB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   2160
   End
   Begin RichTextLib.RichTextBox txtOutput 
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMainServer.frx":2962
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf1 
      Height          =   855
      Left            =   2040
      TabIndex        =   5
      Top             =   3720
      Width           =   855
      _cx             =   22742500
      _cy             =   22742500
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmMainServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This application is published free to any interested developers. Conditions: _
Images and logos as well as the name Neophile and NeoChat are the registered _
properties of Neophile Digital Solutions. Any improper use or replication thereof _
will result in prosecution. Images were include to demonstrate how a VB app can _
easily be transformed into an interesting and visually stimulating program.
'All code is free to use for study purposes and also feel free to copy it into _
your own projects if you wish. Just remember to give credit where credit is due.
'This was my first attempt at using the Winsock control. If you have any _
sugestions, comments or questions, please email me at ernst@neophile.co.za .
'It would also be my first submision to Planet-Source-Code, so I would _
appreciate your votes. Thanks to everyone for publishing such great source _
code, without which I would truly be lost sometimes, and to Planet-Source-Code _
for providing this platform that benefits all of us strugling developers!

Dim strState As String 'Variable to store winsock control's state
Dim strColor As String 'Color string for text
Dim blBold As Boolean 'Text bold boolean
Dim blItalic As Boolean 'Text Italic boolean

Private intMax As Long 'Variable to count elements in Winsock Control array
Dim i As Integer 'Counter

Private Sub cmdOption_Click()
Timer1.Enabled = False
    frmSvrOptions.Show 'vbModal
    
End Sub

Private Sub cmdSend_Click()

    If tcpClient.State = 7 Then '7 = Connected
        tcpClient.SendData sColor & "~" & bBold & "~" & bItalic & "~" & sUserName & " :->  " & txtSend.Text
      '  txtOutput.Text = "Me:> " & txtSend.Text & vbNewLine & txtOutput.Text
   
        txtSend.Text = ""
        swf1.Play
    End If

End Sub

Private Sub Command1_Click()
    Unload Me
    End
End Sub

Private Sub Command2_Click()
Me.WindowState = 1

End Sub


Private Sub Form_Activate()
'swf1.LoadMovie 1, App.Path & "\graphics\flash\svrtraingle.swf"
'swf1.Movie = App.Path & "\graphics\flash\svrtraingle.swf"
    Me.Caption = "NeoChat V.1.0.0 (TCP Client)"
    txtSend.SetFocus
End Sub



Private Sub Form_DblClick()
    'Me.Top = 100
End Sub

Private Sub Form_GotFocus()
    Me.Caption = "NeoChat V.1.0.0 (TCP Client)"
    txtSend.SetFocus
End Sub

Private Sub Form_Load()
    
       intMax = 0
   tcpServer(0).LocalPort = 1001
   tcpServer(0).Listen
   
    ' The name of the Winsock control is tcpClient.
    ' Note: to specify a remote host, you can use
    ' either the IP address (ex: "121.111.1.1") or
    ' the computer's "friendly" name, as shown here.
    tcpClient.RemoteHost = tcpClient.LocalIP
    tcpClient.RemotePort = 1001
'If tcpClient.State = 0 Then
    ' Invoke the Connect method to initiate a
    ' connection.
    tcpClient.Connect
'End If
 
  
    MakeRound Me, 25
    swf1.Movie = App.Path & "\svrtraingle.swf"

    swf1.Play
    
    
   'MsgBox tcpServer(0).LocalHostName & ", " & tcpServer(0).LocalIP
End Sub

Private Sub cmdConnect_Click()

If tcpClient.State = 0 Then
    ' Invoke the Connect method to initiate a
    ' connection.
    tcpClient.Connect
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveForm Me
    Me.Caption = "NeoChat V.1.0.0 (TCP Client)"

End Sub

Private Sub Form_Paint()
    
    Me.Caption = "NeoChat V.1.0.0 (TCP Client)"
'    txtSend.SetFocus

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim myForm As Form
For Each myForm In Forms
    Unload myForm
Next
End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Error: " & Description
End Sub

Private Sub Timer1_Timer()

If Not tcpClient.State = 0 Then
'    cmdConnect.Enabled = False
'    cmdConnect.Visible = False
Else
 '   cmdConnect.Visible = True
 '   cmdConnect.Enabled = True
End If


    If tcpClient.State = 0 Then
        strState = "Closed"
        swf1.Loop = False
    ElseIf tcpClient.State = 1 Then
        strState = "Open"
        swf1.Loop = False
    ElseIf tcpClient.State = 2 Then
        strState = "Listening"
    swf1.Loop = True
    ElseIf tcpClient.State = 3 Then
        strState = "Connection Pending"
        swf1.Loop = True
    ElseIf tcpClient.State = 4 Then
        strState = "Resolving Host"
        swf1.Loop = True
    ElseIf tcpClient.State = 5 Then
        strState = "Host Resolved"
        swf1.Loop = True
    ElseIf tcpClient.State = 6 Then
        strState = "Connecting"
        swf1.Loop = True
    ElseIf tcpClient.State = 7 Then
        strState = "Connected"
        swf1.Loop = False
    ElseIf tcpClient.State = 8 Then
        strState = "Peer is closing the connection"
        swf1.Loop = True
        tcpClient.Close
    ElseIf tcpClient.State = 9 Then
        strState = "Error"
        tcpClient.Close
        swf1.Loop = True
    End If
    If swf1.Loop = True Then
        swf1.Play
    End If
    
    Label1.Caption = "Status: " & strState & " " & tcpServer.Count

    
End Sub

Private Sub txtOutput_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtSend_Change()
'    tcpClient.SendData txtSend.Text
End Sub

Private Sub tcpClient_DataArrival _
(ByVal bytesTotal As Long)
    Dim strData As String
    Dim iColor As Integer
    Dim sBit As String
    'MsgBox Me.hwnd
    sBit = ""
    iColor = 0
    tcpClient.GetData strData
    Do Until sBit = "~"
    If iColor > Len(strData) Then GoTo SkipColor
        iColor = iColor + 1
        sBit = Mid$(strData, iColor, 1)
    Loop
    strColor = Mid$(strData, 1, iColor - 1)
    
    strData = Mid$(strData, iColor + 1, Len(strData) - (iColor))
    sBit = ""
    iColor = 0
    Do Until sBit = "~"
        iColor = iColor + 1
        sBit = Mid$(strData, iColor, 1)
    Loop
    
    blBold = Mid$(strData, 1, iColor - 1)
    strData = Mid$(strData, iColor + 1, Len(strData) - (iColor))
    
    sBit = ""
    iColor = 0
    Do Until sBit = "~"
        iColor = iColor + 1
        sBit = Mid$(strData, iColor, 1)
    Loop
    blItalic = Mid$(strData, 1, iColor - 1)
    strData = Mid$(strData, iColor + 1, Len(strData) - (iColor))
    
SkipColor:
    txtOutput.SelStart = Len(txtOutput.Text)
    txtOutput.SelColor = strColor
    txtOutput.SelBold = blBold
    txtOutput.SelItalic = blItalic
    txtOutput.SelText = vbNewLine & strData & vbNewLine
    'txtOutput.Text = txtOutput.Text & vbNewLine & strData
    
   
    
    If Not GetActiveWindow = Me.hwnd Then
        Beep 600, 100
        FlashWindow Me.hwnd, 1
        Me.Caption = "New Message"
    End If
    
End Sub

Private Sub tcpServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
 ' Check if the control's State is closed. If not,
    ' close the connection before accepting the new
    ' connection.
    'If tcpServer.State <> sckClosed Then _
    'tcpServer.Close
    ' Accept the request with the requestID
    ' parameter.
    'tcpServer.Accept requestID
    'MsgBox "conected"
    If Index = 0 Then
      intMax = intMax + 1
      Load tcpServer(intMax)
      tcpServer(intMax).LocalPort = 0
      tcpServer(intMax).Accept requestID
      'Load txtData(intMax)
   End If

End Sub
Private Sub tcpServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   ' Declare a variable for the incoming data.
    ' Invoke the GetData method and set the Text
    ' property of a TextBox named txtOutput to
    ' the data.
    Dim strData As String
    tcpServer(Index).GetData strData
  '  txtOutput.Text = strData & vbNewLine & txtOutput.Text
    
    'pass the message to all clients, the client will interpret it
    For i = 1 To tcpServer.UBound
        If tcpServer(i).State = 7 Then
            tcpServer(i).SendData strData
        Else
            tcpServer(i).Close
        End If
    Next
    
    
End Sub
Private Sub tcpServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'MsgBox "error: " & Description
End Sub
