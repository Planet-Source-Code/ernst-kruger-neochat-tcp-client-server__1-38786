VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "NeoChat V.1.0.0 (TCP Client)"
   ClientHeight    =   4710
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5205
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmClient.frx":030A
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOption 
      DownPicture     =   "frmClient.frx":1486
      Height          =   255
      Left            =   3000
      Picture         =   "frmClient.frx":1729
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   180
      Width           =   1215
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
      DisabledPicture =   "frmClient.frx":19D2
      DownPicture     =   "frmClient.frx":1CA2
      Height          =   495
      Left            =   240
      Picture         =   "frmClient.frx":1F80
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "frmClient.frx":225A
      Height          =   255
      Left            =   4410
      Picture         =   "frmClient.frx":23B7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      DownPicture     =   "frmClient.frx":2519
      Height          =   255
      Left            =   4680
      Picture         =   "frmClient.frx":26A9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   2160
   End
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H00808080&
      DisabledPicture =   "frmClient.frx":2840
      DownPicture     =   "frmClient.frx":2B98
      Height          =   495
      Left            =   240
      MaskColor       =   &H0000C000&
      Picture         =   "frmClient.frx":2F05
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2760
      Width           =   4695
   End
   Begin RichTextLib.RichTextBox txtOutput 
      Height          =   2175
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmClient.frx":3278
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf1 
      Height          =   855
      Left            =   2040
      TabIndex        =   7
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
      TabIndex        =   2
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmClient"
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

Private Sub cmdOption_Click()

    frmOptions.Show vbModal
    
End Sub

Private Sub cmdSend_Click()

    If tcpClient.State = 7 Then '7 = connected
        'Send string with color and text variables
        tcpClient.SendData sColor & "~" & bBold & "~" & bItalic & "~" & sUserName & " :->  " & txtSend.Text
      '  txtOutput.Text = "Me:> " & txtSend.Text & vbNewLine & txtOutput.Text
        
        'reset input textbox
        txtSend.Text = ""
        'Play flash movie
        swf1.Play
    End If

End Sub

Private Sub Command1_Click()
    Unload Me
    End
End Sub

Private Sub Command2_Click()
'minimize window
Me.WindowState = 1

End Sub


Private Sub Form_Activate()
    
    Me.Caption = "NeoChat V.1.0.0 (TCP Client)"
    txtSend.SetFocus

End Sub

Private Sub Form_GotFocus()
    
    Me.Caption = "NeoChat V.1.0.0 (TCP Client)"
    txtSend.SetFocus
    
End Sub

Private Sub Form_Load()
    
    Set frmSplash = Nothing
    ' The name of the Winsock control is tcpClient.
    ' Note: to specify a remote host, you can use
    ' either the IP address (ex: "121.111.1.1") or
    ' the computer's "friendly" name, as shown here.
    tcpClient.RemoteHost = sServer
    tcpClient.RemotePort = 1001
        
 
  
    MakeRound Me, 25
    swf1.Movie = App.Path & "\traingle.swf"
    
    swf1.Play
End Sub

Private Sub cmdConnect_Click()

If tcpClient.State = 0 Then
    ' Invoke the Connect method to initiate a
    ' connection.
    tcpClient.Connect
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'call sub routine in module to move frameles form SEE modMain
    MoveForm Me
    Me.Caption = "NeoChat V.1.0.0 (TCP Client)"

End Sub

Private Sub Form_Paint()
    
    Me.Caption = "NeoChat V.1.0.0 (TCP Client)"

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Loop through all forms and close all
Dim myForm As Form
For Each myForm In Forms
    Unload myForm
Next

End Sub

Private Sub Timer1_Timer()

If Not tcpClient.State = 0 Then '0 = Closed - Display Connect button(cmdConnect)
    cmdConnect.Enabled = False
    cmdConnect.Visible = False
Else
    cmdConnect.Visible = True
    cmdConnect.Enabled = True
End If

    'Get tcpClient state
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
    
    'Update label1
    Label1.Caption = "Status: " & strState

    
End Sub

Private Sub txtOutput_KeyDown(KeyCode As Integer, Shift As Integer)
'disallow any input in txtOutput
KeyCode = 0
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
'disallow any input in txtOutput
KeyAscii = 0
End Sub

Private Sub tcpClient_DataArrival _
(ByVal bytesTotal As Long)
    
    Dim strData As String
    Dim iColor As Integer 'Place holder for color value
    Dim sBit As String 'Variable to evaluate each character in a string
    

    sBit = "" 'Variable to evaluate each character in a string
    iColor = 0 'Place holder for color value
    tcpClient.GetData strData 'Retrieves the current block of data and stores it in a variable of type variant
    Do Until sBit = "~" 'Get the color value at the start of the string
    If iColor > Len(strData) Then GoTo SkipColor 'something is wrong
        iColor = iColor + 1
        sBit = Mid$(strData, iColor, 1)
    Loop
    strColor = Mid$(strData, 1, iColor - 1)
    
    strData = Mid$(strData, iColor + 1, Len(strData) - (iColor)) 'The rest of it ia our message
    sBit = ""
    iColor = 0
    Do Until sBit = "~"
        iColor = iColor + 1
        sBit = Mid$(strData, iColor, 1)
    Loop
    
    'Check BOLD value
    blBold = Mid$(strData, 1, iColor - 1)
    strData = Mid$(strData, iColor + 1, Len(strData) - (iColor))
    
    sBit = ""
    iColor = 0
    Do Until sBit = "~"
        iColor = iColor + 1
        sBit = Mid$(strData, iColor, 1)
    Loop
    
    'Check Italic value
    blItalic = Mid$(strData, 1, iColor - 1)
    strData = Mid$(strData, iColor + 1, Len(strData) - (iColor))
    
SkipColor:
    txtOutput.SelStart = Len(txtOutput.Text)
    txtOutput.SelColor = strColor
    txtOutput.SelBold = blBold
    txtOutput.SelItalic = blItalic
    txtOutput.SelText = vbNewLine & strData & vbNewLine
    'txtOutput.Text = txtOutput.Text & vbNewLine & strData
    
   
    'Get user's attention when message arrives
    If Not GetActiveWindow = Me.hwnd Then
        Beep 600, 100
        FlashWindow Me.hwnd, 1
        Me.Caption = "New Message"
    End If
    
End Sub


