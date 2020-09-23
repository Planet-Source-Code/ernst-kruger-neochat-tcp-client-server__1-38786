VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSvrOptions 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSvrOptions.frx":0000
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      Picture         =   "frmSvrOptions.frx":131C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      DownPicture     =   "frmSvrOptions.frx":1375
      Height          =   255
      Left            =   4680
      Picture         =   "frmSvrOptions.frx":1502
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   255
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox chkBold 
      DownPicture     =   "frmSvrOptions.frx":1693
      Height          =   375
      Left            =   2160
      Picture         =   "frmSvrOptions.frx":1864
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   375
   End
   Begin VB.CheckBox chkItl 
      DownPicture     =   "frmSvrOptions.frx":1A36
      Height          =   375
      Left            =   3240
      Picture         =   "frmSvrOptions.frx":1BDF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   375
   End
   Begin VB.CheckBox chkServer 
      BackColor       =   &H000080FF&
      Caption         =   "Server"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4200
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Host IP:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmSvrOptions"
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

'This form is identical to frmOptions except for color.

Dim strColor As String
Dim bChanged As Boolean
Dim iSave As Integer
Dim bRestart As Boolean

Private Sub chkBold_Click()
bChanged = True
End Sub

Private Sub chkItl_Click()
bChanged = True
End Sub

Private Sub chkServer_Click()
bChanged = True
bRestart = True
End Sub

Private Sub cmdColor_Click()
bChanged = True
    dlg.ShowColor
    strColor = dlg.Color
    cmdColor.BackColor = strColor
'    Label3.Caption = strColor
End Sub

Private Sub Command1_Click()
If bChanged = True Then
iSave = MsgBox("Save settings?", vbYesNo + vbQuestion, "Settings Changed")

    If iSave = vbYes Then
        If Not txtName.Text = "" Then
            If Not txtName = sUserName Then
                WriteINI sINIFile, "Settings", "UserName", txtName.Text
                sUserName = sGetINI(sINIFile, "Settings", "UserName", "?")
            End If
        End If
        If Not strColor = "" Then
            If Not strColor = sColor Then
                WriteINI sINIFile, "Settings", "Color", strColor
                sColor = sGetINI(sINIFile, "Settings", "Color", "0")
            End If
        End If
        If chkBold.Value = 0 Then
            WriteINI sINIFile, "Settings", "Bold", "False"
        Else
            WriteINI sINIFile, "Settings", "Bold", "True"
        End If
        If chkItl.Value = 0 Then
            WriteINI sINIFile, "Settings", "Italic", "False"
        Else
            WriteINI sINIFile, "Settings", "Italic", "True"
        End If
        If Not txtHost.Text = "" Then
            WriteINI sINIFile, "Settings", "Server", txtHost.Text
            sServer = sGetINI(sINIFile, "Settings", "Server", "?")
        End If
        If chkServer.Value = 0 Then
            WriteINI sINIFile, "Settings", "IsServer", "False"
        Else
            WriteINI sINIFile, "Settings", "IsServer", "True"
        End If
        bBold = sGetINI(sINIFile, "Settings", "Bold", "False")
        bItalic = sGetINI(sINIFile, "Settings", "Italic", "False")
        bIsServer = sGetINI(sINIFile, "Settings", "IsServer", "False")
        
    ElseIf iSave = vbNo Then

    End If

End If
    Unload Me
    
    If Not bRestart Then
        frmMainServer.Show
    Else
    Dim myForm As Form
        For Each myForm In Forms
            Unload myForm
            frmSplash.Show
        Next myForm
    End If
End Sub

Private Sub Form_Load()
bRestart = False
    MakeRound Me, 25
    bChanged = False
    txtName.Text = sGetINI(sINIFile, "Settings", "UserName", "?")
    txtHost.Text = sGetINI(sINIFile, "Settings", "Server", "?")
    
    cmdColor.BackColor = sColor
    
    If bBold = False Then
        chkBold.Value = 0
    Else
        chkBold.Value = 1
    End If
    
    If bItalic = False Then
        chkItl.Value = 0
    Else
        chkItl.Value = 1
    End If
    If bIsServer = False Then
        chkServer.Value = 0
    Else
        chkServer.Value = 1
    End If
    
End Sub

Private Sub txtHost_KeyPress(KeyAscii As Integer)
bChanged = True
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
bChanged = True
End Sub

