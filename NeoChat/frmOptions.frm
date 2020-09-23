VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOptions.frx":0000
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkServer 
      BackColor       =   &H00FF0000&
      Caption         =   "Server"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.CheckBox chkItl 
      DownPicture     =   "frmOptions.frx":120C
      Height          =   375
      Left            =   3240
      Picture         =   "frmOptions.frx":13B1
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   375
   End
   Begin VB.CheckBox chkBold 
      DownPicture     =   "frmOptions.frx":1558
      Height          =   375
      Left            =   2160
      Picture         =   "frmOptions.frx":1728
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      DownPicture     =   "frmOptions.frx":18FA
      Height          =   255
      Left            =   4680
      Picture         =   "frmOptions.frx":1A87
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4200
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      Picture         =   "frmOptions.frx":1C18
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1455
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
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'SEE modINI for details on using ini files to store variables
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
        frmClient.Show
    Else
    Dim myForm As Form
        For Each myForm In Forms
            Unload myForm
            frmSplash.Show
        Next myForm
    End If
End Sub



Private Sub Form_Load()
'See modINI for details on retreiving variables from ini files
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
