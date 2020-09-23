VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   1080
      Top             =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Top             =   3960
      Width           =   45
   End
End
Attribute VB_Name = "frmSplash"
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

'This Splash screen retreives variables from the ini file to determine which form to open, _
and loads other user variables


'For mor on ini files see modINI
Option Explicit
Dim i As Integer


Private Sub Form_Load()
i = 0

    sINIFile = App.Path & "\neodata.INI"
    'set a reference reference to the INI file. see modINI for all INI functions.
    
    sUserName = sGetINI(sINIFile, "Settings", "UserName", "?")
    sServer = sGetINI(sINIFile, "Settings", "Server", "?")
    bIsServer = sGetINI(sINIFile, "Settings", "IsServer", "False")
    sColor = sGetINI(sINIFile, "Settings", "Color", "0")
    bBold = sGetINI(sINIFile, "Settings", "Bold", "False")
    bItalic = sGetINI(sINIFile, "Settings", "Italic", "False")
    
End Sub




Private Sub Timer2_Timer()
If i = 5 Then
Timer2.Enabled = False
i = 0
    'check if a username exists. if not, create one.
    
    If sUserName = "?" Then
        'username does not exist, so ask for one
       frmOptions.Show
       Unload Me
    Else
        If bIsServer = True Then
            frmMainServer.Show
            Unload Me
    
        Else

            frmClient.Show
            Unload Me
        End If
  
    End If
Else

Label1.Caption = "User :" & sUserName & " Server: " & sServer


    i = i + 1
End If
End Sub
