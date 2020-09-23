Attribute VB_Name = "modMain"
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

'This is the main module. I use it to store create session variables, and for various _
API calls _
The Following API functions are used:
    'Beep - this a simple api call to play a beep sound through the system speaker. _
        I use this to notify the user of a new message when the main forms are _
        minimized or don't have focus
    'GetActiveWindow - this one is also very straight forward. It retreives the _
        handle of the current active window, and if it isn't one of the main _
        forms the program will notify the user when a new message arives
    'FlashWindow - this function normally flashes a forms' title bar and taskbar _
        icon, but as none of my forms have title bars it only flashes the _
        taskbar. This is also used to alert the user when a new message arives
    'ReleaseCapture - this function is essential for moving borderless forms, _
        see the MoveForm funtion in this module. frmClient and frmMainServer _
        both call this function in their MouseDown events.
    'SendMessage - also used in the MoveForm function for borderless forms
    'CreateRoundRectRgn - Used to round the corners of a form. Implemented in _
        the MakeRound function in this module which is called from almost all _
        the forms, in their Load events.
    'SetWindowRgn - also used in the MakeRound function, you could even create _
        a perfectly round form using this funtion
'All this is repeated at each function
Option Explicit

Public sINIFile As String 'Stores path to INI file
Public sUserName As String 'Stores current username
Public sColor As String 'stores color text to be displayed
Public ncount As Integer 'just a counter
Public sServer As String 'the ip address or computer name of the host computer
Public bIsServer As Boolean 'true on the host computer
Public bBold As Boolean 'if text should be displayed boldface
Public bItalic As Boolean 'for text to be displayed italic
Public iINI As Integer 'another counter

Public Declare Function Beep Lib "kernel32" _
(ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'Beep - this a simple api call to play a beep sound through the system speaker. _
        I use this to notify the user of a new message when the main forms are _
        minimized or don't have focus


Public Declare Function GetActiveWindow Lib "user32" () As Long
'GetActiveWindow - this one is also very straight forward. It retreives the _
        handle of the current active window, and if it isn't one of the main _
        forms the program will notify the user when a new message arives
        

Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
'FlashWindow - this function normally flashes a forms' title bar and taskbar _
        icon, but as none of my forms have title bars it only flashes the _
        taskbar. This is also used to alert the user when a new message arives
        
        
Private Declare Function ReleaseCapture Lib "user32" () As Long
'ReleaseCapture - this function is essential for moving borderless forms, _
        see the MoveForm funtion in this module. frmClient and frmMainServer _
        both call this function in their MouseDown events.
        

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As Long
'SendMessage - also used in the MoveForm function for borderless forms
Const WM_NCLBUTTONDOWN = &HA1


Declare Function CreateRoundRectRgn Lib "gdi32" _
        (ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long
'CreateRoundRectRgn - Used to round the corners of a form. Implemented in _
        the MakeRound function in this module which is called from almost all _
        the forms, in their Load events.
        
        
Declare Function SetWindowRgn Lib "user32" _
        (ByVal hwnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
 'SetWindowRgn - also used in the MakeRound function, you could even create _
        a perfectly round form using this funtion
        
'Add this module to any app to create great looking borderless forms. Remember _
to set the ShowInTaskbar property on your forms if you want to add minimize _
functionality.
        
Public Sub MoveForm(f As Form)
'Uses ReleaseCapture and SendMessage API functions to move borderless froms. _
See it action in the MouseDown event on frmClient and frmMainServer.

    ReleaseCapture
    SendMessage f.hwnd, WM_NCLBUTTONDOWN, 2, 0
    
End Sub


Public Sub MakeRound(pForm As Form, lValue As Long)
'Uses CreateRoundRectRgn and SetWindowRgn API calls to round the edges of a form. _
I added rounded GIFs with transparent corners to each form picture _
property. This is an easy way to create great looking graphical interfaces to _
any application. I work with two graphic designers who are always keen _
to show off their talents. Thanks for all the images Marc!

    Dim lRet As Long
    Dim l As Long
    Dim llWidth As Long
    Dim llHeight As Long
            
    'Get Form size in pixels
    llWidth = pForm.Width / Screen.TwipsPerPixelX
    llHeight = pForm.Height / Screen.TwipsPerPixelY
    
    'Create Form with Rounded Corners
    lRet = CreateRoundRectRgn(0, 0, llWidth, llHeight, _
                              lValue, lValue)
                              
    l = SetWindowRgn(pForm.hwnd, lRet, True)
End Sub

