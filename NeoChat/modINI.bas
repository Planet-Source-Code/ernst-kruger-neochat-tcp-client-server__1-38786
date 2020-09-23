Attribute VB_Name = "modINI"
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

'This module contains the INI file (neodata.INI) functions, for storing settings like username _
Host IP etc. It was my first attempt at using INI files, and I don't know how _
I ever managed without them. It uses only 2 API functions, namely _
GetPrivateProfileString and WritePrivateProfileString to read and write _
data to and from the INI file (neodata.INI) respectively

Option Explicit 'All variables must be declared

Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
'GetPrivateProfileString: This API function simply reads data stored in an INI file
'It is called from the sGetINI function in this module. It needs a reference to _
the physical path of your INI file, the section to read from within the INI file (neodata.INI), _
the key or setting which you want to retreive as well as a default value _
in case no value had been entered previously or the INI file (neodata.INI) doesn't exist.

Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) _
    As Long
'WritePrivateProfileString: This API function writes data to _
the INI file (neodata.INI). It needs basically the same info as _
GetPrivateProfileString (above) but instead of _
the default value it needs an actual value to store

Public Function sGetINI(sINIFile As String, sSection As String, sKey _
As String, sDefault As String) As String
' this is where you call the GetPrivateProfileString API function. With values _
supplied in frmSplash, frmOptions etc.

    Dim sTemp As String * 256 'this string will contain the returned value from _
                                the INI file (neodata.INI)
    
    Dim nLength As Integer 'this will store the length of the sTemp string _
                            to be returned.

    sTemp = Space$(256) 'sets sTemp to a 256 character long blank string
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
    'call the GetPrivateProfileString API and determine the length of the _
    return string
    
    sGetINI = Left$(sTemp, nLength)
    'contains the returned string, shortened from 256 to the appropriate length
    
End Function

Public Function WriteINI(sINIFile As String, sSection As String, sKey _
As String, sValue As String)
'This function writes a value to the INI file (neodata.INI) and if no INI file exists it _
creates one. This function is also called from frmSplash, frmOptions and _
frmSvrOptions

    Dim n As Integer 'used as a counter
    Dim sTemp As String 'the value to be written to the INI file (neodata.INI)

    sTemp = sValue

    For n = 1 To Len(sValue) 'this For loop removes Carriage and Linefeed _
                                characters from the value to be written (sValue)
        If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf Then _
        Mid$(sValue, n) = ""
    Next n
    
    n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
    'write the data to the INI file (neodata.INI)

End Function
