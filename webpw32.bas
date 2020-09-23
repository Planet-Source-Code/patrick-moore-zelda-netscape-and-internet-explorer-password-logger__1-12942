Attribute VB_Name = "webPWS32"
Option Explicit
'**********************************
'* CODE BY: PATRICK MOORE (ZELDA) *
'* Feel free to re-distribute or  *
'* Use in your own projects.      *
'* Giving credit to me would be   *
'* nice :)                        *
'*                                *
'* Please vote for me if you find *
'* this code useful :]   -Patrick *
'**********************************
'http://members.nbci.com/erx931/VB/
'
'PS: Please look for more submissions to PSC by me
'    I've recently been working on a lot of them.
'    :))  All my submissions are under author name
'    "Patrick Moore (Zelda)"

'Define Find Window functions
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'Define Send Message functions
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'Define Get Text constants
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

'Define the Web Info type
Public Type WebInfo
    Username As String
    Password As String
    Server As String
End Type

Public Function GetText(WindowHandle As Long) As String
Dim Buffer As String, TextLength As Long

'Get the length of the text
TextLength = SendMessage(WindowHandle, WM_GETTEXTLENGTH, 0&, 0&)

'Fill a string with blank characters, the length
'of the text
Buffer = String(TextLength, 0&)

'Grab the actual text
Call SendMessageByString(WindowHandle, WM_GETTEXT, TextLength + 1, Buffer)

GetText = Buffer
End Function

Public Function FindIELogin() As Long
'Find the IE password-required login window
Dim Win As Long
Win = FindWindow("#32770", "Enter Network Password")
FindIELogin = Win
End Function

Public Function FindNetscapeLogin() As Long
'Find the Netscape password-required login window
Dim Win As Long
Win = FindWindow("#32770", "Username and Password Required")
FindNetscapeLogin = Win
End Function

Public Function GetLoginInfo() As WebInfo
Dim IEWin As Long, NetscapeWin As Long
Dim Edit As Long, Edit2 As Long, TheStatic As Long
Dim User As String, Pass As String, Server As String

'Find the IE login window
IEWin = FindIELogin

'Find the Netscape login window
NetscapeWin = FindNetscapeLogin

'See which we're working with

If IEWin > 0 Then
    'If it's IE, use this code
    
    'Find the first textbox
    Edit = FindWindowEx(IEWin, 0&, "Edit", vbNullString)
    
    'Find the second textbox
    Edit2 = FindWindowEx(IEWin, Edit, "Edit", vbNullString)
    
    'Find the icon
    TheStatic = FindWindowEx(IEWin, 0&, "Static", vbNullString)
    
    'Find a label
    TheStatic = FindWindowEx(IEWin, TheStatic, "Static", vbNullString)
    
    'Find the next label
    TheStatic = FindWindowEx(IEWin, TheStatic, "Static", vbNullString)
    
    'Find the next label
    TheStatic = FindWindowEx(IEWin, TheStatic, "Static", vbNullString)
    
    'Set the server to the label's caption
    Server = GetText(TheStatic)
    
    'Find the next label
    TheStatic = FindWindowEx(IEWin, TheStatic, "Static", vbNullString)
    
    'Set the server to the label's caption AND the previous
    'value of Static
    Server = Server & ";" & GetText(TheStatic)
    
    'Set the user to the first textbox's text
    User = GetText(Edit)
    
    'Set the password to the second textbox's text
    Pass = GetText(Edit2)
End If

If NetscapeWin > 0 Then
    'If it's Netscape, use this code
    
    'Find the first textbox
    Edit = FindWindowEx(NetscapeWin, 0&, "Edit", vbNullString)
    
    'Find the second textbox
    Edit2 = FindWindowEx(NetscapeWin, Edit, "Edit", vbNullString)
    
    'Find the label
    TheStatic = FindWindowEx(NetscapeWin, 0&, "Static", vbNullString)
    
    'Set the Server as the label's caption
    Server = GetText(TheStatic)
    
    'Trim the Server so it shows only the servers
    Server = Mid(Server, InStr(Server, "for") + 4, Len(Server))
    Server = Left(Server, Len(Server) - 1)
    
    'Set the User to the first textbox's text
    User = GetText(Edit)
    
    'Set the password to the second textbox's text
    Pass = GetText(Edit2)
End If
    
    
GetLoginInfo.Username = User
GetLoginInfo.Password = Pass
GetLoginInfo.Server = Server
End Function
