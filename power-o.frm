VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Shut"
   ClientHeight    =   6225
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Left click for counting down.Right click for extra options"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   0
      TabIndex        =   31
      Text            =   "write the correct password here"
      Top             =   1200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Options"
      Height          =   2295
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Width           =   3735
      Begin VB.OptionButton logoff 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Log Off User"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox thepass 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   29
         ToolTipText     =   "Hold shift button to view the password"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox passwordprotect 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Password Protected (Force)"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "When is checked enables the password protection for shutting/restarting the computer"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton restartcomputer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Restart The Computer"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "When selected restarts the computer"
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton shutdown 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Shut Down The Computer"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "When selected shuts down the computer"
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Label writethepass 
         BackStyle       =   0  'Transparent
         Caption         =   "Write The Password Bellow:"
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select specific time"
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   3735
      Begin VB.OptionButton temp 
         Caption         =   "Option12"
         Height          =   195
         Left            =   1200
         TabIndex        =   32
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton moreopt 
         Caption         =   "More Options"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4 hours"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3 hours"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2 hours"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1,5 hour"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 hour"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "30 minutes"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "20 minutes"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "15 minutes"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "10 minutes"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5 minutes"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1 minute"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Timer Timer2 
      Left            =   2880
      Top             =   3840
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Stop the counting!"
      Top             =   960
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   3600
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Seconds adjustment  .Left click to add/Right click to abstract"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Minutes adjustment  .Left click to add/Right click to abstract"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   480
      MaxLength       =   3
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "Hours adjustment .Left click to add/Right click to abstract"
      Top             =   600
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   3400
      X2              =   3400
      Y1              =   0
      Y2              =   260
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3444
      TabIndex        =   37
      Top             =   0
      Width           =   135
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Label closed 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3666
      TabIndex        =   36
      Top             =   0
      Width           =   255
   End
   Begin VB.Label headd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Power-O"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   20
      X1              =   0
      X2              =   4320
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label inl 
      BackStyle       =   0  'Transparent
      Caption         =   "In:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      ToolTipText     =   "This will set the countdown time  e.g. In: 0:10:33"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label At 
      BackStyle       =   0  'Transparent
      Caption         =   "At:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      ToolTipText     =   "This will set the time in 24h format e.g. At: 15:05:44"
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "The remaining time.With click changes color"
      Top             =   480
      Width           =   2745
   End
   Begin VB.Label Label5 
      Caption         =   "15"
      Height          =   135
      Left            =   2880
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Hours      Minutes   Seconds"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Type MYVERSION
lMajorVersion As Long
lMinorVersion As Long
lExtraInfo As Long
End Type
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Dim winversion As String
Dim curx, cury As Integer
Dim comms() As String
Dim fpass As Boolean
Dim ftime As Boolean
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const SPI_SCREENSAVERRUNNING = 97
Dim trys As Integer
Private Declare Function FlashWindow _
Lib "user32" ( _
ByVal hwnd As Long, _
ByVal bInvert As Long _
) As Long
Private Sub saveset()
On Error GoTo wth
Open App.Path & "\settings.txt" For Output As #1
If cap.em.Checked = True Then
Print #1, "1"
Else
Print #1, "0"
End If
If cap.ud.Checked = True Then
Print #1, "1"
Else
Print #1, "0"
End If
wth:
Close #1
End Sub
Private Sub loadset()
On Error GoTo ex
Dim aka As String
Open App.Path & "\settings.txt" For Input As #1
Line Input #1, aka
cap.em.Checked = aka
easYmove = aka
Line Input #1, aka
cap.ud.Checked = aka
Close #1
ex:
End Sub

Private Sub restore()
Label6.ForeColor = &HFF&
Form1.BackColor = &HC0C0C0
Frame1.BackColor = &HC0C0C0
Option1.BackColor = &HC0C0C0
Option2.BackColor = &HC0C0C0
Option3.BackColor = &HC0C0C0
Option4.BackColor = &HC0C0C0
Option5.BackColor = &HC0C0C0
Option6.BackColor = &HC0C0C0
Option7.BackColor = &HC0C0C0
Option8.BackColor = &HC0C0C0
Option9.BackColor = &HC0C0C0
Option10.BackColor = &HC0C0C0
Option11.BackColor = &HC0C0C0
Frame2.BackColor = &HC0C0C0
shutdown.BackColor = &HC0C0C0
restartcomputer.BackColor = &HC0C0C0
passwordprotect.BackColor = &HC0C0C0
logoff.BackColor = &HC0C0C0
End Sub
Private Sub change()
If Label6.ForeColor = &HFF& Then
Label6.ForeColor = &H40&
Form1.BackColor = &HFFFF&
Frame1.BackColor = &HFFFF&
Option1.BackColor = &HFFFF&
Option2.BackColor = &HFFFF&
Option3.BackColor = &HFFFF&
Option4.BackColor = &HFFFF&
Option5.BackColor = &HFFFF&
Option6.BackColor = &HFFFF&
Option7.BackColor = &HFFFF&
Option8.BackColor = &HFFFF&
Option9.BackColor = &HFFFF&
Option10.BackColor = &HFFFF&
Option11.BackColor = &HFFFF&
Frame2.BackColor = &HFFFF&
shutdown.BackColor = &HFFFF&
restartcomputer.BackColor = &HFFFF&
passwordprotect.BackColor = &HFFFF&
logoff.BackColor = &HFFFF&
Else
Label6.ForeColor = &HFF&
Form1.BackColor = &HC0C0C0
Frame1.BackColor = &HC0C0C0
Option1.BackColor = &HC0C0C0
Option2.BackColor = &HC0C0C0
Option3.BackColor = &HC0C0C0
Option4.BackColor = &HC0C0C0
Option5.BackColor = &HC0C0C0
Option6.BackColor = &HC0C0C0
Option7.BackColor = &HC0C0C0
Option8.BackColor = &HC0C0C0
Option9.BackColor = &HC0C0C0
Option10.BackColor = &HC0C0C0
Option11.BackColor = &HC0C0C0
Frame2.BackColor = &HC0C0C0
shutdown.BackColor = &HC0C0C0
restartcomputer.BackColor = &HC0C0C0
passwordprotect.BackColor = &HC0C0C0
logoff.BackColor = &HC0C0C0
End If
End Sub
Private Sub At_Click()
At.ForeColor = &HFF&
inl.ForeColor = &H80000011
If Text1.Text > 23 Then Text1.Text = 0
If Text2.Text > 59 Then Text2.Text = 0
If Text3.Text > 59 Then Text3.Text = 0
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Option10.Enabled = False
Option11.Enabled = False
 cap.Caption = "Do it at:" & Label6.Caption
  
End Sub

Private Sub closed_Click()
Call saveset
End
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Form1.Height = 1260 Then
If moreopt.Caption = "Less Options" Then
Form1.Height = 6225
Else
Form1.Height = 3840
End If
Else
 Form1.Height = 1260

End If
End If
End Sub
Private Sub Command2_Click()

 Dim aa As String
Dim a As String

 If Timer1.Enabled = False Then Exit Sub

If passwordprotect.Value = 1 Then
   If thepass.Text <> "" Then
If Text4.Visible = False Then
   Form1.Height = 1515
   Text4.Visible = True
   
   Timer1.Enabled = True
Else
Text4.Visible = False
Form1.Height = 1260
End If
  Exit Sub
  Else
   Timer1.Enabled = False
   Command1.Enabled = True
   Text1.Enabled = True
   Text2.Enabled = True
   Text3.Enabled = True
   cap.Caption = "Power-O"
   
   closed.Visible = True
   fpass = False
   ftime = False
   Call restore
   Exit Sub
   End If
Else
   Timer2.Interval = 0
   Label5.Caption = 15
   cap.Caption = "Power-O"
   
   Timer1.Enabled = False

  
  End If
Dim ret As Integer
     Dim pOld As Boolean
     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
headd.Caption = Label6.Caption
Call restore
End Sub
Function split(the_commands As String) As String

Dim were() As Integer

ReDim were(0)
Dim tnext As Integer

Dim i As Integer
For i = 1 To Len(the_commands)
 If (Mid(the_commands, i, 1) = "-") Or (Mid(the_commands, i, 1) = "/") Then
 ReDim Preserve were(UBound(were) + 1)
  were(UBound(were)) = i
 End If
Next i

If UBound(were) > 0 Then
ReDim comms(UBound(were))

For i = 1 To UBound(were)
  If UBound(were) = i Then
  tnext = Len(the_commands) + 2
  Else
  tnext = were(i + 1)
  End If
  comms(i) = Mid(the_commands, were(i) + 1, tnext - were(i) - 2)

Next i
End If
End Function
Private Sub Form_Load()
Dim myVer As MYVERSION
    Dim strTmp As String
    Dim vers As Long
    myVer = WindowsVersion()

    If myVer.lMajorVersion = 4 Then
        If myVer.lExtraInfo = VER_PLATFORM_WIN32_NT Then
            strTmp = "NT"
        ElseIf myVer.lExtraInfo = VER_PLATFORM_WIN32_WINDOWS Then
            vers = myVer.lMinorVersion
            If vers <= 10 Then
                strTmp = "98"
            Else
                strTmp = "98"
            End If
        End If
       Else
        strTmp = "XP"
    End If
winversion = strTmp

Call loadset
If cap.ud.Checked = True Then
'no right click menu
gHW = Text1.hwnd
Hook
gHW = Text2.hwnd
Hook
gHW = Text3.hwnd
Hook
Form1.Text1.ToolTipText = "Hours adjustment .Left click to add/Right click to abstract"
Form1.Text1.MousePointer = 99
Form1.Text2.ToolTipText = "Minutes adjustment .Left click to add/Right click to abstract"
Form1.Text2.MousePointer = 99
Form1.Text3.ToolTipText = "Seconds adjustment .Left click to add/Right click to abstract"
Form1.Text3.MousePointer = 99
Else
Form1.Text1.ToolTipText = "Hours adjustment"
Form1.Text1.MousePointer = 0
Form1.Text2.ToolTipText = "Minutes adjustment"
Form1.Text2.MousePointer = 0
Form1.Text3.ToolTipText = "Seconds adjustment"
Form1.Text3.MousePointer = 0
End If
'for commands
Me.Height = 1260

Dim i As Integer
Dim commands As String
commands = Command()

If commands = "" Then GoTo gg
On Error Resume Next
split (commands)
For i = 1 To UBound(comms)
'commands
 'at command
 If UCase(Mid(comms(i), 1, 2)) = "AT" Then
  If (Mid(comms(i), 4, 2) >= 0) And (Mid(comms(i), 4, 2) < 60) And (Mid(comms(i), 7, 2) >= 0) And (Mid(comms(i), 7, 2) < 60) And (Mid(comms(i), 10, 2) >= 0) And (Mid(comms(i), 10, 2) < 60) Then
   Text1.Text = Mid(comms(i), 4, 2)
   Text2.Text = Mid(comms(i), 7, 2)
   Text3.Text = Mid(comms(i), 10, 2)
   At_Click
   Command1_Click
  End If
 End If
'now command
If UCase(comms(i)) = "NOW" Then Call shut
'min command
If UCase(comms(i)) = "MIN" Then
Me.WindowState = vbMinimized
Me.Label9.Tag = "hide"
Me.Hide
cap.WindowState = vbMinimized
cap.Show
End If
'shutdown command
If UCase(comms(i)) = "SHUTDOWN" Then shutdown.Value = True
'restart command
If UCase(comms(i)) = "RESTART" Then restartcomputer.Value = True
'logoff command
If UCase(comms(i)) = "LOGOFF" Then logoff.Value = True
'force command
If UCase(comms(i)) = "FORCE" Then
passwordprotect.Value = 1
Command1_Click
End If
'set password
If UCase(Mid(comms(i), 1, 8)) = "PASSWORD" Then
thepass.Text = Mid(comms(i), 10, Len(comms(i)) - 8)
passwordprotect.Value = 1
Command1_Click
End If
'in command
If UCase(Mid(comms(i), 1, 2)) = "IN" Then
 If UCase(Mid(comms(i), 4, 3)) = "SEC" Then
  Text3.Text = Mid(comms(i), 8, Len(comms(i)) - 7)
 Else
  If UCase(Mid(comms(i), 4, 3)) = "MIN" Then
   Text2.Text = Mid(comms(i), 8, Len(comms(i)) - 7)
  Else
   If UCase(Mid(comms(i), 4, 4)) = "HOUR" Then
    Text1.Text = Mid(comms(i), 9, Len(comms(i)) - 7)
   Else
    Text1.Text = Mid(comms(i), 4, 2)
    Text2.Text = Mid(comms(i), 7, 2)
    Text3.Text = Mid(comms(i), 10, 2)
   End If
  End If
 End If
Command1_Click
End If
If UCase(comms(i)) = "HIDDEN" Then
Form1.Hide
hiDDen = True
End If
Next i
gg:

fpass = False
ftime = False


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If easYmove = False Then Exit Sub

If Button = 1 Then
curx = X
cury = Y
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If easYmove = False Then Exit Sub
If Button = 1 Then
Form1.Left = Form1.Left + (X - curx)
Form1.Top = Form1.Top + (Y - cury)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If easYmove = False Then Exit Sub
Form1.Refresh
End Sub

Private Sub Form_Terminate()
Call saveset
Dim ret As Integer
     Dim pOld As Boolean
     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
If fpass = True And ftime = True Then Call shut

End Sub
Function shut()
Call saveset
If winversion = "98" Then
If shutdown.Value = True Then Call ShutDownNT(2, passwordprotect.Value)
Else
If shutdown.Value = True Then Call ShutDownNT(1, passwordprotect.Value)
End If
If restartcomputer.Value = True Then Call ShutDownNT(3, passwordprotect.Value)
If logoff.Value = True Then Call ShutDownNT(4, passwordprotect.Value)
End Function

Private Sub Form_Unload(Cancel As Integer)
Call saveset
Unhook
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If easYmove = False Then Exit Sub

If Button = 1 Then
curx = X
cury = Y
End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If easYmove = False Then Exit Sub
If Button = 1 Then
Form1.Left = Form1.Left + (X - curx)
Form1.Top = Form1.Top + (Y - cury)
End If

End Sub

Private Sub headd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
curx = X
cury = Y
Else
cap.PopupMenu cap.mmenu

End If
End Sub

Private Sub headd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Form1.Left = Form1.Left + (X - curx)
Form1.Top = Form1.Top + (Y - cury)

End If
End Sub

Private Sub headd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Refresh
End Sub

Private Sub inl_Click()
inl.ForeColor = &HFF&
At.ForeColor = &H80000011
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
Option10.Enabled = True
Option11.Enabled = True
cap.Caption = "Do it in:" & Label6.Caption
   

End Sub

Private Sub Label2_Change()
 'format the time
 If Len(Label3.Caption) = 1 Then Label3.Caption = "0" & Label3.Caption
 If Len(Label4.Caption) = 1 Then Label4.Caption = "0" & Label4.Caption
 Label6.Caption = Label2.Caption & ":" & Label3.Caption & ":" & Label4.Caption
End Sub

Private Sub Label4_Change()
 'format the time
 If Len(Label4.Caption) = 1 Then Label4.Caption = "0" & Label4.Caption
 Label6.Caption = Label2.Caption & ":" & Label3.Caption & ":" & Label4.Caption
End Sub

Private Sub Label6_Change()
If Label6.Caption = "0:0:0" Then
ftime = False
Else
ftime = True
End If
If Form1.WindowState = vbMinimized Then Exit Sub
 
End Sub


Private Sub Label6_Click()
Call change
End Sub



Private Sub Label9_Click()
Label9.Tag = "hide"

Me.WindowState = vbMinimized
Me.Hide
cap.WindowState = vbMinimized
cap.Show
End Sub

Private Sub moreopt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If moreopt.Caption = "More Options" Then
moreopt.Caption = "Less Options"
Form1.Height = 6225
Else
moreopt.Caption = "More Options"
Form1.Height = 3800
End If
End Sub

Private Sub Option1_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 0
Text2.Text = 1
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option10_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 3
Text2.Text = 0
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option11_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 4
Text2.Text = 0
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub



Private Sub Option2_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 0
Text2.Text = 5
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option3_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 0
Text2.Text = 10
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option4_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 0
Text2.Text = 15
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option5_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 0
Text2.Text = 20
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option6_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 0
Text2.Text = 30
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option7_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 1
Text2.Text = 0
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option8_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 1
Text2.Text = 30
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub Option9_Click()
If Timer1.Enabled = True Then Command1_Click
Text1.Text = 2
Text2.Text = 0
Text3.Text = 0
Label2.Caption = Text1.Text
Label3.Caption = Text2.Text
Label4.Caption = Text3.Text
End Sub

Private Sub passwordprotect_Click()
If passwordprotect.Value = 1 Then
fpass = True
If Timer1.Enabled = True Then
cap.Caption = "Press START to force"
Else
cap.Caption = "Enter a password now"
End If
thepass.Visible = True
writethepass.Visible = True
Else
fpass = False
cap.Caption = "Power-O"
thepass.Visible = False
writethepass.Visible = False
End If

End Sub

Private Sub Text1_Change()
 On Error GoTo re
  If At.ForeColor = &HFF& Then
  If Text1.Text < 24 Then
   Label2.Caption = Int(Text1.Text)
   cap.Caption = "Do it at:" & Label6.Caption
   Else
   GoTo re
   End If
   Else
  If Text1.Text < 999 Then
    Label2.Caption = Int(Text1.Text)
   cap.Caption = "Do it in:" & Label6.Caption
   
   Else
  End If
 End If
re:

End Sub

Private Sub Text1_LostFocus()
 If Not IsNumeric(Text2.Text) Then Text2.Text = 0
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cap.ud.Checked = False Then Exit Sub
On Error Resume Next
If Button = 1 Then
Text1.Text = Text1.Text + 1
Else
If Text1.Text > 0 Then Text1.Text = Text1.Text - 1
End If
End Sub
Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cap.ud.Checked = False Then Exit Sub
On Error Resume Next
If Button = 1 Then
Text2.Text = Text2.Text + 1
Else
If Text2.Text > 0 Then Text2.Text = Text2.Text - 1
End If
End Sub

Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cap.ud.Checked = False Then Exit Sub
On Error Resume Next
If Button = 1 Then
Text3.Text = Text3.Text + 1
Else
If Text3.Text > 0 Then Text3.Text = Text3.Text - 1
End If
End Sub
Private Sub Text2_Change()

 On Error GoTo re
  If Text2.Text > 60 Then
   Text2.Text = 60
  End If
 Label3.Caption = Int(Text2.Text)
cap.Caption = "Do it at:" & Label6.Caption
 Exit Sub
re:
End Sub

Private Sub Text2_LostFocus()
 If Not IsNumeric(Text2.Text) Then Text2.Text = 0
End Sub

Private Sub Text3_Change()

  On Error GoTo re
  If Text3.Text > 60 Then
   Text3.Text = 60
  End If
 Label4.Caption = Int(Text3.Text)
cap.Caption = "Do it at:" & Label6.Caption
 Exit Sub
re:
End Sub

Private Sub Text3_LostFocus()
 If Not IsNumeric(Text2.Text) Then Text2.Text = 0
End Sub

Private Sub Text4_Change()
Dim a As String

If trys >= Len(thepass.Text) * 4 Then
Text4.Text = "Too many attempts.No more"
Text4.Enabled = False
End If
 If thepass.Text = Text4.Text Then
 Form1.Height = 1260
 Timer1.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Visible = False
thepass.Text = ""
thepass.Enabled = True
Command1.Enabled = True
 fpass = False
 
closed.Visible = True

Else
trys = trys + 1
End If

End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub Text4_GotFocus()
Text4.Text = ""
End Sub

Private Sub thepass_KeyDown(KeyCode As Integer, Shift As Integer)
thepass.PasswordChar = ""
End Sub

Private Sub thepass_KeyUp(KeyCode As Integer, Shift As Integer)
thepass.PasswordChar = "*"
End Sub

Private Sub Timer1_Timer()
If inl.ForeColor = &HFF& Then           'in
   Label7.Caption = "Remaining time"
   
    If Label3.Caption = 0 And Label2.Caption = 0 And Label4.Caption < 16 Then
    Label5.Caption = Label4.Caption - 1
    GoTo de
    End If
    If Label4.Caption = 0 And Label3.Caption > 0 Then
    Label3.Caption = Label3.Caption - 1
    Label4.Caption = 60
    End If
    If Label3.Caption = 0 And Label2.Caption > 0 Then
    Label2.Caption = Label2.Caption - 1
    Label3.Caption = 59
    Label4.Caption = 59
    End If
    Label4.Caption = Label4.Caption - 1
    cap.Caption = Label6.Caption
    Exit Sub
    
de:
    
    Form1.Rate = 2
    cap.Caption = "Shut in:" & Label5.Caption & " seconds"
    Label4.Caption = Label4.Caption - 1
    Label5.Caption = Label5.Caption - 1
    Form1.WindowState = vbNormal
 
     If Timer1.Enabled = False Then cap.Caption = "Power-O"

    Dim lngRtn As Long
    lngRtn = FlashWindow(cap.hwnd, CLng(True))
  
      Call change
     If Label5.Caption = -2 Then
    Timer1.Enabled = False
    cap.Caption = "Exit in progess..."
    Call shut
    
    End If
Else 'At
cap.Caption = "At:" & Int(Text1.Text) & ":" & Int(Text2.Text) & ":" & Int(Text3.Text)
Label7.Caption = "Current time"
Label2.Caption = Hour(Time)
Label3.Caption = Minute(Time)
Label4.Caption = Second(Time)
If (Int(Text1.Text) = Int(Hour(Time))) And (Int(Text2.Text) = Int(Minute(Time))) And (Int(Text3.Text) = Int(Second(Time))) Then Call shut

End If

    Command2.Visible = True
End Sub
Private Sub Command1_Click()
Timer1.Enabled = True
Command2.Visible = True
Form1.Height = 1260
If Label3.Caption = 0 And Label2.Caption = 0 And Label4.Caption = 0 Then Label4.Caption = 15
temp.Value = True
If passwordprotect.Value = 1 Then
'no X
Command1.Enabled = False
closed.Visible = False

thepass.Enabled = False
'no CTRL alt delete
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End If
If passwordprotect.Value = 1 Then
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End If
End Sub
Property Let Rate(intPerSecond As Integer)
    Timer2.Interval = 500 / intPerSecond
End Property
Property Let Flash(blnState As Boolean)
    Timer2.Enabled = blnState
End Property

Private Function WindowsVersion() As MYVERSION
    Dim myOS As OSVERSIONINFO, WinVer As MYVERSION
    Dim lResult As Long

    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148

    lResult = GetVersionEx(myOS)

    'Fill user type with pertinent info
    WinVer.lMajorVersion = myOS.dwMajorVersion
    WinVer.lMinorVersion = myOS.dwMinorVersion
    WinVer.lExtraInfo = myOS.dwPlatformId

    WindowsVersion = WinVer


End Function

