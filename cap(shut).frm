VERSION 5.00
Begin VB.Form cap 
   Caption         =   "Shut"
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1320
      Top             =   480
   End
   Begin VB.Menu mmenu 
      Caption         =   "xc"
      Visible         =   0   'False
      Begin VB.Menu ud 
         Caption         =   "Up/Down with click"
      End
      Begin VB.Menu em 
         Caption         =   "Easy move"
      End
      Begin VB.Menu fu 
         Caption         =   "-"
      End
      Begin VB.Menu hhide 
         Caption         =   "Hide(Cant be shown again!)"
      End
      Begin VB.Menu gf 
         Caption         =   "-"
      End
      Begin VB.Menu web 
         Caption         =   "Website"
      End
      Begin VB.Menu mmail 
         Caption         =   "robot@mail.gr"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "cap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub em_Click()
If em.Checked = True Then
em.Checked = False
easYmove = False
Else
em.Checked = True
easYmove = True
End If
End Sub

Private Sub Form_Load()
Me.Icon = Form1.Icon

Me.Caption = Form1.headd.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Interval = 0
Form1.Timer1.Interval = 0
Form1.Timer2.Interval = 0

Unload Me
Unload Form1

End Sub

Private Sub hhide_Click()
Form1.Hide
hiDDen = True
End Sub

Private Sub Timer1_Timer()
If hiDDen = True Then Exit Sub
Form1.headd.Caption = Me.Caption
If Me.WindowState <> vbMinimized Then
If Form1.Label9.Tag <> "hide" Then Exit Sub
Me.Top = Form1.Top
Me.Left = Form1.Left
Me.Hide
Form1.Label9.Tag = ""
Form1.WindowState = vbNormal
Form1.Show
End If

End Sub

Private Sub ud_Click()
If ud.Checked = False Then
ud.Checked = True
Form1.Text1.ToolTipText = "Hours adjustment .Left click to add/Right click to abstract"
Form1.Text1.MousePointer = 99
Form1.Text2.ToolTipText = "Minutes adjustment .Left click to add/Right click to abstract"
Form1.Text2.MousePointer = 99
Form1.Text3.ToolTipText = "Seconds adjustment .Left click to add/Right click to abstract"
Form1.Text3.MousePointer = 99
gHW = Form1.Text1.hwnd
Hook
gHW = Form1.Text2.hwnd
Hook
gHW = Form1.Text3.hwnd
Hook
Else
ud.Checked = False
Form1.Text1.ToolTipText = "Hours adjustment"
Form1.Text1.MousePointer = 0
Form1.Text2.ToolTipText = "Minutes adjustment"
Form1.Text2.MousePointer = 0
Form1.Text3.ToolTipText = "Seconds adjustment"
Form1.Text3.MousePointer = 0
gHW = Form1.Text1.hwnd
Unhook
gHW = Form1.Text2.hwnd
Unhook
gHW = Form1.Text3.hwnd
Unhook
End If
End Sub

Private Sub web_Click()
OpenInternet Me, "http://www.fetix.8m.com", Normal
End Sub
