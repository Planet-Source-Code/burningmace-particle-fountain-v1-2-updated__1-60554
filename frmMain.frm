VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Particle Fountain v1.2"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStatus 
      Interval        =   1000
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer tmrRelease 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer tmrDrops 
      Interval        =   20
      Left            =   720
      Top             =   480
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "0 Particles          FPS: 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuNPPMS 
         Caption         =   "New particles per millisecond:"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDecrease 
         Caption         =   "Decrease NP/ms"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuIncrease 
         Caption         =   "Increase NP/ms"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "Draw width:"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDecDW 
         Caption         =   "Decrease Draw Width"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuIncDW 
         Caption         =   "Increase Draw Width"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIPR 
         Caption         =   "Enable IPR"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoRedraw 
         Caption         =   "Auto Redraw"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **********************************
' ***** Particle Fountain v1.1 *****
' **********************************
' ****** By Graham Sutherland ******
' ******* A.K.A. Burningmace *******
' ********* A.K.A. Trojan1 *********
' **********************************
' ************ (C) 2005 ************
' **********************************

' Please vote and comment on Planet Sourcecode if you like this
' program!
' http://www.planet-source-code.com/

' My System: AMD Athlon 2000 @ 1313MHz, Windows XP Pro SP1
' Performance on my system (running in MSVB IDE):
'   32fps maximum @ <750 Particles
'   Starts to drop below 32fps at 750 Particles (9 NP/ms)
'   13fps @ 4000 Particles, 1px (30 NP/ms)
'   6fps @ 4000 Particles, 10px (30 NP/ms)
'   4fps @ 20000 Particles, 1px (IPR Enabled)
'   2fps @ 20000 Particles, 10px (IPR Enabled)
'   1fps @ 20000 Particles, 10px (IPR + AutoRedraw Enabled)

Private Sub Form_Load()
Init    'Initialise the program
End Sub

Private Sub mnuAutoRedraw_Click()
'Toggle Auto Redraw
Me.Cls 'If you don't clear the screen, odd stuff can happen.
Me.AutoRedraw = Not Me.AutoRedraw
mnuAutoRedraw.Checked = Not mnuAutoRedraw.Checked
End Sub

Private Sub mnuDecDW_Click()
If Me.DrawWidth = 1 Then Exit Sub
Me.DrawWidth = Me.DrawWidth - 1
End Sub

Private Sub mnuDecrease_Click()
NPPMS = NPPMS - 1
If NPPMS < 0 Then NPPMS = 0
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuHelp_Click()
'Display the Readme.txt file

On Error GoTo ErrHandle
If FileExists(App.Path & "\Readme.txt") Then 'Check if it exists
    Shell "notepad.exe " & App.Path & "\Readme.txt", vbNormalFocus 'Run readme.txt in notepad
Else 'If not...
    Call MsgBox("Unable to find file 'Readme.txt' in program directory.", vbCritical, "Particle Fountain - Error") 'Display an error.
End If

Exit Sub 'Skip the error handling if there was no error

ErrHandle:
    'Something went wrong... maybe the user doesn't have notepad!?
    Call MsgBox("An unexpected error occured.", vbCritical, "Particle Fountain - Error")
End Sub

Private Sub mnuIncDW_Click()
If Me.DrawWidth = 10 Then Exit Sub
Me.DrawWidth = Me.DrawWidth + 1
End Sub

Private Sub mnuIncrease_Click()
NPPMS = NPPMS + 1
If NPPMS > 30 Then NPPMS = 30
End Sub

Private Sub mnuIPR_Click()
'Toggle Instant Particle Respawning (see explanation in modFountain)
mnuIPR.Checked = Not mnuIPR.Checked
IPR = Not IPR
End Sub

Private Sub mnuMenu_Click()
'When the user clicks the menu, the number of NP/ms and the draw width
'is updated so that the correct values are shown.
mnuNPPMS.Caption = "New particles per millisecond: " & NPPMS
mnuDrawWidth.Caption = "Draw Width: " & Me.DrawWidth & "px"
End Sub

Private Sub tmrStatus_Timer()
'This timer updates the status.
'The FPS counter works by incrementing 'Frames' by one every time the
'particles are updated. When this sub is called (every 1000ms, or 1s)
'the value of Frames will be how many frames were updated in that
'second. The value is displayed and 'Frames' is reset to 0.
Dim DA As Long
FPS = Frames
Frames = 0
DA = DropsActive
If DA < 0 Then DA = 0 'Sometimes you get -1 particles.
frmMain.lblStatus.Caption = DA & " Particles          FPS: " & FPS 'Update the particle counter
End Sub

Private Sub tmrDrops_Timer()
'This timer refreshes the particles (drops of water)
DrawDrops   'Draw the drops
MoveDrops   'Calculate the drops new positions
End Sub

Private Sub tmrRelease_Timer()
'This timer releases new particles (drops of water)
DropsActive = DropsActive + NPPMS   'Activate some new drops
If DropsActive > 20000 Then DropsActive = 20000   'Particle cap!
End Sub

Public Sub UpdateStatus()
tmrStatus_Timer
End Sub

Public Function FileExists(FileSpec As String) As Boolean
'This is a bit of a hack... I use FileLen to attempt to get the size
'of the file. If the file exists, no error occurs and the function
'returns true. If the file doesn't exists, the "GoTo ErrHandle" code
'is invoked and the function returns false. I could use a
'FileSystemObject from the Microsoft Scripting Runtime library, but
'I can't be bothered to do all the stuff like "Set fso = Nothing"
'when the program is terminated. Besides, this way will filter out
'all files which are not accessible or cause errors when accessed.

On Error GoTo ErrHandle
Dim length As Long
length = FileLen(FileSpec)
FileExists = True
Exit Function

ErrHandle:
    FileExists = False
End Function
