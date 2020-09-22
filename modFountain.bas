Attribute VB_Name = "modFountain"
' **********************************
' ***** Particle Fountain v1.1 *****
' ****** By Graham Sutherland ******
' ************ (C) 2005 ************
' **********************************

Option Explicit 'Basic declaration... I prefer to have this in.

Global NPPMS As Integer 'New Particles Per MilliSecond
'This is basically like particle density.
'Set it too low and it "spurts".
'Set it too high and it will look awful/lag.
'I advise that you set NPPMS to a number between 2 and 14
'If you enter a number which is not an integer, it's usually ignored.

Global Const LowSpecs = False
'Set this to true if using an old operating system, have a very slow
'processor or are experiencing performance problems.
'Setting this to true will enable the DoEvents command in loops,
'allowing the CPU priority to be shifted towards the OS and away from
'this program.
'Note: Enabling this will decrease the performance of this program! It
'is here only to increase the stability of low end systems during
'runtime.

Global IPR As Boolean 'Instant Particle Respawning
'When enabled, particles which go offscreen to instantly respawn
'without being sent back to the stockpile.
'This will cause the number of active particles to steadily climb
'until the particle cap limit is reached.

Type Drop 'Each "drop" is a particle.
    X As Single 'X Position
    Y As Single 'Y Position
    V As Single 'Vertical Velocity
    VX As Single 'Horizontal Velocity (spread)
End Type

Global Drops(20000) As Drop 'Array of 20K drops. ^_^
Global DropsActive As Integer 'Number of active drops.
Global Const Gravity = 9.81 'Gravitational constant for earth
Global Frames As Long 'Frame counter
Global FPS As Long 'Frames per second

Public Sub Init() 'Initialization
Dim i As Integer 'Loop variable
Call Randomize 'Set up the random number generator
DropsActive = 50 'Standard value
NPPMS = 6 'Standard value
IPR = False 'Standard value
frmMain.Hide
For i = 0 To 20000 'Loop through each drop, giving it starting values
    'Some values have variance, adds realism.
    DoEvents 'Free up processing power to the operating system.
    Drops(i).X = ((frmMain.ScaleWidth / 2) + 100) - Int(Rnd * 200) 'In the middle of the form, 100 units either way.
    Drops(i).Y = frmMain.ScaleHeight 'Off the form (at the bottom).
    Drops(i).V = Int(Rnd * 80) + 150 'Basic standard velocity, between 150 and 230.
    Drops(i).VX = 20 - Int(Rnd * 40) 'Horizontal velocity is set as -20 to 20, i.e 20 pixels left or right.
Next i
Call frmMain.UpdateStatus
frmMain.Show
End Sub

Public Sub DrawDrops() 'Draw the drops ^_^
Dim i As Integer 'Loop variable
frmMain.Cls 'Clear the form.
For i = 0 To DropsActive 'Loop through each drop, drawing it.
    If LowSpecs Then DoEvents 'Relinquish processing priority to the OS if enabled.
    frmMain.PSet (Drops(i).X, Drops(i).Y), RGB(200 + Int(Rnd * 30), 200 + Int(Rnd * 30), 255) 'Draw the droplet with a random white-ish-blue colour
Next i
Frames = Frames + 1
End Sub

Public Sub MoveDrops() 'Calculate the new positions of the drops
Dim i As Integer 'Loop variable
For i = 0 To DropsActive 'Loop through each active drop, setting its new values
    If LowSpecs Then DoEvents 'Relinquish processing priority to the OS if enabled.
    Drops(i).Y = Drops(i).Y - Drops(i).V 'Move it up. Remember, Y increases as you go down, so we take the velocity away.
    Drops(i).V = Drops(i).V - Gravity 'Decrement the velocity by the gravitational constant. Basic physics ^_^
    Drops(i).X = Drops(i).X + Drops(i).VX 'Move the particle left or right depending on its horizontal velocity.
    If Drops(i).Y > frmMain.ScaleWidth Then 'If it's off-screen, reset it.
        Drops(i).X = ((frmMain.ScaleWidth / 2) + 100) - Int(Rnd * 200) 'In the middle of the form, 100 units either way.
        Drops(i).Y = frmMain.ScaleHeight 'Off the form (at the bottom).
        Drops(i).V = Int(Rnd * 80) + 150 'Basic standard velocity, between 150 and 230.
        Drops(i).VX = 20 - Int(Rnd * 40) 'Horizontal velocity is set as -20 to 20, i.e 20 pixels left or right.
        If Not IPR Then DropsActive = DropsActive - 1 'If IPR is disabled, send the particle back to the stockpile.
    End If
Next i
End Sub
