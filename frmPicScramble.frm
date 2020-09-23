VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPicScramble 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pic Scramble"
   ClientHeight    =   5160
   ClientLeft      =   735
   ClientTop       =   1800
   ClientWidth     =   11880
   Icon            =   "frmPicScramble.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimeGame 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.PictureBox SourceBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   0
      Picture         =   "frmPicScramble.frx":0442
      ScaleHeight     =   391
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   586
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8790
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_New 
         Caption         =   "New"
      End
      Begin VB.Menu Mnu_ReSize 
         Caption         =   "Resize"
      End
      Begin VB.Menu Mnu_Line 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu_options 
      Caption         =   "Options"
      Begin VB.Menu mnu_timer 
         Caption         =   "Timer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Prefrences 
         Caption         =   "Prefrences"
      End
   End
End
Attribute VB_Name = "frmPicScramble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Its amuzing, what more can I say.

'Enjoy... and if ya have any questions, suggestions
'or comments E-Mail me (addy below)

'Oh yeah, one more thing. I'm working on a rather large program
'one that takes this, and several other "logic" style type puzzles
'and puts em all into one, with lots of cool customizable options
'If anyone wants to join me in working on this (so far it's just me)
'then E-Mail me at mik@dccnet.com

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim SourceMap() As cord
Dim BoxMap() As cord
Dim TotalCol As Integer
Dim TotalRow As Integer
Dim TotalCorrect As Integer
Dim CurrentlySelected As Integer
Dim ScramblePercent As Integer

Dim CurrentTimer As Long
Dim TotalElapsed As Long

Dim TimeLeft As Integer

Private Sub Form_Load()
    ScramblePercent = 100 '% This amount will be the amount to scramble
                         'IE: a Value of 25 will allow for appr 25% of
                         'the grid to be scrambled
    
    SizeAll
    DefaultTime = 180
    TotalCol = 5
    TotalRow = 5
    INI

End Sub
Private Sub SizeAll()
    
    ResizeOffsetX = 6 * Screen.TwipsPerPixelX
    ResizeOffsetY = 52 * Screen.TwipsPerPixelY
    PicBox.Width = SourceBox.Width
    PicBox.Height = SourceBox.Height
    Me.Width = (SourceBox.ScaleWidth * Screen.TwipsPerPixelX) + ResizeOffsetX
    Me.Height = (SourceBox.ScaleHeight * Screen.TwipsPerPixelY) + ResizeOffsetY

End Sub

Private Sub INI()

    Dim TotalBox As Integer
    Dim CurrentBox As Integer
    Dim RandBox As Integer
    Me.Caption = "Pic Scramble"
    
    TotalCorrect = 0
    CurrentlySelected = -1
    TotalBox = (TotalCol * TotalRow)
    TimeLeft = DefaultTime
    
    ReDim SourceMap(TotalBox - 1)

    CurrentBox = 0
    For CounterY = 0 To TotalRow - 1
        For CounterX = 0 To TotalCol - 1
            SourceMap(CurrentBox).x = CounterX
            SourceMap(CurrentBox).y = CounterY
            SourceMap(CurrentBox).Counter = CurrentBox
            CurrentBox = CurrentBox + 1
        Next CounterX
    Next CounterY
    BoxMap() = SourceMap()

    CurrentBox = 0
    ScrambleStep = Int(100 / ScramblePercent)
    For CurrentBox = 0 To TotalBox - 1 Step ScrambleStep
            Randomize Timer
            RandBox = Int(CurrentBox * Rnd)
            OldCounter = BoxMap(CurrentBox).Counter
            BoxMap(CurrentBox).Counter = BoxMap(RandBox).Counter
            BoxMap(RandBox).Counter = OldCounter
            SwapBoxMap CurrentBox, RandBox
    Next CurrentBox

    For CurrentBox = 0 To TotalBox - 1
        If SourceMap(CurrentBox).x = BoxMap(CurrentBox).x And SourceMap(CurrentBox).y = BoxMap(CurrentBox).y Then
            BoxMap(CurrentBox).Correct = 1
            TotalCorrect = TotalCorrect + 1
        End If
    Next CurrentBox
    TimeLeft = DefaultTime
    
    DrawAll
    
    If mnu_timer.Checked = True Then
        TimeGame.Enabled = True
        TimeGame_Timer
        
    Else
        TimeGame.Enabled = False
    End If
    
End Sub
Private Sub SwapBoxMap(Counter1 As Integer, Counter2 As Integer)
    
    OldX = BoxMap(Counter1).x
    Oldy = BoxMap(Counter1).y
    BoxMap(Counter1).x = BoxMap(Counter2).x
    BoxMap(Counter1).y = BoxMap(Counter2).y
    BoxMap(Counter2).x = OldX
    BoxMap(Counter2).y = Oldy

End Sub
Private Sub DrawAll()
    
    Dim OldCaption As String
    Dim CurrentBox As Integer
    
    PicBox.Cls
    GetScale TotalRow, TotalCol, PicBox
    OldCaption = Me.Caption
    Me.Caption = "Loading"
    DoEvents
    
    For Counter = 0 To UBound(BoxMap)
        If DrawEffect = True Then
            CurrentTimer = GetTickCount()
            TotalElapsed = GetTickCount() - CurrentTimer
            Do While TotalElapsed <= DrawSpeed
                TotalElapsed = GetTickCount() - CurrentTimer
                DoEvents
            Loop
        End If
        
        CurrentBox = BoxMap(Counter).Counter
        FillFrom CurrentBox, CurrentBox, , DrawEffect
    
    Next Counter
    PicBox.Refresh
    Me.Caption = OldCaption
    'A simple grid procedure was written for testing
    'Uncomment the line below to see the oh so wonderfull grid
    'DrawGrid TotalRow, TotalCol, QBColor(9), PicBox
End Sub

Private Sub Mnu_Exit_Click()
    
    Unload Me
    
End Sub

Private Sub mnu_New_Click()

    Dim FileName As String
    
    CommonDialog.Filter = "Pictures (*.JPG)|*.JPG"
    CommonDialog.ShowOpen
    FileName = CommonDialog.FileName
    If FileName <> "" Then
        SourceBox.Picture = LoadPicture(FileName)
        PromptScale
        PromptScramble
        SizeAll
        INI
    End If

End Sub
Private Sub PromptScramble()
    Do
        R = InputBox("Percent to be scrambled?", "Scramble Amount", 100)
    Loop While (Trim$(R) = "") Or (Val(R) <= 1) Or (Val(R) > 100)
    ScramblePercent = R
End Sub
Private Sub PromptScale()
        
        If SourceBox.Width > SourceBox.Height Then
            PomptScaleAmount = (Screen.Width / Screen.TwipsPerPixelX) * 0.75
            LargerDirection = SourceBox.Width
        Else
            PomptScaleAmount = (Screen.Height / Screen.TwipsPerPixelY) * 0.75
            LargerDirection = SourceBox.Height
        End If
        
        If LargerDirection > PomptScaleAmount Then
            R = MsgBox("Loaded picture appears to be very large. Would you like the picture to be scaled?", vbYesNo)
            If R = vbYes Then
                RecomendedScale = (PomptScaleAmount / LargerDirection) * 100
                Do
                    R = InputBox("Scale down to what percent?", "Scale Amount", RecomendedScale)
                Loop While (Trim$(R) = "") Or (Val(R) <= 1) Or (Val(R) >= 100)
                ScalePercent = R
                If SourceBox.Width > SourceBox.Height Then
                    PicXScale = SourceBox.Width * (ScalePercent / 100)
                    PicYScale = SourceBox.Height * (PicXScale / SourceBox.Width)
                Else
                    PicYScale = SourceBox.Height * (ScalePercent / 100)
                    PicXScale = SourceBox.Width * (PicYScale / SourceBox.Height)
                End If
                SourceBox.PaintPicture SourceBox.Picture, 0, 0, PicXScale, PicYScale, 0, 0, SourceBox.Width, SourceBox.Height, vbSrcCopy
                SourceBox.Width = PicXScale
                SourceBox.Height = PicYScale
            End If
        End If
End Sub
Private Sub mnu_Prefrences_Click()
    
    R = MsgBox("Entering Prefrences will reset current game, do you wish to proceed?", vbYesNo)
    If R = vbYes Then
        TimeGame.Enabled = False
        frmOptions.Show 1
        INI
    End If

End Sub

Private Sub Mnu_ReSize_Click()
    
    'Error checking for the InputBox's
    'is bit on the primitive side
    'Could use a re-write
    
    Do
        R = InputBox("How Many Box's wide?", "Box Prompt")
    Loop While (Trim$(R) = "") Or (Val(R) <= 0)
    TotalCol = R
    
    Do
        R = InputBox("How Many Box's tall?", "Box Prompt")
    Loop While (Trim$(R) = "") Or (Val(R) <= 0)
    TotalRow = R
    PromptScramble
    INI

End Sub

Private Sub mnu_timer_Click()
    R = MsgBox("Changing Timer setting will reset current game, do you wish to proceed?", vbYesNo)
       
    If R = vbNo Then Exit Sub
    
    If mnu_timer.Checked = True Then
        mnu_timer.Checked = False
    Else
        mnu_timer.Checked = True
    End If
    INI
        
End Sub

Private Sub PicBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Counter As Integer

        If Button = 1 Then
            Point2Box x, y 'Fills the var CurrentBox
            Counter = FindBoxNum(TotalCol, CurrentBox.x, CurrentBox.y)
        
                If CurrentlySelected = -1 Then
                    'clicked on the first box, invert it
                    FillFrom Counter, Counter, NOTSRCCOPY, True
                    CurrentlySelected = Counter
                Else
                    If Counter = CurrentlySelected Then
                        'clicked on the same one
                        'remove the invert (make it look normal again)
                        'and exit sub so that it doesnt mess up our total
                        FillFrom CurrentlySelected, Counter, , True
                        CurrentlySelected = -1
                        Exit Sub
                    End If
            
                    'do the swap
                    FillFrom CurrentlySelected, Counter
                    FillFrom Counter, CurrentlySelected
                    PicBox.Refresh
                    'check if either of those were swaped into the
                    'correct places
                    CheckCorrect CurrentlySelected, Counter
                    CheckCorrect Counter, CurrentlySelected
 
                    'update the array to remember the new locations
                    SwapBoxMap Counter, CurrentlySelected
            
                    CurrentlySelected = -1
                End If
        End If
    DoCheck 'Did ya win? Well, did ya? huh? huh? huh?

End Sub
Private Sub FillFrom(Source As Integer, Dest As Integer, Optional CopyMode As Long = SRCCOPY, Optional RefreshFill As Boolean = False)
    
    Dim SourceX As Integer
    Dim SourceY As Integer
    Dim DestX As Integer
    Dim DestY As Integer

    SourceX = BoxMap(Source).x
    SourceY = BoxMap(Source).y
    DestX = SourceMap(Dest).x
    DestY = SourceMap(Dest).y
            
    FillBox SourceX, SourceY, DestX, DestY, TotalCol, True, CopyMode, SourceBox, PicBox, RefreshFill

End Sub
Private Sub CheckCorrect(Value1 As Integer, Value2 As Integer)
    
    If BoxMap(Value1).x = SourceMap(Value2).x And BoxMap(Value1).y = SourceMap(Value2).y Then
        'the box in this spot is correct
        'update the total
        BoxMap(Value2).Correct = 1
        TotalCorrect = TotalCorrect + 1
    Else
        If BoxMap(Value2).Correct = 1 Then
            'the box in this spot is correct, but now isn't
            'update the total
            BoxMap(Value2).Correct = 0
            TotalCorrect = TotalCorrect - 1
        End If
    End If

End Sub

Private Sub DoCheck()

    Dim TotalBox As Integer

    TotalBox = TotalCol * TotalRow

    If TotalCorrect = TotalBox Then
        'Could be much more impressive later on...
        'Heck, I even gave this its own sub to hold
        'all that "later on" potentially impressive code.
        TimeGame.Enabled = False
        MsgBox "YAY YOU"
    End If

End Sub

Private Sub TimeGame_Timer()
    
    If TimeLeft <= 0 Then
        MsgBox "Times up!"
        INI
    End If
    TimeLeft = TimeLeft - 1
    MinLeft = Fix(TimeLeft / 60)
    SecLeft = (TimeLeft - (MinLeft * 60))
    
    Me.Caption = "Pic Scramble " & "Time Remaining: Min: " & MinLeft & " Sec: " & SecLeft

End Sub
