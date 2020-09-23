VERSION 5.00
Begin VB.Form FrmMatrix 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Assessment XP - Matrix Test"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   Icon            =   "FrmMatrix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrApply 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "FrmMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IntLastXPos As Integer    'For use in Checking the mouse movements
Dim IntLastYPos As Integer 'For use in Checking the mouse movements
Dim IntX As Integer
Dim IntY As Integer

Dim XBefore(2) As Long
Dim YBefore(2) As Long

Dim LOD() As Integer
Dim Lead() As Integer
Dim Letter() As Integer
Dim IntColour() As Integer
Dim BoolUsed() As Boolean
Dim Wait() As Integer

Dim LOD1() As Integer
Dim Lead1() As Integer
Dim Letter1() As Integer
Dim IntColour1() As Integer
Dim BoolUsed1() As Boolean
Dim Wait1() As Integer

Dim LOD2() As Integer
Dim Lead2() As Integer
Dim Letter2() As Integer
Dim IntColour2() As Integer
Dim BoolUsed2() As Boolean
Dim Wait2() As Integer


Dim IntMaxLength As Integer   'The maximum length of the column
Dim IntMaxLngWait As Integer   'The maximum Waiting time Before clearing
Dim IntDropCols As Integer   'The StrNumber of dropping coloumns
Dim IntFadeSpeed As Integer   'The fading speed of the symbols

Dim Size1 As Integer
Dim Size2 As Integer
Dim Size3 As Integer

Dim CW As Long
Dim CH As Long

Private Sub Form_Load()
    Dim IntCurrent As Integer
    Dim IntDoFill As Integer
    Dim IntPNo As Integer
    
    Randomize Timer
     
    '#######################################################################
    'Falling Code Settings
    '#######################################################################
    'This section gets the settings from the registry and stores them in the variables
    IntReloadStyle = 1
    IntMaxLength = 100
    IntMaxLngWait = 200
    IntDropCols = 25       'Retieve the StrNumber of dropping columns
    IntFadeSpeed = 3         'Retieve the fading speed of the columns
    
    FrameCount = 0
    
    TmrApply.Enabled = True
    
    Size1 = 6
    Size2 = 9
    Size3 = 12

End Sub


Private Sub TmrApply_Timer()
    Dim DoRand As Integer
    Dim XR As Integer
    Dim YR As Integer
    Dim Temp As Long
    Dim Loading As Integer
    Dim AddNum As Integer
    
    TmrApply.Enabled = False
    Dim SW As Integer
    Dim SH As Integer
    
    
    
    FrmMatrix.ScaleMode = 4
    FrmMatrix.ScaleWidth = FrmMatrix.ScaleWidth * 10 / Size1
    FrmMatrix.ScaleHeight = FrmMatrix.ScaleHeight * 10 / Size1
    SW = FrmMatrix.ScaleWidth
    SH = FrmMatrix.ScaleHeight
    ReDim LOD(SW + 1, SH + 5) As Integer
    ReDim Lead(SW + 1, SH + 5) As Integer
    ReDim Letter(SW + 1, SH + 5) As Integer
    ReDim IntColour(1 To 2, SW + 1, SH + 5) As Integer
    ReDim Wait(SW + 1, SH + 5) As Integer
    ReDim BoolUsed(SW + 1, SH + 5) As Boolean
    Call SetBefore(0)
    
    FrmMatrix.ScaleMode = 4
    FrmMatrix.ScaleWidth = FrmMatrix.ScaleWidth * 10 / Size2
    FrmMatrix.ScaleHeight = FrmMatrix.ScaleHeight * 10 / Size2
    SW = FrmMatrix.ScaleWidth
    SH = FrmMatrix.ScaleHeight
    ReDim LOD1(SW + 1, SH + 5) As Integer
    ReDim Lead1(SW + 1, SH + 5) As Integer
    ReDim Letter1(SW + 1, SH + 5) As Integer
    ReDim IntColour1(1 To 2, SW + 1, SH + 5) As Integer
    ReDim Wait1(SW + 1, SH + 5) As Integer
    ReDim BoolUsed1(SW + 1, SH + 5) As Boolean
    Call SetBefore(1)

    FrmMatrix.ScaleMode = 4
    FrmMatrix.ScaleWidth = FrmMatrix.ScaleWidth * 10 / Size3
    FrmMatrix.ScaleHeight = FrmMatrix.ScaleHeight * 10 / Size3
    SW = FrmMatrix.ScaleWidth
    SH = FrmMatrix.ScaleHeight
    ReDim LOD2(SW + 1, SH + 5) As Integer
    ReDim Lead2(SW + 1, SH + 5) As Integer
    ReDim Letter2(SW + 1, SH + 5) As Integer
    ReDim IntColour2(1 To 2, SW + 1, SH + 5) As Integer
    ReDim Wait2(SW + 1, SH + 5) As Integer
    ReDim BoolUsed2(SW + 1, SH + 5) As Boolean
    Call SetBefore(2)
    
    
    Font = "Matrix"   'Use the Matrix Font
    
    For DoRand = 1 To IntDropCols 'Create the starting IntDrops
        XR = Int(Rnd * XBefore(0)) + 1  'The IntX position
        YR = Int(Rnd * (YBefore(0) + 5)) + 1   'The IntY position
        LOD(XR, YR) = Int(Rnd * IntMaxLength)      'The Length of the drop
        Lead(XR, YR) = 1 'Make it a Lead symbol
        Letter(XR, YR) = Int(Rnd * 43) + 65         'Set the letter/symbol
        BoolUsed(XR, YR) = True
        IntColour(2, XR, YR) = 255
        IntColour(1, XR, YR) = Rnd * 100 + 100
    Next
    For DoRand = 1 To IntDropCols 'Create the starting IntDrops
        XR = Int(Rnd * XBefore(1)) + 1  'The IntX position
        YR = Int(Rnd * (YBefore(1) + 5)) + 1   'The IntY position
        LOD1(XR, YR) = Int(Rnd * IntMaxLength)      'The Length of the drop
        Lead1(XR, YR) = 1 'Make it a Lead symbol
        Letter1(XR, YR) = Int(Rnd * 43) + 65         'Set the letter/symbol
        BoolUsed1(XR, YR) = True
        IntColour1(2, XR, YR) = 255
        IntColour1(1, XR, YR) = Rnd * 100 + 100
    Next
    For DoRand = 1 To IntDropCols 'Create the starting IntDrops
        XR = Int(Rnd * XBefore(2)) + 1  'The IntX position
        YR = Int(Rnd * (YBefore(2) + 5)) + 1   'The IntY position
        LOD2(XR, YR) = Int(Rnd * IntMaxLength)      'The Length of the drop
        Lead2(XR, YR) = 1 'Make it a Lead symbol
        Letter2(XR, YR) = Int(Rnd * 43) + 65         'Set the letter/symbol
        BoolUsed2(XR, YR) = True
        IntColour2(2, XR, YR) = 255
        IntColour2(1, XR, YR) = Rnd * 100 + 100
    Next
    CW = FrmMatrix.ScaleWidth
    CH = FrmMatrix.ScaleHeight
    Do
        DoEvents
        FrameCount = FrameCount + 1
        Call MoreThanOneColour
        FrmMatrix.Refresh
    Loop Until Timer - StartTime > 5
End Sub

Sub MoreThanOneColour()
    Dim IntDrops As Integer
    Dim IntMakeNew As Integer
    
    FrmMatrix.ScaleMode = 3
    FrmMatrix.Cls
    
    Font.SIZE = Size1
    For IntX = 1 To XBefore(0)
        For IntY = 1 To YBefore(0) + 5
            If BoolUsed(IntX, IntY) <> 0 Then
                If Lead(IntX, IntY) = 1 Then 'Is it Lead
                    If IntY <= YBefore(0) + 4 Then  'Is it smaller than the screen height
                        If LOD(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If Letter(IntX, IntY + 1) <> 0 Then
                                Call Clear(IntX, IntY + 1, 0)
                            End If
                            LOD(IntX, IntY + 1) = LOD(IntX, IntY) - 1
                            Lead(IntX, IntY) = 0
                            Lead(IntX, IntY + 1) = 2
                            Letter(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            BoolUsed(IntX, IntY + 1) = True
                            IntColour(1, IntX, IntY + 1) = Rnd * 100 + 100
                            IntColour(2, IntX, IntY + 1) = 255
                            Wait(IntX, IntY) = IntMaxLngWait / 2 + Rnd(IntMaxLngWait / 2)
                            Call ShowHigh(IntX, IntY + 1, 0)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If Lead(IntX, IntY) = 1 Then
                                Call Clear(IntX, IntY, 0)
                            End If
                        End If
                    Else
                        Call Clear(IntX, IntY, 0)
                    End If
                ElseIf Lead(IntX, IntY) = 2 Then
                    Lead(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                ElseIf Lead(IntX, IntY) = 3 Then
                End If
                If Wait(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    Wait(IntX, IntY) = Wait(IntX, IntY) - 1
                    
                    If Wait(IntX, IntY) = 0 Or IntColour(1, IntX, IntY) = 0 Then
                        Call Clear(IntX, IntY, 0)
                    Else
                        IntColour(1, IntX, IntY) = IntColour(1, IntX, IntY) - IntFadeSpeed
                        IntColour(2, IntX, IntY) = IntColour(2, IntX, IntY) - IntFadeSpeed * 2
                        If IntColour(1, IntX, IntY) < 0 Then
                            IntColour(1, IntX, IntY) = 0
                        End If
                        If IntColour(2, IntX, IntY) < 0 Then
                            IntColour(2, IntX, IntY) = 0
                        End If
                        If IntColour(1, IntX, IntY) = 0 Then
                            Call Clear(IntX, IntY, 0)
                        ElseIf Lead(IntX, IntY) = 0 Then
                            Call ShowColor(IntX, IntY, 0)
                        End If
                    End If
                End If
            End If
        Next
    Next
    If IntDrops < IntDropCols Then
        For IntMakeNew = IntDrops To IntDropCols
            IntX = Int(Rnd * XBefore(0)) + 1
            IntY = Int(Rnd * 5) + 1
            If BoolUsed(IntX, IntY) = True Then
                Call Clear(IntX, IntY, 0)
            End If
            LOD(IntX, IntY) = Int(Rnd * IntMaxLength)
            Lead(IntX, IntY) = 1
            Letter(IntX, IntY) = 64 + Int(Rnd * 26)
            BoolUsed(IntX, IntY) = True
            IntColour(1, IntX, IntY) = Rnd * 100 + 100
            IntColour(2, IntX, IntY) = 255
            Call ShowHigh(IntX, IntY, 0)
        Next
    End If


    Font.SIZE = Size2
    IntDrops = 0
    For IntX = 1 To XBefore(1)
        For IntY = 1 To YBefore(1) + 5
            If BoolUsed1(IntX, IntY) <> 0 Then
                If Lead1(IntX, IntY) = 1 Then 'Is it Lead
                    If IntY <= YBefore(1) + 4 Then  'Is it smaller than the screen height
                        If LOD1(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If Letter1(IntX, IntY + 1) <> 0 Then
                                Call Clear(IntX, IntY + 1, 1)
                            End If
                            LOD1(IntX, IntY + 1) = LOD1(IntX, IntY) - 1
                            Lead1(IntX, IntY) = 0
                            Lead1(IntX, IntY + 1) = 2
                            Letter1(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            BoolUsed1(IntX, IntY + 1) = True
                            IntColour1(1, IntX, IntY + 1) = Rnd * 100 + 100
                            IntColour1(2, IntX, IntY + 1) = 255
                            Wait1(IntX, IntY) = IntMaxLngWait / 2 + Rnd(IntMaxLngWait / 2)
                            Call ShowHigh(IntX, IntY + 1, 1)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If Lead1(IntX, IntY) = 1 Then
                                Call Clear(IntX, IntY, 1)
                            End If
                        End If
                    Else
                        Call Clear(IntX, IntY, 1)
                    End If
                ElseIf Lead1(IntX, IntY) = 2 Then
                    Lead1(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                ElseIf Lead1(IntX, IntY) = 3 Then
                End If
                If Wait1(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    Wait1(IntX, IntY) = Wait1(IntX, IntY) - 1
                    
                    If Wait1(IntX, IntY) = 0 Or IntColour1(1, IntX, IntY) = 0 Then
                        Call Clear(IntX, IntY, 1)
                    Else
                        IntColour1(1, IntX, IntY) = IntColour1(1, IntX, IntY) - (IntFadeSpeed + 1)
                        IntColour1(2, IntX, IntY) = IntColour1(2, IntX, IntY) - (IntFadeSpeed + 1) * 2
                        If IntColour1(1, IntX, IntY) < 0 Then
                            IntColour1(1, IntX, IntY) = 0
                        End If
                        If IntColour1(2, IntX, IntY) < 0 Then
                            IntColour1(2, IntX, IntY) = 0
                        End If
                        If IntColour1(1, IntX, IntY) = 0 Then
                            Call Clear(IntX, IntY, 1)
                        ElseIf Lead1(IntX, IntY) = 0 Then
                            Call ShowColor(IntX, IntY, 1)
                        End If
                    End If
                End If
            End If
        Next
    Next
    If IntDrops < IntDropCols Then
        For IntMakeNew = IntDrops To IntDropCols
            IntX = Int(Rnd * XBefore(1)) + 1
            IntY = Int(Rnd * 5) + 1
            If BoolUsed1(IntX, IntY) = True Then
                Call Clear(IntX, IntY, 1)
            End If

            LOD1(IntX, IntY) = Int(Rnd * IntMaxLength)
            Lead1(IntX, IntY) = 1
            Letter1(IntX, IntY) = 64 + Int(Rnd * 26)
            BoolUsed1(IntX, IntY) = True
            IntColour1(1, IntX, IntY) = Rnd * 100 + 100
            IntColour1(2, IntX, IntY) = 255
            Call ShowHigh(IntX, IntY, 1)
        Next
    End If
    
    
    Font.SIZE = Size3
    IntDrops = 1
    For IntX = 1 To XBefore(2)
        For IntY = 1 To YBefore(2) + 5
            If BoolUsed2(IntX, IntY) <> 0 Then
                If Lead2(IntX, IntY) = 1 Then 'Is it Lead
                    If IntY <= YBefore(2) + 4 Then  'Is it smaller than the screen height
                        If LOD2(IntX, IntY) > 0 Then 'Is there still IntDrops in this column
                            If Letter2(IntX, IntY + 1) <> 0 Then
                                Call Clear(IntX, IntY + 1, 2)
                            End If
                            LOD2(IntX, IntY + 1) = LOD2(IntX, IntY) - 1
                            Lead2(IntX, IntY) = 0
                            Lead2(IntX, IntY + 1) = 2
                            Letter2(IntX, IntY + 1) = Int(Rnd * 43) + 65
                            BoolUsed2(IntX, IntY + 1) = True
                            IntColour2(1, IntX, IntY + 1) = Rnd * 100 + 100
                            IntColour2(2, IntX, IntY + 1) = 255
                            Wait2(IntX, IntY) = IntMaxLngWait / 2 + Rnd(IntMaxLngWait / 2)
                            Call ShowHigh(IntX, IntY + 1, 2)
                        Else    'End of Drop(Kill Letter/Symbol)
                            If Lead2(IntX, IntY) = 1 Then
                                Call Clear(IntX, IntY, 2)
                            End If
                        End If
                    Else
                        Call Clear(IntX, IntY, 2)
                    End If
                ElseIf Lead2(IntX, IntY) = 2 Then
                    Lead2(IntX, IntY) = 1
                    IntDrops = IntDrops + 1
                ElseIf Lead2(IntX, IntY) = 3 Then
                End If
                If Wait2(IntX, IntY) > 0 Then 'Is the Letter/Symbol dieing?
                    Wait2(IntX, IntY) = Wait2(IntX, IntY) - 1
                    
                    If Wait2(IntX, IntY) = 0 Or IntColour2(1, IntX, IntY) = 0 Then
                        Call Clear(IntX, IntY, 2)
                    Else
                        IntColour2(1, IntX, IntY) = IntColour2(1, IntX, IntY) - (IntFadeSpeed + 2)
                        IntColour2(2, IntX, IntY) = IntColour2(2, IntX, IntY) - (IntFadeSpeed + 2) * 2
                        If IntColour2(1, IntX, IntY) < 0 Then
                            IntColour2(1, IntX, IntY) = 0
                        End If
                        If IntColour2(2, IntX, IntY) < 0 Then
                            IntColour2(2, IntX, IntY) = 0
                        End If
                        If IntColour2(1, IntX, IntY) = 0 Then
                            Call Clear(IntX, IntY, 2)
                        ElseIf Lead2(IntX, IntY) = 0 Then
                            Call ShowColor(IntX, IntY, 2)
                        End If
                    End If
                End If
            End If
        Next
    Next
    If IntDrops < IntDropCols Then
        For IntMakeNew = IntDrops To IntDropCols
            IntX = Int(Rnd * XBefore(2)) + 1
            IntY = Int(Rnd * 5) + 1
            If BoolUsed2(IntX, IntY) = True Then
                Call Clear(IntX, IntY, 2)
            End If
            LOD2(IntX, IntY) = Int(Rnd * IntMaxLength)
            Lead2(IntX, IntY) = 1
            Letter2(IntX, IntY) = 64 + Int(Rnd * 26)
            BoolUsed2(IntX, IntY) = True
            IntColour2(1, IntX, IntY) = Rnd * 100 + 100
            IntColour2(2, IntX, IntY) = 255
            Call ShowHigh(IntX, IntY, 2)
        Next
    End If
End Sub

Sub Clear(IntX, IntY, Index) 'Clears a letter by redrawing it as black
    If Index = 0 Then
        BoolUsed(IntX, IntY) = False
        Wait(IntX, IntY) = 0
        IntColour(1, IntX, IntY) = 0
        Lead(IntX, IntY) = 0
    ElseIf Index = 1 Then
        BoolUsed1(IntX, IntY) = False
        Wait1(IntX, IntY) = 0
        IntColour1(1, IntX, IntY) = 0
        Lead1(IntX, IntY) = 0
    ElseIf Index = 2 Then
        BoolUsed2(IntX, IntY) = False
        Wait2(IntX, IntY) = 0
        IntColour2(1, IntX, IntY) = 0
        Lead2(IntX, IntY) = 0
    End If
End Sub

Sub ShowHigh(IntX, IntY, Index) 'Shows a highlighted letter
    If IntY - 4 < 0 Then Exit Sub
    FrmMatrix.ForeColor = vbWhite
    If Index = 0 Then
        Call TextOut(FrmMatrix.hdc, CW / XBefore(0) * IntX, CH / YBefore(0) * (IntY - 4), Chr(Letter(IntX, IntY)), 1)
    ElseIf Index = 1 Then
        Call TextOut(FrmMatrix.hdc, CW / XBefore(1) * IntX, CH / YBefore(1) * (IntY - 4), Chr(Letter1(IntX, IntY)), 1)
    ElseIf Index = 2 Then
        Call TextOut(FrmMatrix.hdc, CW / XBefore(2) * IntX, CH / YBefore(2) * (IntY - 4), Chr(Letter2(IntX, IntY)), 1)
    End If
End Sub

Sub ShowColor(IntX, IntY, Index) 'Shows a Coloured letter
    If IntY - 4 < 0 Then Exit Sub
    If Index = 0 Then
        FrmMatrix.ForeColor = RGB(IntColour(2, IntX, IntY), IntColour(2, IntX, IntY) + IntColour(1, IntX, IntY), IntColour(2, IntX, IntY))
        Call TextOut(FrmMatrix.hdc, CW / XBefore(0) * IntX, CH / YBefore(0) * (IntY - 4), Chr(Letter(IntX, IntY)), 1)
    ElseIf Index = 1 Then
        FrmMatrix.ForeColor = RGB(IntColour1(2, IntX, IntY), IntColour1(2, IntX, IntY) + IntColour1(1, IntX, IntY), IntColour1(2, IntX, IntY))
        Call TextOut(FrmMatrix.hdc, CW / XBefore(1) * IntX, CH / YBefore(1) * (IntY - 4), Chr(Letter1(IntX, IntY)), 1)
    ElseIf Index = 2 Then
        FrmMatrix.ForeColor = RGB(IntColour2(2, IntX, IntY), IntColour2(2, IntX, IntY) + IntColour2(1, IntX, IntY), IntColour2(2, IntX, IntY))
        Call TextOut(FrmMatrix.hdc, CW / XBefore(2) * IntX, CH / YBefore(2) * (IntY - 4), Chr(Letter2(IntX, IntY)), 1)
    End If
End Sub

Sub SetBefore(Index)
    XBefore(Index) = FrmMatrix.ScaleWidth
    YBefore(Index) = FrmMatrix.ScaleHeight
    FrmMatrix.ScaleMode = 3
End Sub


