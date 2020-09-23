VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Assessment XP"
   ClientHeight    =   6240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8415
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicFade 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   6135
      LargeChange     =   20
      Left            =   8160
      SmallChange     =   10
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   -120
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   549
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.PictureBox PicFrame 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   120
         ScaleHeight     =   393
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   536
         TabIndex        =   1
         Top             =   120
         Width           =   8040
         Begin VB.PictureBox PicLabel 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   265
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.Label LblRes 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Result"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label LblMaxVal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   6900
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Shape ShpScore 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            Height          =   255
            Index           =   0
            Left            =   5280
            Top             =   360
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label LblResult 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "50"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   2
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Line LineS1 
            Index           =   0
            Visible         =   0   'False
            X1              =   352
            X2              =   528
            Y1              =   16
            Y2              =   16
         End
         Begin VB.Line LineS2 
            Index           =   0
            Visible         =   0   'False
            X1              =   528
            X2              =   528
            Y1              =   16
            Y2              =   0
         End
      End
   End
   Begin VB.Menu MNUFIle 
      Caption         =   "File"
      Begin VB.Menu MNUSpecs 
         Caption         =   "Computer Specs"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu MNUExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MNUTests 
      Caption         =   "Test"
      Begin VB.Menu MNURAT 
         Caption         =   "Run All Tests"
      End
   End
   Begin VB.Menu MNUHelp 
      Caption         =   "Help"
      Begin VB.Menu MNUAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MNUResults 
         Caption         =   "Results"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Before As Double



Const Rad = (2 * 3.14159265358979) / 360

Dim Scores() As String

Const Layers = 20
Const Dist = 10
Const Detail = 40

Const Connect = True
Const NormCon = True


Dim Radius As Integer
Dim XOffSet As Integer
Dim YOffSet As Integer

Dim Grid(Detail, Layers) As Double
Dim AngOff(Detail, Layers) As Double

Dim ctrX As Integer
Dim ctrY As Integer
Const col = 0
Dim a As Double
Dim b As Double
Dim c As Double
Dim choss As Double
Dim X As Double
Dim Y As Double
Dim xn As Double
Dim n As Double
Dim m As Double


Private Sub Form_Load()
    ReDim Scores(0) As String
    Call ColourPicFrame
    Picture1.BackColor = RGB(240, 240, 240)
    PicFrame.BackColor = RGB(240, 240, 240)
End Sub

Private Sub MNUAbout_Click()
    MsgBox "About:" & vbNewLine & vbNewLine & "Created by Kevin Pfister of DEP Online" & vbNewLine & "Web: www.dep.zion.me.uk" & vbNewLine & "Created in 2004" & vbNewLine & vbNewLine & "Version: " & App.Major & "." & App.Minor & "." & App.Revision, vbOKOnly, "System Assessment XP ~ About"
End Sub

Private Sub MNUExit_Click()
    End
End Sub

Sub NewObjects(Index)
    If LblRes.Count < Index + 1 Then
        Load LblRes(Index)
        Load LblResult(Index)
        Load PicLabel(Index)
        Load ShpScore(Index)
        Load LineS1(Index)
        Load LineS2(Index)
        Load LblMaxVal(Index)
        ReDim Preserve Scores(Index) As String
    End If
    LineS1(Index).Visible = True
    LineS2(Index).Visible = True
    LblMaxVal(Index).Visible = True
    LblRes(Index).Visible = True
    LblResult(Index).Visible = True
    PicLabel(Index).Visible = True
    ShpScore(Index).Visible = True
    LblRes(Index).Top = 50 * Index
    LblResult(Index).Top = 24 + 50 * Index
    PicLabel(Index).Top = 50 * Index
    ShpScore(Index).Top = 24 + 50 * Index
    LineS1(Index).Y1 = 16 + 50 * Index
    LineS1(Index).Y2 = 16 + 50 * Index
    LineS2(Index).Y1 = 50 * Index + 16
    LineS2(Index).Y2 = 50 * Index
    LblMaxVal(Index).Top = LblMaxVal(Index).Top + 50 * Index
    PicFrame.Height = 50 * (Index + 2)
    If PicFrame.Height > Picture1.Height Then
        VScroll1.Max = PicFrame.Height - Picture1.Height
        VScroll1.Enabled = True
    End If
End Sub

Sub ScoreObjects(Index, TestName, Score)
    PicLabel(Index).Cls
    Call BitBlt(PicLabel(Index).hdc, 0, 0, 265, 17, PicFade.hdc, 0, 0, vbSrcCopy)
    PicLabel(Index).Refresh
    PicLabel(Index).Print " " & TestName
    LblResult(Index).Caption = Round(Score, 2)
    Scores(Index) = TestName & "  ~  " & Round(Score, 2)
    If Score Mod 100 <> 0 Then
        LblMaxVal(Index).Caption = (Int(Score / 100) + 1) * 100
        ShpScore(Index).Width = 178 / ((Int(Score / 100) + 1) * 100) * Score
    Else
        LblMaxVal(Index).Caption = Round(Score, 0)
        ShpScore(Index).Width = 178
    End If
End Sub

Private Sub MNURAT_Click()
    Dim Rating As Double
    Dim Score As Double
    Rating = 0 'Starting Rating
    
    MsgBox "Performance Tests:" & vbNewLine & vbNewLine & "It is recommended that you close all other programs including your anti virus before attempting to run tests. This will allow the tests to give a more accurate result on your systems performance" & vbNewLine & vbNewLine & "Click ok to start tests", vbOKOnly, "System Assessment XP ~ Warning"
    
    'Run the tests in turn and gather the results
    
    DoEvents
    Score = MathAddTest
    LblRes(0).Visible = True
    LblResult(0).Visible = True
    PicLabel(0).Visible = True
    ShpScore(0).Visible = True
    LineS1(0).Visible = True
    LineS2(0).Visible = True
    LblMaxVal(0).Visible = True
    Call ScoreObjects(0, "Math - Addition", Score)
    Rating = Rating + Score


    DoEvents
    Score = MathSubTest
    Call NewObjects(1)
    Call ScoreObjects(1, "Math - Subtraction", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = MathDivIntTest
    Call NewObjects(2)
    Call ScoreObjects(2, "Math - Integer Division", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = MathMultIntTest
    Call NewObjects(3)
    Call ScoreObjects(3, "Math - Integer Multiplication", Score)
    Rating = Rating + Score


    DoEvents
    Score = MathDivFltTest
    Call NewObjects(4)
    Call ScoreObjects(4, "Math - Floating Point Division", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = MathMultFltTest
    Call NewObjects(5)
    Call ScoreObjects(5, "Math - Floating Point Multiplication", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = StringBuffer(20)
    Call NewObjects(6)
    Call ScoreObjects(6, "String - 20mb Buffer", Score)
    Rating = Rating + Score


    DoEvents
    Score = StringBuffer(40)
    Call NewObjects(7)
    Call ScoreObjects(7, "String - 40mb Buffer", Score)
    Rating = Rating + Score
    

    DoEvents
    Score = StringBuffer(60)
    Call NewObjects(8)
    Call ScoreObjects(8, "String - 60mb Buffer", Score)
    Rating = Rating + Score
    

    DoEvents
    Score = StringBuffer(80)
    Call NewObjects(9)
    Call ScoreObjects(9, "String - 80mb Buffer", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = GraphicsPsetTest
    Call NewObjects(10)
    Call ScoreObjects(10, "Graphics - PSet Test", Score)
    Rating = Rating + Score
    

    DoEvents
    Score = GraphicsPointTest
    Call NewObjects(11)
    Call ScoreObjects(11, "Graphics - Point Test", Score)
    Rating = Rating + Score


    DoEvents
    Score = GraphicsSetPixelVTest
    Call NewObjects(12)
    Call ScoreObjects(12, "Graphics - SetPixelV Test", Score)
    Rating = Rating + Score
    

    DoEvents
    Score = GraphicsGetPixelTest
    Call NewObjects(13)
    Call ScoreObjects(13, "Graphics - GetPixel Test", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = BitBltTest
    Call NewObjects(14)
    Call ScoreObjects(14, "Graphics - BitBlt", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = TransparentTest
    Call NewObjects(15)
    Call ScoreObjects(15, "Graphics - Transparency", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = DiskWriteTest
    Call NewObjects(16)
    Call ScoreObjects(16, "Disc - Write Speed", Score)
    Rating = Rating + Score


    DoEvents
    Score = DiskReadTest
    Call NewObjects(17)
    Call ScoreObjects(17, "Disc - Read Speed", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = MatrixTest
    Call NewObjects(18)
    Call ScoreObjects(18, "Combined Test - Matrix", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = GraphicsRenderTest
    Call NewObjects(19)
    Call ScoreObjects(19, "Combined Test - Trig Demo", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = Fractals
    Call NewObjects(20)
    Call ScoreObjects(20, "Combined Test - Fractal Patterns", Score)
    Rating = Rating + Score
    

    DoEvents
    Score = FactorTest
    Call NewObjects(21)
    Call ScoreObjects(21, "General - 50,000 Factors Test", Score)
    Rating = Rating + Score
    
    
    DoEvents
    Score = TextBoxWrite
    Call NewObjects(22)
    Call ScoreObjects(22, "Controls - Text Box", Score)
    Rating = Rating + Score


    DoEvents
    Score = ListBoxWrite
    Call NewObjects(23)
    Call ScoreObjects(23, "Controls - List Box", Score)
    Rating = Rating + Score
    
    DoEvents
    Call NewObjects(24)
    Call ScoreObjects(24, "Overall Maths Results", (Val(LblResult(0).Caption) + Val(LblResult(1).Caption) + Val(LblResult(2).Caption) + Val(LblResult(3).Caption) + Val(LblResult(4).Caption) + Val(LblResult(5).Caption)) / 6)
    
    DoEvents
    Call NewObjects(25)
    Call ScoreObjects(25, "Overall Graphics Results", (Val(LblResult(10).Caption) + Val(LblResult(11).Caption) + Val(LblResult(12).Caption) + Val(LblResult(13).Caption) + Val(LblResult(14).Caption) + Val(LblResult(15).Caption)) / 6)


    Call NewObjects(26)
    Call ScoreObjects(26, "Overall Results", Rating / 24)
    
    'Save the results
    
    If Len(App.Path) = 3 Then
        'Root Dir
        Open App.Path & "TestResults.txt" For Output As #1
        Print #1, "Performance Test Results"
        For X = 0 To UBound(Scores())
            Print #1, "Test " & X + 1 & " ~ " & Scores(X)
        Next
        Close
    Else
        Open App.Path & "\TestResults.txt" For Output As #1
        Print #1, "Performance Test Results"
        For X = 0 To UBound(Scores())
            Print #1, Scores(X)
        Next
        Close
    End If
       
    
    MsgBox "All Tests Completed:" & vbNewLine & vbNewLine & "Overall Results: " & Round(Rating / 24, 2), vbOKOnly, "System Assessment XP ~ Score"
End Sub

Function MathAddTest()
    Dim Num As Double
    Before = Timer
    
    Do
        Num = Num + 1
    Loop Until Timer - Before > 1
    
    MathAddTest = (100 / 2830000) * Num
    
End Function

Function MathSubTest()
    Dim Num As Double
    Before = Timer
    
    Do
        Num = Num - 1
    Loop Until Timer - Before > 1
    
    MathSubTest = (100 / 2880000) * Abs(Num)
    
End Function

Function MathDivIntTest()
    Dim Num As Double
    Dim CalcNum As Integer
    Before = Timer
    
    Do
        Num = Num + 1
        CalcNum = 100 / 37
    Loop Until Timer - Before > 1
    
    MathDivIntTest = (100 / 2730000) * Num
End Function

Function MathMultIntTest()
    Dim Num As Double
    Dim CalcNum As Integer
    Before = Timer
    
    Do
        Num = Num + 1
        CalcNum = 97 * 37
    Loop Until Timer - Before > 1
    
    MathMultIntTest = (100 / 2870000) * Num
End Function

Function MathDivFltTest()
    Dim Num As Double
    Dim CalcNum As Double
    Before = Timer
    
    Do
        Num = Num + 1
        CalcNum = 100.5 / 37.3
    Loop Until Timer - Before > 1
    
    MathDivFltTest = (100 / 2800000) * Num
End Function

Function MathMultFltTest()
    Dim Num As Double
    Dim CalcNum As Integer
    Before = Timer
    
    Do
        Num = Num + 1
        CalcNum = 97.7 * 37.8
    Loop Until Timer - Before > 1
    
    MathMultFltTest = (100 / 2730000) * Num
End Function

Function GraphicsPsetTest()
    Dim Num As Double
    FrmTestBox.Show
    DoEvents

    Before = Timer
    Do
        Num = Num + 1
        FrmTestBox.PSet (Rnd * FrmTestBox.Width, Rnd * FrmTestBox.Height), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Loop Until Timer - Before > 1
    Unload FrmTestBox
    GraphicsPsetTest = (100 / 200000) * Num
End Function

Function GraphicsPointTest()
    Dim Num As Double
    Dim CalcNum As Long
    FrmTestBox.Show
    DoEvents

    Before = Timer
    Do
        Num = Num + 1
        CalcNum = FrmTestBox.Point(Rnd * FrmTestBox.Width, Rnd * FrmTestBox.Height)
    Loop Until Timer - Before > 1
    Unload FrmTestBox
    GraphicsPointTest = (100 / 253000) * Num
End Function

Function GraphicsSetPixelVTest()
    Dim Num As Double
    FrmTestBox.Show
    DoEvents

    Before = Timer
    Do
        Num = Num + 1
        Call SetPixelV(FrmTestBox.hdc, Rnd * FrmTestBox.Width, Rnd * FrmTestBox.Height, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
    Loop Until Timer - Before > 1
    Unload FrmTestBox
    GraphicsSetPixelVTest = (100 / 535000) * Num
End Function

Function GraphicsGetPixelTest()
    Dim Num As Double
    Dim CalcNum As Long
    FrmTestBox.Show
    DoEvents

    Before = Timer
    Do
        Num = Num + 1
        CalcNum = GetPixel(FrmTestBox.hdc, Rnd * FrmTestBox.Width, Rnd * FrmTestBox.Height)
    Loop Until Timer - Before > 1
    Unload FrmTestBox
    GraphicsGetPixelTest = (100 / 700000) * Num
End Function

Function MatrixTest()
    FrmMatrix.Show
    DoEvents
    Before = Timer
    StartTime = Before
    Do
        DoEvents
    Loop Until Timer - Before > 5
    Unload FrmMatrix
    MatrixTest = (100 / 240) * FrameCount
End Function

Function DiskWriteTest()
    Dim Num As Double
    
    DoEvents
    Open "C:\Performance.tmp" For Output As #1
    Before = Timer
    Do
        Num = Num + 1
        Print #1, "Performance Test"
    Loop Until Timer - Before > 1
    Close
    Kill "C:\Performance.tmp"
    DiskWriteTest = 100 / 818000 * Num
End Function

Function DiskReadTest()
    Dim Buffer As String
    
    DoEvents
    Open "C:\Performance.tmp" For Output As #1
    For X = 1 To 20 * 1024
        Print #1, String(1024, " ")
    Next
    Close
    
    Before = Timer
    Open "C:\Performance.tmp" For Input As #1
        Buffer = Input(LOF(1), 1)
    Close
    DiskReadTest = Timer - Before
    Kill "C:\Performance.tmp"
    Buffer = ""
    DiskReadTest = 100 / DiskReadTest * 1.12
End Function

Function StringBuffer(Mb)
    Dim Buffer As String
    Before = Timer
    Buffer = String(Mb * 1024 * 1024, " ")
    StringBuffer = Timer - Before
    If Mb = 20 Then
        StringBuffer = 100 / StringBuffer * 0.08
    ElseIf Mb = 40 Then
        StringBuffer = 100 / StringBuffer * 0.16
    ElseIf Mb = 60 Then
        StringBuffer = 100 / StringBuffer * 0.23
    ElseIf Mb = 80 Then
        StringBuffer = 100 / StringBuffer * 0.41
    End If
    
    Buffer = ""
End Function

Function GraphicsRenderTest()
    DoEvents
    FrmTestBox.Show
    XOffSet = FrmTestBox.ScaleWidth / 2
    YOffSet = FrmTestBox.ScaleHeight / 2
    Radius = FrmTestBox.ScaleHeight / 2
    For X = 1 To Layers
        For Y = 1 To Detail
            Grid(Y, X) = Rnd * 1
            AngOff(Y, X) = Rad * (2 - (Rnd * 4))
        Next
    Next
    Running = True
    TotCount = 1
    OffSet = 0
    Before = Timer
    
    Do
        FrmTestBox.Cls
        For Y = Layers To 1 Step -1
            Call Render(Y)
        Next
        FrmTestBox.Refresh
        DoEvents
        For X = 2 To Layers
            For Y = 1 To Detail
                Grid(Y, X - 1) = Grid(Y, X)
                AngOff(Y, X - 1) = AngOff(Y, X)
            Next
        Next
        For Y = 1 To Detail
            Grid(Y, Layers) = Sin(Rad * TotCount / Y) * 5
            AngOff(Y, Layers) = Sin(Rad * TotCount / Y) * 10
        Next
        TotCount = TotCount + 1
        If TotCount > 720 Then
            TotCount = 0
            Running = False
        End If
    Loop Until Running = False
    TimeLen = Timer - Before
    Unload FrmTestBox
    GraphicsRenderTest = 100 / TimeLen * 12.5
End Function

Function FactorTest()
    Dim Factor As String
    Before = Timer
    Open "C:\Performance.tmp" For Output As #1
    For Counter = 1 To 50000
        DoEvents
        Factor = ""
        If Counter Mod 2 = 0 Then
                'IS EVEN
                I = 1
                While (I < Sqr(Counter))
                        If Counter Mod I = 0 Then
                                Factor = Factor + Str$(Counter / I) + ","
                        End If
                        I = I + 1
                Wend
        End If
        If Counter Mod 2 <> 0 Then
                'IS ODD
                I = 1
                While (I < Sqr(Counter))
                        If Counter Mod I = 0 Then
                                Factor = Factor + Str$(Counter / I) + ","
                        End If
                        I = I + 2
                Wend
        End If
        Factor = Factor + "1"
        Print #1, Factor
    Next
    Close
    TimeLen = Timer - Before
    FactorTest = 100 / TimeLen * 2.14
    Kill "C:\Performance.tmp"   'Remove the temp file
End Function

Function BitBltTest()
    Dim Num As Double
    FrmTestBox.Show
    DoEvents
    FrmTestBox.Cls
    Before = Timer
    Do
        Num = Num + 1
        Call BitBlt(FrmTestBox.hdc, Rnd * (2 * FrmTestBox.ScaleWidth) - FrmTestBox.ScaleWidth, Rnd * (2 * FrmTestBox.ScaleHeight) - FrmTestBox.ScaleHeight, 450, 100, FrmTestBox.PicDEP.hdc, 0, 0, vbSrcCopy)
        FrmTestBox.Refresh
    Loop Until Timer - Before > 1
    Unload FrmTestBox
    BitBltTest = 100 / 14300 * Num
End Function

Function TransparentTest()
    Dim Num1 As Byte
    FrmSplash.Show
    FrmSplash.TmrMain.Enabled = False
    DoEvents
    Before = Timer
    Num1 = 0
    For Num = 1 To 254
        Call Mache_Transparent(FrmSplash.hWnd, Num1)
        Num1 = Num1 + 1
    Next
    For Num = 254 To 1 Step -1
        Call Mache_Transparent(FrmSplash.hWnd, Num1)
        Num1 = Num1 - 1
    Next
    
    TimeLen = Timer - Before
    TransparentTest = 100 / TimeLen * 0.16
    Unload FrmSplash
End Function

Function TextBoxWrite()
    Dim Num As Double
    FrmControl.Show
    FrmControl.TxtCheck.Text = ""
    FrmControl.LstCheck.Clear
    DoEvents
    Before = Timer
    Do
        Num = Num + 1
        FrmControl.TxtCheck.Text = FrmControl.TxtCheck.Text & "<Entry>"
    Loop Until Timer - Before > 1
    Unload FrmControl
    TextBoxWrite = 100 / 610 * Num
End Function

Function ListBoxWrite()
    Dim Num As Double
    FrmControl.Show
    FrmControl.TxtCheck.Text = ""
    FrmControl.LstCheck.Clear
    DoEvents
    Before = Timer
    Do
        Num = Num + 1
        FrmControl.LstCheck.AddItem "<Entry>"
    Loop Until Timer - Before > 1
    Unload FrmControl
    ListBoxWrite = 100 / 20500 * Num
End Function

Function Fractals()
    FrmTestBox.Show
    a = 12
    b = 6
    c = 85
    ctrX = FrmTestBox.ScaleWidth / 2
    ctrY = FrmTestBox.ScaleHeight / 2
    DoEvents
    X = 0
    Y = 0
    Z = 0
    n = 0
    Before = Timer
    Do
        
        For Z = 1 To 1000
            If X > 0 Then
                xn = Y - Sqr(Abs(b * X - c))
            ElseIf X < 0 Then
                xn = Y + Sqr(Abs(b * X - c))
            Else
                xn = Y
            End If
            Y = a - X
            X = xn
            SetPixelV FrmTestBox.hdc, Int(X) / 2 + ctrX, Int(Y) / 2 + ctrY, n
            n = n + 1
        Next
        FrmTestBox.Refresh
        DoEvents
    Loop Until n >= 1000000
    TimeLen = Timer - Before
    Fractals = 100 / TimeLen * 1.63
    Unload FrmTestBox
End Function

Private Sub MNUResults_Click()
    MsgBox "Results:" & vbNewLine & "The results are compared to a benchmark pc that is given 100 on each test, the following are the specs to the benchmark pc" & vbNewLine & vbNewLine & "Athlon XP Barton 3000" & vbNewLine & "512mb DDR Ram" & vbNewLine & "ATI Radeon 9800 Pro 128mb Graphics" & vbNewLine & "Running XP Pro", vbOKOnly, "System Assessment XP ~ Results"
End Sub

Private Sub MNUSpecs_Click()
    FrmSpecs.Show
End Sub

Private Sub VScroll1_Change()
    PicFrame.Top = 8 - VScroll1.Value
End Sub

Sub Render(Depth)
    For X = 1 To Detail
        If X <> Detail Then
            FrmTestBox.Line (XOffSet + (Radius - Dist * Depth - 10 * Grid(X, Depth)) * Cos(Rad * (360 / Detail * X) + AngOff(X, Depth) + Rad * OffSet), YOffSet + (Radius - Dist * Depth - 10 * Grid(X, Depth)) * Sin(Rad * (360 / Detail * X) + AngOff(X, Depth) + Rad * OffSet))-(XOffSet + (Radius - Dist * Depth - 10 * Grid(X + 1, Depth)) * Cos(Rad * (360 / Detail * (X + 1)) + AngOff(X + 1, Depth) + Rad * OffSet), YOffSet + (Radius - Dist * Depth - 10 * Grid(X + 1, Depth)) * Sin(Rad * (360 / Detail * (X + 1)) + AngOff(X + 1, Depth) + Rad * OffSet)), vbWhite
        Else
            FrmTestBox.Line (XOffSet + (Radius - Dist * Depth - 10 * Grid(Detail, Depth)) * Cos(Rad * 360 + AngOff(Detail, Depth) + Rad * OffSet), YOffSet + (Radius - Dist * Depth - 10 * Grid(Detail, Depth)) * Sin(Rad * 360 + AngOff(Detail, Depth) + Rad * OffSet))-(XOffSet + (Radius - Dist * Depth - 10 * Grid(1, Depth)) * Cos(Rad * (360 / Detail) + AngOff(1, Depth) + Rad * OffSet), YOffSet + (Radius - Dist * Depth - 10 * Grid(1, Depth)) * Sin(Rad * (360 / Detail) + AngOff(1, Depth) + Rad * OffSet)), vbWhite
        End If
        If Depth <> 1 Then
            FrmTestBox.Line (XOffSet + (Radius - Dist * Depth - 10 * Grid(X, Depth)) * Cos(Rad * (360 / Detail * X) + AngOff(X, Depth) + Rad * OffSet), YOffSet + (Radius - Dist * Depth - 10 * Grid(X, Depth)) * Sin(Rad * (360 / Detail * X) + AngOff(X, Depth) + Rad * OffSet))-(XOffSet + (Radius - Dist * (Depth - 1) - 10 * Grid(X, Depth - 1)) * Cos(Rad * (360 / Detail * X) + AngOff(X, Depth - 1) + Rad * OffSet), YOffSet + (Radius - Dist * (Depth - 1) - 10 * Grid(X, Depth - 1)) * Sin(Rad * (360 / Detail * X) + AngOff(X, Depth - 1) + Rad * OffSet)), vbWhite
        End If
    Next
End Sub

Sub ColourPicFrame()
    DoEvents
    For X = 1 To 265
        Colour = RGB(240, 140 + 100 / 265 * X, 10 + 230 / 265 * X)
        PicFade.Line (X, 0)-(X, 20), Colour
    Next
    DoEvents
End Sub
