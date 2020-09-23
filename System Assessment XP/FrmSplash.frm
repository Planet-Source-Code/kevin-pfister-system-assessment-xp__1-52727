VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSplash.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer TmrMain 
      Interval        =   200
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FadeValue As Byte

Private Sub Form_Load()
    Call Mache_Transparent(FrmSplash.hWnd, 0)
    Me.Visible = True
End Sub

Private Sub TmrMain_Timer()
    If TmrMain.Interval = 200 Then
        TmrMain.Interval = 1
    ElseIf TmrMain.Interval = 100 Then
        FrmMain.Show
        Unload Me
    Else
        FadeValue = FadeValue + 2
        Call Mache_Transparent(FrmSplash.hWnd, FadeValue)
        If FadeValue = 254 Then
            TmrMain.Interval = 100
        End If
    End If
End Sub
