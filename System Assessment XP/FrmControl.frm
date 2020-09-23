VERSION 5.00
Begin VB.Form FrmControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Assessment XP - Control Test Box"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicCheck 
      Height          =   2415
      Left            =   2760
      ScaleHeight     =   2355
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ListBox LstCheck 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox TxtCheck 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "FrmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
