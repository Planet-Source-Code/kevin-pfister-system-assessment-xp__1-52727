VERSION 5.00
Begin VB.Form FrmTestBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Assessment XP - Graphics Test Box"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   Icon            =   "FrmTestBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicDEP 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   240
      Picture         =   "FrmTestBox.frx":0442
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   7200
      Width           =   6750
   End
End
Attribute VB_Name = "FrmTestBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
