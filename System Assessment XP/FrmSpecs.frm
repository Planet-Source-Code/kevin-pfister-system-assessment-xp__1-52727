VERSION 5.00
Begin VB.Form FrmSpecs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer Specs"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   Icon            =   "FrmSpecs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "Done"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label LblUsed 
      BackStyle       =   0  'Transparent
      Caption         =   "Used"
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
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label LblUsed 
      BackStyle       =   0  'Transparent
      Caption         =   "Used"
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
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label LblUsed 
      BackStyle       =   0  'Transparent
      Caption         =   "Used"
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
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.Line LineS2 
      Index           =   2
      X1              =   376
      X2              =   376
      Y1              =   296
      Y2              =   280
   End
   Begin VB.Line LineS1 
      Index           =   2
      X1              =   8
      X2              =   376
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Shape ShpScore 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   120
      Top             =   4560
      Width           =   5535
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
      Index           =   2
      Left            =   4620
      TabIndex        =   15
      Top             =   4200
      Width           =   975
   End
   Begin VB.Line LineS2 
      Index           =   1
      X1              =   376
      X2              =   376
      Y1              =   184
      Y2              =   168
   End
   Begin VB.Line LineS1 
      Index           =   1
      X1              =   8
      X2              =   376
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Shape ShpScore 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   120
      Top             =   2880
      Width           =   5535
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
      Index           =   1
      Left            =   4620
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.Line LineS2 
      Index           =   0
      X1              =   376
      X2              =   376
      Y1              =   72
      Y2              =   56
   End
   Begin VB.Line LineS1 
      Index           =   0
      X1              =   8
      X2              =   376
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Shape ShpScore 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   1200
      Width           =   5535
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
      Left            =   4620
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Physical Memory:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblPhysMem 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label lblPageFile 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3120
      TabIndex        =   9
      Top             =   1800
      Width           =   2475
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Page File:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2115
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Page File:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2115
   End
   Begin VB.Label lblAvailPage 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3120
      TabIndex        =   6
      Top             =   2160
      Width           =   2475
   End
   Begin VB.Label lblTotalVirtual 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3120
      TabIndex        =   5
      Top             =   3480
      Width           =   2475
   End
   Begin VB.Label Label100 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Virtual Memory:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Virtual Memory:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label lblAvailVirtual 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3120
      TabIndex        =   2
      Top             =   3840
      Width           =   2475
   End
   Begin VB.Label lblMemAvail 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   2475
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Physical Memory:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2835
   End
End
Attribute VB_Name = "FrmSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ms As MEMORYSTATUS

Private Sub CmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ms.dwLength = Len(ms)
    GlobalMemoryStatus ms
    lblPhysMem = Format$(ms.dwTotalPhys, "#,###,###,##0 ")
    lblMemAvail = Format$(ms.dwAvailPhys, "#,###,###,##0") & " (" & Format$(ms.dwMemoryLoad / 100, "##0%") & " FREE)"
    lblPageFile = Format$(ms.dwTotalPageFile, "#,###,###,##0 ")
    lblAvailPage = Format$(ms.dwAvailPageFile, "#,###,###,##0 ")
    lblTotalVirtual = Format$(ms.dwTotalVirtual, "#,###,###,##0 ")
    lblAvailVirtual = Format$(ms.dwAvailVirtual, "#,###,###,##0 ")
    
    
    ShpScore(0).Width = 369 / ms.dwTotalPhys * ms.dwAvailPhys
    ShpScore(1).Width = 369 / ms.dwTotalPageFile * (ms.dwTotalPageFile - ms.dwAvailPageFile)
    ShpScore(2).Width = 369 / ms.dwTotalVirtual * (ms.dwTotalVirtual - ms.dwAvailVirtual)
End Sub
