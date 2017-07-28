VERSION 5.00
Begin VB.Form frmGameOver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Over"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGameOver.frx":0000
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1080
      Top             =   480
   End
   Begin VB.Label lblScore8 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 8:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   6300
      TabIndex        =   11
      Top             =   3900
      Width           =   2025
   End
   Begin VB.Label lblScore7 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 7:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   6300
      TabIndex        =   10
      Top             =   3450
      Width           =   2025
   End
   Begin VB.Label lblTotalScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Score:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   1050
      TabIndex        =   9
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label lblScore6 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 6:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   6300
      TabIndex        =   8
      Top             =   3000
      Width           =   2025
   End
   Begin VB.Label lblScore5 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 5:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   6300
      TabIndex        =   7
      Top             =   2550
      Width           =   2025
   End
   Begin VB.Label lblScore4 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 4:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   2175
      TabIndex        =   6
      Top             =   3900
      Width           =   2025
   End
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press R to retry..."
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   2475
      TabIndex        =   5
      Top             =   6240
      Width           =   5550
   End
   Begin VB.Label lblScore3 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 3:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   2175
      TabIndex        =   4
      Top             =   3450
      Width           =   2025
   End
   Begin VB.Label lblScore2 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 2:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   2175
      TabIndex        =   3
      Top             =   3000
      Width           =   2025
   End
   Begin VB.Label lblScore1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   2175
      TabIndex        =   2
      Top             =   2550
      Width           =   2025
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game over... Your score:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2475
      TabIndex        =   1
      Top             =   1680
      Width           =   5550
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blockoid"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   855
      Left            =   3900
      TabIndex        =   0
      Top             =   720
      Width           =   2700
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyR Then
    Score1 = 0
    Score2 = 0
    Score3 = 0
    Score4 = 0
    Score5 = 0
    Score6 = 0
    Score7 = 0
    Score8 = 0
    frmlvl1.Show
    Unload frmGameOver
  End If
End Sub
Private Sub Form_Load()
  lblScore1.Caption = lblScore1.Caption & " " & Score1
  lblScore2.Caption = lblScore2.Caption & " " & Score2
  lblScore3.Caption = lblScore3.Caption & " " & Score3
  lblScore4.Caption = lblScore4.Caption & " " & Score4
  lblScore5.Caption = lblScore5.Caption & " " & Score5
  lblScore6.Caption = lblScore6.Caption & " " & Score6
  lblScore7.Caption = lblScore7.Caption & " " & Score7
  lblScore8.Caption = lblScore8.Caption & " " & Score8
  lblTotalScore.Caption = lblTotalScore.Caption & " " & Score1 + Score2 + Score3 + Score4 + Score5 + Score6 + Score7 + Score8
End Sub

Private Sub Timer1_Timer()

If lblMsg2.Visible = True Then
  lblMsg2.Visible = False
Else
  lblMsg2.Visible = True
End If

End Sub
