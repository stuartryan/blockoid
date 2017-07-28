VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8520
      Top             =   1200
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   1320
      Picture         =   "frmMainMenu.frx":89DB6
      Top             =   3720
      Width           =   8100
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMainMenu.frx":A05F2
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1575
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   9000
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press any key to start..."
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
      TabIndex        =   1
      Top             =   6240
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
      Left            =   3840
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    frmlvl1.Show
    Unload Me
End Sub
Private Sub Timer1_Timer()
If lblMsg.Visible = True Then
  lblMsg.Visible = False
Else
  lblMsg.Visible = True
End If


End Sub
