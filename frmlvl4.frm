VERSION 5.00
Begin VB.Form frmlvl4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 4"
   ClientHeight    =   7500
   ClientLeft      =   16590
   ClientTop       =   7995
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmlvl4.frx":0000
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   2760
   End
   Begin VB.Timer tmrMovement 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   4800
      Top             =   1920
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Shape shpPlayer 
      BorderColor     =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   5040
      Top             =   4320
      Width           =   150
   End
End
Attribute VB_Name = "frmlvl4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Map(0 To 70, 0 To 50) As Byte
Dim PlayerCol As Byte
Dim PlayerRow As Byte
Dim Moving As Boolean
Dim NewCol As Integer
Dim NewRow As Integer
Dim MoveX As Integer
Dim MoveY As Integer
Dim StopAtCol As Byte
Dim StopAtRow As Byte
Dim Time As Byte
Option Explicit
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

tmrTime.Enabled = True

If KeyCode = vbKeyA Then
  If Moving = False Then
    PlayerCol = shpPlayer.Left / 10
    If PlayerCol > 0 Then
      NewCol = PlayerCol - 1
      If PlayerRow < 49 And PlayerRow > 0 And Map(NewCol, PlayerRow) <> 1 Then
        Moving = True
        MoveX = -1
        MoveY = 0
        StopAtCol = 0
        tmrMovement.Enabled = True
      Else
        If Map(NewCol, PlayerRow) <> 1 Then
          PlayerCol = NewCol
          If Map(PlayerCol, PlayerRow) = 2 Then
            tmrTime.Enabled = False
            tmrMovement.Enabled = False
            Score4 = Time * 40
            frmlvl5.Show
            Unload frmlvl4
          End If
          If Map(NewCol, PlayerRow) = 4 Then
            tmrTime.Enabled = False
            tmrMovement.Enabled = False
            Score4 = 0
            frmGameOver.Show
            Unload frmlvl4
          End If
        End If
      End If
    End If
  End If
End If
    
If KeyCode = vbKeyD Then
  If Moving = False Then
    PlayerCol = shpPlayer.Left / 10
    If PlayerCol < 69 Then
      NewCol = PlayerCol + 1
      If PlayerRow < 49 And PlayerRow > 0 And Map(NewCol, PlayerRow) <> 1 Then
        Moving = True
        MoveX = 1
        MoveY = 0
        StopAtCol = 69
        tmrMovement.Enabled = True
      Else
        If Map(NewCol, PlayerRow) <> 1 Then
          PlayerCol = NewCol
          If Map(PlayerCol, PlayerRow) = 2 Then
            tmrTime.Enabled = False
            tmrMovement.Enabled = False
            Score4 = Time * 40
            frmlvl5.Show
            Unload frmlvl4
          End If
          If Map(NewCol, PlayerRow) = 4 Then
            tmrTime.Enabled = False
            tmrMovement.Enabled = False
            Score4 = 0
            frmGameOver.Show
            Unload frmlvl4
          End If
        End If
      End If
    End If
  End If
End If

If KeyCode = vbKeyW Then
  If Moving = False Then
    PlayerRow = shpPlayer.Top / 10
    If PlayerRow > 0 Then
      NewRow = PlayerRow - 1
      If PlayerRow <= 49 And PlayerRow > 0 And Map(PlayerCol, NewRow) <> 1 Then
        Moving = True
        MoveX = 0
        MoveY = -1
        StopAtRow = 0
        StopAtCol = 100
        tmrMovement.Enabled = True
      Else
        If Map(PlayerCol, NewRow) <> 1 Then
          PlayerRow = NewRow
          If Map(PlayerCol, PlayerRow) = 2 Then
            tmrTime.Enabled = False
            tmrMovement.Enabled = False
            Score4 = Time * 40
            frmlvl5.Show
            Unload frmlvl4
          End If
          If Map(PlayerCol, NewRow) = 4 Then
            tmrTime.Enabled = False
            tmrMovement.Enabled = False
            Score4 = 0
            frmGameOver.Show
            Unload frmlvl4
          End If
        End If
      End If
    End If
  End If
End If

If KeyCode = vbKeyS Then
  If Moving = False Then
    PlayerRow = shpPlayer.Top / 10
    If PlayerRow < 49 Then
      NewRow = PlayerRow + 1
      If PlayerRow < 49 And PlayerRow >= 0 And Map(PlayerCol, NewRow) <> 1 Then
        Moving = True
        MoveX = 0
        MoveY = 1
        StopAtRow = 49
        StopAtCol = 100
        tmrMovement.Enabled = True
      Else
        If Map(PlayerCol, NewRow) <> 1 Then
          PlayerRow = NewRow
          If Map(PlayerCol, PlayerRow) = 2 Then
            tmrTime.Enabled = False
            tmrMovement.Enabled = False
            Score4 = Time * 40
            frmlvl5.Show
            Unload frmlvl4
          End If
          If Map(PlayerCol, NewRow) = 4 Then
            tmrTime.Enabled = False
            tmrMovement.Enabled = False
            Score4 = 0
            frmGameOver.Show
            Unload frmlvl4
          End If
        End If
      End If
    End If
  End If
End If

If KeyCode = vbKeyE Then
  If Moving = False Then
    If PlayerCol < 69 And PlayerRow > 0 Then
      NewCol = PlayerCol + 1
      NewRow = PlayerRow - 1
      If Map(NewCol, NewRow) <> 1 Then
        If PlayerRow > 0 And PlayerCol < 69 Then
          Moving = True
          MoveX = 1
          MoveY = -1
          StopAtRow = 0
          StopAtCol = 69
          tmrMovement.Enabled = True
        End If
      End If
    End If
  End If
End If

If KeyCode = vbKeyQ Then
  If Moving = False Then
    If PlayerCol > 0 And PlayerRow > 0 Then
      NewCol = PlayerCol - 1
      NewRow = PlayerRow - 1
      If Map(NewCol, NewRow) <> 1 Then
        If PlayerCol > 0 And PlayerCol Then
          Moving = True
          MoveX = -1
          MoveY = -1
          StopAtRow = 0
          StopAtCol = 0
          tmrMovement.Enabled = True
        End If
      End If
    End If
  End If
End If

If KeyCode = vbKeyZ Then
  If Moving = False Then
    If PlayerCol > 0 And PlayerRow < 49 Then
      NewCol = PlayerCol - 1
      NewRow = PlayerRow + 1
      If Map(NewCol, NewRow) <> 1 Then
        If PlayerRow < 49 And PlayerCol > 0 Then
          Moving = True
          MoveX = -1
          MoveY = 1
          StopAtRow = 49
          StopAtCol = 0
          tmrMovement.Enabled = True
        End If
      End If
    End If
  End If
End If

If KeyCode = vbKeyC Then
  If Moving = False Then
    If PlayerCol < 69 And PlayerRow < 49 Then
      NewCol = PlayerCol + 1
      NewRow = PlayerRow + 1
      If Map(NewCol, NewRow) <> 1 Then
        If PlayerRow < 49 And PlayerCol < 69 Then
          Moving = True
          MoveX = 1
          MoveY = 1
          StopAtRow = 49
          StopAtCol = 69
          tmrMovement.Enabled = True
        End If
      End If
    End If
  End If
End If

If Moving = False Then
  shpPlayer.Top = PlayerRow * 10
  shpPlayer.Left = PlayerCol * 10
End If

End Sub
Private Sub Form_Load()
Dim ColCounter As Byte
Dim RowCounter As Byte

For ColCounter = 0 To 70
  For RowCounter = 0 To 50
    Map(ColCounter, RowCounter) = 0
  Next RowCounter
Next ColCounter

For RowCounter = 0 To 12
  Map(14, RowCounter) = 4
Next RowCounter



For RowCounter = 37 To 49
  Map(14, RowCounter) = 4
Next RowCounter


For RowCounter = 3 To 44
  Map(24, RowCounter) = 1
  Map(23, RowCounter) = 3
  Map(25, RowCounter) = 3
Next RowCounter

Map(24, 2) = 3
Map(24, 45) = 3

For RowCounter = 0 To 5
  Map(44, RowCounter) = 4
Next RowCounter

For ColCounter = 44 To 69
  Map(ColCounter, 21) = 4
Next ColCounter

For ColCounter = 32 To 57
  Map(ColCounter, 40) = 4
Next ColCounter

For RowCounter = 0 To 18
  Map(50, RowCounter) = 4
Next RowCounter

For ColCounter = 58 To 69
  Map(ColCounter, 0) = 2
Next ColCounter


Time = 15


PlayerRow = 49
PlayerCol = 0
shpPlayer.Top = PlayerRow * 10
shpPlayer.Left = PlayerCol * 10

tmrMovement.Enabled = False

Moving = False
End Sub
Private Sub tmrMovement_Timer()
NewRow = PlayerRow + MoveY
NewCol = PlayerCol + MoveX

If NewRow >= 0 And NewRow <= 49 Then
  If NewCol >= 0 And NewCol <= 69 Then
    If Map(NewCol, NewRow) = 1 Then
          NewCol = PlayerCol
      NewRow = PlayerRow
      Moving = False
      tmrMovement.Enabled = False
    End If
    If Map(NewCol, NewRow) = 4 Then
      tmrTime.Enabled = False
      tmrMovement.Enabled = False
      Score4 = 0
      frmGameOver.Show
      Unload frmlvl4
    End If
    If Map(NewCol, NewRow) <> 3 Then
      PlayerCol = NewCol
      PlayerRow = NewRow
      shpPlayer.Top = PlayerRow * 10
      shpPlayer.Left = PlayerCol * 10
    Else
      PlayerCol = NewCol
      PlayerRow = NewRow
      shpPlayer.Top = PlayerRow * 10
      shpPlayer.Left = PlayerCol * 10
      Moving = False
      tmrMovement.Enabled = False
    End If
    If Map(PlayerCol, PlayerRow) = 2 Then
      tmrTime.Enabled = False
      tmrMovement.Enabled = False
      Score4 = Time * 40
      frmlvl5.Show
      Unload frmlvl4
    End If
    If PlayerRow = StopAtRow Or PlayerCol = StopAtCol Then
      Moving = False
      tmrMovement.Enabled = False
    End If
  End If
End If
End Sub
Private Sub tmrTime_Timer()
Time = Time - 1
lblTime.Caption = Time

If Time <= 10 Then
  lblTime.ForeColor = &HFF&
End If

If Time <= 0 Then
  Score4 = 0
  frmGameOver.Show
  Unload frmlvl4
End If
End Sub

