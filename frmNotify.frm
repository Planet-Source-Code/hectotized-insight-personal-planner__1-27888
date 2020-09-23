VERSION 5.00
Begin VB.Form frmNotify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insight Instant Notice"
   ClientHeight    =   2130
   ClientLeft      =   9705
   ClientTop       =   7950
   ClientWidth     =   5280
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFX 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1080
   End
   Begin Insight.xpCommand cmdOK 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "OK"
   End
   Begin Insight.xpCheck cbxAlwaysPlay 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      Caption         =   "Always Allow MP3 To Finish Playing."
      Value           =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   5160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Text Here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmNotify.frx":0000
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Private tb_Action As String
Dim I

Private Sub cbxAlwaysPlay_Click()
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AlwaysPlay", cbxAlwaysPlay.Value

End Sub

Private Sub cmdOK_Click()
SetTrans (220)
tb_Action = "FO"
tmrFX.Enabled = True
'user didnt select to allow to finish playing mp3, so we stop it
If cbxAlwaysPlay.Value = 0 Or cbxAlwaysPlay.Enabled = False Then frmMain.MediaPlayer1.Stop
End Sub

Private Sub Form_Load()

'check to see if media file is mp3 so we can give option to continue playing,
'since we dont want to continue looping a windows sound event, it will never turn stop.
If Right(frmMain.MediaPlayer1.FileName, 1) = "3" Then cbxAlwaysPlay.Enabled = True

'load always play value
cbxAlwaysPlay.Value = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AlwaysPlay")
SetTrans (0)
tb_Action = "FI"
tmrFX.Enabled = True
End Sub

Private Sub tmrFX_Timer()
Select Case tb_Action

Case "FI"


For I = 0 To 240
I = I + 4
If I >= 255 Then I = 255

SetTrans (I)
Next I


tmrFX.Enabled = False

Case "FO"
I = 240
retry:
If I <= 0 Then
Me.Hide
tmrFX.Enabled = False
Else

SetTrans (I)
I = I - 4
GoTo retry
End If

End Select
End Sub

Public Sub SetTrans(nLevel As Byte)

        SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hWnd, 0, nLevel, LWA_ALPHA

End Sub

