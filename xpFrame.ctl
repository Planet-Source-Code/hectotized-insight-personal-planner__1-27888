VERSION 5.00
Begin VB.UserControl xpFrame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "xpFrame.ctx":0000
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   0
      Width           =   975
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   45
         TabIndex        =   2
         Top             =   0
         Width           =   465
      End
   End
End
Attribute VB_Name = "xpFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************
'*  Copyright (C) Outersoft Inc. 2001 - All Rights Reserved        *
'*                                                                 *
'*  FILE:  xpframe.ctl                                             *
'*                                                                 *
'*  DESCRIPTION:                                                   *
'*      A first attempt to create xp theme aware controls          *
'*                                                                 *
'*  SPECIAL THNX:                                                  *
'*        The KPD-Team                                             *
'*        Visit them at http://www.allapi.net/                     *
'*                                                                 *
'*******************************************************************


Option Explicit

'control defaults
Private Const d_Text = "xpFrame"
Private Const d_Enabled = True
Private Const d_Font = Null ' Ambient.Font
Private Const d_Alignment = 0
Private Const d_BackColor = vbWhite
Private Const d_FontColor = vbBlack



Private Const cNormal = GBS_NORMAL

Private c_Font As Font
Private c_FontColor As OLE_COLOR
Private c_BackColor As OLE_COLOR
Private c_Text As String
Private c_Enabled As Boolean


Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long




Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Event Click()

 
Private hTheme As Long
Private cb As RECT, tx As RECT

Private Sub DrawXpControl(nState As Long)
    
    'clear control
    Cls
    
    'set control position
    SetRect cb, 0, 7, UserControl.Width / 15, UserControl.Height / 15
    
    
    'open theme data
    hTheme = OpenThemeData(UserControl.hWnd, "BUTTON")
    
    'if control is enabled we draw normal, if not, then as disabled
    If c_Enabled = True Then 'control enabled
        'draw the frame
        DrawThemeBackground hTheme, UserControl.hDC, BP_GROUPBOX, nState, cb, ByVal 0&
    Else 'control disabled
        DrawThemeBackground hTheme, UserControl.hDC, BP_GROUPBOX, GBS_DISABLED, cb, ByVal 0&
    End If
    
    'show or hide the fakness
    If c_Text = "" Then
    Picture1.Visible = False
    
    Else
    Picture1.Visible = True
    Label1.Caption = c_Text & "  "
    frame1.Caption = c_Text
    Picture1.Width = Label1.Width
    End If
    
    'close theme data
    CloseThemeData hTheme
End Sub
Public Property Get Caption() As String
    Caption = c_Text
End Property

Public Property Let Caption(ByVal nText As String)
    c_Text = nText
    Label1.Caption = c_Text
    frame1.Caption = c_Text
    UserControl_Resize
    PropertyChanged "Caption"
End Property
Public Property Get Enabled() As Boolean
    Enabled = c_Enabled
End Property
Public Property Let Enabled(ByVal nEnabled As Boolean)
    c_Enabled = nEnabled
    UserControl.Enabled = c_Enabled
    Label1.Enabled = c_Enabled
    frame1.Enabled = c_Enabled
    
    Call UserControl_Resize
PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
    Set Font = c_Font
End Property
Public Property Set Font(ByVal vData As Font)
    Set c_Font = vData
    Set frame1.Font = c_Font
    Set Label1.Font = c_Font
    UserControl_Resize
PropertyChanged "Font"
End Property
Public Property Get FontColor() As OLE_COLOR
FontColor = c_FontColor
End Property

Public Property Let FontColor(ByVal New_Color As OLE_COLOR)
c_FontColor = New_Color
frame1.ForeColor = c_FontColor
Label1.ForeColor = c_FontColor
PropertyChanged "FontColor"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = c_BackColor
End Property
Public Property Let BackColor(ByVal vData As OLE_COLOR)
c_BackColor = vData

UserControl.BackColor = c_BackColor
Picture1.BackColor = c_BackColor
frame1.BackColor = c_BackColor

Call UserControl_Resize
PropertyChanged "BackColor"
End Property

Private Sub Frame1_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()

    'We either draw as xp or vb standard and set focus accordingly
    If DrawAsXp = True Then
        frame1.Visible = False
        Picture1.Visible = True
    Else

        With frame1
        .Visible = True
        .Width = Width
        .Height = Height
        End With
        Picture1.Visible = False
    End If
    
    UserControl_Resize
    
End Sub

Private Sub UserControl_InitProperties()

Set Font = Ambient.Font

Caption = d_Text
Enabled = d_Enabled
BackColor = d_BackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", d_Text)
    Enabled = PropBag.ReadProperty("Enabled", d_Enabled)
    BackColor = PropBag.ReadProperty("BackColor", d_BackColor)
    ForeColor = PropBag.ReadProperty("FontColor", d_FontColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", c_Text, d_Text)
    Call PropBag.WriteProperty("Enabled", c_Enabled, d_Enabled)
    Call PropBag.WriteProperty("BackColor", c_BackColor, d_BackColor)
    Call PropBag.WriteProperty("FontColor", c_FontColor, d_FontColor)
    Call PropBag.WriteProperty("Font", c_Font, Ambient.Font)

End Sub
Private Sub UserControl_Resize()

    'We either draw as xp or vb standard and set focus accordingly
    If DrawAsXp = True Then
        DrawXpControl cNormal
    End If
    
    Picture1.Width = Label1.Width
    frame1.Left = 0
    frame1.Top = 0
    frame1.Width = UserControl.Width
    frame1.Height = UserControl.Height
End Sub

'this function is used to determin if os is xp and themes supported and if themes are on
Public Function DrawAsXp() As Boolean
    Dim hLib As Long
    Dim OFind As Boolean, TFind As Boolean, TOn As Boolean
    
    'check to see if theme supported
    hLib = LoadLibrary("uxtheme.dll")
    If hLib <> 0 Then FreeLibrary hLib
    TFind = Not (hLib = 0)

   'now we check to see if windows themes or windows classic is enabled
    If OpenThemeData(UserControl.hWnd, "BUTTON") > 0 Then TOn = True

'   Now we know if we draw as xp or vb standard
    If TFind = False Or TOn = False Then
        DrawAsXp = False
    Else
        DrawAsXp = True
    End If

End Function

