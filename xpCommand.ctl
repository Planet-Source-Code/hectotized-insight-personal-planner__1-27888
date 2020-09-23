VERSION 5.00
Begin VB.UserControl xpCommand 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   ControlContainer=   -1  'True
   ScaleHeight     =   3390
   ScaleWidth      =   7020
   ToolboxBitmap   =   "xpCommand.ctx":0000
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "xpCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************'
'*  Copyright (C) Outersoft Inc. 2001 - All Rights Reserved        *'
'*  outersoftinc@hotmail.com                                       *'
'*                                                                 *'
'*  FILE:  xpCommandButton.ctl                                     *'
'*                                                                 *'
'*  DESCRIPTION:                                                   *'
'*      A first attempt to create xp theme aware controls          *'
'*                                                                 *'
'*  SPECIAL THNX:                                                  *'
'*        The KPD-Team                                             *'
'*        Visit them at http://www.allapi.net/                     *'
'*                                                                 *'
'*******************************************************************'

Option Explicit


'control defaults
Private Const d_Text = "xpCommand"
Private Const d_Enabled = True
Private Const d_BackColor = vbButtonFace




Private Const cHot = PBS_HOT
Private Const cNormal = PBS_NORMAL
Private Const cPressed = PBS_PRESSED
Private Const cDefaulted = PBS_DEFAULTED

Private c_Text As String
Private c_Enabled As Boolean
Private c_Value As Boolean
Private c_BackColor As OLE_COLOR
Private c_Click As Boolean
Private c_Style As Long

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
    SetRect cb, 0, 0, UserControl.Width / 15, UserControl.Height / 15
    
    'set text position
    SetRect tx, 0, 0, UserControl.Width / 15, UserControl.Height / 15
    
    'open theme data
    hTheme = OpenThemeData(UserControl.hWnd, "BUTTON")
    
    'if control is enabled we draw normal, if not, then as disabled
    If c_Enabled = True Then 'control enabled
        'draw the Button
        DrawThemeBackground hTheme, UserControl.hDC, BP_PUSHBUTTON, nState, cb, ByVal 0&
        'draw text
        DrawThemeText hTheme, UserControl.hDC, BP_PUSHBUTTON, nState, c_Text, -1, DT_CENTER Or DT_NOCLIP Or DT_VCENTER Or DT_WORD_ELLIPSIS Or DT_SINGLELINE, 0, tx

    Else 'control disabled
        DrawThemeBackground hTheme, UserControl.hDC, BP_PUSHBUTTON, PBS_DISABLED, cb, ByVal 0&
        'draw text
        DrawThemeText hTheme, UserControl.hDC, BP_PUSHBUTTON, PBS_DISABLED, c_Text, -1, DT_CENTER Or DT_NOCLIP Or DT_VCENTER Or DT_WORD_ELLIPSIS Or DT_SINGLELINE, 0, tx

    End If
    
    
    
    'close theme data
    CloseThemeData hTheme
    
End Sub





Private Sub Command1_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    
    'We either draw as xp or vb standard and set focus accordingly
    If DrawAsXp = True Then
        Command1.Visible = False
    Else

        With Command1
        .Visible = True
        .Width = Width
        .Height = Height
        End With
    End If
    
    UserControl_Resize

End Sub

Private Sub UserControl_InitProperties()
    
Caption = d_Text
Enabled = d_Enabled
BackColor = d_BackColor

End Sub


Private Sub UserControl_LostFocus()

    'exit sub if control disabled
    If c_Enabled = False Then Exit Sub
    
    'We either draw as xp or vb standard and set focus accordingly
    ' DrawAsXp = True Then DrawXpControl cNormal

End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'if control is disabled or no themes then exit sub
    If c_Enabled = False Or DrawAsXp = False Then Exit Sub
    
    'release mouse position capture
    ReleaseCapture
    
    
    'draw button as pressed
    DrawXpControl cPressed

    'capture mouse position
    SetCapture UserControl.hWnd
    
    'we set this so that the mouse move event isnt fired
    c_Click = True
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'if control is disabled or no themes then exit sub
    If c_Enabled = False Or DrawAsXp = False Then Exit Sub
   
    'release mouse position capture
    ReleaseCapture
        
   'we set this so that the mouse move event will fire
    c_Click = False
        
    'draw button as hot
    DrawXpControl cHot

    'capture mouse position
    SetCapture UserControl.hWnd
    
    'raise mouse event
    RaiseEvent Click
    
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'if control is clicked, disabled or no themes then exit sub
    If c_Click = True Or c_Enabled = False Or DrawAsXp = False Then Exit Sub
    
        'if mouse within control then draw as hot, else draw as normal
        If X < 0 Or X > ScaleWidth Or Y < 0 Or Y > ScaleHeight Then
            ReleaseCapture
            DrawXpControl cNormal
        Else
            SetCapture UserControl.hWnd
            DrawXpControl cHot
        End If
    
End Sub

Private Sub UserControl_Resize()
    
    'We either draw as xp or vb standard and set focus accordingly
    If DrawAsXp = True Then
        DrawXpControl cNormal
    End If
    
    'we always resize the vb standard control just in case
    Command1.Width = UserControl.Width
    Command1.Height = UserControl.Height

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'i wont comment anything about the stuff below, pretty standard stuff
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", d_Text)
    Enabled = PropBag.ReadProperty("Enabled", d_Enabled)
    BackColor = PropBag.ReadProperty("BackColor", d_BackColor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", c_Text, d_Text)
    Call PropBag.WriteProperty("Enabled", c_Enabled, d_Enabled)
    Call PropBag.WriteProperty("BackColor", c_BackColor, d_BackColor)

End Sub

Public Property Get Caption() As String
    Caption = c_Text
End Property

Public Property Let Caption(ByVal nText As String)
    c_Text = nText
    Command1.Caption = c_Text
    UserControl_Resize
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = c_Enabled
End Property
Public Property Let Enabled(ByVal nEnabled As Boolean)
    c_Enabled = nEnabled
    
    
    UserControl.Enabled = c_Enabled
    'Command1.Enabled = c_Enabled
    
UserControl_Resize

PropertyChanged "Enabled"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = c_BackColor
End Property
Public Property Let BackColor(ByVal nColor As OLE_COLOR)
    c_BackColor = nColor
    Command1.BackColor = c_BackColor
    UserControl.BackColor = c_BackColor
    Call UserControl_Resize
PropertyChanged "BackColor"
End Property

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
