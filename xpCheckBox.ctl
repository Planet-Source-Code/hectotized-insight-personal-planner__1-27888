VERSION 5.00
Begin VB.UserControl xpCheck 
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
   ToolboxBitmap   =   "xpCheckBox.ctx":0000
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "xpCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************'
'*  Copyright (C) Outersoft Inc. 2001 - All Rights Reserved        *'
'*  outersoftinc@hotmail.com                                       *'
'*                                                                 *'
'*  FILE:  xpCheckbox.ctl                                          *'
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
Private Const d_Text = "xpCheckbox"
Private Const d_Enabled = True
Private Const d_BackColor = vbButtonFace
Private Const d_Value = False
Private Const d_Alignment = 0

Private Const cHot = CBS_CHECKEDHOT
Private Const cNormal = CBS_CHECKEDNORMAL
Private Const cPressed = CBS_CHECKEDPRESSED
Private Const uHot = CBS_UNCHECKEDHOT
Private Const uNormal = CBS_UNCHECKEDNORMAL
Private Const uPressed = CBS_UNCHECKEDPRESSED

Private c_Text As String
Private c_Enabled As Boolean
Private c_Value As Boolean
Private c_BackColor As OLE_COLOR
Private c_Click As Boolean
Private c_Style As Long
Private c_Alignment As Long

Public Enum chkAlign
        LeftJustify = 0
        RightJustify = 1
End Enum

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Event Click()
Event Change()
 
Private hTheme As Long
Private cb As RECT, tx As RECT


Private Sub DrawXpControl(nState As Long)
    
    'clear control
    Cls
    
    'we either draw control according to alignment
    If c_Alignment = 0 Then
        'set control position
        SetRect cb, 0, 0, 20, UserControl.Height / 15
        'set text position
        SetRect tx, 20, 0, UserControl.Width / 15, UserControl.Height / 15
    Else
        SetRect cb, 0, 0, UserControl.Width / 7.5 - 18, UserControl.Height / 15
        SetRect tx, 0, 0, UserControl.Width / 15 - 20, UserControl.Height / 15
    End If
    
    'open theme data
    hTheme = OpenThemeData(UserControl.hWnd, "BUTTON")
    
    'if control is enabled we draw normal, if not, then as disabled
    If c_Enabled = True Then 'control enabled
        'draw the checkbox
        DrawThemeBackground hTheme, UserControl.hDC, BP_CHECKBOX, nState, cb, ByVal 0&
        'draw text
        DrawThemeText hTheme, UserControl.hDC, BP_CHECKBOX, nState, c_Text, -1, DT_LEFT Or DT_VCENTER Or DT_WORD_ELLIPSIS Or DT_SINGLELINE, 0, tx

    Else 'control disabled
        If c_Value = True Then 'draw as disabled checked
            DrawThemeBackground hTheme, UserControl.hDC, BP_CHECKBOX, CBS_CHECKEDDISABLED, cb, ByVal 0&
        Else 'draw as disabled normal unchecked
            DrawThemeBackground hTheme, UserControl.hDC, BP_CHECKBOX, CBS_UNCHECKEDDISABLED, cb, ByVal 0&
        End If
        'draw text
        DrawThemeText hTheme, UserControl.hDC, BP_CHECKBOX, CBS_CHECKEDDISABLED, c_Text, -1, DT_LEFT Or DT_VCENTER Or DT_WORD_ELLIPSIS Or DT_SINGLELINE, 0, tx

    End If
    
    
    'close theme data
    CloseThemeData hTheme
End Sub
Private Sub Check1_Click()
    c_Value = Check1.Value
    RaiseEvent Click
    RaiseEvent Change
End Sub

Private Sub UserControl_Initialize()
    
    'We either draw as xp or vb standard and set focus accordingly
    If DrawAsXp = True Then
        Check1.Visible = False
    Else

        With Check1
        .Visible = True
        .Width = Width
        .Height = Height
        End With
    End If
    
    UserControl_Resize

End Sub

Private Sub UserControl_InitProperties()
    
    'load default settings
    Caption = d_Text
    Enabled = d_Enabled
    BackColor = d_BackColor
    Value = d_Value
    Alignment = d_Alignment

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'if control is disabled or no themes then exit sub
    If c_Enabled = False Or DrawAsXp = False Then Exit Sub
    
    'release mouse position capture
    ReleaseCapture

    'draw button as pressed(checked pressed or unchecked pressed)
    If c_Value = True Then
        DrawXpControl uPressed
    ElseIf c_Value = False Then
        DrawXpControl cPressed
    End If

    'capture mouse position
    SetCapture UserControl.hWnd
    
    'set click to true
    c_Click = True
            
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'if control is clicked, disabled or no themes then exit sub
    If c_Enabled = False Or DrawAsXp = False Then Exit Sub
   
    'release mouse position capture
    ReleaseCapture
   
    'draw button as hot(checked hot or unchecked hot)
    If c_Value = True Then
        c_Click = False
        DrawXpControl uHot
        c_Value = False
    ElseIf c_Value = False Then
        DrawXpControl cHot
        c_Click = False
        c_Value = True
    End If

    'raise mouse event
    RaiseEvent Click

    'capture mouse position
    SetCapture UserControl.hWnd
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'if control is disabled or no themes then exit sub
    If c_Click = True Or c_Enabled = False Or DrawAsXp = False Then Exit Sub
    
    'if checkbox value = 0 (not pressed) then draw controls as not pressed
    If c_Value = False Then
        'if mouse within control then draw as hot, else draw as normal
        If X < 0 Or X > ScaleWidth Or Y < 0 Or Y > ScaleHeight Then
            ReleaseCapture
            DrawXpControl uNormal
        Else
            SetCapture UserControl.hWnd
            DrawXpControl uHot
        End If
    ElseIf c_Value = True Then 'if checkbox value = 1 then draw as pressed
        'if mouse within control then draw as hot, else draw as normal
        If X < 0 Or X > ScaleWidth Or Y < 0 Or Y > ScaleHeight Then
            ReleaseCapture
            DrawXpControl cNormal
        Else
            SetCapture UserControl.hWnd
            DrawXpControl cHot
        End If
    End If
 
End Sub

Private Sub UserControl_Resize()
    
    'We either draw as xp or vb standard and set focus accordingly
    If DrawAsXp = True Then
    
        'draw as checed or as unchecked
        If c_Value = True Then
            DrawXpControl cNormal
        Else
            DrawXpControl uNormal
        End If
    End If
    
    'we always resize the vb standard control just in case
    Check1.Width = UserControl.Width
    Check1.Height = UserControl.Height

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'i wont comment anything about the stuff below, pretty standard stuff
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", c_Text, d_Text)
    Call PropBag.WriteProperty("Enabled", c_Enabled, d_Enabled)
    Call PropBag.WriteProperty("BackColor", c_BackColor, d_BackColor)
    Call PropBag.WriteProperty("Value", c_Value, d_Value)
    Call PropBag.WriteProperty("Alignment", c_Alignment, d_Alignment)

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", d_Text)
    Enabled = PropBag.ReadProperty("Enabled", d_Enabled)
    BackColor = PropBag.ReadProperty("BackColor", d_BackColor)
    Value = PropBag.ReadProperty("Value", d_Value)
    Alignment = PropBag.ReadProperty("Alignment", d_Alignment)

End Sub
Public Property Get Value() As Boolean
    Value = c_Value
End Property
Public Property Let Value(ByVal nVal As Boolean)
    c_Value = nVal
    If c_Value = False Then
        Check1.Value = 0
    Else
        Check1.Value = 1
    End If
    UserControl_Resize
    PropertyChanged "Value"
End Property
Public Property Get Caption() As String
    Caption = c_Text
End Property
Public Property Let Caption(ByVal nText As String)
    c_Text = nText
    Check1.Caption = c_Text
    UserControl_Resize
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = c_Enabled
End Property
Public Property Let Enabled(ByVal nEnabled As Boolean)
    c_Enabled = nEnabled
    UserControl.Enabled = c_Enabled
    Check1.Enabled = c_Enabled
    UserControl_Resize
    PropertyChanged "Enabled"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = c_BackColor
End Property
Public Property Let BackColor(ByVal nColor As OLE_COLOR)
    c_BackColor = nColor
    Check1.BackColor = c_BackColor
    UserControl.BackColor = c_BackColor
    UserControl_Resize
    PropertyChanged "BackColor"
End Property
Public Property Get Alignment() As chkAlign
    Alignment = c_Alignment
End Property
Public Property Let Alignment(ByVal nAlign As chkAlign)
    c_Alignment = nAlign

    Check1.Alignment = c_Alignment
    UserControl_Resize
    PropertyChanged "Alignment"
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
