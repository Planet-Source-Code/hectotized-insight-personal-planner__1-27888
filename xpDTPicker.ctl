VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl xpTimePicker 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000002&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "xpDTPicker.ctx":0000
   Begin VB.PictureBox picDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox picUp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   15
      ScaleHeight     =   2055
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   15
      Width           =   4215
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   735
         Left            =   -30
         TabIndex        =   1
         Top             =   -30
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1296
         _Version        =   393216
         CalendarForeColor=   -2147483634
         CalendarTitleBackColor=   -2147483632
         CalendarTitleForeColor=   -2147483634
         Format          =   19464194
         CurrentDate     =   37182
      End
   End
End
Attribute VB_Name = "xpTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************
'*  Copyright (C) Outersoft Inc. 2001 - All Rights Reserved        *
'*                                                                 *
'*  FILE:  xpTimePicker.ctl                                             *
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

Private ret As Long

'control defaults
Private Const d_Enabled = True
Private Const d_Font = Null ' Ambient.Font
Private Const d_CalendarBackColor = vbRed
Private Const d_CalendarForeColor = vbBlack
Private Const d_CalendarTitleBackColor = vbWhite
Private Const d_CalendarTitleForeColor = vbBlack
Private Const d_CalendarTrailingForeColor = vbBlue

Private Const d_CustomFormat = 0
Private Const d_Format = 0
Private Const d_UpDown = False



'control types
Private Const cFrame = EP_EDITTEXT
Private Const cCombo = CP_DROPDOWNBUTTON



'control parts
'frame
Private Const cFNormal = ETS_NORMAL
Private Const cFDisabled = ETS_DISABLED


''actual dropdown buttons
Private Const uNormal = ABS_UPNORMAL
Private Const uDisabled = ABS_UPDISABLED
Private Const uHot = ABS_UPHOT
Private Const uPressed = ABS_UPPRESSED

Private Const dNormal = ABS_DOWNNORMAL
Private Const dDisabled = ABS_DOWNDISABLED
Private Const dHot = ABS_DOWNHOT
Private Const dPressed = ABS_DOWNPRESSED

Private c_Click As Boolean
Private c_Font As Font

Private c_Enabled As Boolean
Private c_CalendarBackColor As OLE_COLOR
Private c_CalendarTitleBackColor As OLE_COLOR
Private c_CalendarForeColor As OLE_COLOR
Private c_CalendarTitleForeColor As OLE_COLOR
Private c_CalendarTrailingForeColor As OLE_COLOR
Private c_CustomFormat As Long
Private c_Format As Long
Private c_UpDown As Boolean
Private c_Value As Long
Private c_Key As String



Public Enum dtpFormat
        dtpLongDate = 0
        dtpShortDate = 1
        dtpTime = 2
        dtpCustom = 3
End Enum



Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long



Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


 
Private hTheme As Long
Private cb As RECT, tx As RECT

Event Click()
Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)


Private Sub DrawXpOutline(nState As Long)
    'clear control
    Cls
    'set frame position
    SetRect cb, 0, 0, UserControl.Width / 15, UserControl.Height / 15
    'open theme data
    hTheme = OpenThemeData(UserControl.hWnd, "EDIT")
    'draw the checkbox
    
    If c_Enabled = True Then
    DrawThemeBackground hTheme, UserControl.hDC, EP_EDITTEXT, nState, cb, ByVal 0&
    Else
    DrawThemeBackground hTheme, UserControl.hDC, EP_EDITTEXT, cFDisabled, cb, ByVal 0&

    End If

    
    
    'close theme data
    CloseThemeData hTheme
End Sub
Private Sub DrawXpButtonUp(nState As Long)
    'clear control

    picUp.Cls
    'set frame position
    SetRect cb, 0, 0, picUp.Width / 15 + 1, picUp.Height / 15 + 1
    'open theme data
    hTheme = OpenThemeData(picUp.hWnd, "SCROLLBAR")
    'draw the checkbox
    If c_Enabled = True Then
    DrawThemeBackground hTheme, picUp.hDC, SBP_ARROWBTN, nState, cb, ByVal 0&
    Else
    DrawThemeBackground hTheme, picUp.hDC, SBP_ARROWBTN, uDisabled, cb, ByVal 0&
    End If
    
    'close theme data
    CloseThemeData hTheme
End Sub
Private Sub DrawXpButtonDown(nState As Long)
    'clear control

    picDown.Cls
    'set control position
    SetRect cb, 0, 0, picDown.Width / 15 + 1, picDown.Height / 15 + 1
    'open theme data
    hTheme = OpenThemeData(picDown.hWnd, "SCROLLBAR")
    'draw the checkbox
    If c_Enabled = True Then
    DrawThemeBackground hTheme, picDown.hDC, SBP_ARROWBTN, nState, cb, ByVal 0&
    Else
    DrawThemeBackground hTheme, picDown.hDC, SBP_ARROWBTN, dDisabled, cb, ByVal 0&
    End If
    
    'close theme data
    CloseThemeData hTheme
End Sub
Private Sub DTPicker1_Change()

RaiseEvent Change
End Sub

Private Sub DTPicker1_Click()

RaiseEvent Click
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub



Private Sub DTPicker1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If c_Click = True Or c_Enabled = False Or DrawAsXp = False Then Exit Sub

        'if mouse within control then draw as hot, else draw as normal
        If X < 0 Or X > ScaleWidth - 300 Or Y < 0 Or Y > ScaleHeight Then
            ReleaseCapture
            DrawXpButtonDown dNormal
            DrawXpButtonUp uNormal
        Else
            SetCapture DTPicker1.hWnd
            DrawXpButtonDown dHot
            DrawXpButtonUp uHot

        End If
End Sub

Private Sub picDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture


DrawXpButtonDown dPressed

SetCapture picDown.hWnd
c_Click = True
End Sub



Private Sub picDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If c_Click = True Or c_Enabled = False Or DrawAsXp = False Then Exit Sub

        'if mouse within control then draw as hot, else draw as normal
        If X < 0 Or X > picDown.ScaleWidth Or Y < 0 Or Y > picDown.ScaleHeight Then
            ReleaseCapture
            DrawXpButtonDown dNormal
            DrawXpButtonUp uNormal
        Else
            SetCapture picDown.hWnd
            DrawXpButtonDown dHot
            DrawXpButtonUp uHot

        End If
End Sub

Private Sub picDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
DrawXpButtonDown dNormal
DTPicker1.SetFocus


SendKeys c_Key

SetCapture picDown.hWnd
RaiseEvent Click
c_Click = False
End Sub

Private Sub picUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture


DrawXpButtonUp uPressed

SetCapture picUp.hWnd
c_Click = True
End Sub

Private Sub picUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If c_Click = True Or c_Enabled = False Or DrawAsXp = False Then Exit Sub

        'if mouse within control then draw as hot, else draw as normal
        If X < 0 Or X > picUp.ScaleWidth Or Y < 0 Or Y > picUp.ScaleHeight Then
            ReleaseCapture
            DrawXpButtonUp uNormal
            DrawXpButtonDown dNormal
        Else
            SetCapture picUp.hWnd
            DrawXpButtonUp uHot
            DrawXpButtonDown dHot

        End If
End Sub

Private Sub picUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
DrawXpButtonUp uNormal
DTPicker1.SetFocus
SendKeys "{Up}"
SetCapture picUp.hWnd
RaiseEvent Click
c_Click = False
End Sub

Private Sub UserControl_EnterFocus()
DTPicker1.SetFocus
End Sub

Private Sub UserControl_GotFocus()
DTPicker1.SetFocus
End Sub

Private Sub UserControl_InitProperties()
Set Font = Ambient.Font


Enabled = d_Enabled
CalendarBackColor = d_CalendarBackColor
CalendarForeColor = d_CalendarForeColor
CalendarTitleBackColor = d_CalendarTitleBackColor
CalendarTitleForeColor = d_CalendarTitleForeColor
CalendarTrailingForeColor = d_CalendarTrailingForeColor

CustomFormat = d_CustomFormat
Format = d_Format
UpDown = d_UpDown



End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", d_Enabled)

    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    CalendarBackColor = PropBag.ReadProperty("CalendarBackColor", d_CalendarBackColor)
    CalendarForeColor = PropBag.ReadProperty("CalendarForeColor", d_CalendarForeColor)
    CalendarTitleBackColor = PropBag.ReadProperty("CalendarTitleBackColor", d_CalendarTitleBackColor)
    CalendarTitleForeColor = PropBag.ReadProperty("CalendarTitleForeColor", d_CalendarTitleForeColor)
    CalendarTrailingForeColor = PropBag.ReadProperty("CalendarTrailingForeColor", d_CalendarTrailingForeColor)
    CustomFormat = PropBag.ReadProperty("CustomFormat", d_CustomFormat)
    Format = PropBag.ReadProperty("Format", d_Format)
    UpDown = PropBag.ReadProperty("UpDown", d_UpDown)
    'Value = PropBag.ReadProperty("Value", d_Value)
    Enabled = PropBag.ReadProperty("Enabled", d_Enabled)
    




End Sub

Private Sub UserControl_Resize()


If DrawAsXp = True Then
Picture1.Width = UserControl.Width - 30 - picDown.Width
Picture1.Height = UserControl.Height - 30
Picture1.Top = 15
Picture1.Left = 15

If picUp.Visible = False Then picUp.Visible = True
If picDown.Visible = False Then picDown.Visible = True

DTPicker1.Width = UserControl.Width + 350
DTPicker1.Height = Picture1.Height + 50
DTPicker1.Top = -30
DTPicker1.Left = -30

If c_Format = 2 Or c_UpDown = True Then
    c_Key = "{Down}"
    If picUp.Visible = False Then picUp.Visible = True

    picUp.Width = 265
    picUp.Left = UserControl.Width - picUp.Width - 20
    picUp.Top = 15
    picUp.Height = Picture1.Height / 2

    picDown.Width = 265
    picDown.Left = UserControl.Width - picDown.Width - 20
    picDown.Top = Picture1.Height / 2 + 5
    picDown.Height = Picture1.Height / 2
Else
    picUp.Visible = False
    c_Key = "{F4}"
    picDown.Width = 265
    picDown.Left = UserControl.Width - picDown.Width - 10
    picDown.Top = 0
    picDown.Height = Picture1.Height
End If
DrawXpOutline cFNormal
DrawXpButtonUp uNormal
DrawXpButtonDown dNormal

Else

UserControl.Height = DTPicker1.Height
picUp.Visible = False
picDown.Visible = False
Picture1.Top = 0
Picture1.Left = 0
Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height
DTPicker1.Width = UserControl.Width
DTPicker1.Height = Picture1.Height
DTPicker1.Top = 0
DTPicker1.Left = 0
End If

End Sub

Public Property Get Enabled() As Boolean
    Enabled = c_Enabled
End Property
Public Property Let Enabled(ByVal nEnabled As Boolean)
    c_Enabled = nEnabled
    UserControl.Enabled = c_Enabled
    If c_Enabled = False Then
    'DTPicker1.BackColor = UserControl.BackColor
    Else
    'DTPicker1.BackColor = c_BackColor
    End If
    
    DTPicker1.Enabled = c_Enabled

    
    Call UserControl_Resize
PropertyChanged "Enabled"
End Property


Public Property Get Font() As Font
    Set Font = c_Font
End Property
Public Property Set Font(ByVal vData As Font)
    Set c_Font = vData
    Set UserControl.Font = vData
    Set DTPicker1.Font = c_Font
    Call UserControl_Resize
PropertyChanged "Font"
End Property


Public Property Get Value() As String
    Value = DTPicker1.Value
End Property
Public Property Let Value(ByVal vData As String)
    DTPicker1.Value = vData
    Call UserControl_Resize
PropertyChanged "Value"
End Property



Public Property Get CustomFormat() As Long
    CustomFormat = DTPicker1.CustomFormat
End Property
Public Property Let CustomFormat(ByVal nCFormat As Long)
    c_CustomFormat = nCFormat
    DTPicker1.CustomFormat = nCFormat
    Call UserControl_Resize
PropertyChanged "CustomFormat"
End Property

Public Property Get Format() As dtpFormat
    Format = DTPicker1.Format
End Property
Public Property Let Format(ByVal nFormat As dtpFormat)
    c_Format = nFormat
    DTPicker1.Format = nFormat
    Call UserControl_Resize
PropertyChanged "Format"
End Property
Public Sub Refresh()
DTPicker1.Refresh
End Sub

Public Property Get UpDown() As Boolean
    UpDown = DTPicker1.UpDown
End Property
Public Property Let UpDown(ByVal nUpDown As Boolean)
    c_UpDown = nUpDown
    DTPicker1.UpDown = c_UpDown
    Call UserControl_Resize
PropertyChanged "UpDown"
End Property


Public Property Get CalendarBackColor() As OLE_COLOR
    CalendarBackColor = c_CalendarBackColor
End Property
Public Property Let CalendarBackColor(ByVal nColor As OLE_COLOR)
    DTPicker1.CalendarBackColor = nColor
    Call UserControl_Resize
PropertyChanged "CalendarBackColor"
End Property
Public Property Get CalendarForeColor() As OLE_COLOR
CalendarForeColor = c_CalendarForeColor
End Property

Public Property Let CalendarForeColor(ByVal nColor As OLE_COLOR)
c_CalendarForeColor = nColor
DTPicker1.CalendarForeColor = c_CalendarForeColor
PropertyChanged "CalendarForeColor"
End Property


Public Property Get CalendarTitleBackColor() As OLE_COLOR
CalendarTitleBackColor = c_CalendarTitleBackColor
End Property

Public Property Let CalendarTitleBackColor(ByVal nColor As OLE_COLOR)
c_CalendarTitleBackColor = nColor
DTPicker1.CalendarTitleBackColor = c_CalendarTitleBackColor
PropertyChanged "CalendarTitleBackColor"
End Property

Public Property Get CalendarTitleForeColor() As OLE_COLOR
CalendarTitleForeColor = c_CalendarTitleForeColor
End Property

Public Property Let CalendarTitleForeColor(ByVal nColor As OLE_COLOR)
c_CalendarTitleForeColor = nColor
DTPicker1.CalendarTitleForeColor = c_CalendarTitleForeColor
PropertyChanged "CalendarTitleForeColor"
End Property



Public Property Get CalendarTrailingForeColor() As OLE_COLOR
CalendarTrailingForeColor = c_CalendarTrailingForeColor
End Property

Public Property Let CalendarTrailingForeColor(ByVal nColor As OLE_COLOR)
c_CalendarTrailingForeColor = nColor
DTPicker1.CalendarTrailingForeColor = c_CalendarTrailingForeColor
PropertyChanged "CalendarTrailingForeColor"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", c_Enabled, d_Enabled)
    Call PropBag.WriteProperty("Font", c_Font, Ambient.Font)
    
    Call PropBag.WriteProperty("CalendarBackColor", c_CalendarBackColor, d_CalendarBackColor)
    Call PropBag.WriteProperty("CalendarForeColor", c_CalendarForeColor, d_CalendarForeColor)
    Call PropBag.WriteProperty("CalendarTitleBackColor", c_CalendarTitleBackColor, d_CalendarTitleBackColor)
    Call PropBag.WriteProperty("CalendarTitleForeColor", c_CalendarTitleForeColor, d_CalendarTitleForeColor)
    Call PropBag.WriteProperty("CalendarTrailingForeColor", c_CalendarTrailingForeColor, d_CalendarTrailingForeColor)
    Call PropBag.WriteProperty("CustomFormat", c_CustomFormat, d_CustomFormat)
    
    Call PropBag.WriteProperty("Format", c_Format, d_Format)
    Call PropBag.WriteProperty("UpDown", c_UpDown, d_UpDown)
   ' Call PropBag.WriteProperty("Value", c_Value, d_Value)



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


