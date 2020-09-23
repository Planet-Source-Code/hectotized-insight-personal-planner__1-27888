VERSION 5.00
Begin VB.UserControl xpCombo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000016&
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   4800
   ToolboxBitmap   =   "xpComboBox.ctx":0000
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   15
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   15
      ScaleHeight     =   480
      ScaleWidth      =   1200
      TabIndex        =   2
      Top             =   15
      Width           =   1200
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "xpComboBox.ctx":0312
         Left            =   -25
         List            =   "xpComboBox.ctx":0314
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   -25
         Width           =   975
      End
   End
End
Attribute VB_Name = "xpCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************
'*  Copyright (C) Outersoft Inc. 2001 - All Rights Reserved        *
'*                                                                 *
'*  FILE:  xpCombo.ctl                                             *
'*                                                                 *
'*  DESCRIPTION:                                                   *
'*      A first attempt to create xp theme aware controls          *
'*                                                                 *
'*  SPECIAL THNX:                                                  *
'*        The KPD-Team                                             *
'*        Visit them at http://www.allapi.net/                     *
'*                                                                 *
'*******************************************************************

' this is a also a fake combobox. i still use the api to draw the combobox button,
'and the frame, but i had to use the standard vb combobox, since the api only
'draws(paints) the box, ther is nuthing else to create. i also didnt incorporate
'the mouse over. i know i could have done it with a timer, but i didnt feel like it

'i just didnt want anyone to send hatemail claiming this wasnt an actual
'xp combobox. i know some control guru will solve this


Option Explicit

Private ret As Long

'control defaults
Private Const d_Text = "xpCombobox"
Private Const d_Enabled = True
Private Const d_Locked = False
Private Const d_MaxLength = 0
Private Const d_MaskPassword = False

Private Const d_Font = Null ' Ambient.Font
Private Const d_Alignment = 0
Private Const d_BackColor = vbWhite
Private Const d_FontColor = vbBlack

Public Enum cboAlign
        LeftJustify = 0
        RightJustify = 1
        Center = 2
End Enum



'control types
Private Const cFrame = EP_EDITTEXT
Private Const cCombo = CP_DROPDOWNBUTTON



'control parts
'frame
Private Const cFNormal = ETS_NORMAL
Private Const cFDisabled = ETS_DISABLED


''actual dropdown buttons
Private Const cNormal = CBXS_NORMAL
Private Const cDisabled = CBXS_DISABLED
Private Const cHot = CBXS_HOT
Private Const cPressed = CBXS_PRESSED


Private c_MaxLength As Long
Private c_Click As Boolean
Private c_Alignment As Long
Private c_Font As Font
Private c_FontColor As OLE_COLOR
Private c_BackColor As OLE_COLOR
Private c_Text As String
Private c_Enabled As Boolean
Private c_Locked As Boolean

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


 
Private hTheme As Long
Private cb As RECT, tx As RECT

Event Click()
Event Change()



Private Sub DrawXpComboFrame(nState As Long)
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
Private Sub DrawXpComboDrop(nState As Long)
    'clear control
    Pic1.Cls
    'set frame position
    SetRect cb, 0, 0, Pic1.Width / 15, Pic1.Height / 15
    'open theme data
    hTheme = OpenThemeData(Pic1.hWnd, "COMBOBOX")
    'draw the checkbox
    If c_Enabled = True Then
    DrawThemeBackground hTheme, Pic1.hDC, CP_DROPDOWNBUTTON, nState, cb, ByVal 0&
    Else
    DrawThemeBackground hTheme, Pic1.hDC, CP_DROPDOWNBUTTON, cDisabled, cb, ByVal 0&
    End If
    
    'close theme data
    CloseThemeData hTheme
End Sub



Private Sub Combo1_Change()
RaiseEvent Change
End Sub

Private Sub Combo1_Click()
Text1.Text = Combo1.Text
RaiseEvent Click
End Sub

Private Sub Combo1_DropDown()
'MsgBox "FFF"
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture


DrawXpComboDrop cPressed

SetCapture Pic1.hWnd
c_Click = True
End Sub



Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If c_Click = True Or c_Enabled = False Or DrawAsXp = False Then Exit Sub

        'if mouse within control then draw as hot, else draw as normal
        If X < 0 Or X > ScaleWidth Or Y < 0 Or Y > ScaleHeight Then
            ReleaseCapture
            DrawXpComboDrop cNormal
        Else
            SetCapture Pic1.hWnd
            DrawXpComboDrop cHot
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)

        End If

End Sub

Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
DrawXpComboDrop cNormal
Combo1.SetFocus
SendKeys "%{Down}"
SetCapture Pic1.hWnd
c_Click = False
RaiseEvent Click
End Sub



Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if control is disabled or no themes then exit sub
    If c_Enabled = False Or DrawAsXp = False Then Exit Sub
    
        'if mouse within control then draw as hot, else draw as normal
        If X < 0 Or X > ScaleWidth - Pic1.Width - 30 Or Y < 0 Or Y > ScaleHeight Then
            ReleaseCapture
            DrawXpComboDrop cNormal
        Else
            SetCapture Text1.hWnd
            DrawXpComboDrop cHot
        End If
End Sub

Private Sub UserControl_InitProperties()
Set Font = Ambient.Font
Alignment = d_Alignment
ForeColor = d_FontColor
Enabled = d_Enabled
Locked = d_Locked
BackColor = d_BackColor

End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", d_Enabled)
    BackColor = PropBag.ReadProperty("BackColor", d_BackColor)
    ForeColor = PropBag.ReadProperty("ForeColor", d_FontColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text = PropBag.ReadProperty("Text", d_Text)
    Locked = PropBag.ReadProperty("Locked", d_Locked)
    Alignment = PropBag.ReadProperty("Alignment", d_Alignment)
End Sub

Private Sub UserControl_Resize()

If DrawAsXp = True Then
DrawXpComboFrame cFNormal
DrawXpComboDrop cNormal
If Pic1.Visible = False Then Pic1.Visible = True
If Text1.Visible = False Then Text1.Visible = True

Pic1.Left = UserControl.Width - Pic1.Width - 15
UserControl.Height = 30 + Pic1.Height
Text1.Width = UserControl.Width - Pic1.Width - 50
Text1.Left = 40
Text1.Top = 15
Text1.Height = UserControl.Height - 30
Picture1.Width = UserControl.Width - 40
Picture1.Height = UserControl.Height - 30
Picture1.Left = 25
Picture1.Top = 15
Combo1.Width = Picture1.Width + 45
Combo1.Top = -25
Combo1.Left = -25
Else

Pic1.Visible = False
Text1.Visible = False
Picture1.Top = 0
Picture1.Left = 0
Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height

Combo1.Width = UserControl.Width
UserControl.Height = Combo1.Height
Combo1.Left = 0
Combo1.Top = 0
End If



End Sub
Public Property Get Alignment() As cboAlign
    Alignment = c_Alignment
End Property
Public Property Let Alignment(ByVal vAlign As cboAlign)
    c_Alignment = vAlign
    Text1.Alignment = c_Alignment
    Call UserControl_Resize
    
PropertyChanged "Alignment"
End Property
Public Property Get Enabled() As Boolean
    Enabled = c_Enabled
End Property
Public Property Let Enabled(ByVal nEnabled As Boolean)
    c_Enabled = nEnabled
    UserControl.Enabled = c_Enabled
    If c_Enabled = False Then
    Text1.BackColor = UserControl.BackColor
    Else
    Text1.BackColor = c_BackColor
    End If
    
    Text1.Enabled = c_Enabled
    Pic1.Enabled = c_Enabled
    
    Call UserControl_Resize
PropertyChanged "Enabled"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = c_BackColor
End Property
Public Property Let BackColor(ByVal vData As OLE_COLOR)
c_BackColor = vData

Text1.BackColor = c_BackColor
Combo1.BackColor = c_BackColor

Call UserControl_Resize
PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = c_FontColor
End Property
Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    c_FontColor = vData
    Combo1.ForeColor = c_FontColor
      Text1.ForeColor = c_FontColor
PropertyChanged "ForeColor"
End Property
Public Property Get Font() As Font
    Set Font = c_Font
End Property
Public Property Set Font(ByVal vData As Font)
    Set c_Font = vData
    Set UserControl.Font = vData
    Set Text1.Font = c_Font
        Set Combo1.Font = c_Font
    Call UserControl_Resize
PropertyChanged "Font"
End Property
Public Property Get Text() As String
    Text = c_Text
End Property
Public Property Let Text(ByVal vData As String)
    c_Text = vData
    Text1.Text = vData
    Combo1.Text = vData
PropertyChanged "Text"
End Property

Public Sub AddItem(nItem As String)
    Combo1.AddItem nItem
End Sub
Public Sub Clear()
    Combo1.Clear
End Sub
Public Sub Refresh()
    Combo1.Refresh
End Sub
Public Sub RemoveItem(Index As Integer)
    Combo1.RemoveItem Index
End Sub


Public Property Get ListIndex() As String
    ListIndex = Combo1.ListIndex
End Property
Public Property Let ListIndex(ByVal vData As String)
    Combo1.ListIndex = vData
PropertyChanged "ListIndex"
End Property

Public Property Get List(nIndex As String)
    List = Combo1.List(nIndex)
End Property

Public Property Get ListCount() As Integer
    ListCount = Combo1.ListCount
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Property Let ListCount(ByVal vData As String)
'    Combo1.ListCount = vData
'PropertyChanged "ListCount"
'End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Locked() As Boolean
    Locked = c_Locked
End Property
Public Property Let Locked(ByVal vData As Boolean)
    c_Locked = vData
    Text1.Locked = c_Locked
    Call UserControl_Resize
PropertyChanged "Locked"
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", c_Enabled, d_Enabled)
    Call PropBag.WriteProperty("BackColor", c_BackColor, d_BackColor)
    Call PropBag.WriteProperty("Locked", c_Locked, d_Locked)
    Call PropBag.WriteProperty("Font", c_Font, Ambient.Font)
    Call PropBag.WriteProperty("Text", c_Text, d_Text)
    Call PropBag.WriteProperty("ForeColor", c_FontColor, d_FontColor)
    Call PropBag.WriteProperty("Alignment", c_Alignment, d_Alignment)

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
