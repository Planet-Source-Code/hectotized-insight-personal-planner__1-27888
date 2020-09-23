VERSION 5.00
Begin VB.UserControl xpTextbox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "xpTextbox.ctx":0000
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "xpTextbox.ctx":0312
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "xpTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************
'*  Copyright (C) Outersoft Inc. 2001 - All Rights Reserved        *
'*                                                                 *
'*  FILE:  xpTexbox.ctl                                             *
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
Private Const d_Text = "xpText"
Private Const d_Enabled = True
Private Const d_Locked = False
Private Const d_MaxLength = 0
Private Const d_SelText = ""
Private Const d_SelLength = 0
Private Const d_SelStart = 0
Private Const d_MaskPassword = False
Private Const d_Font = Null ' Ambient.Font
Private Const d_Alignment = 0
Private Const d_BackColor = vbWhite
Private Const d_FontColor = vbBlack


'api to draw control as hot, selected, defaulted
Private Const cNormal = ETS_NORMAL
Private Const cDisabled = ETS_DISABLED

Public Enum tbxAlign
        LeftJustify = 0
        RightJustify = 1
        Center = 2
End Enum



Private c_SelText As String
Private c_SelLength As Long
Private c_SelStart As Long
Private c_MaxLength As Long
Private c_MaskPassword As Boolean
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



Private Sub DrawXpControl(nState As Long)
    
    'clear control
    Cls
    
    'set control position
    SetRect cb, 0, 0, UserControl.Width / 15, UserControl.Height / 15
    
   
    'open theme data
    hTheme = OpenThemeData(UserControl.hWnd, "EDIT")
    
    'if control is enabled we draw normal, if not, then as disabled
    If c_Enabled = True Then 'control enabled
        'draw the checkbox
        DrawThemeBackground hTheme, UserControl.hDC, EP_EDITTEXT, nState, cb, ByVal 0&
    Else 'control disabled
        DrawThemeBackground hTheme, UserControl.hDC, BP_CHECKBOX, CBS_CHECKEDDISABLED, cb, ByVal 0&
    End If
    
    'Write text
    Text1.Text = c_Text
    
    'close theme data
    CloseThemeData hTheme
    
End Sub



Private Sub UserControl_InitProperties()
'Text1.MultiLine = True
Set Font = Ambient.Font

Text = d_Text
Enabled = d_Enabled
Locked = d_Locked
BackColor = d_BackColor
Alignment = d_Alignment
MaxLength = d_MaxLength
SelLength = d_SelLength
SelStart = d_SelStart
SelText = d_SelText
Alignment = d_Alignment
MaskPassword = d_MaskPassword
End Sub


Private Sub UserControl_Resize()


DrawXpControl cNormal


'here we draw as xp or not
If DrawAsXp = True Then
    Text1.Appearance = 0
    Text1.BorderStyle = 0
    Text1.Left = 30
    Text1.Top = 15
    Text1.Width = UserControl.Width - 60
    Text1.Height = UserControl.Height - 30

Else 'if not xp we want our 3d borders back
    Text1.Appearance = 1
    Text1.BorderStyle = 1
    
    Text1.Left = 0
    Text1.Top = 0
    Text1.Width = UserControl.Width
    Text1.Height = UserControl.Height

End If

    If c_MaskPassword = False Then
    Text1.PasswordChar = ""
    Set Text1.Font = UserControl.Font
    Else
    Text1.PasswordChar = Label1.Caption
    Set Text1.Font = Label1.Font
    End If


End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text = PropBag.ReadProperty("Text", d_Text)
    Enabled = PropBag.ReadProperty("Enabled", d_Enabled)
    Locked = PropBag.ReadProperty("Locked", d_Locked)
    MaxLength = PropBag.ReadProperty("MaxLength", d_MaxLength)
    MaskPassword = PropBag.ReadProperty("MaskPassword", d_MaskPassword)
    SelLength = PropBag.ReadProperty("SelLength", d_SelLength)
    SelStart = PropBag.ReadProperty("SelStart", d_SelStart)
    SelText = PropBag.ReadProperty("SelText", d_SelText)
    ForeColor = PropBag.ReadProperty("ForeColor", d_FontColor)
    BackColor = PropBag.ReadProperty("BackColor", d_BackColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Alignment = PropBag.ReadProperty("Alignment", d_Alignment)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Locked", c_Locked, d_Locked)
    Call PropBag.WriteProperty("MaxLength", c_MaxLength, d_MaxLength)
    Call PropBag.WriteProperty("MaskPassword", c_MaskPassword, d_MaskPassword)
    Call PropBag.WriteProperty("Font", c_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", c_BackColor, d_BackColor)
    Call PropBag.WriteProperty("ForeColor", c_FontColor, d_FontColor)
    Call PropBag.WriteProperty("Alignment", c_Alignment, d_Alignment)
    Call PropBag.WriteProperty("SelLength", c_SelLength, d_SelLength)
    Call PropBag.WriteProperty("SelStart", c_SelStart, d_SelStart)
    Call PropBag.WriteProperty("SelText", c_SelText, d_SelText)
    Call PropBag.WriteProperty("Enabled", c_Enabled, d_Enabled)
    Call PropBag.WriteProperty("Text", c_Text, d_Text)
End Sub

Public Property Get Font() As Font
    Set Font = c_Font
End Property
Public Property Set Font(ByVal vData As Font)
    Set c_Font = vData
    Set Text1.Font = c_Font
    Set UserControl.Font = c_Font
    UserControl_Resize
PropertyChanged "Font"
End Property
Public Property Get FontColor() As OLE_COLOR
FontColor = c_FontColor
End Property

Public Property Let FontColor(ByVal New_Color As OLE_COLOR)
c_FontColor = New_Color
Text1.ForeColor = c_FontColor
PropertyChanged "FontColor"
End Property

Public Property Get Alignment() As tbxAlign
    Alignment = c_Alignment
End Property
Public Property Let Alignment(ByVal vAlign As tbxAlign)
    c_Alignment = vAlign
    Text1.Alignment = c_Alignment
    Call UserControl_Resize
    
PropertyChanged "Alignment"
End Property

Public Property Get Locked() As Boolean
    Locked = c_Locked
End Property
Public Property Let Locked(ByVal vLocked As Boolean)
    c_Locked = vLocked
    Text1.Locked = c_Locked
    Call UserControl_Resize
    
PropertyChanged "Locked"
End Property



Public Property Get MaxLength() As Long
    MaxLength = c_MaxLength
End Property
Public Property Let MaxLength(ByVal vMaxLength As Long)
    c_MaxLength = vMaxLength
    Text1.MaxLength = c_MaxLength
    Call UserControl_Resize
    
PropertyChanged "MaxLength"
End Property

Public Property Get SelStart() As Long
    SelStart = c_SelStart
End Property
Public Property Let SelStart(ByVal nSelStart As Long)
    c_SelStart = nSelStart
    Text1.SelStart = c_SelStart
    Call UserControl_Resize
    
PropertyChanged "SelStart"
End Property
Public Property Get SelLength() As Long
    SelLength = c_SelLength
End Property
Public Property Let SelLength(ByVal nSelLength As Long)
    c_SelLength = nSelLength
    Text1.SelLength = c_SelLength
    Call UserControl_Resize
    
PropertyChanged "SelLength"
End Property
Public Property Get SelText() As String
    SelText = c_SelText
End Property
Public Property Let SelText(ByVal nSelText As String)
    c_SelText = nSelText
    Text1.SelText = c_SelText
    Call UserControl_Resize
    
PropertyChanged "SelText"
End Property

Public Property Get MaskPassword() As Boolean
    MaskPassword = c_MaskPassword
End Property
Public Property Let MaskPassword(ByVal vMaskPassword As Boolean)
    c_MaskPassword = vMaskPassword
    If c_MaskPassword = False Then
    Text1.PasswordChar = ""
    Set Text1.Font = UserControl.Font
    Else
    Text1.PasswordChar = Label1.Caption
    Set Text1.Font = Label1.Font
    End If
    
    Call UserControl_Resize
    
PropertyChanged "MaskPassword"
End Property




Public Property Get Text() As String
    Text = Text1.Text
End Property

Public Property Let Text(ByVal nText As String)
    RaiseEvent Change
    c_Text = nText
    Text1.Text = c_Text

    UserControl_Resize
    PropertyChanged "Text"

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

    
    Call UserControl_Resize
PropertyChanged "Enabled"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = c_BackColor
End Property
Public Property Let BackColor(ByVal vData As OLE_COLOR)
c_BackColor = vData


Text1.BackColor = c_BackColor

Call UserControl_Resize
PropertyChanged "BackColor"
End Property

Private Sub Text1_Change()
RaiseEvent Change
End Sub

Private Sub Text1_Click()
RaiseEvent Click
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

