VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insight - Personal Planner"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   10590
   FillColor       =   &H8000000D&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin Insight.xpFrame frame1 
      Height          =   5655
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   9975
      Caption         =   "Add Instant Notice"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Insight.xpTimePicker dtpTime 
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   0
         Format          =   2
      End
      Begin Insight.xpCombo cbAudio 
         Height          =   285
         Left            =   1680
         TabIndex        =   37
         Top             =   1320
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   503
         Locked          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "(None)"
         Alignment       =   2
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483646
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthColumns    =   2
         MonthBackColor  =   -2147483633
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   49872897
         TitleBackColor  =   -2147483645
         TitleForeColor  =   -2147483634
         TrailingForeColor=   -2147483636
         CurrentDate     =   37136
         MinDate         =   -52196
      End
      Begin Insight.xpTextbox tbxNotifyText 
         Height          =   525
         Left            =   120
         TabIndex        =   33
         Top             =   4530
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
      End
      Begin Insight.xpTextbox setDate 
         Height          =   285
         Left            =   2280
         TabIndex        =   32
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         Locked          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Text            =   "Date"
      End
      Begin Insight.xpTextbox setTime 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         Locked          =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Text            =   "Time"
      End
      Begin Insight.xpCommand cmdClear 
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         Top             =   5160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Clear Notice"
         Enabled         =   0   'False
         BackColor       =   -2147483643
      End
      Begin Insight.xpCommand cmdSave 
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   5160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Set Notice"
         BackColor       =   -2147483643
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Insight Notice Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Notice Message:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   4320
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Notice Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Notice Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Notice Sound:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   22
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Insight Notice Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5655
      Left            =   7440
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   6155
      Begin Insight.xpFrame xpFrame1 
         Height          =   2535
         Left            =   0
         TabIndex        =   12
         Top             =   3120
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   4471
         Caption         =   "Misc Settings"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Insight.xpCheck cbxCleanUp 
            Height          =   255
            Left            =   600
            TabIndex        =   35
            Top             =   1560
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   450
            Caption         =   "Automaticaly Delete Expired Notices On Start Up."
            BackColor       =   16777215
         End
         Begin Insight.xpCheck cbxStartUp 
            Height          =   255
            Left            =   600
            TabIndex        =   27
            Top             =   1200
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            Caption         =   "On Start Up, Load to Systray Even If No Active Alerts."
            BackColor       =   16777215
         End
         Begin Insight.xpCheck cbxAlarmMode 
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   840
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "Enable Alarm Clock Mode"
            BackColor       =   16777215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "**This Will Only Accur If iNotice Or Your Computer Was Shut Off Durring A Scheduled Notice."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000010&
            Height          =   495
            Left            =   600
            TabIndex        =   36
            Top             =   1920
            Width           =   4575
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Other Settings"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   3495
         End
      End
      Begin Insight.xpFrame frame7 
         Height          =   3015
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   5318
         Caption         =   "Sound Options"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Insight.xpTextbox xplMP3Dir 
            Height          =   285
            Left            =   360
            TabIndex        =   30
            Top             =   1440
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
         End
         Begin Insight.xpCheck chkMp3Finish 
            Height          =   255
            Left            =   600
            TabIndex        =   29
            Top             =   2520
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   450
            Caption         =   "Always Allow MP3 To Finish Playing After Notifing."
            BackColor       =   16777215
         End
         Begin Insight.xpCommand cmdBrowse 
            Height          =   375
            Left            =   4560
            TabIndex        =   10
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "Browse"
            BackColor       =   -2147483643
         End
         Begin Insight.xpOpt optSoundDir 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            Caption         =   "Use Windows Default Sound Directory"
            BackColor       =   16777215
            Value           =   -1  'True
         End
         Begin Insight.xpOpt optSoundDir 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            Caption         =   "Use MP3 Music Directory As Default"
            BackColor       =   16777215
         End
         Begin VB.Label LBLdsd 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Sound Directory:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   3495
         End
      End
   End
   Begin Insight.xpFrame Frame2 
      Height          =   5655
      Left            =   2160
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   9975
      Caption         =   "Current Notices"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ListView lvAlerts 
         Height          =   4455
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Alert Date"
            Object.Width           =   4587
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Alert Time"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Alert Message"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Alert Sound"
            Object.Width           =   6068
         EndProperty
      End
      Begin Insight.xpCommand cmdChange 
         Height          =   375
         Left            =   3000
         TabIndex        =   28
         Top             =   5160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Change"
         Enabled         =   0   'False
         BackColor       =   -2147483643
      End
      Begin Insight.xpCommand cmdDelete 
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   5160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Delete"
         Enabled         =   0   'False
         BackColor       =   -2147483643
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Instant Notices:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   3495
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   9180
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6227
            MinWidth        =   5116
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3229
            MinWidth        =   2118
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9119
            MinWidth        =   8008
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pichook 
      Height          =   615
      Left            =   9240
      Picture         =   "frmMain.frx":1CFA
      ScaleHeight     =   555
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   1800
      Width           =   375
   End
   Begin VB.FileListBox File1 
      Height          =   3450
      Left            =   10200
      TabIndex        =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   8760
      Top             =   1080
   End
   Begin VB.Timer tmrCaption 
      Interval        =   60
      Left            =   9120
      Top             =   2880
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   6495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   11456
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New Instant Notice"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "View Instant Notices"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Insight.xpCommand cmdHide 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   6720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Hide"
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   4
      Left            =   10200
      Picture         =   "frmMain.frx":39F4
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   3
      Left            =   9480
      Picture         =   "frmMain.frx":3F7E
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   2
      Left            =   9960
      Picture         =   "frmMain.frx":4508
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   2775
      Left            =   9840
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      Left            =   9480
      Picture         =   "frmMain.frx":4A92
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IconImage 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   9720
      Picture         =   "frmMain.frx":501C
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const WM_CLOSE = &H10


''''start browse for folder'''''''
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEFORCOMPUTER = &H1000

Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long


Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
    End Type
    ''''''''end browse for folder'''''

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private TrayI As NOTIFYICONDATA

Private f_Change As Long

Dim I As Long
Dim sIcon, msg
Dim CountIt As String, FirstAlert As String


Private Sub cbAudio_Click()
'since user clicked hour combobox then we enable clear button
cmdSave.Enabled = True
cmdClear.Enabled = True
'MonthView1.SetFocus
'MsgBox cbAudio.Text
End Sub

Private Sub cbxAlarmMode_Click()

On Error Resume Next
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AlarmMode", cbxAlarmMode.Value

If cbxAlarmMode.Value = True Then
optSoundDir(0).Enabled = False
optSoundDir(1).Enabled = False
xplMP3Dir.Enabled = False
cmdBrowse.Enabled = False
chkMp3Finish.Enabled = False
frmNotify.cbxAlwaysPlay.Enabled = False
frame7.Enabled = False
LBLdsd.Enabled = False
Else

optSoundDir(0).Enabled = True
optSoundDir(1).Enabled = True
xplMP3Dir.Enabled = True
cmdBrowse.Enabled = True
chkMp3Finish.Enabled = True
frmNotify.cbxAlwaysPlay.Enabled = True
frame7.Enabled = True
LBLdsd.Enabled = True
End If

End Sub

Private Sub cbxCleanUp_Click()
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AutoCleanUp", cbxCleanUp.Value

End Sub

Private Sub cbxStartUp_Click()
If cbxStartUp.Value = True Then
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AutoStart", cbxStartUp.Value

SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Insight", App.Path & "\" & App.EXEName & ".exe"
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Insight"
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AutoStart", cbxStartUp.Value

End If
End Sub

Private Sub chkMp3Finish_Click()
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AlwaysPlay", chkMp3Finish.Value

End Sub

Private Sub cmdBrowse_Click()

xplMP3Dir.Text = BrowseFolder
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "MP3Directory", xplMP3Dir.Text

End Sub

Private Sub cmdClear_Click()
On Error Resume Next
cbAudio.Text = "(None)"
tbxNotifyText.Text = ""
MonthView1.Value = Date
setDate.Text = "Date"
setTime.Text = "Time"
dtpTime.Value = Time

'if user selected change but then cleared, we wanna unselect the notice
If f_Change >= 1 Then
lvAlerts.SelectedItem.Selected = False
f_Change = 0
End If
cmdClear.Enabled = False
cmdSave.Enabled = False

'dtpTime.SetFocus
End Sub

Private Sub cmdDelete_Click()

On Error Resume Next

If lvAlerts.SelectedItem.Selected = False Then
MsgBox "Please Select Notice To Delete"
Exit Sub
End If


'Delete alert from registry
DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & lvAlerts.SelectedItem.Text & " - " & lvAlerts.SelectedItem.SubItems(1)

'Delete alert from listview
lvAlerts.ListItems.Remove (lvAlerts.SelectedItem.Index)

'update labels
StatusBar.Panels(1).Text = "You have a total of " & lvAlerts.ListItems.Count & " active notices."

'aDD TO SYSTRAY
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd 'Link the trayicon to this picturebox
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = pichook.Picture
    TrayI.szTip = "You have a total of " & lvAlerts.ListItems.Count & " active notices." & Chr$(0)
 
    'Create the icon
    Shell_NotifyIcon NIM_MODIFY, TrayI
    
'if no alerts, then disable buttons
If lvAlerts.ListItems.Count = 0 Then
cmdDelete.Enabled = False
cmdChange.Enabled = False
'Switch to Set Alert tab to modify the settings
TabStrip.Tabs(1).Selected = True

End If

End Sub


Private Sub cmdHide_Click()
Me.Hide
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
'if user selected change notice and is now saving it, we wanna delete the
'old notice

If f_Change >= 1 Then

'Delete alert from registry
DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & lvAlerts.SelectedItem.Text & " - " & lvAlerts.SelectedItem.SubItems(1)

'Delete alert from listview
lvAlerts.ListItems.Remove (f_Change)

'update labels
StatusBar.Panels(1).Text = "You have a total of " & lvAlerts.ListItems.Count & " active notices."

'aDD TO SYSTRAY
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd 'Link the trayicon to this picturebox
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = pichook.Picture
    TrayI.szTip = "You have a total of " & lvAlerts.ListItems.Count & " active notices." & Chr$(0)
 
    'Create the icon
    Shell_NotifyIcon NIM_MODIFY, TrayI
    
'if no alerts, then disable buttons
If lvAlerts.ListItems.Count = 0 Then
cmdDelete.Enabled = False
cmdChange.Enabled = False
End If
If lvAlerts.ListItems.Count >= 1 Then lvAlerts.SelectedItem.Selected = False

f_Change = 0
End If
'end update notice






'dont save alert if user didnt set date
If setDate.Text = "Date" Then MsgBox "You Must Set The Alert Date!!": Exit Sub
If setTime.Text = "Time" Then MsgBox "You Must Set The Alert Time!!": Exit Sub

'user selected invalid time
If NoticeExpired(ShortDate(setDate.Text), setTime.Text) = True Then
MsgBox "You Selected A Time That Has Already Past, Please Select New Time"
dtpTime.SetFocus
Exit Sub
End If


'if timer was turned off, then we turn it on
If Timer2.Enabled = False Then Timer2.Enabled = True


'load new alert to list view

If cbAudio.List(cbAudio.ListIndex) = "" Then sIcon = 2 Else sIcon = 1
With lvAlerts.ListItems.Add(, , setDate.Text, , sIcon)
    .SubItems(1) = setTime.Text
    .SubItems(2) = tbxNotifyText.Text
    If sIcon = 1 Then
    .SubItems(3) = File1.Path & "\" & cbAudio.List(cbAudio.ListIndex)
    End If

End With

'save new alert to registry
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & setDate.Text & " - " & setTime.Text, "AlertTime", setTime.Text
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & setDate.Text & " - " & setTime.Text, "AlertDate", setDate.Text
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & setDate.Text & " - " & setTime.Text, "AlertMessage", tbxNotifyText.Text
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & setDate.Text & " - " & setTime.Text, "AlertSound", File1.Path & "\" & cbAudio.List(cbAudio.ListIndex)

'send click to clear set alert input settings and disable save button
cmdClear_Click '

'enable delete button
If cmdDelete.Enabled = False Then cmdDelete.Enabled = True
If cmdChange.Enabled = False Then cmdChange.Enabled = True
'update labels
StatusBar.Panels(1).Text = "You have a total of " & lvAlerts.ListItems.Count & " active notices."


End Sub






Private Sub cmdChange_Click()
On Error Resume Next

If lvAlerts.SelectedItem.Selected = False Then
    MsgBox "Please Select Notice To Change"
    Exit Sub
End If

f_Change = lvAlerts.SelectedItem.Index


setTime.Text = lvAlerts.SelectedItem.SubItems(1)

setDate.Text = lvAlerts.SelectedItem.Text

'Load sound event to cbAudio
For I = 1 To cbAudio.ListCount

'If lvAlerts.SelectedItem.SubItems(3) = "" Then
cbAudio.Text = "(None)"
'ElseIf cbAudio.List(i) = lvAlerts.SelectedItem.SubItems(3) Then
'cbAudio.ListIndex = i
'End If

Next I

'Load text message from list view to set alert message
tbxNotifyText.Text = lvAlerts.SelectedItem.SubItems(2)

'load date to monthview1
MonthView1.Value = ShortDate(lvAlerts.SelectedItem.Text)
'you edited the shortdate sub

dtpTime.Value = lvAlerts.SelectedItem.SubItems(1)
'Switch to Set Alert tab to modify the settings
TabStrip.Tabs(1).Selected = True
End Sub
Private Function BrowseFolder()
On Error Resume Next

'Opens a Treeview control that displays
    '     the directories in a computer
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "Instant Notice"


    With tBrowseInfo
        .hwndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
        
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        
        BrowseFolder = sBuffer
        
    End If
    
    
    
    
    
End Function
Private Function ShortDate(nvalue As String)
On Error Resume Next

If nvalue = "" Then Exit Function

'this is to convert date (Tuesday, September 28, 2001) to (09/28/2001)
Dim syear, smonth, sday


syear = Right(nvalue, 4)

sday = Left(Right(nvalue, 8), 2)

smonth = Right(nvalue, Len(nvalue) - InStr(nvalue, ",") - 1)
smonth = Left(smonth, Len(smonth) - 9)

Select Case RTrim(LTrim(smonth))

Case "January"
smonth = 1
Case "February"
smonth = 2
Case "March"
smonth = 3
Case "April"
smonth = 4
Case "May"
smonth = 5
Case "June"
smonth = 6
Case "July"
smonth = 7
Case "August"
smonth = 8
Case "September"
smonth = 9
Case "October"
smonth = 10
Case "November"
smonth = 11
Case "December"
smonth = 12
End Select

ShortDate = smonth & "/" & sday & "/" & syear

End Function


Private Function NoticeExpired(comDate As String, comTime As String) As Boolean
On Error Resume Next
Dim nDate As String, tDate As String, nyear As String, tyear As String
Dim ntime As String, ttime As String

nDate = Format(comDate, "MM/DD/YYYY")
tDate = Format(Date, "MM/DD/YYYY")
ntime = Format(comTime, "HH:mm")
ttime = Format(Time, "HH:mm")

'years
nyear = Right(nDate, 4)
tyear = Right(tDate, 4)



'yesterday or older we flag
If nDate < tDate And nyear <= tyear Then
NoticeExpired = True
Exit Function
End If

'today but with older time we flag
If nDate = tDate And nyear = tyear And ntime < ttime Then
NoticeExpired = True
Exit Function
End If



'if we got this far, the notice is valid
NoticeExpired = False


End Function




Private Sub dtpTime_Change()
setTime.Text = Format(dtpTime.Value, "h:mm AM/PM")
End Sub

Private Sub dtpTime_Click()
cmdSave.Enabled = True
cmdClear.Enabled = True
End Sub

Private Sub Form_Load()

'here we check if previously loaded, if it is, then we show that one and kill this one
If App.PrevInstance = True Then
ShowWindow GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "hWnd"), SW_SHOWNORMAL
End
End If


On Error Resume Next

'here we save our hwnd, i know thers a better way of doing this
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "hWnd", Me.hWnd


Dim I, z

Dim rTime As String, rDate As String
    RemoveMenu GetSystemMenu(Me.hWnd, 0), 6, MF_BYPOS
    RemoveMenu GetSystemMenu(Me.hWnd, 0), 5, MF_BYPOS

TabStrip.Left = 75
TabStrip.Top = 120

frame1.Top = 720
Frame2.Top = 720
Frame3.Top = 720
frame1.Left = 240
Frame2.Left = 240
Frame3.Left = 240
Me.Height = 7995
Me.Width = 6720

    ' Load pictures into the ImageList.
    For I = 0 To 4
        ImageList1.ListImages.Add , , IconImage(I).Picture
    Next I
    
    


        
    
'load sound directory if any
xplMP3Dir.Text = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "MP3Directory")

'load user selected default
optSoundDir(1).Value = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "UseMP3Dir")
optSoundDir(0).Value = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "UseWaveDir")

If optSoundDir(0).Value = False And optSoundDir(1).Value = False Then optSoundDir(0).Value = True

'now we check for user default directory, if not set then we set to windows default
If optSoundDir(1).Value = False Then
File1.Pattern = "*.wav"
File1.Path = "C:\WINDOWS\Media"
xplMP3Dir.Enabled = False
cmdBrowse.Enabled = False
Else
File1.Pattern = "*.mp3"
File1.Path = xplMP3Dir.Text
xplMP3Dir.Enabled = True
cmdBrowse.Enabled = True
End If

'we load autoclean up
cbxCleanUp.Value = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AutoCleanUp")
    
'if user selected to autoclean, then we do it
If cbxCleanUp.Value = True Then
retry:


CountIt = CountRegKeys(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts")

For I = 0 To CountIt - 1

FirstAlert = GetRegKey(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts", I)


rDate = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & FirstAlert, "AlertDate")
rTime = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & FirstAlert, "AlertTime")

'if notice is expired we erase and start check over
If NoticeExpired(ShortDate(rDate), rTime) = True Then
DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & FirstAlert

GoTo retry
End If

Next I
End If


'get the amount of notices
CountIt = CountRegKeys(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts")

'disable delete button if not notices
If CountIt = 0 Then cmdDelete.Enabled = False


'load saved alerts to listview
For I = 0 To CountIt - 1
FirstAlert = GetRegKey(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts", I)



rDate = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & FirstAlert, "AlertDate")
rTime = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & FirstAlert, "AlertTime")



If GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & FirstAlert, "AlertSound") = "" Then sIcon = 2 Else sIcon = 1

    With lvAlerts.ListItems.Add(, , rDate, , sIcon)
        .SubItems(1) = rTime
        .SubItems(2) = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & FirstAlert, "AlertMessage")
        .SubItems(3) = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & FirstAlert, "AlertSound")
    End With


Next I
    



'load alarm mode setting
cbxAlarmMode.Value = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AlarmMode")

'aDD TO SYSTRAY
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd 'Link the trayicon to this picturebox
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = pichook.Picture
    TrayI.szTip = "You have a total of " & CountIt & " active notices." & Chr$(0)
 
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI


StatusBar.Panels(1).Text = "You have a total of " & CountIt & " active notices."
    
    
    
    
'If no saved alerts then kill timer
If lvAlerts.ListItems.Count = 0 Then Timer2.Enabled = False
    
'load always play
chkMp3Finish.Value = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AlwaysPlay")


'set both date and time pickers to current time and date
dtpTime.Value = Time

MonthView1.Value = Date

'set cbx if Insight is loaded at start up
cbxStartUp.Value = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "AutoStart")


'add audio fils to combobox
For I = 0 To File1.ListCount - 1
File1.ListIndex = I
cbAudio.AddItem File1.FileName
Next I


'enable or disable delete and  change
If lvAlerts.ListItems.Count = 0 Then
cmdDelete.Enabled = False
cmdChange.Enabled = False
Else
cmdDelete.Enabled = True
cmdChange.Enabled = True

End If

'cbHour.SetFocus


End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd
    TrayI.uId = 1&
    'Delete the icon
    Shell_NotifyIcon NIM_DELETE, TrayI
    c_CANCEL = False

Timer2.Enabled = False


End

End Sub




Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
'check for valid date
If MonthView1.Value < Date Then
MsgBox "Sorry, Insight Cannot Be Set For A Date That Has Already Past", vbExclamation, "Insight Error"
MonthView1.SetFocus
Exit Sub
End If


setDate.Text = Format(MonthView1.Value, "Long Date")

'since user clicked monthview then we enable clear button
cmdSave.Enabled = True
cmdClear.Enabled = True

End Sub
Private Sub HideTabs()

frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub



Private Sub optSoundDir_Click(Index As Integer)
On Error Resume Next

Select Case Index

Case 0
optSoundDir(0).Value = True
optSoundDir(1).Value = False
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "UseWaveDir", optSoundDir(0).Value
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "UseMP3Dir", 0
xplMP3Dir.Enabled = False
cmdBrowse.Enabled = False

Case 1
optSoundDir(0).Value = False
optSoundDir(1).Value = True

SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "UseMP3Dir", optSoundDir(1).Value
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "UseWaveDir", 0
xplMP3Dir.Enabled = True
cmdBrowse.Enabled = True
If xplMP3Dir.Text = "" Then cmdBrowse_Click

End Select

cbAudio.Clear
If optSoundDir(1).Value = False Then
File1.Pattern = "*.wav"
File1.Path = "C:\WINDOWS\Media"
Else
File1.Pattern = "*.mp3"
File1.Path = xplMP3Dir.Text
End If




For I = 0 To File1.ListCount - 1
File1.ListIndex = I
cbAudio.AddItem File1.FileName
Next I

cbAudio.Text = "(None)"

End Sub

Private Sub TabStrip_Click()
HideTabs
Select Case TabStrip.SelectedItem.Index
Case 1
frame1.Visible = True
Case 2
Frame2.Visible = True
Case 3
Frame3.Visible = True

End Select
End Sub



Private Sub tbxNotifyText_Change()
'since user clicked hour combobox then we enable clear button
cmdSave.Enabled = True
cmdClear.Enabled = True
End Sub

Private Sub Timer2_Timer()
On Error Resume Next

Dim stTime, stAP
For I = 1 To lvAlerts.ListItems.Count




If lvAlerts.ListItems(I).SubItems(1) = Format(Time, "h:mm AM/PM") And lvAlerts.ListItems(I).Text = Format(Date, "Long Date") Then



If cbxAlarmMode.Value = False Then
'''Check to see if we have an audible alert
    If lvAlerts.ListItems(I).SubItems(3) > "" Then
    MediaPlayer1.FileName = lvAlerts.ListItems(I).SubItems(3)
    'MediaPlayer1.Volume = 127
    MediaPlayer1.PlayCount = 1
    MediaPlayer1.Play
    End If

Else
    MediaPlayer1.FileName = "C:\WINDOWS\Media\notify.wav"
    'MediaPlayer1.Volume = 127
    MediaPlayer1.PlayCount = 0
    MediaPlayer1.Play



End If





'''Check to see if we have a text message to display
    If lvAlerts.ListItems(I).SubItems(2) > "" Then
    frmNotify.lblText = lvAlerts.ListItems(I).SubItems(2)
    frmNotify.Show
    Else
    frmNotify.lblText = "Insight was set to alert you of sumthing,, but since you didnt say what, Insight can not tell you. You must have got high."
    frmNotify.Show
    End If



'Delete alert from registry
DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight\Alerts\" & lvAlerts.ListItems(I).Text & " - " & lvAlerts.ListItems(I).SubItems(1)

'Delete alert from listview
lvAlerts.ListItems.Remove I

If lvAlerts.ListItems.Count = 0 Then cmdDelete.Enabled = False


End If

Next I
    


End Sub

Private Sub tmrCaption_Timer()
StatusBar.Panels(2).Text = Time
StatusBar.Panels(3).Text = Format(Date, "Long Date")
End Sub


Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msg = X / Screen.TwipsPerPixelX
    If msg = WM_LBUTTONDBLCLK Then  'If the user dubbel-clicked on the icon
       frmMain.Show
    ElseIf msg = WM_RBUTTONUP Then  'Right click
        Me.PopupMenu frmMenus.mnuPopUp
    End If
End Sub





Private Sub xplMP3Dir_Change()
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\Insight", "MP3Directory", xplMP3Dir.Text
End Sub


