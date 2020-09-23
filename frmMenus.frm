VERSION 5.00
Begin VB.Form frmMenus 
   Caption         =   "mnuPop"
   ClientHeight    =   2385
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "This form is used for the systray pop up only, since i dont want the menus visible on my actual forms."
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "SysTray"
      Begin VB.Menu mnuPop 
         Caption         =   "Set New iNotice Alert"
         Index           =   1
      End
      Begin VB.Menu mnuPop 
         Caption         =   "View iNotice Alerts"
         Index           =   2
      End
      Begin VB.Menu mnuPop 
         Caption         =   "View iNotice Settings"
         Index           =   3
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Unload iNotice"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuPop_Click(Index As Integer)
    Select Case Index
        Case 1  'new alert
            frmMain.Show
            frmMain.TabStrip.Tabs(1).Selected = True
        Case 2  'saved alert
            frmMain.Show
            frmMain.TabStrip.Tabs(2).Selected = True
        Case 3  'settings
            frmMain.Show
            frmMain.TabStrip.Tabs(3).Selected = True
        Case 5 ' exit

            Unload frmMain
            c_CANCEL = True
            Unload frmMenus
    End Select
End Sub
