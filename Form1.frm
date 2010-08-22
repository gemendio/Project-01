VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Room Scheduler"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNormalMode 
         Caption         =   "Normal Mode"
      End
      Begin VB.Menu mnuActiveMode 
         Caption         =   "Active Mode"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSchedule 
      Caption         =   "&Schedule"
      Begin VB.Menu mnuAddNewSched 
         Caption         =   "Add New Schedule"
      End
      Begin VB.Menu mnuViewScheds 
         Caption         =   "View Schedules"
      End
   End
   Begin VB.Menu mnuStudent 
      Caption         =   "St&udent"
      Begin VB.Menu mnuAddNewStudent 
         Caption         =   "Add New Student"
      End
      Begin VB.Menu mnuViewStudents 
         Caption         =   "View Students"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuExit_Click()
    End
End Sub
