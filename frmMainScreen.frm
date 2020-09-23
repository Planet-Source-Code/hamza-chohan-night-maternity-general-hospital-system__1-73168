VERSION 5.00
Begin VB.MDIForm frmMainScreen 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Nighat Maternity Home And General Hospital"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11640
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   8295
      Left            =   0
      Picture         =   "frmMainScreen.frx":0000
      ScaleHeight     =   8235
      ScaleWidth      =   11580
      TabIndex        =   0
      Top             =   0
      Width           =   11640
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu miLogin 
         Caption         =   "Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu miLogout 
         Caption         =   "Logout"
         Shortcut        =   ^O
      End
      Begin VB.Menu miSep 
         Caption         =   "-"
      End
      Begin VB.Menu miquit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu miPatient 
         Caption         =   "Patient"
         Begin VB.Menu miPPersonal 
            Caption         =   "Patient Personal"
            Shortcut        =   ^P
         End
         Begin VB.Menu miPMedical 
            Caption         =   "Patient Medical"
            Shortcut        =   ^M
         End
         Begin VB.Menu miPD 
            Caption         =   "Patient Discharge"
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu miSep2 
         Caption         =   "-"
      End
      Begin VB.Menu miEmployee 
         Caption         =   "Employee"
         Begin VB.Menu miEPersonal 
            Caption         =   "Employee Personal"
            Shortcut        =   ^E
         End
         Begin VB.Menu miEJob 
            Caption         =   "Employee Job"
            Shortcut        =   ^J
         End
      End
   End
   Begin VB.Menu mnuEntry 
      Caption         =   "&Entery"
      Begin VB.Menu miEM 
         Caption         =   "Emergency Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu miSep5 
         Caption         =   "-"
      End
      Begin VB.Menu miOPT 
         Caption         =   "Operation Theatre"
      End
      Begin VB.Menu miSep1 
         Caption         =   "-"
      End
      Begin VB.Menu miOP 
         Caption         =   "Outdoor Patient"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuMedical 
      Caption         =   "&Medical"
      Begin VB.Menu miMS 
         Caption         =   "Medical Store"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuSecurity 
      Caption         =   "&Security"
      Begin VB.Menu mnuChange 
         Caption         =   "Change Password"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu miWindow 
      Caption         =   "&Window"
      Begin VB.Menu miTH 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu miTV 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu miCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu miAI 
         Caption         =   "Arrange Icon"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu miAP 
         Caption         =   "About Project"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
 frmMainScreen.mnuEntry.Enabled = False
 frmMainScreen.mnuMedical.Enabled = False
 frmMainScreen.mnuView.Enabled = False
 frmMainScreen.miLogout.Enabled = False
 frmMainScreen.mnuChange.Enabled = False

End Sub

Private Sub miAI_Click()
frmMainScreen.Arrange vbArrangeIcons

End Sub

Private Sub miAP_Click()
frmHelp.Show


End Sub

Private Sub miCascade_Click()
frmMainScreen.Arrange vbCascade

End Sub

Private Sub miEJob_Click()
frmEmployeeJob.Show


End Sub

Private Sub miEM_Click()
frmEmergency.Show


End Sub

Private Sub miEPersonal_Click()
frmEmployeePersonal.Show


End Sub

Private Sub miLogin_Click()
frmPass.Show

End Sub

Private Sub miLogout_Click()
 frmMainScreen.mnuEntry.Enabled = False
 frmMainScreen.mnuMedical.Enabled = False
 frmMainScreen.mnuView.Enabled = False
 frmMainScreen.miLogin.Enabled = True
 frmMainScreen.miLogout.Enabled = False
 frmMainScreen.mnuChange.Enabled = False
End Sub

Private Sub miMS_Click()
frmMedical.Show


End Sub

Private Sub miOP_Click()
frmOPD.Show


End Sub

Private Sub miOPT_Click()
frmOperation.Show

End Sub

Private Sub miPD_Click()
frmDischarge.Show


End Sub

Private Sub miPMedical_Click()
frmPatientMedical.Show


End Sub

Private Sub miPPersonal_Click()
frmPatientPersonal.Show


End Sub

Private Sub miquit_Click()
End

End Sub

Private Sub miTH_Click()
frmMainScreen.Arrange vbTileHorizontal

End Sub

Private Sub miTV_Click()
frmMainScreen.Arrange vbTileVertical

End Sub

Private Sub mnuChange_Click()
frmChangePass.Show
End Sub

