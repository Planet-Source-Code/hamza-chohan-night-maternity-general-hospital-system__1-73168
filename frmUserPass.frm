VERSION 5.00
Begin VB.Form frmPass 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Screen"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      Begin VB.Data Data 
         Connect         =   "Ms Access;pwd=nmhbahoo"
         DatabaseName    =   "Hospital.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SYSTEMUSERS"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3240
         Picture         =   "frmUserPass.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Press Button and go Main Screen"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2040
         Picture         =   "frmUserPass.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Press Button to Login"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtp 
         DataSource      =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtu 
         DataSource      =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter User Password:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter User Login:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txtu.Text = Data.Recordset!UserName And txtp.Text = Data.Recordset!Password Then
Unload Me
 frmMainScreen.mnuEntry.Enabled = True
 frmMainScreen.mnuMedical.Enabled = True
 frmMainScreen.mnuView.Enabled = True
 frmMainScreen.miLogin.Enabled = False
 frmMainScreen.miLogout.Enabled = True
 frmMainScreen.mnuChange.Enabled = True
MsgBox "WELCOME TO NIGHT MATERNITY HOME AND GENERAL HOSPITAL SYSTEM", vbInformation, "WELCOME"
Else
MsgBox "User Not Found", vbCritical, "User Information"
txtp.SetFocus
 frmMainScreen.mnuEntry.Enabled = False
 frmMainScreen.mnuMedical.Enabled = False
 frmMainScreen.mnuView.Enabled = False
 frmMainScreen.miLogin.Enabled = True
 frmMainScreen.miLogout.Enabled = False
 frmMainScreen.mnuChange.Enabled = False
 
End If
End Sub

Private Sub Command2_Click()
Unload frmPass
End Sub

