VERSION 5.00
Begin VB.Form frmChangePass 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtu 
         DataSource      =   "Data"
         Height          =   495
         Left            =   2640
         TabIndex        =   0
         Top             =   360
         Width           =   3375
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
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtp2 
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
         Left            =   2640
         TabIndex        =   3
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Change"
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
         Left            =   3360
         Picture         =   "frmChangePass.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Press Button to Login"
         Top             =   2280
         Width           =   1215
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
         Left            =   4680
         Picture         =   "frmChangePass.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Press Button and go Main Screen"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Data Data 
         Connect         =   "Ms Access;pwd=nmhbahoo"
         DatabaseName    =   "Hospital.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   -120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SYSTEMUSERS"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter Username:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter Password:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter Confirm Password:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txtp.Text = txtp2.Text Then
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace

Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Hospital.mdb")
    PwdString = "nmhbahoo"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("SYSTEMUSERS", dbOpenTable)
rs.Edit
rs("PASSWORD") = txtp2.Text
rs("USERNAME") = txtu.Text
rs.Update
UsernameAndPasswordLastShow = MsgBox("Remember your password and your username!" & Chr(13) & Chr(13) & Chr(13) & "Username = " & txtu.Text & Chr(13) & "Password = " & txtp.Text, vbInformation, "Warning")
Unload Me
Else
MsgBox "Passwords don't match!", vbCritical, "Warning!"
txtp.Text = vbNullString
txtp2.Text = vbNullString
End If
End Sub

Private Sub Command2_Click()
Unload Me
App.TaskVisible = False
End Sub
