VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmployeeJob 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Nighat Maternity Home And General Hospital     (Employee Job Information)"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data dtHospital 
      BackColor       =   &H00E0E0E0&
      Connect         =   "Ms Access;pwd=nmhbahoo"
      DatabaseName    =   "Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EMPLOYEEJOB"
      Top             =   1920
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   10
      Top             =   6960
      Width           =   11415
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search by National Identity Card"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7560
         Picture         =   "frmEmployeeJob.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Click on This Button to Search the Record By National Identity Card"
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1560
         Picture         =   "frmEmployeeJob.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Press the Button and Modify/Change the Records"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6360
         Picture         =   "frmEmployeeJob.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Back to Main Screen"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5160
         Picture         =   "frmEmployeeJob.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Press the Button to Cancel the All Operations"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3960
         Picture         =   "frmEmployeeJob.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Press the Button to Save the Records."
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2760
         Picture         =   "frmEmployeeJob.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Press the Button and Delete the Records."
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         Picture         =   "frmEmployeeJob.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Add the New Record"
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msfHospital 
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   11415
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   3480
         TabIndex        =   32
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2640
         Width           =   11175
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   9240
         TabIndex        =   29
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   9240
         TabIndex        =   28
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   9240
         TabIndex        =   27
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   9240
         TabIndex        =   26
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cmbJobstatus 
         Height          =   315
         Left            =   3480
         TabIndex        =   25
         Top             =   2160
         Width           =   2895
      End
      Begin VB.ComboBox cmbJob 
         Height          =   315
         Left            =   3480
         TabIndex        =   24
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox cmbSex 
         Height          =   315
         Left            =   9240
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbEmpNic 
         Height          =   315
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3480
         TabIndex        =   21
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Identity Number:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Sex:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6480
         TabIndex        =   19
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Total Pay:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6480
         TabIndex        =   18
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Allowances:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6480
         TabIndex        =   17
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Basic Salary:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6480
         TabIndex        =   16
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Working Timing:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6480
         TabIndex        =   15
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Job:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Job Status:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee National Identity No:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   3375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   9960
      Picture         =   "frmEmployeeJob.frx":1DCE
      ScaleHeight     =   2115
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Employee Job Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   6375
   End
End
Attribute VB_Name = "frmEmployeeJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbEmpNic_Click()
Text1.Text = cmbEmpNic.Text
cmbEmpNic.Visible = False
Text1.Visible = True
dtHospital.RecordSource = "select * from EMPLOYEEJOB where EMP_IDNO=Val('" + Text1 + "')"
dtHospital.Refresh
If dtHospital.Recordset.RecordCount > 0 Then
      Text2 = dtHospital.Recordset.Fields("EMP_NAME")
      Text3 = dtHospital.Recordset.Fields("EMP_TIMING")
      Text4 = dtHospital.Recordset.Fields("EMP_SAL")
      Text5 = dtHospital.Recordset.Fields("EMP_ALL")
      Text6 = dtHospital.Recordset.Fields("EMP_TPAY")
      Text8 = dtHospital.Recordset.Fields("EMP_NICNO")
      cmbJob.Text = dtHospital.Recordset.Fields("EMP_JOB")
      cmbJobstatus.Text = dtHospital.Recordset.Fields("EMP_JOBSTATUS")
      cmbSex.Text = dtHospital.Recordset.Fields("EMP_SEX")
   Else
        MsgBox "Record Not Found, Please Try Again.", vbOKCancel
End If

End Sub

Private Sub cmdAdd_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text7.Text = "Add"
dtHospital.RecordSource = "select MAX(EMP_IDNO) from EMPLOYEEJOB"
dtHospital.Refresh
Text1 = dtHospital.Recordset.Fields(0) + 1

End Sub

Private Sub cmdCancel_Click()
Call Form_Load

End Sub

Private Sub cmdClose_Click()
frmMainScreen.Show
Unload Me


End Sub

Private Sub cmdDelete_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text7.Text = "Delete"
dtHospital.RecordSource = "select EMP_IDNO from EMPLOYEEJOB"
dtHospital.Refresh
cmbEmpNic.Clear
Do Until dtHospital.Recordset.EOF
          cmbEmpNic.AddItem dtHospital.Recordset.Fields("EMP_IDNO")
          dtHospital.Recordset.MoveNext
Loop
cmbEmpNic.Visible = True
cmbEmpNic.SetFocus

End Sub

Private Sub cmdModify_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text7.Text = "Modify"
dtHospital.RecordSource = "select EMP_IDNO from EMPLOYEEJOB"
dtHospital.Refresh
cmbEmpNic.Clear
Do Until dtHospital.Recordset.EOF
          cmbEmpNic.AddItem dtHospital.Recordset.Fields("EMP_IDNO")
          dtHospital.Recordset.MoveNext
Loop
cmbEmpNic.Visible = True
cmbEmpNic.SetFocus

          
End Sub

Private Sub cmdSave_Click()
dtHospital.RecordSource = "EMPLOYEEJOB"
dtHospital.Refresh
If Text7.Text = "Add" Then
         dtHospital.Recordset.AddNew
         dtHospital.Recordset.Fields("EMP_IDNO") = Val(Text1)
         dtHospital.Recordset.Fields("EMP_NAME") = Text2
         dtHospital.Recordset.Fields("EMP_TIMING") = Text3
         dtHospital.Recordset.Fields("EMP_SAL") = Val(Text4)
         dtHospital.Recordset.Fields("EMP_ALL") = Val(Text5)
         dtHospital.Recordset.Fields("EMP_TPAY") = Val(Text6)
         dtHospital.Recordset.Fields("EMP_NICNO") = Text8
         dtHospital.Recordset.Fields("EMP_JOB") = cmbJob.Text
         dtHospital.Recordset.Fields("EMP_JOBSTATUS") = cmbJobstatus.Text
         dtHospital.Recordset.Fields("EMP_SEX") = cmbSex.Text
         dtHospital.Recordset.Update
         dtHospital.Refresh
End If
If Text7.Text = "Modify" Then
         dtHospital.Database.Execute "update EMPLOYEEJOB set EMP_IDNO=Val('" + Text1 + "'),EMP_NAME ='" + Text2 + "',EMP_TIMING = '" + Text3 + "',EMP_SAL = Val('" + Text4 + "'),EMP_ALL = Val('" + Text5 + "'),EMP_TPAY = Val('" + Text6 + "'),EMP_NICNO = '" + Text8 + "',EMP_JOB ='" + cmbJob.Text + "',EMP_JOBSTATUS = '" + cmbJobstatus.Text + "',EMP_SEX ='" + cmbSex.Text + "' where EMP_IDNO=Val('" + Text1 + "')"
         dtHospital.Refresh
End If
If Text7.Text = "Delete" Then
         dtHospital.Database.Execute "delete from EMPLOYEEJOB where EMP_IDNO=Val('" + Text1 + "')'"
         dtHospital.Refresh
End If
Call Form_Load

End Sub

Private Sub cmdSearch_Click()
dtHospital.RecordSource = "select * from EMPLOYEEJOB where EMP_NICNO=('" + Text8 + "')"
dtHospital.Refresh
If dtHospital.Recordset.RecordCount > 0 Then
      Text1 = dtHospital.Recordset.Fields("EMP_IDNO")
      Text2 = dtHospital.Recordset.Fields("EMP_NAME")
      Text3 = dtHospital.Recordset.Fields("EMP_TIMING")
      Text4 = dtHospital.Recordset.Fields("EMP_SAL")
      Text5 = dtHospital.Recordset.Fields("EMP_ALL")
      Text6 = dtHospital.Recordset.Fields("EMP_TPAY")
      cmbJob.Text = dtHospital.Recordset.Fields("EMP_JOB")
      cmbJobstatus.Text = dtHospital.Recordset.Fields("EMP_JOBSTATUS")
      cmbSex.Text = dtHospital.Recordset.Fields("EMP_SEX")
   Else
        MsgBox "Record Not Found, Please Try Again.", vbOKCancel
End If

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
cmbSex.Text = ""
cmbJob.Text = ""
cmbJobstatus.Text = ""
cmbEmpNic.Text = ""


cmbSex.AddItem "Male"
cmbSex.AddItem "Female"

cmbJob.AddItem "Doctor"
cmbJob.AddItem "Nurse"
cmbJob.AddItem "Boy"
cmbJob.AddItem "Cleaner'"
cmbJob.AddItem "Dai"
cmbJob.AddItem "Operator"
cmbJob.AddItem "Dispenser"

cmbJobstatus.AddItem "Permanent"
cmbJobstatus.AddItem "Visitor"

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbEmpNic.Visible = False


'==========================
'grid Coding.
'==========================

msfHospital.Cols = 1
msfHospital.Rows = 2
msfHospital.Cols = 11
msfHospital.FixedCols = 1
msfHospital.FixedRows = 1
msfHospital.ColWidth(1) = 2000
msfHospital.ColWidth(2) = 2500
msfHospital.ColWidth(3) = 1500
msfHospital.ColWidth(4) = 2000
msfHospital.ColWidth(5) = 2000
msfHospital.ColWidth(6) = 1000
msfHospital.ColWidth(7) = 1000
msfHospital.ColWidth(8) = 1000
msfHospital.ColWidth(9) = 1000
msfHospital.ColWidth(10) = 1500
msfHospital.Row = 0
msfHospital.Col = 1
msfHospital.Text = "Employee No"
msfHospital.Col = 2
msfHospital.Text = "Employee Name"
msfHospital.Col = 3
msfHospital.Text = "Timing"
msfHospital.Col = 4
msfHospital.Text = "Employee Job "
msfHospital.Col = 5
msfHospital.Text = "Job Status"
msfHospital.Col = 6
msfHospital.Text = "Salary"
msfHospital.Col = 7
msfHospital.Text = "Allowance"
msfHospital.Col = 8
msfHospital.Text = "Net Pay"
msfHospital.Col = 9
msfHospital.Text = "Sex"
msfHospital.Col = 10
msfHospital.Text = "NIC NO"

dtHospital.RecordSource = "EMPLOYEEJOB"
dtHospital.Refresh
dtHospital.Recordset.MoveNext
msfHospital.Rows = dtHospital.Recordset.RecordCount + 1
Do Until dtHospital.Recordset.EOF
       msfHospital.Col = 1
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_IDNO")
       msfHospital.Col = 2
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_NAME")
       msfHospital.Col = 3
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_TIMING")
       msfHospital.Col = 4
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_JOB")
       msfHospital.Col = 5
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_JOBSTATUS")
       msfHospital.Col = 6
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_SAL")
       msfHospital.Col = 7
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_ALL")
       msfHospital.Col = 8
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_TPAY")
       msfHospital.Col = 9
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_SEX")
       msfHospital.Col = 10
       msfHospital.Text = dtHospital.Recordset.Fields("EMP_NICNO")
       dtHospital.Recordset.MoveNext
       msfHospital.Row = msfHospital.Row + 1
Loop


End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Text6.Text = Val(Text4) + Val(Text5)
End If

End Sub
