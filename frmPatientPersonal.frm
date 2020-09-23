VERSION 5.00
Begin VB.Form frmPatientPersonal 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Nighat Maternity Home And General Hospital"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data dtHospital 
      BackColor       =   &H00E0E0E0&
      Connect         =   "Ms Access;pwd=nmhbahoo"
      DatabaseName    =   "Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   11415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   6600
      Width           =   11415
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
         Left            =   3120
         Picture         =   "frmPatientPersonal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   7920
         Picture         =   "frmPatientPersonal.frx":0442
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
         Left            =   6720
         Picture         =   "frmPatientPersonal.frx":0884
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
         Left            =   5520
         Picture         =   "frmPatientPersonal.frx":0CC6
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
         Left            =   4320
         Picture         =   "frmPatientPersonal.frx":1108
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
         Left            =   1920
         Picture         =   "frmPatientPersonal.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add the New Record"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   11415
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   4680
         Width           =   10095
      End
      Begin VB.TextBox Text10 
         Height          =   1335
         Left            =   6000
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   3240
         Width           =   4815
      End
      Begin VB.TextBox Text9 
         Height          =   1335
         Left            =   720
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   3240
         Width           =   5055
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   8400
         TabIndex        =   31
         Top             =   2400
         Width           =   2775
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         Left            =   8400
         TabIndex        =   30
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   8400
         TabIndex        =   29
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   8400
         TabIndex        =   28
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   8400
         TabIndex        =   27
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2400
         TabIndex        =   26
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2640
         TabIndex        =   25
         Top             =   1920
         Width           =   2535
      End
      Begin VB.ComboBox cmbSex 
         Height          =   315
         Left            =   2400
         TabIndex        =   24
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox cmbPatNo 
         Height          =   315
         Left            =   2400
         TabIndex        =   23
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2400
         TabIndex        =   22
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient City/Village:"
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
         Left            =   5280
         TabIndex        =   20
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Contact Address."
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
         Left            =   5760
         TabIndex        =   19
         Top             =   2880
         Width           =   5055
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Permanent Address."
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
         Left            =   720
         TabIndex        =   18
         Top             =   2880
         Width           =   5055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Age:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Checkup Date:"
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
         Left            =   5280
         TabIndex        =   16
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Sex:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Date of Birth:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Father/Husband Name:"
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
         Left            =   5280
         TabIndex        =   13
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Status:"
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
         Left            =   5280
         TabIndex        =   12
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Name:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Phone Number:"
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
         Left            =   5280
         TabIndex        =   10
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Number:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Patient Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmPatientPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPatno_Click()
Text1.Text = cmbPatno.Text
cmbPatno.Visible = False
Text1.Visible = True
dtHospital.RecordSource = "select * from PATIENT_PERSONAL where PAT_NO=val('" + Text1 + "')"
dtHospital.Refresh
If dtHospital.Recordset.RecordCount > 0 Then
  Text2 = dtHospital.Recordset.Fields("PAT_NAME")
  Text3 = dtHospital.Recordset.Fields("PAT_DOB")
  Text4 = dtHospital.Recordset.Fields("PAT_AGE")
  Text5 = dtHospital.Recordset.Fields("PAT_CDATE")
  Text6 = dtHospital.Recordset.Fields("PAT_FNAME")
  Text7 = dtHospital.Recordset.Fields("PAT_PHONENO")
  Text8 = dtHospital.Recordset.Fields("PAT_CITY")
  Text9 = dtHospital.Recordset.Fields("PAT_PADD")
  Text10 = dtHospital.Recordset.Fields("PAT_CADD")
  cmbSex.Text = dtHospital.Recordset.Fields("PAT_SEX")
  cmbStatus.Text = dtHospital.Recordset.Fields("PAT_STATUS")
Else
        MsgBox "Record not Found. Plz try Again", vbOKCancel
End If

End Sub

Private Sub cmdAdd_Click()

cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text11.Text = "Add"
dtHospital.RecordSource = "select MAX(PAT_NO) from PATIENT_PERSONAL"
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
Text11.Text = "Delete"
cmbPatno.Clear
dtHospital.RecordSource = "select PAT_NO from PATIENT_PERSONAL"
dtHospital.Refresh
Do Until dtHospital.Recordset.EOF
     cmbPatno.AddItem dtHospital.Recordset.Fields("PAT_NO")
     dtHospital.Recordset.MoveNext
Loop
cmbPatno.Visible = True
cmbPatno.SetFocus

End Sub

Private Sub cmdModify_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text11.Text = "Modify"
dtHospital.RecordSource = "select PAT_NO from PATIENT_PERSONAL"
dtHospital.Refresh
cmbPatno.Clear
Do Until dtHospital.Recordset.EOF
     cmbPatno.AddItem dtHospital.Recordset.Fields("PAT_NO")
     dtHospital.Recordset.MoveNext
Loop
cmbPatno.Visible = True
cmbPatno.SetFocus


End Sub

Private Sub cmdSave_Click()
dtHospital.RecordSource = "PATIENT_PERSONAL"
dtHospital.Refresh
If Text11.Text = "Add" Then
        dtHospital.Recordset.AddNew
        dtHospital.Recordset.Fields("PAT_NO") = Val(Text1)
        dtHospital.Recordset.Fields("PAT_NAME") = Text2
        dtHospital.Recordset.Fields("PAT_DOB") = Text3
        dtHospital.Recordset.Fields("PAT_AGE") = Val(Text4)
        dtHospital.Recordset.Fields("PAT_CDATE") = Text5
        dtHospital.Recordset.Fields("PAT_FNAME") = Text6
        dtHospital.Recordset.Fields("PAT_PHONENO") = Text7
        dtHospital.Recordset.Fields("PAT_CITY") = Text8
        dtHospital.Recordset.Fields("PAT_PADD") = Text9
        dtHospital.Recordset.Fields("PAT_CADD") = Text10
        dtHospital.Recordset.Fields("PAT_SEX") = cmbSex.Text
        dtHospital.Recordset.Fields("PAT_STATUS") = cmbStatus.Text
        dtHospital.Recordset.Update
        dtHospital.Refresh
End If

If Text11.Text = "Modify" Then
       dtHospital.Database.Execute "update PATIENT_PERSONAL set PAT_NO=Val('" + Text1 + "'), PAT_NAME = '" + Text2 + "', PAT_DOB = '" + Text3 + "', PAT_AGE = Val('" + Text4 + "'), PAT_CDATE ='" + Text5 + "', PAT_FNAME ='" + Text6 + "', PAT_PHONENO = '" + Text7 + "', PAT_CITY ='" + Text8 + "', PAT_PADD ='" + Text9 + "', PAT_CADD ='" + Text10 + "',PAT_SEX = '" + cmbSex.Text + "',PAT_STATUS ='" + cmbStatus.Text + "' where PAT_NO=val('" + Text1 + "')"
       dtHospital.Refresh
End If

If Text11.Text = "Delete" Then
      dtHospital.Database.Execute "delete from PATIENT_PERSONAL where PAT_NO=val('" + Text1 + "')"
      dtHospital.Refresh
End If

Call Form_Load

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
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
cmbPatno.Text = ""
cmbSex.Text = ""
cmbStatus.Text = ""

cmbSex.AddItem "Male"
cmbSex.AddItem "Female"
cmbSex.AddItem "Child"
cmbSex.AddItem "Girl"
cmbSex.AddItem "Boy"

cmbStatus.AddItem "Government"
cmbStatus.AddItem "Civil"
cmbStatus.AddItem "Private"
cmbStatus.AddItem "Sami Government"
cmbStatus.AddItem "Relative"

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbPatno.Visible = False

End Sub
