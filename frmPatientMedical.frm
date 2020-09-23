VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPatientMedical 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Nighat Maternity Home And General Hospital             (Patient Medical Information)"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data dtHospitalPM 
      Connect         =   "Ms Access;pwd=nmhbahoo"
      DatabaseName    =   "Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PATIENT_MEDICAL"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid msfHospital 
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton cmdSearchRoom 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search By Room No"
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
         Left            =   9480
         Picture         =   "frmPatientMedical.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Click on This button to Search By Room No"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search By Patient No"
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
         Left            =   7440
         Picture         =   "frmPatientMedical.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click on Button for Search the Patient No and all Record"
         Top             =   240
         Width           =   2055
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
         Left            =   240
         Picture         =   "frmPatientMedical.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Add the New Record"
         Top             =   240
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
         Left            =   2640
         Picture         =   "frmPatientMedical.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Press the Button and Delete the Records."
         Top             =   240
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
         Left            =   3840
         Picture         =   "frmPatientMedical.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Press the Button to Save the Records."
         Top             =   240
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
         Left            =   5040
         Picture         =   "frmPatientMedical.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Press the Button to Cancel the All Operations"
         Top             =   240
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
         Left            =   6240
         Picture         =   "frmPatientMedical.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Back to Main Screen"
         Top             =   240
         Width           =   1215
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
         Left            =   1440
         Picture         =   "frmPatientMedical.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Press the Button and Modify/Change the Records"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   1695
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   3240
         Width           =   4575
      End
      Begin VB.TextBox Text8 
         Height          =   735
         Left            =   7440
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   8760
         TabIndex        =   35
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   8760
         TabIndex        =   34
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   8760
         TabIndex        =   33
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cmbOperation 
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   8760
         TabIndex        =   31
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cmbPatSex 
         Height          =   315
         Left            =   2160
         TabIndex        =   30
         Top             =   1920
         Width           =   3015
      End
      Begin VB.ComboBox cmbPatMNo 
         Height          =   315
         Left            =   2160
         TabIndex        =   29
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2160
         TabIndex        =   28
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   27
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   480
         Width           =   3015
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ward and Room Information"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   1815
         Left            =   6360
         TabIndex        =   18
         Top             =   3120
         Width           =   5295
         Begin VB.ComboBox cmbWard 
            Height          =   315
            Left            =   2640
            TabIndex        =   40
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   2640
            TabIndex        =   39
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   2640
            TabIndex        =   38
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Patient Room No:"
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
            Left            =   480
            TabIndex        =   25
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Patient Bed Number:"
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
            Left            =   480
            TabIndex        =   24
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Patient Ward Name:"
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
            Left            =   480
            TabIndex        =   23
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Admission Date:"
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
         Left            =   5400
         TabIndex        =   22
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label10 
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
         Left            =   5400
         TabIndex        =   21
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Operation:"
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
         TabIndex        =   20
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Medicines:"
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
         Left            =   5400
         TabIndex        =   19
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Reference Doctor Name:"
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
         Left            =   5400
         TabIndex        =   17
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Medical History:"
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
         Left            =   1680
         TabIndex        =   16
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label5 
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
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         TabIndex        =   14
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
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
         Left            =   5400
         TabIndex        =   13
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label2 
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
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmPatientMedical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPatMNo_Click()
Text1.Text = cmbPatMNo.Text
cmbPatMNo.Visible = False
Text1.Visible = True
dtHospitalPM.RecordSource = "select * from PATIENT_MEDICAL where PAT_NO=Val('" + Text1 + "')"
dtHospitalPM.Refresh
If dtHospitalPM.Recordset.RecordCount > 0 Then
       Text2 = dtHospitalPM.Recordset.Fields("PAT_NAME")
       Text3 = dtHospitalPM.Recordset.Fields("PAT_AGE")
       Text4 = dtHospitalPM.Recordset.Fields("PAT_CHDATE")
       Text5 = dtHospitalPM.Recordset.Fields("PAT_ADMIND")
       Text6 = dtHospitalPM.Recordset.Fields("PAT_FNAME")
       Text7 = dtHospitalPM.Recordset.Fields("PAT_REFDNAME")
       Text8 = dtHospitalPM.Recordset.Fields("PAT_DES")
       Text9 = dtHospitalPM.Recordset.Fields("PAT_HIS")
       Text10 = dtHospitalPM.Recordset.Fields("PAT_BEDNO")
       Text11 = dtHospitalPM.Recordset.Fields("PAT_ROOMNO")
       cmbPatSex.Text = dtHospitalPM.Recordset.Fields("PAT_SEX")
       cmbOperation.Text = dtHospitalPM.Recordset.Fields("PAT_OP")
       cmbWard.Text = dtHospitalPM.Recordset.Fields("PAT_WNAME")
 Else
    MsgBox "Record not Found. Plz Try Again.", vbOKCancel
End If

End Sub

Private Sub cmdAdd_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text12.Text = "Add"
dtHospitalPM.RecordSource = "select MAX(PAT_NO) from PATIENT_MEDICAL"
dtHospitalPM.Refresh
Text1 = dtHospitalPM.Recordset.Fields(0) + 1

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
Text12.Text = "Delete"
cmbPatMNo.Clear
dtHospitalPM.RecordSource = "select PAT_NO from PATIENT_MEDICAL"
dtHospitalPM.Refresh
Do Until dtHospitalPM.Recordset.EOF
        cmbPatMNo.AddItem dtHospitalPM.Recordset.Fields("PAT_NO")
        dtHospitalPM.Recordset.MoveNext
Loop
cmbPatMNo.Visible = True
cmbPatMNo.SetFocus


End Sub

Private Sub cmdModify_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text12.Text = "Modify"
dtHospitalPM.RecordSource = "select PAT_NO from PATIENT_MEDICAL"
dtHospitalPM.Refresh
cmbPatMNo.Clear
Do Until dtHospitalPM.Recordset.EOF
        cmbPatMNo.AddItem dtHospitalPM.Recordset.Fields("PAT_NO")
        dtHospitalPM.Recordset.MoveNext
Loop
cmbPatMNo.Visible = True
cmbPatMNo.SetFocus


End Sub

Private Sub cmdSave_Click()
dtHospitalPM.RecordSource = "PATIENT_MEDICAL"
dtHospitalPM.Refresh
If Text12.Text = "Add" Then
        dtHospitalPM.Recordset.AddNew
        dtHospitalPM.Recordset.Fields("PAT_NO") = Val(Text1)
        dtHospitalPM.Recordset.Fields("PAT_NAME") = Text2
        dtHospitalPM.Recordset.Fields("PAT_AGE") = Val(Text3)
        dtHospitalPM.Recordset.Fields("PAT_CHDATE") = Text4
        dtHospitalPM.Recordset.Fields("PAT_ADMIND") = Text5
        dtHospitalPM.Recordset.Fields("PAT_FNAME") = Text6
        dtHospitalPM.Recordset.Fields("PAT_REFDNAME") = Text7
        dtHospitalPM.Recordset.Fields("PAT_DES") = Text8
        dtHospitalPM.Recordset.Fields("PAT_HIS") = Text9
        dtHospitalPM.Recordset.Fields("PAT_BEDNO") = Val(Text10)
        dtHospitalPM.Recordset.Fields("PAT_ROOMNO") = Val(Text11)
        dtHospitalPM.Recordset.Fields("PAT_SEX") = cmbPatSex.Text
        dtHospitalPM.Recordset.Fields("PAT_OP") = cmbOperation.Text
        dtHospitalPM.Recordset.Fields("PAT_WNAME") = cmbWard.Text
        dtHospitalPM.Recordset.Update
        dtHospitalPM.Refresh
End If

If Text12.Text = "Modify" Then
       dtHospitalPM.Database.Execute "update PATIENT_MEDICAL set PAT_NO=Val('" + Text1 + "'), PAT_NAME ='" + Text2 + "', PAT_AGE = Val('" + Text3 + "'), PAT_CHDATE ='" + Text4 + "', PAT_ADMIND ='" + Text5 + "', PAT_FNAME ='" + Text6 + "', PAT_REFDNAME ='" + Text7 + "', PAT_DES ='" + Text8 + "', PAT_HIS ='" + Text9 + "', PAT_BEDNO = Val('" + Text10 + "'),PAT_ROOMNO = Val('" + Text11 + "'), PAT_SEX ='" + cmbPatSex.Text + "',PAT_OP = '" + cmbOperation.Text + "', PAT_WNAME = '" + cmbWard.Text + "' where PAT_NO =Val('" + Text1 + "')"
       dtHospitalPM.Refresh
End If

If Text12.Text = "Delete" Then
       dtHospitalPM.Database.Execute "delete from PATIENT_MEDICAL where PAT_NO=Val('" + Text1 + "' )"
       dtHospitalPM.Refresh
End If

Call Form_Load

        
        
        
End Sub

Private Sub cmdSearch_Click()
dtHospitalPM.RecordSource = "select * from PATIENT_MEDICAL where PAT_NO=Val('" + Text1 + "')"
dtHospitalPM.Refresh
If dtHospitalPM.Recordset.RecordCount > 0 Then
       Text2 = dtHospitalPM.Recordset.Fields("PAT_NAME")
       Text3 = dtHospitalPM.Recordset.Fields("PAT_AGE")
       Text4 = dtHospitalPM.Recordset.Fields("PAT_CHDATE")
       Text5 = dtHospitalPM.Recordset.Fields("PAT_ADMIND")
       Text6 = dtHospitalPM.Recordset.Fields("PAT_FNAME")
       Text7 = dtHospitalPM.Recordset.Fields("PAT_REFDNAME")
       Text8 = dtHospitalPM.Recordset.Fields("PAT_DES")
       Text9 = dtHospitalPM.Recordset.Fields("PAT_HIS")
       Text10 = dtHospitalPM.Recordset.Fields("PAT_BEDNO")
       Text11 = dtHospitalPM.Recordset.Fields("PAT_ROOMNO")
       cmbPatSex.Text = dtHospitalPM.Recordset.Fields("PAT_SEX")
       cmbOperation.Text = dtHospitalPM.Recordset.Fields("PAT_OP")
       cmbWard.Text = dtHospitalPM.Recordset.Fields("PAT_WNAME")
 Else
    MsgBox "Record not Found. Plz Try Again.", vbOKCancel
End If


End Sub

Private Sub cmdSearchRoom_Click()
dtHospitalPM.RecordSource = "select * from PATIENT_MEDICAL where PAT_ROOMNO=Val('" + Text11 + "')"
dtHospitalPM.Refresh
If dtHospitalPM.Recordset.RecordCount > 0 Then
       Text1 = dtHospitalPM.Recordset.Fields("PAT_NO")
       Text2 = dtHospitalPM.Recordset.Fields("PAT_NAME")
       Text3 = dtHospitalPM.Recordset.Fields("PAT_AGE")
       Text4 = dtHospitalPM.Recordset.Fields("PAT_CHDATE")
       Text5 = dtHospitalPM.Recordset.Fields("PAT_ADMIND")
       Text6 = dtHospitalPM.Recordset.Fields("PAT_FNAME")
       Text7 = dtHospitalPM.Recordset.Fields("PAT_REFDNAME")
       Text8 = dtHospitalPM.Recordset.Fields("PAT_DES")
       Text9 = dtHospitalPM.Recordset.Fields("PAT_HIS")
       Text10 = dtHospitalPM.Recordset.Fields("PAT_BEDNO")
       cmbPatSex.Text = dtHospitalPM.Recordset.Fields("PAT_SEX")
       cmbOperation.Text = dtHospitalPM.Recordset.Fields("PAT_OP")
       cmbWard.Text = dtHospitalPM.Recordset.Fields("PAT_WNAME")
 Else
    MsgBox "Record not Found. Plz Try Again.", vbOKCancel
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
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
cmbPatSex.Text = ""
cmbWard.Text = ""
cmbOperation.Text = ""
cmbPatMNo.Text = ""

cmbPatSex.AddItem "Male"
cmbPatSex.AddItem "Female"

cmbOperation.AddItem "Yes"
cmbOperation.AddItem "No"

cmbWard.AddItem "ENT Ward"
cmbWard.AddItem "Surgical Ward"
cmbWard.AddItem "Children Ward"
cmbWard.AddItem "Urology Ward"
cmbWard.AddItem "Children Ward"
cmbWard.AddItem "Cardiology Ward"
cmbWard.AddItem "ICCU Room"
cmbWard.AddItem "Gynae Ward"

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbPatMNo.Visible = False


'==========================
'grid Coding.
'==========================

msfHospital.Cols = 1
msfHospital.Rows = 2
msfHospital.Cols = 15
msfHospital.FixedCols = 1
msfHospital.FixedRows = 1
msfHospital.ColWidth(1) = 1000
msfHospital.ColWidth(2) = 2000
msfHospital.ColWidth(3) = 2000
msfHospital.ColWidth(4) = 1000
msfHospital.ColWidth(5) = 1500
msfHospital.ColWidth(6) = 500
msfHospital.ColWidth(7) = 1500
msfHospital.ColWidth(8) = 1500
msfHospital.ColWidth(9) = 3000
msfHospital.ColWidth(10) = 1000
msfHospital.ColWidth(11) = 1000
msfHospital.ColWidth(12) = 2500
msfHospital.ColWidth(13) = 1500
msfHospital.ColWidth(14) = 1500
msfHospital.Row = 0
msfHospital.Col = 1
msfHospital.Text = "Patient No"
msfHospital.Col = 2
msfHospital.Text = "Name"
msfHospital.Col = 3
msfHospital.Text = "D/o,W/o"
msfHospital.Col = 4
msfHospital.Text = "CheckUp"
msfHospital.Col = 5
msfHospital.Text = "Admission"
msfHospital.Col = 6
msfHospital.Text = "Age"
msfHospital.Col = 7
msfHospital.Text = "Doctor"
msfHospital.Col = 8
msfHospital.Text = "Medicines"
msfHospital.Col = 9
msfHospital.Text = "History"
msfHospital.Col = 10
msfHospital.Text = "BedNo"
msfHospital.Col = 11
msfHospital.Text = "RoomNo"
msfHospital.Col = 12
msfHospital.Text = "Ward"
msfHospital.Col = 13
msfHospital.Text = "Sex"
msfHospital.Col = 14
msfHospital.Text = "Operation"

dtHospitalPM.RecordSource = "PATIENT_MEDICAL"
dtHospitalPM.Refresh
dtHospitalPM.Recordset.MoveNext
msfHospital.Rows = dtHospitalPM.Recordset.RecordCount + 1
Do Until dtHospitalPM.Recordset.EOF
       msfHospital.Col = 1
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_NO")
       msfHospital.Col = 2
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_NAME")
       msfHospital.Col = 3
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_FNAME")
       msfHospital.Col = 4
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_CHDATE")
       msfHospital.Col = 5
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_ADMIND")
       msfHospital.Col = 6
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_AGE")
       msfHospital.Col = 7
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_REFDNAME")
       msfHospital.Col = 8
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_DES")
       msfHospital.Col = 9
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_HIS")
       msfHospital.Col = 10
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_BEDNO")
       msfHospital.Col = 11
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_ROOMNO")
       msfHospital.Col = 12
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_WNAME")
       msfHospital.Col = 13
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_SEX")
       msfHospital.Col = 14
       msfHospital.Text = dtHospitalPM.Recordset.Fields("PAT_OP")
       dtHospitalPM.Recordset.MoveNext
       msfHospital.Row = msfHospital.Row + 1
Loop

End Sub
