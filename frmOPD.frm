VERSION 5.00
Begin VB.Form frmOPD 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outdoor Patient Information"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   6120
      Width           =   8655
   End
   Begin VB.Data datHospital 
      Connect         =   "Ms Access;pwd=nmhbahoo"
      DatabaseName    =   "Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   1320
      TabIndex        =   21
      Top             =   6720
      Width           =   8655
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
         Left            =   2040
         Picture         =   "frmOPD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   6840
         Picture         =   "frmOPD.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   5640
         Picture         =   "frmOPD.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   4440
         Picture         =   "frmOPD.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   3240
         Picture         =   "frmOPD.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   840
         Picture         =   "frmOPD.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Add the New Record"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4695
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   8655
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   4200
         TabIndex        =   20
         Top             =   4080
         Width           =   3975
      End
      Begin VB.ComboBox cmbUS 
         Height          =   315
         Left            =   4200
         TabIndex        =   19
         Top             =   3120
         Width           =   3975
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   3600
         Width           =   3975
      End
      Begin VB.ComboBox cmbSex 
         Height          =   315
         Left            =   4200
         TabIndex        =   17
         Top             =   2640
         Width           =   3975
      End
      Begin VB.ComboBox cmbSlip 
         Height          =   315
         Left            =   4200
         TabIndex        =   16
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   4680
         TabIndex        =   13
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Attendent Doctor Fee:"
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
         TabIndex        =   10
         Top             =   4080
         Width           =   3375
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Ultera-Sound Fee:"
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
         TabIndex        =   9
         Top             =   3600
         Width           =   3375
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Ultera-Sound:"
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
         TabIndex        =   8
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outdoor Patient Sex:"
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
         TabIndex        =   7
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outdoor Patient Checkup Date:"
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
         TabIndex        =   6
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outdoor Partient Number:"
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
         TabIndex        =   5
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outdoor Patient Father's/ Husband's Name:"
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
         TabIndex        =   4
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outdoor Patient Name:"
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
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outdoor Patient Slip No:"
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
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Image Image1 
      Height          =   2070
      Left            =   9960
      Picture         =   "frmOPD.frx":198C
      Top             =   0
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Outdoor Patient Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmOPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSlip_Click()
Text1 = cmbSlip.Text
cmbSlip.Visible = False
Text1.Visible = True
datHospital.RecordSource = "SELECT * FROM OPD WHERE SLIP_NO=val('" + Text1 + "')"
datHospital.Refresh
If datHospital.Recordset.RecordCount > 0 Then
    
         Text2 = datHospital.Recordset.Fields("PATIENT_NAME")
         Text3 = datHospital.Recordset.Fields("PATIENT_FNAME")
         Text4 = datHospital.Recordset.Fields("PATIENT_NO")
         Text5 = datHospital.Recordset.Fields("PATIENT_CDATE")
         Text6 = datHospital.Recordset.Fields("PATIENT_USFEE")
         Text7 = datHospital.Recordset.Fields("DOC_FEE")
         cmbSex.Text = datHospital.Recordset.Fields("PATIENT_SEX")
         cmbUS.Text = datHospital.Recordset.Fields("PATIENT_US")
Else
    MsgBox "Record not Found. Please Try again.", vbOKCancel
End If

End Sub

Private Sub cmdAdd_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text8.Text = "Add"
datHospital.RecordSource = "SELECT MAX(SLIP_NO) FROM OPD"
datHospital.Refresh
Text1 = datHospital.Recordset.Fields(0) + 1

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
Text8.Text = "Delete"
cmbSlip.Clear
datHospital.RecordSource = "SELECT SLIP_NO FROM OPD"
datHospital.Refresh
Do Until datHospital.Recordset.EOF
      cmbSlip.AddItem datHospital.Recordset.Fields("SLIP_NO")
      datHospital.Recordset.MoveNext
Loop
cmbSlip.Visible = True
cmbSlip.SetFocus

End Sub

Private Sub cmdModify_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text8.Text = "Modify"
datHospital.RecordSource = "SELECT SLIP_NO FROM OPD"
datHospital.Refresh
cmbSlip.Clear
Do Until datHospital.Recordset.EOF
      cmbSlip.AddItem datHospital.Recordset.Fields("SLIP_NO")
      datHospital.Recordset.MoveNext
Loop
cmbSlip.Visible = True
cmbSlip.SetFocus


End Sub

Private Sub cmdSave_Click()

datHospital.RecordSource = "OPD"
datHospital.Refresh
If Text8.Text = "Add" Then
         datHospital.Recordset.AddNew
         datHospital.Recordset.Fields("SLIP_NO") = Val(Text1)
         datHospital.Recordset.Fields("PATIENT_NAME") = Text2
         datHospital.Recordset.Fields("PATIENT_FNAME") = Text3
         datHospital.Recordset.Fields("PATIENT_NO") = Val(Text4)
         datHospital.Recordset.Fields("PATIENT_CDATE") = Text5
         datHospital.Recordset.Fields("PATIENT_USFEE") = Text6
         datHospital.Recordset.Fields("DOC_FEE") = Val(Text7)
         datHospital.Recordset.Fields("PATIENT_SEX") = cmbSex.Text
         datHospital.Recordset.Fields("PATIENT_US") = cmbUS.Text
         datHospital.Recordset.Update
         datHospital.Refresh
End If

If Text8.Text = "Modify" Then
     datHospital.Database.Execute "UPDATE  OPD SET SLIP_NO=VAL('" + Text1 + "'),PATIENT_NAME = ' " + Text2 + " ', PATIENT_FNAME ='" + Text3 + "',  PATIENT_NO = Val('" + Text4 + "'), PATIENT_CDATE ='" + Text5 + "', PATIENT_USFEE = '" + Text6 + "', DOC_FEE = Val('" + Text7 + "'), PATIENT_SEX = '" + cmbSex.Text + "',PATIENT_US = '" + cmbUS.Text + "' WHERE SLIP_NO=val('" + Text1 + "')"
     datHospital.Refresh
End If

If Text8.Text = "Delete" Then
     datHospital.Database.Execute "DELETE FROM OPD WHERE SLIP_NO=val('" + Text1 + "')"
     datHospital.Refresh
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
cmbSlip.Text = ""
cmbUS.Text = ""
cmbSex.Text = ""

cmbUS.AddItem "Yes"
cmbUS.AddItem "No"

cmbSex.AddItem "Female"
cmbSex.AddItem "Male"
cmbSex.AddItem "Child"

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbSlip.Visible = False



End Sub
