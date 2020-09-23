VERSION 5.00
Begin VB.Form frmEmergency 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emergency Patient Information "
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   2040
      TabIndex        =   28
      Top             =   6720
      Width           =   7695
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
         Picture         =   "frmEmergency.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add the New Record"
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
         Left            =   2640
         Picture         =   "frmEmergency.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Press the Button and Delete the Records."
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
         Left            =   3840
         Picture         =   "frmEmergency.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Press the Button to Save the Records."
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
         Left            =   5040
         Picture         =   "frmEmergency.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Press the Button to Cancel the All Operations"
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
         Left            =   6240
         Picture         =   "frmEmergency.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Back to Main Screen"
         Top             =   360
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
         Picture         =   "frmEmergency.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Press the Button and Modify/Change the Records"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   6120
      Width           =   7695
   End
   Begin VB.Data datHospital 
      Connect         =   "Ms Access;pwd=nmhbahoo"
      DatabaseName    =   "Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4935
      Left            =   2040
      TabIndex        =   16
      Top             =   1080
      Width           =   7815
      Begin VB.ComboBox cmbSlipno 
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.ComboBox cmbOperation 
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Top             =   3720
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3840
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   2760
         Width           =   3375
      End
      Begin VB.ComboBox cmbAdmission 
         Height          =   315
         Left            =   3840
         TabIndex        =   7
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emergency Slip Number:"
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
         Top             =   360
         Width           =   3375
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
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label3 
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
         Left            =   480
         TabIndex        =   23
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Father/ Husband Name:"
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
         TabIndex        =   22
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   21
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emergency Doctor Name:"
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
         TabIndex        =   20
         Top             =   2760
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Admission:"
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
         TabIndex        =   19
         Top             =   3240
         Width           =   3375
      End
      Begin VB.Label Label8 
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
         Left            =   480
         TabIndex        =   18
         Top             =   3720
         Width           =   3375
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emergecy Diagnosis:"
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
         TabIndex        =   17
         Top             =   4200
         Width           =   3375
      End
   End
   Begin VB.Image Image2 
      Height          =   2070
      Left            =   9960
      Picture         =   "frmEmergency.frx":198C
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   2070
      Left            =   0
      Picture         =   "frmEmergency.frx":F16E
      Top             =   0
      Width           =   1995
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emergency Patient Information"
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
      Left            =   2760
      TabIndex        =   26
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frmEmergency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSlipno_Click()
Text1.Text = cmbSlipno.Text
cmbSlipno.Visible = False
Text1.Visible = True
datHospital.RecordSource = " SELECT * from EMERGENCY where EMG_SLIPNO=val('" + Text1 + "')"
datHospital.Refresh
If datHospital.Recordset.RecordCount > 0 Then
      Text2 = datHospital.Recordset.Fields("PAT_NO")
      Text3 = datHospital.Recordset.Fields("PAT_NAME")
      Text4 = datHospital.Recordset.Fields("PAT_FNAME")
      Text5 = datHospital.Recordset.Fields("PAT_AGE")
      Text6 = datHospital.Recordset.Fields("EMG_DNAME")
      cmbOperation.Text = datHospital.Recordset.Fields("PAT_OPER")
      cmbAdmission.Text = datHospital.Recordset.Fields("PAT_ADMINYN")
      Text7 = datHospital.Recordset.Fields("EMG_DES")
 Else
     MsgBox "Record not found, Please Try Again."
End If

End Sub

Private Sub cmdAdd_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text8.Text = "Add"
datHospital.RecordSource = "Select Max(EMG_SLIPNO)from EMERGENCY"
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
cmbSlipno.Clear
datHospital.RecordSource = "SELECT EMG_SLIPNO FROM EMERGENCY"
datHospital.Refresh

Do Until datHospital.Recordset.EOF
      cmbSlipno.AddItem datHospital.Recordset.Fields("EMG_SLIPNO")
      datHospital.Recordset.MoveNext
Loop
cmbSlipno.Visible = True
cmbSlipno.SetFocus
End Sub

Private Sub cmdModify_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text8.Text = "Modify"
datHospital.RecordSource = "SELECT EMG_SLIPNO FROM EMERGENCY"
datHospital.Refresh
cmbSlipno.Clear
Do Until datHospital.Recordset.EOF
      cmbSlipno.AddItem datHospital.Recordset.Fields("EMG_SLIPNO")
      datHospital.Recordset.MoveNext
Loop
cmbSlipno.Visible = True
cmbSlipno.SetFocus
End Sub

Private Sub cmdSave_Click()
datHospital.RecordSource = "EMERGENCY"
datHospital.Refresh
If Text8.Text = "Add" Then
      datHospital.Recordset.AddNew
      datHospital.Recordset.Fields("EMG_SLIPNO") = Val(Text1)
      datHospital.Recordset.Fields("PAT_NO") = Text2
      datHospital.Recordset.Fields("PAT_NAME") = Text3
      datHospital.Recordset.Fields("PAT_FNAME") = Text4
      datHospital.Recordset.Fields("PAT_AGE") = Val(Text5)
      datHospital.Recordset.Fields("EMG_DNAME") = Text6
      datHospital.Recordset.Fields("PAT_OPER") = cmbOperation.Text
      datHospital.Recordset.Fields("PAT_ADMINYN") = cmbAdmission.Text
      datHospital.Recordset.Fields("EMG_DES") = Text7
      datHospital.Recordset.Update
      datHospital.Refresh
End If
If Text8.Text = "Modify" Then
      datHospital.Database.Execute "update EMERGENCY SET EMG_SLIPNO=VAL('" + Text1 + "'),PAT_NO = '" + Text2 + "',  PAT_NAME ='" + Text3 + "',  PAT_FNAME ='" + Text4 + "',PAT_AGE = Val('" + Text5 + "'), EMG_DNAME ='" + Text6 + "', PAT_OPER = '" + cmbOperation.Text + "', PAT_ADMINYN ='" + cmbAdmission.Text + "',EMG_DES = '" + Text7 + "' WHERE EMG_SLIPNO=VAL('" + Text1 + "')"
      datHospital.Refresh
End If

If Text8.Text = "Delete" Then
       datHospital.Database.Execute "Delete form EMERGENCY where EMG_SLIPNO=val('" + Text1 + "')"
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
cmbSlipno.Text = ""
cmbAdmission.Text = ""
cmbOperation.Text = ""

cmbAdmission.AddItem "Yes"
cmbAdmission.AddItem "No"

cmbOperation.AddItem "Yes"
cmbOperation.AddItem "No"

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbSlipno.Visible = False




End Sub
