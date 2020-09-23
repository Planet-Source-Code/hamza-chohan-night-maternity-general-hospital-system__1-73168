VERSION 5.00
Begin VB.Form frmDischarge 
   BackColor       =   &H00E0E0E0&
   Caption         =   " Discharge Patient Information"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data dtHos 
      BackColor       =   &H00E0E0E0&
      Connect         =   "Ms Access;pwd=nmhbahoo"
      DatabaseName    =   "Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DISCHARGE"
      Top             =   8040
      Visible         =   0   'False
      Width           =   8655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   32
      Top             =   6240
      Width           =   11415
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search By Patient No "
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
         Picture         =   "frmDischarge.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Click on This Button for Search the Discharge Patient"
         Top             =   360
         Width           =   3495
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
         Picture         =   "frmDischarge.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Left            =   6240
         Picture         =   "frmDischarge.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Left            =   5040
         Picture         =   "frmDischarge.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   36
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
         Left            =   3840
         Picture         =   "frmDischarge.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Left            =   2640
         Picture         =   "frmDischarge.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Left            =   240
         Picture         =   "frmDischarge.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Add the New Record"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   11415
      Begin VB.TextBox Text14 
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
         Height          =   615
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2760
         Width           =   5415
      End
      Begin VB.TextBox Text13 
         Height          =   1335
         Left            =   7680
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox Text12 
         Height          =   1335
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   3720
         Width           =   3855
      End
      Begin VB.TextBox Text11 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   9120
         TabIndex        =   26
         Top             =   2280
         Width           =   2055
      End
      Begin VB.ComboBox cmbPatyn 
         Height          =   315
         Left            =   8640
         TabIndex        =   25
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox cmbPatno 
         Height          =   315
         Left            =   2520
         TabIndex        =   24
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   8640
         TabIndex        =   23
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   8640
         TabIndex        =   22
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   8640
         TabIndex        =   21
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Attendent Doctor Name: "
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
         TabIndex        =   14
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Follow Up:"
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
         Left            =   8760
         TabIndex        =   13
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Treatment:"
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
         Left            =   4800
         TabIndex        =   12
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Disch. On Request Yes/ No:"
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
         TabIndex        =   11
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Condition at Discharge:"
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
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Diagnosis:"
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
         Left            =   960
         TabIndex        =   9
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Date of Discharge:"
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
         TabIndex        =   8
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Date of Admission:"
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
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Weight:"
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
         Top             =   2760
         Width           =   2415
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
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient W/o. D/o Name:"
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
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient Address:"
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
         Top             =   1800
         Width           =   2415
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
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2415
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Patient Discharge Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   615
      Left            =   2760
      TabIndex        =   31
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmDischarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPatno_Click()
Text1 = cmbPatno.Text
cmbPatno.Visible = False
Text1.Visible = True
dtHos.RecordSource = "select * from DISCHARGE where PAT_NO=Val('" + Text1 + "')"
dtHos.Refresh
If dtHos.Recordset.RecordCount > 0 Then
      Text2 = dtHos.Recordset.Fields("PAT_NAME")
      Text3 = dtHos.Recordset.Fields("PAT_FNAME")
      Text4 = dtHos.Recordset.Fields("PAT_ADDRESS")
      Text5 = dtHos.Recordset.Fields("PAT_AGE")
      Text6 = dtHos.Recordset.Fields("PAT_WEIGHT")
      Text7 = dtHos.Recordset.Fields("DO_ADMIN")
      Text8 = dtHos.Recordset.Fields("DO_DISCH")
      Text9 = dtHos.Recordset.Fields("PAT_CADIS")
      Text10 = dtHos.Recordset.Fields("PATT_DOCTNAME")
      Text11 = dtHos.Recordset.Fields("PAT_DIAG")
      Text12 = dtHos.Recordset.Fields("PAT_TMEN")
      Text13 = dtHos.Recordset.Fields("PAT_FUP")
      cmbPatyn.Text = dtHos.Recordset.Fields("PAT_DOREQ")
   Else
      MsgBox "Record Not Found, Plz Try Again", vbOKCancel
End If

End Sub

Private Sub cmdAdd_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text14.Text = "Add"
dtHos.RecordSource = "select MAX(PAT_NO) from DISCHARGE"
dtHos.Refresh
Text1 = dtHos.Recordset.Fields(0) + 1

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
Text14.Text = "Delete"
cmbPatno.Clear
dtHos.RecordSource = "select PAT_NO from DISCHARGE"
dtHos.Refresh
Do Until dtHos.Recordset.EOF
          cmbPatno.AddItem dtHos.Recordset.Fields("PAT_NO")
          dtHos.Recordset.MoveNext
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
Text14.Text = "Modify"
dtHos.RecordSource = "select PAT_NO from DISCHARGE"
dtHos.Refresh
cmbPatno.Clear
Do Until dtHos.Recordset.EOF
          cmbPatno.AddItem dtHos.Recordset.Fields("PAT_NO")
          dtHos.Recordset.MoveNext
Loop
cmbPatno.Visible = True
cmbPatno.SetFocus

End Sub

Private Sub cmdSave_Click()
dtHos.RecordSource = "DISCHARGE"
dtHos.Refresh
If Text14.Text = "Add" Then
      dtHos.Recordset.AddNew
      dtHos.Recordset.Fields("PAT_NO") = Val(Text1)
      dtHos.Recordset.Fields("PAT_NAME") = Text2
      dtHos.Recordset.Fields("PAT_FNAME") = Text3
      dtHos.Recordset.Fields("PAT_ADDRESS") = Text4
      dtHos.Recordset.Fields("PAT_AGE") = Val(Text5)
      dtHos.Recordset.Fields("PAT_WEIGHT") = Text6
      dtHos.Recordset.Fields("DO_ADMIN") = Text7
      dtHos.Recordset.Fields("DO_DISCH") = Text8
      dtHos.Recordset.Fields("PAT_CADIS") = Text9
      dtHos.Recordset.Fields("PATT_DOCTNAME") = Text10
      dtHos.Recordset.Fields("PAT_DIAG") = Text11
      dtHos.Recordset.Fields("PAT_TMEN") = Text12
      dtHos.Recordset.Fields("PAT_FUP") = Text13
      dtHos.Recordset.Fields("PAT_DOREQ") = cmbPatyn.Text
      dtHos.Recordset.Update
      dtHos.Refresh
End If
 
If Text14.Text = "Modify" Then
      dtHos.Database.Execute "update DISCHARGE set PAT_NO=Val('" + Text1 + "'),PAT_NAME ='" + Text2 + "',PAT_FNAME ='" + Text3 + "',PAT_ADDRESS ='" + Text4 + "',PAT_AGE = Val('" + Text5 + "'),PAT_WEIGHT ='" + Text6 + "',DO_ADMIN ='" + Text7 + "',DO_DISCH ='" + Text8 + "',PAT_CADIS ='" + Text9 + "',PATT_DOCTNAME ='" + Text10 + "',PAT_DIAG ='" + Text11 + "',PAT_TMEN ='" + Text12 + "',PAT_FUP ='" + Text13 + "',PAT_DOREQ ='" + cmbPatyn.Text + "' where PAT_NO=Val('" + Text1 + "')"
      dtHos.Refresh
End If

If Text14.Text = "Delete" Then
      dtHos.Database.Execute "delete from DISCHARGE where PAT_NO=Val('" + Text1 + "')"
      dtHos.Refresh
End If
Call Form_Load

End Sub

Private Sub cmdSearch_Click()
dtHos.RecordSource = "select * from DISCHARGE where PAT_NO=Val('" + Text1 + "')"
dtHos.Refresh
If dtHos.Recordset.RecordCount > 0 Then
      Text2 = dtHos.Recordset.Fields("PAT_NAME")
      Text3 = dtHos.Recordset.Fields("PAT_FNAME")
      Text4 = dtHos.Recordset.Fields("PAT_ADDRESS")
      Text5 = dtHos.Recordset.Fields("PAT_AGE")
      Text6 = dtHos.Recordset.Fields("PAT_WEIGHT")
      Text7 = dtHos.Recordset.Fields("DO_ADMIN")
      Text8 = dtHos.Recordset.Fields("DO_DISCH")
      Text9 = dtHos.Recordset.Fields("PAT_CADIS")
      Text10 = dtHos.Recordset.Fields("PATT_DOCTNAME")
      Text11 = dtHos.Recordset.Fields("PAT_DIAG")
      Text12 = dtHos.Recordset.Fields("PAT_TMEN")
      Text13 = dtHos.Recordset.Fields("PAT_FUP")
      cmbPatyn.Text = dtHos.Recordset.Fields("PAT_DOREQ")
   Else
      MsgBox "Record Not Found, Plz Try Again", vbOKCancel
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
Text13.Text = ""
Text14.Text = ""
cmbPatno.Text = ""
cmbPatyn.Text = ""

cmbPatyn.AddItem "Yes"
cmbPatyn.AddItem "No"


cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbPatno.Visible = False

End Sub
