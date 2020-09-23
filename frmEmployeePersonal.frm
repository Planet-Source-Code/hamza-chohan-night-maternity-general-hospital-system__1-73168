VERSION 5.00
Begin VB.Form frmEmployeePersonal 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Nighat Maternity Home And General Hospital        ( Employee Personal Information )"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   6840
      Width           =   11415
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
         Picture         =   "frmEmployeePersonal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   4320
         Picture         =   "frmEmployeePersonal.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   5520
         Picture         =   "frmEmployeePersonal.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   6720
         Picture         =   "frmEmployeePersonal.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   7920
         Picture         =   "frmEmployeePersonal.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   3120
         Picture         =   "frmEmployeePersonal.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Press the Button and Modify/Change the Records"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Data dtHos 
      Caption         =   "Data1"
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
      RecordSource    =   "EMPLOYEEPERSONAL"
      Top             =   6720
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   11415
      Begin VB.ComboBox cmbSexEmp 
         Height          =   315
         Left            =   2760
         TabIndex        =   39
         Top             =   2760
         Width           =   2895
      End
      Begin VB.ComboBox cmbEmployeeNo 
         Height          =   315
         Left            =   2760
         TabIndex        =   38
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
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
         TabIndex        =   37
         Top             =   5280
         Width           =   5415
      End
      Begin VB.TextBox Text14 
         Height          =   975
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   4080
         Width           =   5415
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   9000
         TabIndex        =   35
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   9000
         TabIndex        =   34
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   9000
         TabIndex        =   33
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   9000
         TabIndex        =   32
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   9000
         TabIndex        =   31
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   9000
         TabIndex        =   30
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   9000
         TabIndex        =   29
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   2055
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   3720
         Width           =   5295
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2760
         TabIndex        =   25
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2760
         TabIndex        =   24
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Country:"
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
         TabIndex        =   21
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee City:"
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
         TabIndex        =   20
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Specialization:"
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
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Qualification:"
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
         TabIndex        =   18
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Phone Number:"
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
         TabIndex        =   17
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Fax Number:"
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
         TabIndex        =   16
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label8 
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
         Left            =   5760
         TabIndex        =   15
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Permanent Address:"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Contact Address:"
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
         Left            =   6960
         TabIndex        =   13
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Date of Birth:"
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
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Age:"
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
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Father Name:"
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
         TabIndex        =   10
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Identity No:"
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
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Employee Personal Information"
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
      Left            =   1440
      TabIndex        =   40
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmEmployeePersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbEmployeeNo_Click()
Text1.Text = cmbEmployeeNo.Text
cmbEmployeeNo.Visible = False
Text1.Visible = True
dtHos.RecordSource = "select * from EMPLOYEEPERSONAL where EMP_ID=Val('" + Text1 + "')"
dtHos.Refresh
If dtHos.Recordset.RecordCount > 0 Then
     Text2 = dtHos.Recordset.Fields("EMP_NAME")
     Text3 = dtHos.Recordset.Fields("EMP_FNAME")
     Text4 = dtHos.Recordset.Fields("EMP_DOB")
     Text5 = dtHos.Recordset.Fields("EMP_AGE")
     Text6 = dtHos.Recordset.Fields("EMP_PADD")
     Text7 = dtHos.Recordset.Fields("EMP_NICNO")
     Text8 = dtHos.Recordset.Fields("EMP_FAXNO")
     Text9 = dtHos.Recordset.Fields("EMP_PHONENO")
     Text10 = dtHos.Recordset.Fields("EMP_QULI")
     Text11 = dtHos.Recordset.Fields("EMP_SEP")
     Text12 = dtHos.Recordset.Fields("EMP_CITY")
     Text13 = dtHos.Recordset.Fields("EMP_COUNTRY")
     Text14 = dtHos.Recordset.Fields("EMP_CADD")
     cmbSexEmp.Text = dtHos.Recordset.Fields("EMP_SEX")
   Else
      MsgBox "Record Not Found. Plz Try Again", vbOKCancel
End If

     
End Sub

Private Sub cmdAdd_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text15.Text = "Add"
dtHos.RecordSource = "select MAX(EMP_ID) from EMPLOYEEPERSONAL"
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
Text15.Text = "Delete"
cmbEmployeeNo.Clear
dtHos.RecordSource = "select EMP_ID from EMPLOYEEPERSONAL"
dtHos.Refresh
Do Until dtHos.Recordset.EOF
      cmbEmployeeNo.AddItem dtHos.Recordset.Fields("EMP_ID")
      dtHos.Recordset.MoveNext
Loop
cmbEmployeeNo.Visible = True
cmbEmployeeNo.SetFocus

End Sub

Private Sub cmdModify_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text15.Text = "Modify"
dtHos.RecordSource = "select EMP_ID from EMPLOYEEPERSONAL"
dtHos.Refresh
cmbEmployeeNo.Clear
Do Until dtHos.Recordset.EOF
      cmbEmployeeNo.AddItem dtHos.Recordset.Fields("EMP_ID")
      dtHos.Recordset.MoveNext
Loop
cmbEmployeeNo.Visible = True
cmbEmployeeNo.SetFocus

End Sub

Private Sub cmdSave_Click()
dtHos.RecordSource = "EMPLOYEEPERSONAL"
dtHos.Refresh
If Text15.Text = "Add" Then
       dtHos.Recordset.AddNew
       dtHos.Recordset.Fields("EMP_ID") = Val(Text1)
       dtHos.Recordset.Fields("EMP_NAME") = Text2
       dtHos.Recordset.Fields("EMP_FNAME") = Text3
       dtHos.Recordset.Fields("EMP_DOB") = Text4
       dtHos.Recordset.Fields("EMP_AGE") = Val(Text5)
       dtHos.Recordset.Fields("EMP_PADD") = Text6
       dtHos.Recordset.Fields("EMP_NICNO") = Text7
       dtHos.Recordset.Fields("EMP_FAXNO") = Text8
       dtHos.Recordset.Fields("EMP_PHONENO") = Text9
       dtHos.Recordset.Fields("EMP_QULI") = Text10
       dtHos.Recordset.Fields("EMP_SEP") = Text11
       dtHos.Recordset.Fields("EMP_CITY") = Text12
       dtHos.Recordset.Fields("EMP_COUNTRY") = Text13
       dtHos.Recordset.Fields("EMP_CADD") = Text14
       dtHos.Recordset.Fields("EMP_SEX") = cmbSexEmp.Text
       dtHos.Recordset.Update
       dtHos.Refresh
End If
If Text15.Text = "Modify" Then
      dtHos.Database.Execute "update EMPLOYEEPERSONAL set EMP_ID=Val('" + Text1 + "'),EMP_NAME ='" + Text2 + "', EMP_FNAME ='" + Text3 + "',EMP_DOB ='" + Text4 + "',EMP_AGE = Val('" + Text5 + "'), EMP_PADD ='" + Text6 + "',EMP_NICNO ='" + Text7 + "',EMP_FAXNO ='" + Text8 + "',EMP_PHONENO ='" + Text9 + "', EMP_QULI ='" + Text10 + "', EMP_SEP ='" + Text11 + "',EMP_CITY ='" + Text12 + "',EMP_COUNTRY ='" + Text13 + "',EMP_CADD ='" + Text14 + "',EMP_SEX ='" + cmbSexEmp.Text + "' where EMP_ID=Val('" + Text1 + "')"
      dtHos.Refresh
End If

If Text15.Text = "Delete" Then
      dtHos.Database.Execute "delete from EMPLOYEEPERSONAL where EMP_ID=Val('" + Text1 + "')"
      dtHos.Refresh
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
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
cmbSexEmp.Text = ""
cmbEmployeeNo.Text = ""

cmbSexEmp.AddItem "Male"
cmbSexEmp.AddItem "Female"

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbEmployeeNo.Visible = False

End Sub
