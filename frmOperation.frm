VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOperation 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Nighat Maternity Home And General Hospital       ( Operation Theatre Information)"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data dtHos 
      Caption         =   "Data1"
      Connect         =   "Ms Access;pwd=nmhbahoo"
      DatabaseName    =   "Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "OPERATIONTH"
      Top             =   6480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid msfHospital 
      Height          =   1455
      Left            =   120
      TabIndex        =   26
      Top             =   5520
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   120
      TabIndex        =   25
      Top             =   6960
      Width           =   11655
      Begin VB.CommandButton cmdSearchOPNo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search Operation No"
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
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Click on This Button to Search the Operation and It's All Patient Information"
         Top             =   360
         Width           =   3975
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
         Picture         =   "frmOperation.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmOperation.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "frmOperation.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "frmOperation.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "frmOperation.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Picture         =   "frmOperation.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Press the Button and Modify/Change the Records"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox cmbDA 
         Height          =   315
         Left            =   9360
         TabIndex        =   14
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox cmbSex 
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Top             =   2760
         Width           =   2535
      End
      Begin VB.ComboBox cmbNumber 
         Height          =   315
         Left            =   3360
         TabIndex        =   45
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text16 
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         Height          =   975
         Left            =   6000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   4080
         Width           =   5535
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   9360
         TabIndex        =   16
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   9360
         TabIndex        =   15
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   9360
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   9600
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   9600
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   4440
         Width           =   4215
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Patient Death/Alive:"
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
         Left            =   6000
         TabIndex        =   43
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Name:"
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
         TabIndex        =   42
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Diagnosis."
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
         Left            =   7320
         TabIndex        =   41
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Patient History."
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
         Left            =   840
         TabIndex        =   40
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Charges:"
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
         Left            =   6000
         TabIndex        =   39
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Patient Sex:"
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
         TabIndex        =   38
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Patient Age:"
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
         TabIndex        =   37
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Reference Doctor Name:"
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
         Left            =   6000
         TabIndex        =   36
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Date:"
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
         Left            =   6000
         TabIndex        =   35
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Boy Name:"
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
         TabIndex        =   34
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Dai/ Nurse Name:"
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
         TabIndex        =   33
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Sister Name:"
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
         Left            =   6000
         TabIndex        =   32
         Top             =   2760
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Antehisia Doctor Name:"
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
         Left            =   6000
         TabIndex        =   31
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Attendent Doctor Name: "
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
         Left            =   6000
         TabIndex        =   30
         Top             =   1320
         Width           =   3735
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
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Patient Name:"
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
         TabIndex        =   28
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operation Number:"
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
         TabIndex        =   27
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbNumber_Click()
Text1.Text = cmbNumber.Text
cmbNumber.Visible = False
Text1.Visible = True
dtHos.RecordSource = "select * from OPERATIONTH where OP_NO=Val('" + Text1 + "')"
dtHos.Refresh
If dtHos.Recordset.RecordCount > 0 Then
      Text2 = dtHos.Recordset.Fields("OP_NAME")
      Text3 = dtHos.Recordset.Fields("OP_PAT_NAME")
      Text4 = dtHos.Recordset.Fields("OP_PAT_FNAME")
      Text5 = dtHos.Recordset.Fields("OP_PAT_AGE")
      Text6 = dtHos.Recordset.Fields("OP_HNURSNAME")
      Text7 = dtHos.Recordset.Fields("OP_BOY")
      Text8 = dtHos.Recordset.Fields("OP_PAT_HIST")
      Text9 = dtHos.Recordset.Fields("OP_PAT_DATE")
      Text10 = dtHos.Recordset.Fields("OP_PAT_REF")
      Text11 = dtHos.Recordset.Fields("OP_ATT_DOCT")
      Text12 = dtHos.Recordset.Fields("OP_ATH")
      Text13 = dtHos.Recordset.Fields("OP_SISNAME")
      Text14 = dtHos.Recordset.Fields("OP_CHARGE")
      Text15 = dtHos.Recordset.Fields("OP_DES")
      cmbSex.Text = dtHos.Recordset.Fields("OP_PAT_SEX")
      cmbDA.Text = dtHos.Recordset.Fields("OP_PAT_DL")
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
Text16.Text = "Add"
dtHos.RecordSource = "select MAX(OP_NO) from OPERATIONTH"
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
Text16.Text = "Delete"
cmbNumber.Clear
dtHos.RecordSource = "select OP_NO from OPERATIONTH"
dtHos.Refresh
Do Until dtHos.Recordset.EOF
         cmbNumber.AddItem dtHos.Recordset.Fields("OP_NO")
         dtHos.Recordset.MoveNext
Loop
cmbNumber.Visible = True
cmbNumber.SetFocus

End Sub

Private Sub cmdModify_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text16.Text = "Modify"
dtHos.RecordSource = "select OP_NO from OPERATIONTH"
dtHos.Refresh
cmbNumber.Clear
Do Until dtHos.Recordset.EOF
         cmbNumber.AddItem dtHos.Recordset.Fields("OP_NO")
         dtHos.Recordset.MoveNext
Loop
cmbNumber.Visible = True
cmbNumber.SetFocus

End Sub

Private Sub cmdSave_Click()
dtHos.RecordSource = "OPERATIONTH"
dtHos.Refresh
If Text16.Text = "Add" Then
       dtHos.Recordset.AddNew
       dtHos.Recordset.Fields("OP_NO") = Val(Text1)
       dtHos.Recordset.Fields("OP_NAME") = Text2
       dtHos.Recordset.Fields("OP_PAT_NAME") = Text3
       dtHos.Recordset.Fields("OP_PAT_FNAME") = Text4
       dtHos.Recordset.Fields("OP_PAT_AGE") = Val(Text5)
       dtHos.Recordset.Fields("OP_HNURSNAME") = Text6
       dtHos.Recordset.Fields("OP_BOY") = Text7
       dtHos.Recordset.Fields("OP_PAT_HIST") = Text8
       dtHos.Recordset.Fields("OP_PAT_DATE") = Text9
       dtHos.Recordset.Fields("OP_PAT_REF") = Text10
       dtHos.Recordset.Fields("OP_ATT_DOCT") = Text11
       dtHos.Recordset.Fields("OP_ATH") = Text12
       dtHos.Recordset.Fields("OP_SISNAME") = Text13
       dtHos.Recordset.Fields("OP_CHARGE") = Val(Text14)
       dtHos.Recordset.Fields("OP_DES") = Text15
       dtHos.Recordset.Fields("OP_PAT_SEX") = cmbSex.Text
       dtHos.Recordset.Fields("OP_PAT_DL") = cmbDA.Text
       dtHos.Recordset.Update
       dtHos.Refresh
End If

If Text16.Text = "Modify" Then
      dtHos.Database.Execute "update OPERATIONTH set OP_NO=Val('" + Text1 + "'), OP_NAME ='" + Text2 + "', OP_PAT_NAME ='" + Text3 + "', OP_PAT_FNAME ='" + Text4 + "', OP_PAT_AGE = Val('" + Text5 + "'), OP_HNURSNAME ='" + Text6 + "', OP_BOY ='" + Text7 + "', OP_PAT_HIST ='" + Text8 + "',OP_PAT_DATE ='" + Text9 + "', OP_PAT_REF ='" + Text10 + "', OP_ATT_DOCT ='" + Text11 + "', OP_ATH ='" + Text12 + "', OP_SISNAME ='" + Text13 + "', OP_CHARGE = Val('" + Text14 + "'), OP_DES ='" + Text15 + "',OP_PAT_SEX ='" + cmbSex.Text + "', OP_PAT_DL ='" + cmbDA.Text + "' where OP_NO=Val('" + Text1 + "')"
      dtHos.Refresh
End If

If Text16.Text = "Delete" Then
     dtHos.Database.Execute "delete from OPERATIONTH where OP_NO =Val('" + Text1 + "')"
     dtHos.Refresh
End If
Call Form_Load

       
End Sub

Private Sub cmdSearchOPNo_Click()
Text16.Text = "Search"
dtHos.RecordSource = "select * from OPERATIONTH where OP_NO=Val('" + Text1 + "')"
dtHos.Refresh
If dtHos.Recordset.RecordCount > 0 Then
      Text1 = dtHos.Recordset.Fields("OP_NO")
      Text2 = dtHos.Recordset.Fields("OP_NAME")
      Text3 = dtHos.Recordset.Fields("OP_PAT_NAME")
      Text4 = dtHos.Recordset.Fields("OP_PAT_FNAME")
      Text5 = dtHos.Recordset.Fields("OP_PAT_AGE")
      Text6 = dtHos.Recordset.Fields("OP_HNURSNAME")
      Text7 = dtHos.Recordset.Fields("OP_BOY")
      Text8 = dtHos.Recordset.Fields("OP_PAT_HIST")
      Text9 = dtHos.Recordset.Fields("OP_PAT_DATE")
      Text10 = dtHos.Recordset.Fields("OP_PAT_REF")
      Text11 = dtHos.Recordset.Fields("OP_ATT_DOCT")
      Text12 = dtHos.Recordset.Fields("OP_ATH")
      Text13 = dtHos.Recordset.Fields("OP_SISNAME")
      Text14 = dtHos.Recordset.Fields("OP_CHARGE")
      Text15 = dtHos.Recordset.Fields("OP_DES")
      cmbSex.Text = dtHos.Recordset.Fields("OP_PAT_SEX")
      cmbDA.Text = dtHos.Recordset.Fields("OP_PAT_DL")
  Else
       MsgBox "Record not Found. Plz try Again", vbOKCancel
End If


End Sub

Private Sub cmdSearchPatNo_Click()
dtHos.RecordSource = "select * from OPERATIONTH where OP_PAT_DATE='" + Text9 + "'"
dtHos.Refresh
If dtHos.Recordset.RecordCount > 0 Then
      Text1 = dtHos.Recordset.Fields("OP_NO")
      Text2 = dtHos.Recordset.Fields("OP_NAME")
      Text3 = dtHos.Recordset.Fields("OP_PAT_NAME")
      Text4 = dtHos.Recordset.Fields("OP_PAT_FNAME")
      Text5 = dtHos.Recordset.Fields("OP_PAT_AGE")
      Text6 = dtHos.Recordset.Fields("OP_HNURSNAME")
      Text7 = dtHos.Recordset.Fields("OP_BOY")
      Text8 = dtHos.Recordset.Fields("OP_PAT_HIST")
      Text10 = dtHos.Recordset.Fields("OP_PAT_REF")
      Text11 = dtHos.Recordset.Fields("OP_ATT_DOCT")
      Text12 = dtHos.Recordset.Fields("OP_ATH")
      Text13 = dtHos.Recordset.Fields("OP_SISNAME")
      Text14 = dtHos.Recordset.Fields("OP_CHARGE")
      Text15 = dtHos.Recordset.Fields("OP_DES")
      cmbSex.Text = dtHos.Recordset.Fields("OP_PAT_SEX")
      cmbDA.Text = dtHos.Recordset.Fields("OP_PAT_DL")
  Else
       MsgBox "Record not Found. Plz try Again", vbOKCancel
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
Text15.Text = ""
cmbSex.Text = ""
cmbNumber.Text = ""
cmbDA.Text = ""

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbNumber.Visible = False

cmbSex.AddItem "Female"
cmbSex.AddItem "Male"

cmbDA.AddItem "Alive"
cmbDA.AddItem "Death"

Text9.Text = Date

'==========================
'grid Coding.
'==========================

msfHospital.Cols = 1
msfHospital.Rows = 2
msfHospital.Cols = 18
msfHospital.FixedCols = 1
msfHospital.FixedRows = 1
msfHospital.ColWidth(1) = 1000
msfHospital.ColWidth(2) = 2500
msfHospital.ColWidth(3) = 2500
msfHospital.ColWidth(4) = 3000
msfHospital.ColWidth(5) = 800
msfHospital.ColWidth(6) = 2500
msfHospital.ColWidth(7) = 2500
msfHospital.ColWidth(8) = 4500
msfHospital.ColWidth(9) = 2500
msfHospital.ColWidth(10) = 2000
msfHospital.ColWidth(11) = 2500
msfHospital.ColWidth(12) = 1700
msfHospital.ColWidth(13) = 1700
msfHospital.ColWidth(14) = 1500
msfHospital.ColWidth(15) = 1500
msfHospital.ColWidth(16) = 2500
msfHospital.ColWidth(17) = 2500
msfHospital.Row = 0
msfHospital.Col = 1
msfHospital.Text = "Number"
msfHospital.Col = 2
msfHospital.Text = "Operation Name"
msfHospital.Col = 3
msfHospital.Text = "Patient Name"
msfHospital.Col = 4
msfHospital.Text = "Father/Husband Name"
msfHospital.Col = 5
msfHospital.Text = "Age"
msfHospital.Col = 6
msfHospital.Text = "Nurse Name"
msfHospital.Col = 7
msfHospital.Text = "Helper Boy"
msfHospital.Col = 8
msfHospital.Text = "Patient History"
msfHospital.Col = 9
msfHospital.Text = "Operation Date"
msfHospital.Col = 10
msfHospital.Text = "Reference"
msfHospital.Col = 11
msfHospital.Text = "Attendent Doctor"
msfHospital.Col = 12
msfHospital.Text = "Anthsis"
msfHospital.Col = 13
msfHospital.Text = "Sister Name"
msfHospital.Col = 14
msfHospital.Text = " Charges"
msfHospital.Col = 15
msfHospital.Text = "Patient Sex"
msfHospital.Col = 16
msfHospital.Text = "Medicine"
msfHospital.Col = 17
msfHospital.Text = "Death/Alive"

dtHos.RecordSource = "OPERATIONTH"
dtHos.Refresh
dtHos.Recordset.MoveNext
msfHospital.Rows = dtHos.Recordset.RecordCount + 1
Do Until dtHos.Recordset.EOF
       msfHospital.Col = 1
       msfHospital.Text = dtHos.Recordset.Fields("OP_NO")
       msfHospital.Col = 2
       msfHospital.Text = dtHos.Recordset.Fields("OP_NAME")
       msfHospital.Col = 3
       msfHospital.Text = dtHos.Recordset.Fields("OP_PAT_NAME")
       msfHospital.Col = 4
       msfHospital.Text = dtHos.Recordset.Fields("OP_PAT_FNAME")
       msfHospital.Col = 5
       msfHospital.Text = dtHos.Recordset.Fields("OP_PAT_AGE")
       msfHospital.Col = 6
       msfHospital.Text = dtHos.Recordset.Fields("OP_HNURSNAME")
       msfHospital.Col = 7
       msfHospital.Text = dtHos.Recordset.Fields("OP_BOY")
       msfHospital.Col = 8
       msfHospital.Text = dtHos.Recordset.Fields("OP_PAT_HIST")
       msfHospital.Col = 9
       msfHospital.Text = dtHos.Recordset.Fields("OP_PAT_DATE")
       msfHospital.Col = 10
       msfHospital.Text = dtHos.Recordset.Fields("OP_PAT_REF")
       msfHospital.Col = 11
       msfHospital.Text = dtHos.Recordset.Fields("OP_ATT_DOCT")
       msfHospital.Col = 12
       msfHospital.Text = dtHos.Recordset.Fields("OP_ATH")
       msfHospital.Col = 13
       msfHospital.Text = dtHos.Recordset.Fields("OP_SISNAME")
       msfHospital.Col = 14
       msfHospital.Text = dtHos.Recordset.Fields("OP_CHARGE")
       msfHospital.Col = 15
       msfHospital.Text = dtHos.Recordset.Fields("OP_PAT_SEX")
       msfHospital.Col = 16
       msfHospital.Text = dtHos.Recordset.Fields("OP_DES")
       msfHospital.Col = 17
       msfHospital.Text = dtHos.Recordset.Fields("OP_PAT_DL")
       dtHos.Recordset.MoveNext
       msfHospital.Row = msfHospital.Row + 1
Loop



End Sub
