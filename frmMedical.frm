VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMedical 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Medical Store"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data dtHospital 
      Connect         =   "Ms Access;pwd=nmhbahoo"
      DatabaseName    =   "Hospital.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MEDICALSTORE"
      Top             =   8160
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   9120
      Picture         =   "frmMedical.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Press the Button and Modify/Change the Records"
      Top             =   6480
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
      Left            =   10320
      Picture         =   "frmMedical.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Back to Main Screen"
      Top             =   7440
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
      Left            =   9120
      Picture         =   "frmMedical.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Press the Button to Cancel the All Operations"
      Top             =   7440
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
      Left            =   7920
      Picture         =   "frmMedical.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Press the Button to Save the Records."
      Top             =   7440
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
      Left            =   10320
      Picture         =   "frmMedical.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Press the Button and Delete the Records."
      Top             =   6480
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
      Left            =   7920
      Picture         =   "frmMedical.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Add the New Record"
      Top             =   6480
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid msfHospital 
      Height          =   2055
      Left            =   120
      TabIndex        =   31
      Top             =   6360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bill Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   7215
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   3600
         TabIndex        =   27
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   3600
         TabIndex        =   26
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   3600
         TabIndex        =   25
         ToolTipText     =   "Please Click Here to Check The Credit Balance."
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3600
         TabIndex        =   24
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bill Date:"
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
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Debit:"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Credit:"
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
         TabIndex        =   15
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mdeicine Total Price:"
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
         TabIndex        =   14
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Medicine Price Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2535
      Left            =   7440
      TabIndex        =   9
      Top             =   3720
      Width           =   4335
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   3000
         TabIndex        =   29
         ToolTipText     =   "Press Enter to Check the Price of Per Tablet/Capsool/Eng"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   3000
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Per Tab/Inj/Cap in Rs:"
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
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pata (10 Tablet)in Rs:"
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
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pack Price in Rs:"
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
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   7440
      TabIndex        =   5
      Top             =   1440
      Width           =   4335
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         ToolTipText     =   "Please Click Here to Check the Current Stock of Medicine"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Others:"
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
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sales:"
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
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Qutinty:"
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
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Medicin Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   7215
      Begin VB.ComboBox cmbCname 
         Height          =   315
         Left            =   3720
         TabIndex        =   39
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Medicine Order Date:"
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
         TabIndex        =   4
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Medicine Company Name:"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Medicine Name:"
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
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nighat Maternity Medical Store"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmMedical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCname_Click()

Text2.Text = cmbCname.Text
cmbCname.Visible = False
Text2.Visible = True
dtHospital.RecordSource = "select * from MEDICALSTORE where M_CNAME='" + Text2 + "' "
dtHospital.Refresh
If dtHospital.Recordset.RecordCount > 0 Then
     Text1 = dtHospital.Recordset.Fields("M_NAME")
     Text2 = dtHospital.Recordset.Fields("M_CNAME")
     Text3 = dtHospital.Recordset.Fields("ORDER_DATE")
     Text4 = dtHospital.Recordset.Fields("M_QTY")
     Text5 = dtHospital.Recordset.Fields("M_SALES")
     Text6 = dtHospital.Recordset.Fields("M_OTHER")
     Text7 = dtHospital.Recordset.Fields("C_TOTAL")
     Text8 = dtHospital.Recordset.Fields("COMP_CD")
     Text9 = dtHospital.Recordset.Fields("COMP_DEB")
     Text10 = dtHospital.Recordset.Fields("BILL_DATE")
     Text11 = dtHospital.Recordset.Fields("PACK_PICE")
     Text12 = dtHospital.Recordset.Fields("PATA_PRICE")
     Text13 = dtHospital.Recordset.Fields("PER_TAB")
Else
     MsgBox "Record not found and plz try Again.", vbOKCancel
End If

End Sub

Private Sub cmdAdd_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text14.Text = "Add"
dtHospital.RecordSource = "select M_CNAME from MEDICALSTORE"
dtHospital.Refresh

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
cmbCname.Clear
dtHospital.RecordSource = "select M_CNAME from MEDICALSTORE"
dtHospital.Refresh
Do Until dtHospital.Recordset.EOF
     cmbCname.AddItem dtHospital.Recordset.Fields("M_CNAME")
     dtHospital.Recordset.MoveNext
Loop
cmbCname.Visible = True
cmbCname.SetFocus

End Sub

Private Sub cmdModify_Click()
cmdAdd.Enabled = False
cmdModify.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
Text14.Text = "Modify"
dtHospital.RecordSource = "select M_CNAME from MEDICALSTORE"
dtHospital.Refresh
cmbCname.Clear
Do Until dtHospital.Recordset.EOF
     cmbCname.AddItem dtHospital.Recordset.Fields("M_CNAME")
     dtHospital.Recordset.MoveNext
Loop
cmbCname.Visible = True
cmbCname.SetFocus


End Sub

Private Sub cmdSave_Click()
dtHospital.RecordSource = "MEDICALSTORE"
dtHospital.Refresh
If Text14.Text = "Add" Then
         dtHospital.Recordset.AddNew
         dtHospital.Recordset.Fields("M_NAME") = Text1
         dtHospital.Recordset.Fields("M_CNAME") = Text2
         dtHospital.Recordset.Fields("ORDER_DATE") = Text3
         dtHospital.Recordset.Fields("M_QTY") = Val(Text4)
         dtHospital.Recordset.Fields("M_SALES") = Val(Text5)
         dtHospital.Recordset.Fields("M_OTHER") = Val(Text6)
         dtHospital.Recordset.Fields("C_TOTAL") = Val(Text7)
         dtHospital.Recordset.Fields("COMP_CD") = Val(Text8)
         dtHospital.Recordset.Fields("COMP_DEB") = Val(Text9)
         dtHospital.Recordset.Fields("BILL_DATE") = Text10
         dtHospital.Recordset.Fields("PACK_PICE") = Val(Text11)
         dtHospital.Recordset.Fields("PATA_PRICE") = Val(Text12)
         dtHospital.Recordset.Fields("PER_TAB") = Val(Text13)
         dtHospital.Recordset.Update
         dtHospital.Refresh
End If
If Text14.Text = "Modify" Then
     dtHospital.Database.Execute "update MEDICALSTORE set M_CNAME='" + Text2 + "',M_NAME = '" + Text1 + "', ORDER_DATE ='" + Text3 + "',M_QTY = Val('" + Text4 + "'),M_SALES = Val('" + Text5 + "'),M_OTHER = Val('" + Text6 + "'),C_TOTAL = Val('" + Text7 + "'),COMP_CD = Val('" + Text8 + "'),COMP_DEB = Val('" + Text9 + "'),BILL_DATE ='" + Text10 + "',PACK_PICE = Val('" + Text11 + " '),PATA_PRICE = Val('" + Text12 + "'),PER_TAB = Val('" + Text13 + "') where M_CNAME = '" + Text2 + "'"
     dtHospital.Refresh
End If

If Text14.Text = "Delete" Then
     dtHospital.Database.Execute "delete from MEDICALSTORE where M_CNAME='" + Text2 + "'"
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
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmbCname.Visible = False

'==========================
'grid Coding.
'==========================

msfHospital.Cols = 1
msfHospital.Rows = 2
msfHospital.Cols = 14
msfHospital.FixedCols = 1
msfHospital.FixedRows = 1
msfHospital.ColWidth(1) = 2000
msfHospital.ColWidth(2) = 2500
msfHospital.ColWidth(3) = 1000
msfHospital.ColWidth(4) = 1000
msfHospital.ColWidth(5) = 1500
msfHospital.ColWidth(6) = 1500
msfHospital.ColWidth(7) = 1500
msfHospital.ColWidth(8) = 1500
msfHospital.ColWidth(9) = 1500
msfHospital.ColWidth(10) = 1500
msfHospital.ColWidth(11) = 1500
msfHospital.ColWidth(12) = 1500
msfHospital.ColWidth(13) = 1500
msfHospital.Row = 0
msfHospital.Col = 1
msfHospital.Text = "Medicine Name"
msfHospital.Col = 2
msfHospital.Text = "Company Name"
msfHospital.Col = 3
msfHospital.Text = "Order Date"
msfHospital.Col = 4
msfHospital.Text = "Quantity"
msfHospital.Col = 5
msfHospital.Text = "Sales"
msfHospital.Col = 6
msfHospital.Text = "Others"
msfHospital.Col = 7
msfHospital.Text = "Total Balance"
msfHospital.Col = 8
msfHospital.Text = "Credit"
msfHospital.Col = 9
msfHospital.Text = "debit"
msfHospital.Col = 10
msfHospital.Text = "Bill Date"
msfHospital.Col = 11
msfHospital.Text = "Pack"
msfHospital.Col = 12
msfHospital.Text = "Pata"
msfHospital.Col = 13
msfHospital.Text = "Per Tab/Cap/Inj"

dtHospital.RecordSource = "MEDICALSTORE"
dtHospital.Refresh
dtHospital.Recordset.MoveNext
msfHospital.Rows = dtHospital.Recordset.RecordCount + 1
Do Until dtHospital.Recordset.EOF
       msfHospital.Col = 1
       msfHospital.Text = dtHospital.Recordset.Fields("M_NAME")
       msfHospital.Col = 2
       msfHospital.Text = dtHospital.Recordset.Fields("M_CNAME")
       msfHospital.Col = 3
       msfHospital.Text = dtHospital.Recordset.Fields("ORDER_DATE")
       msfHospital.Col = 4
       msfHospital.Text = dtHospital.Recordset.Fields("M_QTY")
       msfHospital.Col = 5
       msfHospital.Text = dtHospital.Recordset.Fields("M_SALES")
       msfHospital.Col = 6
       msfHospital.Text = dtHospital.Recordset.Fields("M_OTHER")
       msfHospital.Col = 7
       msfHospital.Text = dtHospital.Recordset.Fields("C_TOTAL")
       msfHospital.Col = 8
       msfHospital.Text = dtHospital.Recordset.Fields("COMP_CD")
       msfHospital.Col = 9
       msfHospital.Text = dtHospital.Recordset.Fields("COMP_DEB")
       msfHospital.Col = 10
       msfHospital.Text = dtHospital.Recordset.Fields("BILL_DATE")
       msfHospital.Col = 11
       msfHospital.Text = dtHospital.Recordset.Fields("PACK_PICE")
       msfHospital.Col = 12
       msfHospital.Text = dtHospital.Recordset.Fields("PATA_PRICE")
       msfHospital.Col = 13
       msfHospital.Text = dtHospital.Recordset.Fields("PER_TAB")
       dtHospital.Recordset.MoveNext
       msfHospital.Row = msfHospital.Row + 1
Loop


End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text13.Text = Val(Text12) / 10
End If

End Sub

Private Sub Text6_Click()
Text6.Text = Val(Text4) - Val(Text5)

End Sub

Private Sub Text8_Click()
Text8.Text = Val(Text7) - Val(Text9)

End Sub
