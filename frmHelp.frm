VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Project Copy Rights"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
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
      Left            =   9600
      Picture         =   "frmHelp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8080&
      Caption         =   "Mr. Bahar Hussain"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      Caption         =   "System Engineer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   2055
      Left            =   9240
      Picture         =   "frmHelp.frx":0442
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "naveed_zubary@hotmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "hussain_bahar@hotmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Contact Us on Mail:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   " Phone No (092-0471-614016 && 625623)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmHelp.frx":4DC88
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   " Kacha Koat Road Jhang Sadar."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Bahoo System Software Jhang"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "All Rights Reserved By: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Bahoo System Software Jhang"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 frmMainScreen.Show
 Unload Me
 

End Sub

