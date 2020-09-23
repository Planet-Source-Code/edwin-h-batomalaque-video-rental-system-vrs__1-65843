VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmloading 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmloading.frx":0000
   ScaleHeight     =   5130
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   7080
      Top             =   4440
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   3000
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      MouseIcon       =   "frmloading.frx":7F2F2
      Scrolling       =   1
   End
   Begin VB.Label lblinit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   600
      TabIndex        =   5
      Top             =   4560
      Width           =   45
   End
   Begin VB.Label lblcomplete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   5760
      TabIndex        =   3
      Top             =   4560
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   255
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3450
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   3720
      Width           =   1395
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright  C  2006, Edwin Batomalaque, Philippines"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   3480
      Width           =   4260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Video Rental System V06.0521"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   3240
      Width           =   2205
   End
End
Attribute VB_Name = "frmloading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
lblcomplete.Caption = "0 % Complete"
End Sub



Private Sub Timer2_Timer()

ProgressBar1.Value = ProgressBar1.Value + 1
lblcomplete.Caption = ProgressBar1.Value & "% Completed"
 Select Case ProgressBar1.Value
 Case Is = 2
 lblinit.Caption = "Initializing."
 Case Is = 4
 lblinit.Caption = "Initializing.."
 Case Is = 8
 lblinit.Caption = "Initializing..."
 Case Is = 12
 lblinit.Caption = "Initializing...."
 Case Is = 14
 lblinit.Caption = "Initializing....."
 Case Is = 18
 lblinit.Caption = "Initializing......"
  Case Is = 20
 lblinit.Caption = "Loading all Forms"
 Case Is = 30
 lblinit.Caption = "Generating Main Menu"
 Case Is = 50
 lblinit.Caption = "Analizing Data Memory"
 Case Is = 60
 lblinit.Caption = "Preparing Movie List"
 Case Is = 90
 lblinit.Caption = "Finalizing the System"
 Case Is = 100
 lblinit.Caption = "Please Wait.."
 Unload Me
 frmlogin.Show
End Select

End Sub


