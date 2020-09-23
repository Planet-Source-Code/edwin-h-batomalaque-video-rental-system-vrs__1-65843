VERSION 5.00
Begin VB.Form frmauthor 
   Caption         =   "About the Author"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   Icon            =   "frmauthor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Edwin H. Batomalaque                                   "
      BeginProperty Font 
         Name            =   "0 Bills Holiday DNA"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Its ME"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   990
   End
   Begin VB.Label Label4 
      Caption         =   "Email Address:   edzeielief4@yahoo.com   Tel. # 301-6551  or Cel. # 0906"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "For comments and suggestion please send at: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   $"frmauthor.frx":27A2
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   -360
      X2              =   6360
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmauthor.frx":284B
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008080FF&
      BorderWidth     =   10
      FillColor       =   &H00FF0000&
      Height          =   2415
      Left            =   3510
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2670
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   3600
      Picture         =   "frmauthor.frx":28E1
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmauthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
mdifrmmain.Enabled = True
End Sub

Private Sub Form_Load()
mdifrmmain.Enabled = False
End Sub

Private Sub Form_Terminate()
mdifrmmain.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
End Sub
