VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fmrborrower 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2700
      Left            =   0
      Picture         =   "frmborrower.frx":0000
      ScaleHeight     =   2640
      ScaleWidth      =   11970
      TabIndex        =   0
      Top             =   0
      Width           =   12030
      Begin VB.TextBox txtidno 
         Height          =   435
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   2340
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "OK"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2070
         TabIndex        =   2
         Top             =   1275
         Width           =   855
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   465
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   820
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   706
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "   Borrowers ID"
         TabPicture(0)   =   "frmborrower.frx":A3FC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
      End
   End
End
Attribute VB_Name = "fmrborrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
