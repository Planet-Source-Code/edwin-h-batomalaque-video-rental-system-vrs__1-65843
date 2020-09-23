VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmaddmovie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VRS - Movies"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "frmaddmovie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3135
      Picture         =   "frmaddmovie.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6315
      Width           =   975
   End
   Begin VB.CommandButton cmdlcancel 
      Caption         =   "Canc&el"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2145
      Picture         =   "frmaddmovie.frx":22B4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6300
      Width           =   975
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   165
      Picture         =   "frmaddmovie.frx":2B7E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6300
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1140
      Picture         =   "frmaddmovie.frx":3448
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6300
      Width           =   975
   End
   Begin VB.TextBox txtitemid 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Zurich Cn BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   4530
      Left            =   105
      TabIndex        =   17
      Top             =   1455
      Width           =   6165
      Begin VB.ComboBox cbostatus 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmaddmovie.frx":3A44
         Left            =   3315
         List            =   "frmaddmovie.frx":3A4E
         TabIndex        =   12
         Top             =   4140
         Width           =   975
      End
      Begin VB.ComboBox cbocategory 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmaddmovie.frx":3A5B
         Left            =   1695
         List            =   "frmaddmovie.frx":3A7A
         TabIndex        =   4
         Top             =   1350
         Width           =   1815
      End
      Begin VB.TextBox txtdirector 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1710
         TabIndex        =   8
         Top             =   3060
         Width           =   3780
      End
      Begin MSComCtl2.DTPicker datepurchase 
         Height          =   375
         Left            =   1740
         TabIndex        =   9
         Top             =   3660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   19595265
         CurrentDate     =   38844
      End
      Begin VB.TextBox txttitle 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1695
         MaxLength       =   50
         TabIndex        =   2
         Top             =   495
         Width           =   3825
      End
      Begin VB.TextBox txtmcast 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2160
         Width           =   3810
      End
      Begin VB.TextBox txtseccast 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2580
         Width           =   3825
      End
      Begin VB.TextBox txtprice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4860
         MaxLength       =   8
         TabIndex        =   10
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtnoofdays 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1695
         MaxLength       =   1
         TabIndex        =   5
         Top             =   1725
         Width           =   495
      End
      Begin VB.TextBox txtnoofcd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1740
         MaxLength       =   1
         TabIndex        =   11
         Top             =   4080
         Width           =   495
      End
      Begin VB.ComboBox cmdformat 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmaddmovie.frx":3ADA
         Left            =   1695
         List            =   "frmaddmovie.frx":3AE7
         TabIndex        =   3
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Director"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   180
         TabIndex        =   29
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   2460
         TabIndex        =   28
         Top             =   4200
         Width           =   525
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "&Format"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   210
         TabIndex        =   27
         Top             =   1110
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "&Movie Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   255
         TabIndex        =   26
         Top             =   585
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Catego&ry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   195
         TabIndex        =   25
         Top             =   1470
         Width           =   750
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "&Main Cast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   180
         TabIndex        =   24
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary  Cast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   165
         TabIndex        =   23
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Item &Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   3660
         TabIndex        =   22
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of &Days"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   180
         TabIndex        =   21
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purchased"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   180
         TabIndex        =   20
         Top             =   3720
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "N&o. of Disc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   180
         TabIndex        =   19
         Top             =   4200
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   570
      Left            =   -60
      TabIndex        =   0
      Top             =   -15
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1005
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   917
      TabCaption(0)   =   "  Add Movies"
      TabPicture(0)   =   "frmaddmovie.frx":3AFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Shape Shape3 
      Height          =   630
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   645
      Width           =   6135
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Height          =   4530
      Left            =   150
      TabIndex        =   30
      Top             =   1560
      Width           =   6180
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   6075
      Shape           =   3  'Circle
      Top             =   6345
      Width           =   300
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   6375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item ID #"
      BeginProperty Font 
         Name            =   "Helvetica LT Std Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   405
      TabIndex        =   18
      Top             =   855
      Width           =   1035
   End
End
Attribute VB_Name = "frmaddmovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbocategory_KeyPress(KeyAscii As Integer)
Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If
End Sub

Private Sub cbostatus_KeyPress(KeyAscii As Integer)
Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If
End Sub

Private Sub cmdclose_Click()
Unload Me
frmitemlist.Enabled = True
End Sub



Private Sub cmdformat_KeyPress(KeyAscii As Integer)
Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If
End Sub

Private Sub cmdlcancel_Click()
clear
disable


cmdsave.Enabled = False
cmdlcancel.Enabled = False

cmdnew.Enabled = True
cmdClose.Enabled = True

End Sub

Private Sub cmdnew_Click()

Set db = OpenDatabase("D:\Video Rental System (VRS)\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movies")
If rs.RecordCount = 0 Then
txtitemid.Text = "000001"
Else
rs.MoveLast
txtitemid.Text = Format(rs!itemid + 1, "000000")
End If
enable
txttitle.SetFocus

cmdsave.Enabled = True
cmdlcancel.Enabled = True
cmdnew.Enabled = False
cmdClose.Enabled = False


End Sub

Private Sub cmdsave_Click()
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movies where itemid  ='" & txtitemid & "'")
If txttitle = "" Or cbocategory = "" Or cmdformat = "" Or cbostatus = "" Or txtprice = "" Or txtnoofcd = "" Or txtnoofdays = "" Then
MsgBox ("Pls. Complete the Field"), vbInformation, "Fields Missing Data"
Else
If rs.RecordCount = 0 Then
rs.AddNew
rs!itemid = txtitemid.Text
rs!movtitle = txttitle.Text
rs!movcategory = cbocategory
rs!movmainchar = txtmcast.Text
rs!movsecchar = txtseccast.Text
rs!movformt = cmdformat.Text
rs!movstatus = cbostatus.Text
rs!movdirector = txtdirector.Text
rs!movdpurchase = datepurchase
rs!movamount = txtprice.Text
rs!movcopies = txtnoofcd.Text
rs!movnodays = txtnoofdays.Text

rs.Update
rs.Close
db.Close
frmitemlist.adomovies.Refresh
frmitemlist.dgridmovies.Refresh

Item_report.Hide
Item_report.Refresh

End If

clear

disable

cmdsave.Enabled = False
cmdlcancel.Enabled = False

cmdnew.Enabled = True
cmdClose.Enabled = True

End If

End Sub


Sub clear()

txtitemid.Text = ""
txttitle.Text = ""
cbocategory.Text = ""
txtmcast.Text = ""
txtseccast.Text = ""
cmdformat.Text = ""
cbostatus.Text = ""
txtdirector.Text = ""
txtprice.Text = ""
txtnoofcd.Text = ""
txtnoofdays.Text = ""

End Sub

Sub enable()


txttitle.Enabled = True
cbocategory.Enabled = True
txtmcast.Enabled = True
txtseccast.Enabled = True
cmdformat.Enabled = True
cbostatus.Enabled = True
txtdirector.Enabled = True
datepurchase.Enabled = True
txtprice.Enabled = True
txtnoofcd.Enabled = True
txtnoofdays.Enabled = True


End Sub

Sub disable()
txttitle.Enabled = False
cbocategory.Enabled = False
txtmcast.Enabled = False
txtseccast.Enabled = False
cmdformat.Enabled = False
cbostatus.Enabled = False
txtdirector.Enabled = False
datepurchase.Enabled = False
txtprice.Enabled = False
txtnoofcd.Enabled = False
txtnoofdays.Enabled = False

End Sub


Private Sub Form_Load()
frmitemlist.Enabled = False

End Sub

Private Sub Form_Terminate()
frmitemlist.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmitemlist.Enabled = True
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 13) Then
KeyAscii = 0
End If
End Sub

Private Sub txttitle_LostFocus()
txttitle.Text = StrConv(txttitle, vbProperCase)
End Sub
