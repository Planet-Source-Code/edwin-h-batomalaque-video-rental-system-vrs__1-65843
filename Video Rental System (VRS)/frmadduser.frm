VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmadduser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "VRS- Create Account"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6165
   ControlBox      =   0   'False
   Icon            =   "frmadduser.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmadduser.frx":1E72
   ScaleHeight     =   6165
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5025
      TabIndex        =   18
      Top             =   3870
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5025
      TabIndex        =   7
      Top             =   4425
      Width           =   975
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      Height          =   495
      Left            =   5025
      TabIndex        =   6
      Top             =   3330
      Width           =   975
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   5025
      TabIndex        =   8
      Top             =   4995
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dbgriduser 
      Bindings        =   "frmadduser.frx":15C97
      Height          =   2220
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   3916
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "empid"
         Caption         =   "User ID #"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "username"
         Caption         =   "User Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "password"
         Caption         =   "Password"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "level"
         Caption         =   "Level / Status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1800
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adouser 
      Height          =   375
      Left            =   3720
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Video Rental System (VRS)\VRS Database\vrs.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Video Rental System (VRS)\VRS Database\vrs.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "security"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   5550
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      Caption         =   "Verify your password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   960
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2760
      Begin VB.TextBox txtpas2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   135
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Pls. Fillup the field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1860
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5760
      Begin VB.ComboBox cbolevel 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmadduser.frx":15CAD
         Left            =   3120
         List            =   "frmadduser.frx":15CB7
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2400
      End
      Begin VB.TextBox txtempid 
         Enabled         =   0   'False
         Height          =   345
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtpas 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   2385
      End
      Begin VB.TextBox txtus 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status/Level"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   3120
         TabIndex        =   12
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
   End
   Begin TabDlg.SSTab stab 
      Height          =   465
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   820
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   758
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User"
      TabPicture(0)   =   "frmadduser.frx":15CD5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
End
Attribute VB_Name = "frmadduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcx_Click()

End Sub




Private Sub cmdo_KeyPress(KeyAscii As Integer)

End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub cbolevel_KeyPress(KeyAscii As Integer)
Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If


End Sub

Private Sub cmdAdd_Click()

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from security")
If rs.RecordCount = 0 Then
txtempid.Text = "0001"
Else
rs.MoveLast
txtempid.Text = Format(rs!empid + 1, "0000")
cmdadd.Enabled = False
End If

cmdsave.Enabled = True
enable_entries
cmdcancel.Enabled = True
cmdexit.Enabled = False
cmdadd.Enabled = False


End Sub

Private Sub cmdcancel_Click()
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdadd.Enabled = True
disable_entries
cmdexit.Enabled = True

clear
End Sub

Private Sub cmdDelete_Click()
Set db = OpenDatabase("D:\Video Rental System (VRS)\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from security where val(empid) = '" & Val(dbgriduser.Columns(0)) & "'")
 Prompt$ = "Do you really want to delete this user?"
    reply = MsgBox(Prompt$, vbOKCancel, "Delete Record")
    If reply = vbOK Then
        rs.Delete
        rs.MoveNext
        rs.Close
        db.Close
        dbgriduser.Refresh
        adouser.Refresh
        
        
    End If

End Sub

Private Sub cmdexit_Click()
Unload Me
mdifrmmain.Enabled = True
End Sub



Private Sub Command1_Click()

End Sub

Private Sub cmdsave_Click()

If txtempid = "" Or cbolevel = "" Or txtus = "" Or txtpas = "" Then
MsgBox ("Missing Data, Complete the Fields"), vbInformation, "Invalid"
Else

Set db = OpenDatabase("D:\Video Rental System (VRS)\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from security where Password = '" & txtpas & "'")
If txtpas2 = txtpas = 0 Then
MsgBox "Password Not Match, try again", vbCritical, "Password"
txtpas.SetFocus
Else
If rs.RecordCount = 0 Then
rs.AddNew
rs!UserName = txtus
rs!Password = txtpas
rs!Level = cbolevel
rs!empid = txtempid
rs.Update
rs.Close
db.Close
dbgriduser.Refresh
adouser.Refresh

cmdsave.Enabled = False
cmdadd.Enabled = True
clear
disable_entries
MsgBox "You are now registered user", vbInformation, "Regestered User"
Else
MsgBox "Password are existing, Do You want to replace current password ", vbYesNo, "Existing Password"
If vbYes = True Then
rs.Edit
rs!UserName = txtus
rs!Password = txtpas
rs!Level = cbolevel
rs!empid = txtempid
rs.Update
rs.Close
db.Close
dbgriduser.Refresh
adouser.Refresh

cmdsave.Enabled = False
cmdadd.Enabled = True
clear
disable_entries
Else
txtus.Text = ""
txtpas.Text = ""
txtpas2.Text = ""
txtempid.Text = ""
cbolevel.Text = ""
txtus.SetFocus

End If
End If
End If

End If

End Sub

Private Sub cmdsave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set db = OpenDatabase("D:\Edwin Visual\user.mdb")
Set rs = db.OpenRecordset("Select *from access where Password = '" & txtpas & "'")
If rs.RecordCount = 0 Then
rs.AddNew
rs!UserName = txtus
rs!Password = txtpas
rs!Level = cbolevel
rs!empid = txtempid
rs.Update
rs.Close
db.Close
dbgriduser.Refresh
adouser.Refresh


cmdsave.Enabled = False
cmdadd.Enabled = True
clear
disable_entries
MsgBox "You are now registered user", vbInformation, "Regesitered User"
Else

MsgBox "Password are existing, Do You want to replace current password ", vbYesNo, "Existing Password"
If vbYes = True Then
rs.Edit
rs!UserName = txtus
rs!Password = txtpas
rs!Level = cbolevel
rs!empid = txtempid
rs.Update
rs.Close
db.Close
dbgriduser.Refresh
adouser.Refresh


cmdsave.Enabled = False
cmdadd.Enabled = True
clear
disable_entries
Else
txtus.Text = ""
txtpas.Text = ""
txtpas2.Text = ""
txtempid.Text = ""
cbolevel.Text = ""
txtus.SetFocus
End If
End If
End If
End Sub


Private Sub Form_Load()
mdifrmmain.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdifrmmain.lblstatus.Caption = "Security Congfiguration......... Add User Account"

End Sub

Private Sub Form_Terminate()
mdifrmmain.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
End Sub

Private Sub txtpas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtpas2.SetFocus
End If
End Sub

Private Sub txtpas2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtpas2 = txtpas Then
cmdo.SetFocus
Else
MsgBox "Password Not Match, try again", vbCritical, "Password"
End If
End If
End Sub

Private Sub txtus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtpas.SetFocus
End If
End Sub


Sub clear()
txtus.Text = ""
txtpas.Text = ""
txtpas2.Text = ""
txtempid.Text = ""
cbolevel.Text = ""
End Sub


Sub enable_entries()
txtempid.Enabled = True
txtus.Enabled = True
txtpas.Enabled = True
txtpas2.Enabled = True
cbolevel.Enabled = True
End Sub

Sub disable_entries()

txtempid.Enabled = False
txtus.Enabled = False
txtpas.Enabled = False
txtpas2.Enabled = False
cbolevel.Enabled = False

End Sub
