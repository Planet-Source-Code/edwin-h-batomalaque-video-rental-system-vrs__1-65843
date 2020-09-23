VERSION 5.00
Begin VB.Form frmlogin 
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14925
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frmlogin.frx":0000
   ScaleHeight     =   9855
   ScaleWidth      =   14925
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7290
      TabIndex        =   1
      Top             =   5415
      Width           =   2055
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7275
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5850
      Width           =   2055
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   5010
      TabIndex        =   3
      Top             =   7545
      Width           =   735
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8970
      TabIndex        =   4
      Top             =   7545
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "USER ACCESS"
      BeginProperty Font 
         Name            =   "Bauhaus Md BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   6270
      TabIndex        =   0
      Top             =   3000
      Width           =   2115
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   5865
      Width           =   1245
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   5505
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   5010
      Picture         =   "frmlogin.frx":2CCF82
      Stretch         =   -1  'True
      Top             =   3585
      Width           =   675
   End
   Begin VB.Image bats 
      Appearance      =   0  'Flat
      Height          =   4575
      Left            =   4890
      Picture         =   "frmlogin.frx":2CD3C4
      Stretch         =   -1  'True
      Top             =   3465
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   50
      Height          =   5175
      Left            =   3810
      Shape           =   3  'Circle
      Top             =   3105
      Width           =   7215
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bats_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdifrmmain.lblstatus.Caption = "User Access ....... Only registered users can access the system"
End Sub

Private Sub cmdcancel_Click()
 LoginSucceeded = False
    Me.Hide
     txtuser.Text = ""
        txtpass.Text = ""
        mdifrmmain.StatusBar1.Panels(2) = ""
        mdifrmmain.StatusBar1.Panels(4) = ""
End Sub




Private Sub cmdOK_Click()
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from security where Username ='" + txtuser + "'")
If rs.RecordCount = 0 Then
MsgBox "User Name is Invalid!", , "Login"
        txtuser.Text = ""
        txtuser.SetFocus
       SendKeys "{Home}+{End}"
Exit Sub
Else
    Set db = OpenDatabase("D:\Video Rental System (VRS)\VRS Database\vrs.mdb")
        Set rs = db.OpenRecordset("Select *from security where Password ='" + txtpass + "'")
        If rs.RecordCount = 1 Then
        LoginSucceeded = True
        Me.Hide
        MsgBox "Congratulations!!!! Access Granted", vbInformation, "Login Success"
        mdifrmmain.member.Enabled = True
        mdifrmmain.mnutransaction.Enabled = True
        mdifrmmain.mnuadmin.Enabled = True
        mdifrmmain.mnureport.Enabled = True
        mdifrmmain.cmdadmin.Enabled = True
        mdifrmmain.cmdmember.Enabled = True
        mdifrmmain.cmdsearch.Enabled = True
        mdifrmmain.user.Enabled = True
        mdifrmmain.help.Enabled = True
        mdifrmmain.item.Enabled = True
        mdifrmmain.cmdadd.Enabled = True
        mdifrmmain.cmdrent.Enabled = True
        mdifrmmain.cmdreturn.Enabled = True
        txtuser.Text = ""
        txtpass.Text = ""
        
        
        mdifrmmain.StatusBar1.Panels(2).Text = rs!UserName
        mdifrmmain.StatusBar1.Panels(2).Text = StrConv(mdifrmmain.StatusBar1.Panels(2).Text, vbUpperCase)
        mdifrmmain.StatusBar1.Panels(4).Text = rs!Level
        mdifrmmain.StatusBar1.Panels(8).Text = Time
        mdifrmmain.StatusBar1.Panels(6).Text = Date
        mdifrmmain.user.Caption = "Log Off" + " " + mdifrmmain.StatusBar1.Panels(2)
        
        
        Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
        Set rs = db.OpenRecordset("Select *from userlogin")
        If rs.RecordCount = 0 Then
        mdifrmmain.control_num.Caption = "001"
        Else
        rs.MoveLast
        mdifrmmain.control_num.Caption = Format(rs!control_num + 1, "000")
        End If
        
        Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
        Set rs = db.OpenRecordset("Select *from userlogin where control_num ='" + mdifrmmain.control_num + "'")
        rs.AddNew
        rs!control_num = mdifrmmain.control_num.Caption
        rs!UserName = mdifrmmain.StatusBar1.Panels(2).Text
        rs!datelogin = Date
        rs!timelogin = mdifrmmain.StatusBar1.Panels(8).Text
        rs.Update
        rs.Close
        db.Close

        Exit Sub
    
    Else
        MsgBox "Password is Invalid!", , "Login"
        txtpass.Text = ""
        txtpass.SetFocus
       SendKeys "{Home}+{End}"
 
    End If
End If

End Sub

Private Sub Form_Load()
bats.Visible = True
frmlogin.Show
End Sub


Private Sub Timer1_Timer()

End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
SendKeys h1
Call cmdOK_Click

End If
End Sub

Private Sub txtuser_GotFocus()
  SendKeys hl
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
txtpass.SetFocus
  SendKeys hl
End If
End Sub
