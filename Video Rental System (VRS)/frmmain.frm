VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8b.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdifrmmain 
   BackColor       =   &H8000000F&
   Caption         =   "One Touch World Cinema (VRS)"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10020
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "frmmain.frx":1E72
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   360
      Top             =   600
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      _Version        =   393216
      MouseIcon       =   "frmmain.frx":2CEDF4
      Begin VB.CommandButton cmdreturn 
         Enabled         =   0   'False
         Height          =   495
         Left            =   4335
         Picture         =   "frmmain.frx":2D0C76
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Return"
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton cmdrent 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3555
         Picture         =   "frmmain.frx":2D10C3
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Rent"
         Top             =   30
         Width           =   800
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   405
         Left            =   14565
         TabIndex        =   10
         ToolTipText     =   "Connected"
         Top             =   105
         Width           =   570
         _cx             =   1005
         _cy             =   714
         FlashVars       =   ""
         Movie           =   "D:\Video Rental System (VRS)\Images\LOADING.SWF"
         Src             =   "D:\Video Rental System (VRS)\Images\LOADING.SWF"
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
         Profile         =   0   'False
         ProfileAddress  =   ""
         ProfilePort     =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   14520
         TabIndex        =   11
         Top             =   60
         Width           =   660
      End
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   255
      End
      Begin VB.CommandButton cmdadd 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2760
         Picture         =   "frmmain.frx":2D1575
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "List of Items"
         Top             =   30
         Width           =   800
      End
      Begin VB.Frame Frame1 
         Height          =   580
         Left            =   5160
         TabIndex        =   6
         Top             =   -45
         Width           =   9345
         Begin VB.Label lblstatus 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   360
            TabIndex        =   8
            Top             =   240
            Width           =   45
         End
         Begin VB.Label control_num 
            Height          =   375
            Left            =   3480
            TabIndex        =   14
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   29
         Width           =   75
      End
      Begin VB.CommandButton cmdadmin 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   495
         Left            =   1755
         Picture         =   "frmmain.frx":2D1ABB
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Administration"
         Top             =   30
         Width           =   1000
      End
      Begin VB.CommandButton cmdsearch 
         Enabled         =   0   'False
         Height          =   495
         Left            =   960
         Picture         =   "frmmain.frx":2D1DC5
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Search Video"
         Top             =   30
         Width           =   800
      End
      Begin VB.CommandButton cmdmember 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   500
         Left            =   180
         Picture         =   "frmmain.frx":2D3527
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Membership"
         Top             =   30
         Width           =   800
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8010
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   13
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2408
            MinWidth        =   1235
            Picture         =   "frmmain.frx":2D4369
            Text            =   "User Name"
            TextSave        =   "User Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1746
            MinWidth        =   1587
            Picture         =   "frmmain.frx":2D4683
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            Picture         =   "frmmain.frx":2D57CE
            Text            =   "Date"
            TextSave        =   "Date"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Time Log-in"
            TextSave        =   "Time Log-in"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Programmed by:  Edwin Batomalaque"
            TextSave        =   "Programmed by:  Edwin Batomalaque"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   18
            MinWidth        =   18
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu member 
         Caption         =   "&Membership"
         Begin VB.Menu config 
            Caption         =   "C&onfiguration"
            Shortcut        =   {F2}
         End
      End
      Begin VB.Menu item 
         Caption         =   "It&ems"
         Begin VB.Menu additems 
            Caption         =   "&Configuration"
            Shortcut        =   {F9}
         End
         Begin VB.Menu searchmovie 
            Caption         =   "&Search Movie"
         End
      End
      Begin VB.Menu user 
         Caption         =   "L&og Off"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu rent 
         Caption         =   "&Rent"
         Shortcut        =   {F4}
      End
      Begin VB.Menu return 
         Caption         =   "Re&turn"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "Re&ports"
      Begin VB.Menu consol 
         Caption         =   "Co&nsoladated Report"
         Begin VB.Menu sales 
            Caption         =   "Sales Reports"
         End
      End
      Begin VB.Menu customer 
         Caption         =   "Cu&stomer"
         Begin VB.Menu allcustomer 
            Caption         =   "R&egistered Member"
         End
      End
      Begin VB.Menu items 
         Caption         =   "Ite&ms"
         Begin VB.Menu allitems 
            Caption         =   "All Items"
         End
         Begin VB.Menu rentedvideos 
            Caption         =   "All Rented Videos"
         End
         Begin VB.Menu unreturn 
            Caption         =   "Unreturn Item"
         End
      End
      Begin VB.Menu user_access 
         Caption         =   "User Access  Report"
      End
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "&Administration"
      Begin VB.Menu acc 
         Caption         =   "Creat&e Account"
         Shortcut        =   {F6}
      End
      Begin VB.Menu penalty 
         Caption         =   "Penalty Rate"
      End
      Begin VB.Menu rentlimits 
         Caption         =   "Rental Limits"
      End
      Begin VB.Menu mem 
         Caption         =   "Membership Fee"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu system 
         Caption         =   "About the System"
      End
      Begin VB.Menu author 
         Caption         =   "About the Author"
      End
   End
End
Attribute VB_Name = "mdifrmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub acc_Click()
frmadduser.Show
End Sub

Private Sub additems_Click()
frmitemlist.Show
End Sub

Private Sub allcustomer_Click()
Cust_report.Refresh
Cust_report.Show

End Sub

Private Sub allitems_Click()

Item_report.Refresh
Item_report.Show

End Sub

Private Sub author_Click()
frmauthor.Show
End Sub

Private Sub cmdAdd_Click()
If mdifrmmain.StatusBar1.Panels(4).Text = "Administration" Then
frmitemlist.Show
    Exit Sub
  Else
    MsgBox "Access Denied!", vbCritical, "Restricted Area"
  End If

End Sub

Private Sub cmdadd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblstatus.Caption = "List of Items"

End Sub

Private Sub cmdadmin_Click()

If mdifrmmain.StatusBar1.Panels(4).Text = "Administration" Then
    frmadduser.Show
    Exit Sub
  Else
    MsgBox "Access Denied!", vbCritical, "Restricted Area"
  End If

End Sub

Private Sub cmdadmin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblstatus.Caption = cmdadmin.ToolTipText
End Sub

Private Sub cmdmember_Click()
frmmembership.Show
End Sub



Private Sub cmdmember_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblstatus.Caption = cmdmember.ToolTipText

End Sub


Private Sub cmdrent_Click()
frmrent.Show

End Sub

Private Sub cmdrent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblstatus.Caption = "Rent"
End Sub

Private Sub cmdreturn_Click()
frmreturn.Show

End Sub

Private Sub cmdreturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblstatus.Caption = "Return"
End Sub

Private Sub cmdsearch_Click()
frmsearchvideo.Show
End Sub

Private Sub cmdsearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblstatus.Caption = cmdsearch.ToolTipText
End Sub

Private Sub Command3_Click()

End Sub

Private Sub config_Click()



frmmembership.Show



End Sub

Private Sub del_Click()
frmdelete.Show
End Sub

Private Sub File_Click()

End Sub

Private Sub Exit_Click()
        Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
        Set rs = db.OpenRecordset("Select *from userlogin where control_num ='" + mdifrmmain.control_num + "'")
        If rs.RecordCount = 0 Then
        End
        Else
        rs.Edit
        rs!control_num = mdifrmmain.control_num.Caption
        rs!datelogout = Date
        rs!timelogout = Time
        rs.Update
        rs.Close
        db.Close
        End If
        
End
End Sub

Private Sub help_Click()
lblstatus.Caption = "Help"
End Sub

Private Sub item_Click()
lblstatus.Caption = "Open the items Form"
If mdifrmmain.StatusBar1.Panels(4).Text = "Administration" Then
    Exit Sub
  Else
    MsgBox "Access Denied!", vbCritical, "Restricted Area"
  End If
  
End Sub


Private Sub MDIForm_Load()

member.Enabled = False
mnutransaction.Enabled = False
mnureport.Enabled = False
mnuadmin.Enabled = False
item.Enabled = False
help.Enabled = False
user.Enabled = False
frmlogin.Show
frmlogin.txtuser.SetFocus

End Sub

Private Sub prof_Click()
DataReport1.Show
End Sub

Private Sub stud_Click()
frmsearchvideo.Show
End Sub

Private Sub uprecord_Click()
Set db = OpenDatabase("D:\Edwin Visual\stud_info.mdb")
Set rs = db.OpenRecordset("Select *from tblStudInfo where idno ='" & txtidno & "'")
If rs.RecordCount = 0 Then
Else

rs.Edit
rs!idno = txtidno
rs!lname = txtLname
rs!fname = txtFname
rs!mname = txtMname
rs!address = txtAdd
rs!bdate = DTBdate
rs!age = lblage
rs.Update
Call clear
End If

End Sub
Sub clear()
txtidno.Text = ""
txtLname.Text = ""
txtFname.Text = ""
txtMname.Text = ""
txtAdd.Text = ""
lblage.Caption = ""
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblstatus.Caption = ""
End Sub

Private Sub mem_Click()
frmpenalty.Show
frmpenalty.Caption = "Membership Fee"
End Sub

Private Sub member_Click()
lblstatus.Caption = "Membership Form"
End Sub

Private Sub mnuadmin_Click()
    lblstatus.Caption = "Administration"
 If mdifrmmain.StatusBar1.Panels(4).Text = "Administration" Then
    Exit Sub
  Else
    MsgBox "Access Denied!", vbCritical, "Restricted Area"
  End If
End Sub

Private Sub mnuFile_Click()
lblstatus.Caption = "File"
End Sub

Private Sub mnureport_Click()
lblstatus.Caption = "Reports"
End Sub

Private Sub mnutransaction_Click()
lblstatus.Caption = "Transactions"
End Sub

Private Sub penalty_Click()
frmpenalty.Show
frmpenalty.Caption = "Penalty Rate"
End Sub

Private Sub rent_Click()
mdifrmmain.Enabled = False

lblstatus.Caption = "Rent Movie CD's"
frmrent.Show
End Sub

Private Sub rentedvideos_Click()

all_rented.Refresh
all_rented.Show


End Sub

Private Sub rentlimits_Click()
frmpenalty.Show
frmpenalty.Caption = "Rental Limits"

End Sub

Private Sub return_Click()
lblstatus.Caption = "Return Movie CD's"

frmreturn.Show

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case SSTab1.Caption
Case Is = "Search Movies"
frmsearchvideo.Show
Case Is = "Membership"
frmmembership.Show

End Select

End Sub

Private Sub sales_Click()
sales_report.Refresh
sales_report.Show

End Sub

Private Sub searchmovie_Click()

If mdifrmmain.StatusBar1.Panels(4).Text = "Administration" Then
frmitemlist.Show
    Exit Sub
  Else
    MsgBox "Access Denied!", vbCritical, "Restricted Area"
  End If
  
End Sub

Private Sub system_Click()
frmAbout.Show
End Sub

Private Sub Timer1_Timer()
frmlogin.Show
frmlogin.txtuser.SetFocus
Timer1.Enabled = False
member.Enabled = False
mnutransaction.Enabled = False
mnureport.Enabled = False
mnuadmin.Enabled = False
cmdmember.Enabled = False
cmdsearch.Enabled = False
cmdadmin.Enabled = False
cmdrent.Enabled = False
cmdreturn.Enabled = False
cmdadd.Enabled = False
item.Enabled = False
help.Enabled = False
user.Enabled = False


End Sub



Private Sub unreturn_Click()
unreturnmovies.Refresh
unreturnmovies.Show
End Sub

Private Sub user_access_Click()
user_login.Show
End Sub

Private Sub user_Click()
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
        Set rs = db.OpenRecordset("Select *from userlogin where control_num ='" + mdifrmmain.control_num + "'")
        rs.Edit
        rs!control_num = mdifrmmain.control_num.Caption
        rs!datelogout = Date
        rs!timelogout = Time
        rs.Update
        rs.Close
        db.Close
        
lblstatus.Caption = "Log Off" + " " + StatusBar1.Panels(2)
Prompt$ = "Do you really want to log-off?"
    reply = MsgBox(Prompt$, vbOKCancel, "Log-off" + " " + StatusBar1.Panels(2))
    If reply = vbOK Then
   Me.Timer1.Enabled = True
    mdifrmmain.StatusBar1.Panels(2) = "Waiting .."
    mdifrmmain.StatusBar1.Panels(4) = "Waiting .."
    
    

End If

Unload frmadduser
Unload frmmembership
Unload frmreturn
Unload frmrent
Unload frmsearchmember
Unload frmsearchvideo
Unload frmmembercash

End Sub
