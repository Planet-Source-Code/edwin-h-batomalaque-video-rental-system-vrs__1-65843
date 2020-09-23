VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsearchcust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrowers Rented Details"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Rented Movies"
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   6075
         TabIndex        =   8
         Top             =   2730
         Width           =   900
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Return"
         Height          =   495
         Left            =   5055
         TabIndex        =   7
         Top             =   2730
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid dbgridrental 
         Bindings        =   "frmsearchcust.frx":0000
         Height          =   2340
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Double Click the Item to return"
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4128
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   65535
         Enabled         =   0   'False
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "custnum"
            Caption         =   "custnum"
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
            DataField       =   "itemid"
            Caption         =   "itemid"
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
            DataField       =   "title"
            Caption         =   "title"
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
            DataField       =   "movformt"
            Caption         =   "movformt"
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
         BeginProperty Column04 
            DataField       =   "dateborrowed"
            Caption         =   "dateborrowed"
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
         BeginProperty Column05 
            DataField       =   "duedate"
            Caption         =   "duedate"
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
         BeginProperty Column06 
            DataField       =   "status"
            Caption         =   "status"
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
         BeginProperty Column07 
            DataField       =   "amount"
            Caption         =   "amount"
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
         BeginProperty Column08 
            DataField       =   "custname"
            Caption         =   "custname"
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
         BeginProperty Column09 
            DataField       =   "rented_time"
            Caption         =   "rented_time"
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
            ScrollBars      =   2
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtidno 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc adorent 
      Height          =   480
      Left            =   5160
      Top             =   3240
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.Label lblborrower 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Borrowers Name"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Borrowers ID Number"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   780
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   285
      Width           =   6855
   End
End
Attribute VB_Name = "frmsearchcust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from customers where custnum  ='" & txtidno & "'")
    If rs.RecordCount = 0 Then
    MsgBox ("Borrower's ID Not Exist"), vbInformation, "Not Found"
    txtidno.Text = ""
    txtidno.SetFocus

    Else
            lblborrower.Caption = rs!custlname + ", " + rs!custfname + " " + rs!custmname
    End If
 
    Set db = OpenDatabase("D:\Video Rental System (VRS)\VRS Database\vrs.mdb")
    Set rs = db.OpenRecordset("Select *from movierental")
        If rs.RecordCount = txtidno.Text Then
        Else

        End If

Call ado
If dbgridrental.ApproxCount = 0 Then
MsgBox ("No Current Movies Rented"), vbInformation, "Item Not Found"
dbgridrental.Enabled = False
Else
dbgridrental.Enabled = True
End If

End Sub

Private Sub Command1_Click()
Call data


End Sub

Private Sub Command2_Click()

Unload Me
frmrent.Enabled = True

End Sub


Private Sub dbgridrental_DblClick()

Call data
Unload Me
frmrent.Enabled = True

End Sub

Private Sub Form_Load()
frmrent.Enabled = False
End Sub

Private Sub Form_Terminate()
frmrent.Enabled = True
End Sub

Private Sub txtidno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtidno.Text = Format(txtidno.Text, "00000")
Call cmdOK_Click
End If

End Sub

Private Sub txtidno_LostFocus()

txtidno.Text = Format(txtidno.Text, "00000")
End Sub


Sub ado()

adorent.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Video Rental System (VRS)\VRS Database\vrs.mdb;Persist Security Info=False"
adorent.RecordSource = "Select *from movierental where custnum = '" + txtidno + "'"
adorent.Refresh
dbgridrental.Refresh

End Sub


Sub data()
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movierental where val(itemid) = '" & Val(dbgridrental.Columns(1)) & "'")

frmreturn.txtitemid.Text = rs!itemid
frmreturn.lblidno.Caption = rs!custnum
frmreturn.lblitemid.Caption = rs!itemid
frmreturn.lbltime.Caption = rs!rented_time
frmreturn.lblTitle.Caption = rs!Title
frmreturn.lblstatus.Caption = rs!status
frmreturn.lblformat.Caption = rs!movformt
frmreturn.lbldborrowed.Caption = rs!dateborrowed
frmreturn.lblduedate.Caption = rs!duedate
frmreturn.lbldatereturn.Caption = Date
frmreturn.lblamount.Caption = rs!Amount
frmreturn.lblborrower.Caption = lblborrower.Caption
frmreturn.lblstat = "Return"


frmreturn.cmdreturn.Enabled = True
frmreturn.cmdclear.Enabled = True

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from return")
If rs.RecordCount = 0 Then
frmreturn.txtnum.Text = "0001"
Else
rs.MoveLast
frmreturn.txtnum.Text = Format(rs!num + 1, "0000")
End If


Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from penalty")

 frmreturn.lblnum_penalty.Caption = CDate(frmreturn.lbldatereturn.Caption) - CDate(frmreturn.lblduedate.Caption)
  If frmreturn.lblnum_penalty.Caption < 1 Then
    frmreturn.lblnum_penalty.Caption = 0
    frmreturn.lblstat.Caption = "Returned"
    frmreturn.lbltot_penalty.Caption = 0
  Else
    frmreturn.lblstat.Caption = "Returned"
    frmreturn.lbltot_penalty.Caption = Val(frmreturn.lblnum_penalty.Caption) * Val(rs!penalty)
    frmreturn.lbltot_penalty.Caption = Format(frmreturn.lbltot_penalty.Caption, "#,###.00")
  End If


Unload Me


End Sub
