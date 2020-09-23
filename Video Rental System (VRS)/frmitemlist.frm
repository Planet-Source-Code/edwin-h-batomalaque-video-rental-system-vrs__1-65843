VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmitemlist 
   Caption         =   "VRS - List of Movies"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   Icon            =   "frmitemlist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "Add &New"
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
      Left            =   240
      Picture         =   "frmitemlist.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      Left            =   1245
      Picture         =   "frmitemlist.frx":273C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   945
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   2220
      Picture         =   "frmitemlist.frx":3006
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   990
   End
   Begin MSAdodcLib.Adodc adomovies 
      Height          =   375
      Left            =   7680
      Top             =   720
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
      RecordSource    =   "movies"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   855
      Left            =   3240
      Picture         =   "frmitemlist.frx":3CD0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   8895
      Begin MSDataGridLib.DataGrid dgridmovies 
         Bindings        =   "frmitemlist.frx":424C
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16776960
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "itemid"
            Caption         =   " Item ID #"
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
            DataField       =   "movtitle"
            Caption         =   "      Movie Title"
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
            DataField       =   "movcategory"
            Caption         =   "   Category"
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
            DataField       =   "movmainchar"
            Caption         =   "     Main Character"
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
            DataField       =   "movsecchar"
            Caption         =   "Secondary Character"
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
            DataField       =   "movdirector"
            Caption         =   "     Director"
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
            DataField       =   "movformt"
            Caption         =   "Format"
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
            DataField       =   "movstatus"
            Caption         =   "    Status"
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
            DataField       =   "movdpurchase"
            Caption         =   "Date Purchased"
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
            DataField       =   "movcopies"
            Caption         =   "No. of Copies"
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
         BeginProperty Column10 
            DataField       =   "movnodays"
            Caption         =   "No. of Days"
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
         BeginProperty Column11 
            DataField       =   "movamount"
            Caption         =   "Amount"
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
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   953
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   882
      TabCaption(0)   =   "Movies Database"
      TabPicture(0)   =   "frmitemlist.frx":4264
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   165
      Shape           =   4  'Rounded Rectangle
      Top             =   4740
      Width           =   4155
   End
End
Attribute VB_Name = "frmitemlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movies where val(itemid)  ='" & Val(dgridmovies.Columns(0)) & "'")
 Prompt$ = "Do you really want to delete this tapes?"
    reply = MsgBox(Prompt$, vbOKCancel, "Delete Record")
    If reply = vbOK Then
        rs.Delete
        rs.MoveNext
        rs.Close
        db.Close
        adomovies.Refresh
        dgridmovies.Refresh
   End If
   
End Sub

Private Sub cmdEdit_Click()

frmeditmovie.Show
frmeditmovie.enable
frmeditmovie.cmdsave.Enabled = True
frmeditmovie.cmdlcancel.Enabled = True


Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movies where val(itemid)  ='" & Val(dgridmovies.Columns(0)) & "'")
If rs.RecordCount = 0 Then
MsgBox ("No Selected Movies"), vbCritical, "Pls. Select One"
Else
frmeditmovie.txtitemid.Text = rs!itemid
frmeditmovie.txttitle.Text = rs!movtitle
frmeditmovie.cbocategory = rs!movcategory
frmeditmovie.txtmcast.Text = rs!movmainchar
frmeditmovie.txtseccast.Text = rs!movsecchar
frmeditmovie.cmdformat.Text = rs!movformt
frmeditmovie.cbostatus.Text = rs!movstatus
frmeditmovie.txtdirector.Text = rs!movdirector
frmeditmovie.datepurchase = rs!movdpurchase
frmeditmovie.txtprice.Text = rs!movamount
frmeditmovie.txtnoofcd.Text = rs!movcopies
frmeditmovie.txtnoofdays.Text = rs!movnodays

End If

End Sub

Private Sub cmdnew_Click()
frmaddmovie.Show
End Sub



Private Sub Command4_Click()
Unload Me
mdifrmmain.Enabled = True
End Sub

Private Sub Form_Load()
mdifrmmain.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdifrmmain.lblstatus.Caption = "List of All Items...... You can edit, add , delete and save penalty rate"
End Sub

Private Sub Form_Terminate()
mdifrmmain.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
End Sub
