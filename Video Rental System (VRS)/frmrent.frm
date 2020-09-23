VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmrent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VRS - Rent"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "frmrent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adorented 
      Height          =   330
      Left            =   9120
      Top             =   5520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "rented"
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
   Begin VB.CommandButton cmditem 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtformat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   420
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1470
      Width           =   1125
   End
   Begin VB.CommandButton cmdcan 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   540
      Left            =   3225
      Picture         =   "frmrent.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5340
      Width           =   1515
   End
   Begin VB.TextBox txtamount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1470
      Width           =   960
   End
   Begin VB.CommandButton cmdover 
      Caption         =   "Overide "
      Enabled         =   0   'False
      Height          =   390
      Left            =   9675
      TabIndex        =   7
      Top             =   1500
      Width           =   735
   End
   Begin VB.Frame frameborrow 
      BackColor       =   &H00808000&
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
      Height          =   1875
      Left            =   3120
      TabIndex        =   21
      Top             =   3000
      Width           =   3735
      Begin VB.TextBox txtidno 
         Height          =   435
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   2340
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "OK"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2070
         TabIndex        =   3
         Top             =   1275
         Width           =   855
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   495
         Left            =   -120
         TabIndex        =   30
         Top             =   -15
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   873
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   706
         BackColor       =   -2147483638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "   Borrowers ID Number"
         TabPicture(0)   =   "frmrent.frx":3CE4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
      End
   End
   Begin VB.CommandButton cmdrent 
      Caption         =   "Rent"
      Enabled         =   0   'False
      Height          =   435
      Left            =   9675
      TabIndex        =   6
      Top             =   960
      Width           =   705
   End
   Begin VB.Frame Frame2 
      Caption         =   "List of Rented Movies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3015
      Left            =   120
      TabIndex        =   26
      Top             =   2040
      Width           =   10365
      Begin VB.CommandButton cmdreceipt 
         Caption         =   "Print Receipts"
         Enabled         =   0   'False
         Height          =   495
         Left            =   8880
         TabIndex        =   36
         Top             =   2400
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dbgridrent 
         Bindings        =   "frmrent.frx":4B36
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777088
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "custnum"
            Caption         =   "Borrower ID"
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
            Caption         =   "Item ID #"
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
            Caption         =   "Movie Title"
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
         BeginProperty Column04 
            DataField       =   "dateborrowed"
            Caption         =   "Date Borrowed"
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
            Caption         =   "Due Date"
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
            Caption         =   "Status"
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
         BeginProperty Column08 
            DataField       =   "custname"
            Caption         =   "Borrowers Name"
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
            ScrollBars      =   2
            RecordSelectors =   0   'False
            BeginProperty Column00 
               DividerStyle    =   0
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               DividerStyle    =   0
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column04 
               DividerStyle    =   0
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column05 
               DividerStyle    =   0
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column06 
               DividerStyle    =   0
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column07 
               DividerStyle    =   0
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               DividerStyle    =   0
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtrunningamount 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox status 
         Height          =   615
         Left            =   7320
         TabIndex        =   38
         Top             =   1680
         Width           =   975
      End
      Begin MSAdodcLib.Adodc adotemp 
         Height          =   450
         Left            =   600
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   794
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
         RecordSource    =   "tempview"
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
      Begin VB.Label lblitemsno 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   8880
         TabIndex        =   19
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "Total Item Rented"
         Height          =   375
         Left            =   8880
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Running Amount"
         Height          =   195
         Left            =   8880
         TabIndex        =   27
         Top             =   435
         Width           =   1185
      End
   End
   Begin VB.TextBox txtstatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Enabled         =   0   'False
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
      Height          =   375
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1500
      Width           =   735
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   555
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   555
      Left            =   4755
      Picture         =   "frmrent.frx":4B4C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5325
      Width           =   1410
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Print All Rented"
      Enabled         =   0   'False
      Height          =   540
      Left            =   1740
      Picture         =   "frmrent.frx":69BE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5340
      Width           =   1500
   End
   Begin VB.CommandButton cmdnew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "New Borrowers"
      Enabled         =   0   'False
      Height          =   555
      Left            =   300
      MaskColor       =   &H00C0E0FF&
      Picture         =   "frmrent.frx":8830
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5325
      Width           =   1455
   End
   Begin VB.TextBox txtitemid 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "All Rented Movies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   10260
      Begin MSDataGridLib.DataGrid dbgriddetail 
         Bindings        =   "frmrent.frx":A6A2
         Height          =   1995
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   3519
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16774388
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "num"
            Caption         =   "Number"
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
            DataField       =   "custnum"
            Caption         =   "Borrowers ID"
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
            DataField       =   "itemid"
            Caption         =   "Item ID"
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
            DataField       =   "title"
            Caption         =   "Movie Title"
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
         BeginProperty Column05 
            DataField       =   "dateborrowed"
            Caption         =   "Date Borrowed"
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
            DataField       =   "duedate"
            Caption         =   "Due Date"
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
            DataField       =   "status"
            Caption         =   "Status"
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
            DataField       =   "amount"
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
         BeginProperty Column09 
            DataField       =   "custname"
            Caption         =   "Borrowers Name"
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
            DataField       =   "datereturn"
            Caption         =   "Date Returned"
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
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbrented 
         Bindings        =   "frmrent.frx":A6BA
         Height          =   1290
         Left            =   7440
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2275
         _Version        =   393216
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "invoice_num"
            Caption         =   "invoice_num"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "tot_items"
            Caption         =   "tot_items"
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
            DataField       =   "payments"
            Caption         =   "payments"
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
            DataField       =   "day_paid"
            Caption         =   "day_paid"
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
            DataField       =   "cashier"
            Caption         =   "cashier"
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbgridtemp 
         Bindings        =   "frmrent.frx":A6D2
         Height          =   1335
         Left            =   1440
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2355
         _Version        =   393216
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
         ColumnCount     =   8
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
         BeginProperty Column07 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
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
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   630
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1111
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   970
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   Borrowers Name"
      TabPicture(0)   =   "frmrent.frx":A6E8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblborrower"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin TabDlg.SSTab SSTab2 
         Height          =   690
         Left            =   7170
         TabIndex        =   31
         Top             =   -60
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1217
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabHeight       =   970
         TabCaption(0)   =   " ID #"
         TabPicture(0)   =   "frmrent.frx":AFC2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblidno"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Label lblidno 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1095
            TabIndex        =   32
            Top             =   150
            Width           =   2055
         End
      End
      Begin VB.Label lblborrower 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   2235
         TabIndex        =   23
         Top             =   75
         Width           =   4335
      End
   End
   Begin MSAdodcLib.Adodc adodetail 
      Height          =   495
      Left            =   6240
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
      BackColor       =   255
      ForeColor       =   8454143
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Video Rental System (VRS)\VRS Database\vrs.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Video Rental System (VRS)\VRS Database\vrs.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "items"
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
   Begin VB.TextBox txttotal 
      Height          =   285
      Left            =   9360
      TabIndex        =   35
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtdue 
      Height          =   525
      Left            =   6720
      TabIndex        =   34
      Top             =   3240
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adorent 
      Height          =   480
      Left            =   7680
      Top             =   5400
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
   Begin MSDataGridLib.DataGrid dgridmovies 
      Bindings        =   "frmrent.frx":CE44
      Height          =   855
      Left            =   4680
      TabIndex        =   37
      Top             =   6720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "itemid"
         Caption         =   "Item Number"
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
         Caption         =   "Movie Title"
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
         Caption         =   "Category"
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
         Caption         =   "Main Cast"
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
         Caption         =   "Secondary Cast"
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
         DataField       =   "movformat"
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
      BeginProperty Column06 
         DataField       =   "movstatus"
         Caption         =   "Status"
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
         DataField       =   "movdirector"
         Caption         =   "Director"
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
      BeginProperty Column10 
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
      BeginProperty Column11 
         DataField       =   "movpenalty"
         Caption         =   "Penalty Rate"
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
      BeginProperty Column12 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2594.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adomovies 
      Height          =   375
      Left            =   9120
      Top             =   7800
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
   Begin VB.TextBox txtinvoice 
      Height          =   375
      Left            =   5640
      TabIndex        =   39
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtnum 
      Height          =   375
      Left            =   5760
      TabIndex        =   41
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Format"
      Height          =   195
      Left            =   6240
      TabIndex        =   33
      Top             =   1560
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BorderColor     =   &H80000007&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   870
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   1485
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Left            =   8040
      TabIndex        =   29
      Top             =   1620
      Width           =   540
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFC0&
      BorderColor     =   &H80000007&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   870
      Left            =   4710
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   1485
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0FFFF&
      BorderColor     =   &H80000007&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   870
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   1485
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFC0&
      BorderColor     =   &H80000007&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   870
      Left            =   1740
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   1485
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      BorderColor     =   &H80000007&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   870
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5265
      Width           =   1485
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Left            =   4560
      TabIndex        =   25
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Movie Title"
      Height          =   195
      Left            =   4560
      TabIndex        =   24
      Top             =   1080
      Width           =   780
   End
   Begin VB.Shape Shape1 
      Height          =   915
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   795
      Width           =   4275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   825
   End
End
Attribute VB_Name = "frmrent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcan_Click()
clear

txtitemid.SetFocus

End Sub

Private Sub cmdcancel_Click()
Unload Me
mdifrmmain.Enabled = True
End Sub


Private Sub cmdclose_Click()
mdifrmmain.Enabled = True
Unload Me

            Do
                        Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
                        Set rs = db.OpenRecordset("Select *from tempview")
                            If rs.RecordCount = 1 Then
                            rs.Delete
                                rs.MoveLast
                                rs.Close
                                db.Close
                                frmrent.adotemp.Refresh
                                frmrent.dbgridtemp.Refresh
                            Else
                            GoTo edz:
                            End If
            Loop
                
edz:

End Sub

Private Sub cmditem_Click()


Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movies where itemid = '" & txtitemid & "'")
If rs.RecordCount = 0 Then
MsgBox ("Item Not Found"), vbInformation, "Not Found"
txtitemid.Text = ""
txtitemid.SetFocus
Else
txtitemid.Text = rs!itemid
txttitle.Text = rs!movtitle
txtstatus.Text = rs!movstatus
txtamount = rs!movamount
txtformat = rs!movformt
txtdue.Text = Format(Date + Val(rs!movnodays))

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from items")
If rs.RecordCount = 0 Then
txtnum.Text = "0001"
Else
rs.MoveLast
txtnum.Text = Format(rs!num + 1, "0000")
End If


Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from rented")
If rs.RecordCount = 0 Then
txtinvoice.Text = "00000001"
Else
rs.MoveLast
txtinvoice.Text = Format(rs!invoice_num + 1, "00000000")
End If

End If

cmdrent.Enabled = True
cmdover.Enabled = True
cmdrent.SetFocus


End Sub

Private Sub cmdnew_Click()
 
txtrunningamount.Text = ""
 frameborrow.Visible = True
 lblidno.Caption = ""
 lblborrower.Caption = ""
 txtidno.SetFocus
 
 Call ado
 


End Sub

Private Sub cmdOK_Click()
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from customers where custnum  ='" & txtidno & "'")
    If rs.RecordCount = 0 Then
    MsgBox ("Borrower's ID Not Exist"), vbInformation, "Not Found"
    txtidno.Text = ""
    txtidno.SetFocus
        Else
            frameborrow.Visible = False
            lblborrower.Caption = rs!custlname + ", " + rs!custfname + " " + rs!custmname
            lblidno.Caption = rs!custnum
            txtidno.Text = ""
            enable
            txtitemid.SetFocus
            
    End If
 
    Set db = OpenDatabase("D:\Video Rental System (VRS)\VRS Database\vrs.mdb")
    Set rs = db.OpenRecordset("Select *from movierental")
        If rs.RecordCount = lblidno.Caption Then
        Else

        End If

Call ado

dbgriddetail.Enabled = True


End Sub


Private Sub cmdover_Click()
txtamount.Locked = False
End Sub

Private Sub cmdreceipt_Click()
frmmembercash.Show
frmmembercash.Caption = "Rental Fees"
frmmembercash.cmdprint.Caption = "Print"
frmmembercash.lblamount.Text = txtrunningamount.Text
frmmembercash.lblidno.Caption = lblidno.Caption
frmmembercash.lblborrower.Caption = lblborrower.Caption
frmmembercash.lblitemscount.Caption = lblitemsno.Caption

End Sub

Private Sub cmdrent_Click()

status = "OUT"

If txtstatus.Text = "OUT" Then
MsgBox ("The Movies Is Out"), vbCritical, "Cannot be Rented"
clear
txtitemid.SetFocus

Else
cmdnew.Enabled = False
cmdreceipt.Enabled = True

    If txtitemid.Text = "" Then
        MsgBox ("No CD's Selected"), vbCritical, "Cannot Rent"
        txtitemid.SetFocus
        Else
                            
        
        Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
        Set rs = db.OpenRecordset("Select *from penalty")

            If Val(lblitemsno.Caption) >= Val(rs!rent_limits) Then
                MsgBox ("Cannot rent more than  ") + lblitemsno.Caption + (" items"), vbCritical, "Limit Exceeds"

            Else

            Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
            Set rs = db.OpenRecordset("Select *from movies where itemid = '" & txtitemid & "'")
                rs.Edit
                rs!itemid = txtitemid.Text
                rs!movstatus = status.Text
                rs.Update
                rs.Close
                db.Close
                adomovies.Refresh
                dgridmovies.Refresh
               
                
                                
                Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
                Set rs = db.OpenRecordset("Select *from items where num  ='" & txtnum & "'")
                    rs.AddNew
                    rs!num = txtnum.Text
                    rs!custnum = lblidno.Caption
                    rs!itemid = txtitemid.Text
                    rs!Title = txttitle.Text
                    rs!movformt = txtformat.Text
                    rs!dateborrowed = Date
                    rs!duedate = txtdue.Text
                    rs!Amount = txtamount.Text
                    rs!custname = lblborrower.Caption
                    rs!status = txtstatus.Text
                    rs.Update
                    rs.Close
                    db.Close
                    frmrent.adodetail.Refresh
                    frmrent.adodetail.Recordset.Update
                    frmrent.dbgriddetail.Refresh
                    txtnum.Text = ""
                    
    
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movierental where itemid  ='" & txtitemid & "'")
    rs.AddNew
    rs!custnum = lblidno.Caption
    rs!itemid = txtitemid.Text
    rs!Title = txttitle.Text
    rs!status = status.Text
    rs!movformt = txtformat
    rs!dateborrowed = Date
    rs!duedate = txtdue.Text
    rs!Amount = txtamount.Text
    rs!custname = lblborrower.Caption
    rs!rented_time = Time
        rs.Update
        rs.Close
        db.Close
        adorent.Refresh
        dbgridrent.Refresh
        
        

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from tempview where itemid  ='" & txtitemid & "'")
    rs.AddNew
    rs!custnum = lblidno.Caption
    rs!itemid = txtitemid.Text
    rs!Title = txttitle.Text
    rs!movformt = txtformat
    rs!dateborrowed = Date
    rs!duedate = txtdue.Text
    rs!Amount = txtamount.Text
    rs!custname = lblborrower.Caption
        rs.Update
        rs.Close
        db.Close
        adotemp.Refresh
        dbgridtemp.Refresh
        
                txtamount.Locked = True
                txtrunningamount.Text = Val(txtrunningamount.Text) + Val(txtamount.Text)
                txtrunningamount.Text = Format(txtrunningamount.Text, "#,###0.00")

                txttotal.Text = 1
                lblitemsno.Caption = Val(lblitemsno.Caption) + Val(txttotal.Text)


                clear
                txtitemid.SetFocus
          
          
               
  
                
    End If
 End If
    
End If

End Sub



Private Sub Form_Load()
mdifrmmain.Enabled = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdifrmmain.lblstatus.Caption = "Rent Videos ........ Enter Borrowers ID and the Item ID for the videos you want to rent."
End Sub

Private Sub Form_Terminate()
mdifrmmain.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Do
                        Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
                        Set rs = db.OpenRecordset("Select *from tempview")
                            If rs.RecordCount = 1 Then
                            rs.Delete
                                rs.MoveLast
                                rs.Close
                                db.Close
                                frmrent.adotemp.Refresh
                                frmrent.dbgridtemp.Refresh
                            Else
                            GoTo edz:
                            End If
            Loop
                
edz:
End Sub

Private Sub txtidno_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 13) Then
        KeyAscii = 0
    End If
    

If KeyAscii = 13 Then
txtidno.Text = Format(txtidno.Text, "00000")
Call cmdOK_Click
End If
End Sub

Private Sub txtidno_LostFocus()
txtidno.Text = Format(txtidno.Text, "00000")
End Sub


Private Sub txtitemid_Change()
cmditem.Enabled = True

End Sub

Private Sub txtitemid_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtitemid.Text = Format(txtitemid.Text, "000000")
Call cmditem_Click
End If

End Sub

Private Sub txtitemid_LostFocus()
txtitemid.Text = Format(txtitemid.Text, "000000")
End Sub





'----------- sub function ---------------------'

Sub enable()
cmdnew.Enabled = True
cmdprint.Enabled = True
cmdcan.Enabled = True
cmdcancel.Enabled = True
cmdClose.Enabled = True
dbgridrent.Enabled = True
lblborrower.Enabled = True
txtitemid.Enabled = True
txttitle.Enabled = True
txtstatus.Enabled = True
txtrunningamount.Enabled = True
lblitemsno.Enabled = True

End Sub




Sub ado()
adorent.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Video Rental System (VRS)\VRS Database\vrs.mdb;Persist Security Info=False"
adorent.RecordSource = "Select *from movierental where custnum = '" + lblidno + "'"
adorent.Refresh
dbgridrent.Refresh

End Sub


Sub clear()

txtitemid.Text = ""
txttitle.Text = ""
txtformat.Text = ""
txtamount.Text = ""
txtstatus.Text = ""

End Sub
