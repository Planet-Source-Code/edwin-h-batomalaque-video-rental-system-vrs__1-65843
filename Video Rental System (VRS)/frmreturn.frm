VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmreturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VRS - Return"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   Icon            =   "frmreturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Borrowers Details"
      Height          =   540
      Left            =   3240
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frmreturn.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5340
      Width           =   1470
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Close"
      Height          =   540
      Left            =   4725
      Picture         =   "frmreturn.frx":3CE4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5340
      Width           =   1470
   End
   Begin VB.CommandButton cmditem 
      Caption         =   "OK"
      Height          =   375
      Left            =   3315
      TabIndex        =   2
      Top             =   1020
      Width           =   855
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00008000&
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   9285
      Begin VB.Label lbltot_penalty 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   6600
         TabIndex        =   29
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblnum_penalty 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   6600
         TabIndex        =   28
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblstat 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   6600
         TabIndex        =   27
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbldatereturn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1680
         TabIndex        =   26
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblduedate 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1680
         TabIndex        =   25
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lbldborrowed 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label18 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Borrowed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Returned"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4440
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Days Penalty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Penalty Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   540
      Left            =   1740
      Picture         =   "frmreturn.frx":5B56
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5340
      Width           =   1500
   End
   Begin VB.CommandButton cmdreturn 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Return"
      Enabled         =   0   'False
      Height          =   555
      Left            =   300
      MaskColor       =   &H00C0E0FF&
      Picture         =   "frmreturn.frx":79C8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5325
      Width           =   1455
   End
   Begin VB.TextBox txtitemid 
      Height          =   375
      Left            =   1155
      TabIndex        =   1
      Top             =   1020
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unreturned  Movies"
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
      Width           =   9300
      Begin MSDataGridLib.DataGrid dbgridreturn 
         Bindings        =   "frmreturn.frx":983A
         Height          =   1935
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777088
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
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
            Caption         =   "Borrowers ID #"
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
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            RecordSelectors =   0   'False
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
               ColumnWidth     =   764.787
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
      TabIndex        =   7
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
      TabCaption(0)   =   "  Return Movies"
      TabPicture(0)   =   "frmreturn.frx":9852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Label lbldate 
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
         Left            =   8280
         TabIndex        =   15
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.TextBox txtdue 
      Height          =   525
      Left            =   6720
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adoreturn 
      Height          =   480
      Left            =   6240
      Top             =   6720
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   847
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
      RecordSource    =   "movierental"
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
   Begin VB.TextBox txtnum 
      Height          =   375
      Left            =   5160
      TabIndex        =   39
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label lblitemid 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   6720
      TabIndex        =   14
      Top             =   1860
      Width           =   2475
   End
   Begin VB.Label lblamount 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   37
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   195
      Left            =   3240
      TabIndex        =   36
      Top             =   2520
      Width           =   540
   End
   Begin VB.Label lblformat 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   35
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Movie Format"
      Height          =   195
      Left            =   285
      TabIndex        =   34
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label lblborrower 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   5640
      TabIndex        =   33
      Top             =   780
      Width           =   3615
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Borrowers Name"
      Height          =   195
      Left            =   4320
      TabIndex        =   32
      Top             =   960
      Width           =   1170
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "ID #"
      Height          =   195
      Left            =   5160
      TabIndex        =   31
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label lblidno 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   6720
      TabIndex        =   17
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   5880
      TabIndex        =   16
      Top             =   2520
      Width           =   450
   End
   Begin VB.Label Label3 
      Caption         =   "Item ID #"
      Height          =   255
      Left            =   5850
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   1455
      TabIndex        =   12
      Top             =   1875
      Width           =   4200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rented Movies"
      Height          =   255
      Left            =   285
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
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
   Begin VB.Shape Shape1 
      Height          =   930
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   705
      Width           =   9255
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
      Left            =   195
      TabIndex        =   6
      Top             =   1140
      Width           =   825
   End
   Begin VB.Label lbltime 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   6960
      TabIndex        =   38
      Top             =   7080
      Width           =   975
   End
End
Attribute VB_Name = "frmreturn"
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

End Sub


Private Sub cmdclear_Click()
clear
txtitemid.SetFocus
End Sub

Private Sub cmdclose_Click()
Unload Me
mdifrmmain.Enabled = True
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
Set rs = db.OpenRecordset("Select *from movierental where itemid = '" & txtitemid & "'")
If rs.RecordCount = 0 Then
MsgBox ("Item Not Found"), vbInformation, "Not Found"
txtitemid.Text = ""
txtitemid.SetFocus

Else

lblidno.Caption = rs!custnum
lblitemid.Caption = rs!itemid
lbltime.Caption = rs!rented_time
lbltitle.Caption = rs!Title
lblstatus.Caption = rs!status
lblformat.Caption = rs!movformt
lbldborrowed.Caption = rs!dateborrowed
lblduedate.Caption = rs!duedate
lbldatereturn.Caption = Date
lblamount.Caption = rs!Amount

cmdreturn.Enabled = True
cmdclear.Enabled = True

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from return")
If rs.RecordCount = 0 Then
txtnum.Text = "0001"
Else
rs.MoveLast
txtnum.Text = Format(rs!num + 1, "0000")
End If



Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from customers where custnum  ='" & lblidno & "'")
lblborrower.Caption = rs!custlname + ", " + rs!custfname + " " + rs!custmname
   
    
lblstat = "Return"

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from penalty")

 lblnum_penalty.Caption = CDate(lbldatereturn.Caption) - CDate(lblduedate.Caption)
  If lblnum_penalty.Caption < 1 Then
    lblnum_penalty.Caption = 0
    lblstat.Caption = "Returned"
    lbltot_penalty.Caption = 0
  Else
    lblstat.Caption = "Returned"
    lbltot_penalty.Caption = Val(lblnum_penalty.Caption) * Val(rs!penalty)
    lbltot_penalty.Caption = Format(lbltot_penalty.Caption, "#,###.00")
  End If



cmdreturn.SetFocus


End If


End Sub





Private Sub txtidno_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 13) Then
        KeyAscii = 0
    End If
    

If KeyAscii = 13 Then
txtIdno.Text = Format(txtIdno.Text, "00000")
Call cmdOK_Click
End If
End Sub

Private Sub txtidno_LostFocus()
txtIdno.Text = Format(txtIdno.Text, "00000")
End Sub


Private Sub cmdreturn_Click()

If txtitemid.Text = "" Or lblborrower.Caption = "" Or lblidno.Caption = "" Or lbltitle.Caption = "" Then
MsgBox ("Enter Item ID"), vbCritical, "Cannot Find CD"
Else
    Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
    Set rs = db.OpenRecordset("Select *from return where num  ='" & txtnum & "'")
                    rs.AddNew
                    rs!num = txtnum.Text
                    rs!custnum = lblidno.Caption
                    rs!itemid = txtitemid.Text
                    rs!Title = lbltitle.Caption
                    rs!movformt = lblformat.Caption
                    rs!dateborrowed = lbldborrowed.Caption
                    rs!duedate = lblduedate.Caption
                    rs!Amount = lblamount.Caption
                    rs!custname = lblborrower.Caption
                    rs!datereturn = lbldatereturn.Caption
                    rs!penalty_num = lblnum_penalty.Caption
                    rs!tot_penalty = lbltot_penalty.Caption
                    rs.Update
                    rs.Close
                    db.Close
                    txtnum.Text = ""

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movies where itemid = '" & txtitemid & "'")
        rs.Edit
        rs!itemid = txtitemid.Text
        rs!movstatus = "IN"
            rs.Update
            rs.Close
            db.Close
            frmitemlist.adomovies.Refresh
            frmitemlist.dgridmovies.Refresh
            
       
        
          
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from movierental where itemid = '" & txtitemid & "'")
    rs.Delete
    rs.MoveNext
    rs.Close
    db.Close
    dbgridreturn.Refresh
    adoreturn.Refresh
    
        
    txtnum.Text = ""

 
 If lbltot_penalty.Caption >= 1 Then
       
               Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
                Set rs = db.OpenRecordset("Select *from tempview where itemid  ='" & txtitemid & "'")
                    rs.AddNew
                    rs!custnum = lblidno.Caption
                    rs!itemid = txtitemid.Text
                    rs!Title = lbltitle.Caption
                    rs!movformt = lblformat.Caption
                    rs!dateborrowed = Date
                    rs!duedate = lblduedate.Caption
                    rs!custname = lblborrower.Caption
                        rs.Update
                        rs.Close
                        db.Close
                        frmrent.adotemp.Refresh
                        frmrent.dbgridtemp.Refresh
                
                Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
                Set rs = db.OpenRecordset("Select *from membership_fee")
                If rs.RecordCount = 0 Then
                        frmmembercash.lbltrans.Caption = "00000001"
                Else
                        rs.MoveLast
                        frmmembercash.lbltrans.Caption = Format(rs!invoice_num + 1, "00000000")
                End If
              
                    Call cmdreceipt
                    frmmembercash.txtcash.SetFocus
               
           
     Else
        MsgBox ("Successfully Returned"), vbInformation, "Returned Items"
        txtnum.Text = ""
        clear
        txtitemid.SetFocus
        cmdreturn.Enabled = False
        cmdclear.Enabled = False
        End If
   
  End If


End Sub

Private Sub Command1_Click()
frmsearchcust.Show

End Sub

Private Sub Form_Load()
mdifrmmain.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdifrmmain.lblstatus.Caption = "Return Video CD's .......... Pay Attention to the borrowers if the CD is in good condition before you return"
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Form_Terminate()
Unload Me
 mdifrmmain.Enabled = True
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

Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True

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




'--------------------- Sub Command ------------------------------'



Sub clear()

txtitemid.Text = ""
lblborrower.Caption = ""
lblidno.Caption = ""
lblitemid.Caption = ""
lbltitle.Caption = ""
lblstatus.Caption = ""
lblformat.Caption = ""
lbldborrowed.Caption = ""
lblduedate.Caption = ""
lbldatereturn.Caption = ""
lblamount.Caption = ""
lblstat.Caption = ""
lblnum_penalty.Caption = ""
lbltot_penalty.Caption = ""
lbltime.Caption = ""
txtnum.Text = ""


End Sub

Sub cmdreceipt()

frmmembercash.Show
frmmembercash.cmdprint.Caption = "Print"
frmmembercash.Caption = "Penalty Fee"
frmmembercash.lblamount.Text = lbltot_penalty.Caption
frmmembercash.lblidno = lblidno.Caption
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from customers where custnum = '" & lblidno.Caption & "'")
frmmembercash.lblborrower.Caption = rs!custlname + ", " + rs!custfname + " " + rs!custmname
End Sub

