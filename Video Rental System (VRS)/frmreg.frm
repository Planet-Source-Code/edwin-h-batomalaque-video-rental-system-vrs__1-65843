VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmembership 
   BackColor       =   &H00C0C000&
   Caption         =   "VRS - Membership"
   ClientHeight    =   10500
   ClientLeft      =   270
   ClientTop       =   1725
   ClientWidth     =   11400
   DrawStyle       =   5  'Transparent
   FillColor       =   &H80000012&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmreg.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtnname 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   9600
      TabIndex        =   30
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtamount 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   9600
      TabIndex        =   29
      Top             =   3960
      Width           =   2100
   End
   Begin VB.OptionButton Opothersproof 
      Caption         =   "Others"
      Enabled         =   0   'False
      Height          =   255
      Left            =   12360
      TabIndex        =   26
      Top             =   2490
      Width           =   1455
   End
   Begin VB.OptionButton Optel 
      Caption         =   "Telephone Bill"
      Enabled         =   0   'False
      Height          =   255
      Left            =   12360
      TabIndex        =   25
      Top             =   2220
      Width           =   1455
   End
   Begin VB.OptionButton Opwater 
      Caption         =   "Water Bill"
      Enabled         =   0   'False
      Height          =   255
      Left            =   12360
      TabIndex        =   24
      Top             =   1965
      Width           =   1455
   End
   Begin VB.OptionButton Opelectric 
      Caption         =   "Electric Bill"
      Enabled         =   0   'False
      Height          =   255
      Left            =   12360
      TabIndex        =   23
      Top             =   1725
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Available Requirements"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2235
      Left            =   9120
      TabIndex        =   65
      Top             =   1080
      Width           =   5505
      Begin VB.TextBox txtproof 
         Height          =   330
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1680
         Width           =   1965
      End
      Begin VB.OptionButton opothersid 
         Caption         =   "Others"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1455
         Width           =   975
      End
      Begin VB.OptionButton opsssid 
         Caption         =   "SSS ID"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1155
         Width           =   1575
      End
      Begin VB.OptionButton opdriveid 
         Caption         =   "Drivers License ID"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   900
         Width           =   1815
      End
      Begin VB.OptionButton opstudentid 
         Caption         =   "Student ID"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   660
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtval 
         Height          =   315
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1725
         Width           =   1815
      End
      Begin VB.Label Label28 
         Caption         =   "Proof of Billing"
         Height          =   255
         Left            =   2880
         TabIndex        =   67
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Valid ID'S"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox txttelrelative 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6810
      TabIndex        =   17
      Top             =   5010
      Width           =   2085
   End
   Begin VB.TextBox txtaddrelative 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1425
      TabIndex        =   18
      Top             =   5430
      Width           =   7485
   End
   Begin VB.TextBox txtnamerelative 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1440
      TabIndex        =   16
      Top             =   5040
      Width           =   4260
   End
   Begin VB.TextBox txtIdno 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "News705 BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   2025
      TabIndex        =   0
      Top             =   390
      Width           =   2775
   End
   Begin VB.ComboBox cbostatus 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmreg.frx":9059
      Left            =   3720
      List            =   "frmreg.frx":9069
      TabIndex        =   15
      Top             =   4365
      Width           =   1965
   End
   Begin VB.TextBox txtcmobile 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6720
      TabIndex        =   10
      Top             =   3840
      Width           =   2130
   End
   Begin VB.ComboBox cbosex 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmreg.frx":9091
      Left            =   1320
      List            =   "frmreg.frx":909B
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.ComboBox cboreligion 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmreg.frx":90AD
      Left            =   3720
      List            =   "frmreg.frx":90CF
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   3600
      Width           =   1965
   End
   Begin VB.TextBox txtctelephone 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "!(999) 000-0000;0;_"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   315
      Left            =   6720
      TabIndex        =   13
      Top             =   4200
      Width           =   2115
   End
   Begin VB.ComboBox cbonationality 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmreg.frx":913D
      Left            =   3720
      List            =   "frmreg.frx":9165
      TabIndex        =   12
      Top             =   3960
      Width           =   1980
   End
   Begin VB.Frame Frame2 
      Caption         =   "Registered Customer"
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
      Height          =   2235
      Left            =   360
      TabIndex        =   45
      Top             =   6735
      Width           =   11535
      Begin MSDataGridLib.DataGrid dbgridmember 
         Bindings        =   "frmreg.frx":91E0
         Height          =   1890
         Left            =   120
         TabIndex        =   37
         Top             =   285
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3334
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         ColumnCount     =   23
         BeginProperty Column00 
            DataField       =   "custnum"
            Caption         =   "Customer ID #"
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
            DataField       =   "custlname"
            Caption         =   "Last Name"
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
            DataField       =   "custfname"
            Caption         =   "First Name"
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
            DataField       =   "custmname"
            Caption         =   "Middle Name"
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
            DataField       =   "custaddress"
            Caption         =   "Address"
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
            DataField       =   "custcity"
            Caption         =   "City"
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
            DataField       =   "custstatus"
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
            DataField       =   "custsex"
            Caption         =   "Sex"
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
            DataField       =   "custpostcode"
            Caption         =   "Zip Code"
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
            DataField       =   "custregdate"
            Caption         =   "Registered Date"
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
            DataField       =   "custbday"
            Caption         =   "Birth Day"
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
            DataField       =   "custage"
            Caption         =   "Age"
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
            DataField       =   "custreligion"
            Caption         =   "Religion"
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
         BeginProperty Column13 
            DataField       =   "custzipcode"
            Caption         =   "Zip Code"
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
         BeginProperty Column14 
            DataField       =   "custnamerelative"
            Caption         =   "Name of Relative"
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
         BeginProperty Column15 
            DataField       =   "custmobile"
            Caption         =   "Mobile Number"
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
         BeginProperty Column16 
            DataField       =   "custtelephone"
            Caption         =   "Telephone Number"
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
         BeginProperty Column17 
            DataField       =   "custvalidid"
            Caption         =   "Valid ID"
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
         BeginProperty Column18 
            DataField       =   "custproof"
            Caption         =   "Proof of Billing"
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
         BeginProperty Column19 
            DataField       =   "custtelrelative"
            Caption         =   "Tel of Relative"
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
         BeginProperty Column20 
            DataField       =   "custaddrelative"
            Caption         =   "custaddrelative"
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
         BeginProperty Column21 
            DataField       =   "custnationality"
            Caption         =   "Nationality"
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
         BeginProperty Column22 
            DataField       =   "custstate"
            Caption         =   "State"
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column22 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9240
      MaskColor       =   &H000000FF&
      Picture         =   "frmreg.frx":91FC
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6180
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7425
      MaskColor       =   &H000000FF&
      Picture         =   "frmreg.frx":A46E
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6180
      UseMaskColor    =   -1  'True
      Width           =   1770
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Update Record"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5670
      MaskColor       =   &H000000FF&
      Picture         =   "frmreg.frx":C2E0
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6180
      UseMaskColor    =   -1  'True
      Width           =   1740
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Edit Record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3900
      MaskColor       =   &H000000FF&
      Picture         =   "frmreg.frx":E152
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6180
      UseMaskColor    =   -1  'True
      Width           =   1740
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Save Record"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2160
      MaskColor       =   &H000000FF&
      Picture         =   "frmreg.frx":FFC4
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6180
      UseMaskColor    =   -1  'True
      Width           =   1725
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Add New Record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   390
      MaskColor       =   &H00C0E0FF&
      Picture         =   "frmreg.frx":11E36
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6180
      Width           =   1740
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Personal Information"
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
      Height          =   2295
      Left            =   360
      TabIndex        =   39
      Top             =   1080
      Width           =   8655
      Begin VB.TextBox txtzip 
         Enabled         =   0   'False
         Height          =   330
         Left            =   6600
         TabIndex        =   7
         Top             =   1830
         Width           =   1815
      End
      Begin VB.ComboBox cbostate 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmreg.frx":13CA8
         Left            =   1440
         List            =   "frmreg.frx":13CD0
         TabIndex        =   6
         Top             =   1860
         Width           =   3015
      End
      Begin VB.TextBox txtcity 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   1440
         Width           =   6975
      End
      Begin VB.TextBox txtLname 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   615
         Width           =   2775
      End
      Begin VB.TextBox txtAdd 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1050
         Width           =   6975
      End
      Begin VB.TextBox txtMname 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5535
         TabIndex        =   3
         Top             =   600
         Width           =   2880
      End
      Begin VB.TextBox txtFname 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3075
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   5040
         TabIndex        =   71
         Top             =   1950
         Width           =   1140
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State/Country"
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
         TabIndex        =   51
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
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
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "City/Province"
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
         TabIndex        =   44
         Top             =   1590
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         TabIndex        =   42
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
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
         Left            =   5520
         TabIndex        =   41
         Top             =   375
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Left            =   3120
         TabIndex        =   40
         Top             =   375
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc adomembership 
      Height          =   375
      Left            =   6480
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "customers"
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
   Begin MSComCtl2.DTPicker DTBdate 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   53411841
      CurrentDate     =   38781
   End
   Begin VB.Image imagestatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   2220
      Left            =   12240
      Picture         =   "frmreg.frx":13D4C
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2100
   End
   Begin VB.Image imagepic 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   12120
      Picture         =   "frmreg.frx":158E2
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nick Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9240
      TabIndex        =   70
      Top             =   4680
      Width           =   960
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Membership Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   9240
      TabIndex        =   69
      Top             =   3600
      Width           =   2025
   End
   Begin VB.Label lbldatereg 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      Height          =   300
      Left            =   12480
      TabIndex        =   68
      Top             =   435
      Width           =   2175
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.#"
      Height          =   195
      Left            =   6120
      TabIndex        =   64
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   840
      TabIndex        =   63
      Top             =   5520
      Width           =   570
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   840
      TabIndex        =   62
      Top             =   5160
      Width           =   420
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "In Case of Emergency pls. contact"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   480
      TabIndex        =   61
      Top             =   4800
      Width           =   2850
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   12000
      TabIndex        =   38
      Top             =   8520
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   885
      Index           =   6
      Left            =   10920
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3120
      TabIndex        =   60
      Top             =   4440
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   480
      TabIndex        =   59
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   480
      TabIndex        =   58
      Top             =   4080
      Width           =   330
   End
   Begin VB.Label lblage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1350
      TabIndex        =   11
      Top             =   3990
      Width           =   1155
   End
   Begin VB.Label label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   5880
      TabIndex        =   57
      Top             =   3600
      Width           =   1365
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2760
      TabIndex        =   56
      Top             =   4080
      Width           =   915
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3000
      TabIndex        =   55
      Top             =   3720
      Width           =   675
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   480
      TabIndex        =   54
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6120
      TabIndex        =   53
      Top             =   3960
      Width           =   465
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. #"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6120
      TabIndex        =   52
      Top             =   4320
      Width           =   420
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   2580
      Left            =   9165
      Top             =   3450
      Width           =   5520
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Membership"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   10440
      TabIndex        =   50
      Top             =   480
      Width           =   1950
   End
   Begin VB.Label grand 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   11760
      TabIndex        =   49
      Top             =   9120
      Width           =   2415
   End
   Begin VB.Label txtcourse 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   12120
      TabIndex        =   48
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label txtmajor 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   12120
      TabIndex        =   47
      Top             =   1800
      Width           =   2565
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      FillColor       =   &H00404040&
      Height          =   8925
      Left            =   120
      Top             =   255
      Width           =   14790
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   885
      Index           =   5
      Left            =   9165
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   1785
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   885
      Index           =   4
      Left            =   7395
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   1785
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   885
      Index           =   3
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   1755
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   885
      Index           =   2
      Left            =   3870
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   1770
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   885
      Index           =   1
      Left            =   2115
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   1770
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   885
      Index           =   0
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "ID Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   360
      TabIndex        =   43
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   360
      Top             =   3450
      Width           =   8670
   End
End
Attribute VB_Name = "frmmembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbonationality_GotFocus()
cbonationality.BackColor = &HFFFFC0
cbonationality.ForeColor = &HFF&

End Sub

Private Sub cbonationality_KeyPress(KeyAscii As Integer)
Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If
If KeyAscii = 13 Then
cbosex.SetFocus
End If

End Sub

Private Sub cbonationality_LostFocus()
cbonationality.BackColor = &HFFFFFF
cbonationality.ForeColor = &H80000012
End Sub

Private Sub cboreligion_GotFocus()
cboreligion.BackColor = &HFFFFC0
cboreligion.ForeColor = &HFF&

End Sub

Private Sub cboreligion_KeyPress(KeyAscii As Integer)
Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If
If KeyAscii = 13 Then
cbonationality.SetFocus
End If


End Sub

Private Sub cboreligion_LostFocus()

cboreligion.BackColor = &HFFFFFF
cboreligion.ForeColor = &H80000012

If cboreligion.Text = "Others" Then
txtreligion.Enabled = True
End If

End Sub

Private Sub cbosex_GotFocus()
cbosex.BackColor = &HFFFFC0
cbosex.ForeColor = &HFF&

End Sub

Private Sub cbosex_KeyPress(KeyAscii As Integer)
Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If
If KeyAscii = 13 Then
cbostatus.SetFocus
End If

End Sub

Private Sub cbosex_LostFocus()

cbosex.BackColor = &HFFFFFF
cbosex.ForeColor = &H80000012

End Sub

Private Sub cbostate_GotFocus()
cbostate.BackColor = &HFFFFC0
cbostate.ForeColor = &HFF&

End Sub

Private Sub cbostate_KeyPress(KeyAscii As Integer)
Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If

If KeyAscii = 13 Then
DTBdate.SetFocus
End If


End Sub

Private Sub cbostate_LostFocus()
cbostate.BackColor = &HFFFFFF
cbostate.ForeColor = &H80000012

End Sub

Private Sub cbostatus_GotFocus()
cbostatus.BackColor = &HFFFFC0
cbostatus.ForeColor = &HFF&

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

Private Sub cbostatus_LostFocus()

cbostatus.BackColor = &HFFFFFF
cbostatus.ForeColor = &H80000012

End Sub

Private Sub cmdAdd_Click()
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from customers")
If rs.RecordCount = 0 Then
txtidno.Text = "00001"
Else
rs.MoveLast
txtidno.Text = Format(rs!custnum + 1, "00000")
End If
'enable entries
enable


'disable command buttons
cmdadd.Enabled = False
cmdEdit.Enabled = False
cmdUpdate.Enabled = False
cmdClose.Enabled = False
txtidno.Enabled = False

'enable command buttons
cmdcancel.Enabled = True
cmdsave.Enabled = True


txtLname.SetFocus
lbldatereg.Caption = Date

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from penalty")
txtamount.Text = rs!mem_fee

enable_option


End Sub


Private Sub cmdcancel_Click()

'enable buttons
cmdadd.Enabled = True
cmdEdit.Enabled = True
cmdClose.Enabled = True



'disable buttons
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdUpdate.Enabled = False


'clear entries

entries

'disable entries
disable

End Sub


Private Sub cmdclose_Click()

Unload frmmembership
mdifrmmain.Show

End Sub



Private Sub cmdDelete_Click()


Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from customers where val(custnum)  ='" & Val(dbgridmember.Columns(0)) & "'")
 Prompt$ = "Do you really want to delete the members information?"
    reply = MsgBox(Prompt$, vbOKCancel, "Delete Record")
    If reply = vbOK Then
        rs.Delete
        rs.MoveNext
        rs.Close
        db.Close
        adomembership.Refresh
        dbgridmember.Refresh
   End If
   
End Sub

Private Sub cmdEdit_Click()
frmsearchmember.Show
'enable cmd buttons
cmdUpdate.Enabled = True
cmdcancel.Enabled = True

'disable cmd buttons
cmdadd.Enabled = False
cmdsave.Enabled = False
cmdEdit.Enabled = False
cmdClose.Enabled = False




txtidno.Enabled = False

End Sub


Private Sub cmdsave_Click()
    
If txtLname.Text = "" Or txtFname.Text = "" Or txtMname.Text = "" Or txtAdd.Text = "" Or cbosex.Text = "" Or cbostate.Text = "" Or lblage = "" Then
    MsgBox "Missing Data! Do not leave a blank textfield.", vbInformation, "Information"
Else

    Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb ")
    Set rs = db.OpenRecordset("Select *from customers where custlname + custfname + custmname + Format(custbday)  = '" & txtLname + txtFname + txtMname + Format(DTBdate) & "'")
        If rs.RecordCount = 0 Then
            rs.AddNew
            rs!custnum = txtidno.Text
            rs!custlname = txtLname.Text
            rs!custfname = txtFname.Text
            rs!custmname = txtMname.Text
            rs!custaddress = txtAdd.Text
            rs!custage = lblage.Caption
            rs!custcity = txtcity.Text
            rs!custstate = cbostate.Text
            rs!custbday = DTBdate
            rs!custreligion = cboreligion.Text
            rs!custnationality = cbonationality.Text
            rs!custsex = cbosex.Text
            rs!custstatus = cbostatus.Text
            rs!custzipcode = txtzip.Text
            rs!custnamerelative = txtnamerelative.Text
            rs!custtelrelative = txttelrelative.Text
            rs!custaddrelative = txtaddrelative.Text
            rs!custmobile = txtcmobile.Text
            rs!custtelephone = txtctelephone.Text
            rs!custvalidid = txtval.Text
            rs!custproof = txtproof.Text
            rs!custaddrelative = txtaddrelative.Text
            rs!custregdate = lbldatereg.Caption
            rs!custvalidid = txtval.Text
            rs!custnname = txtnname.Text
            rs!custamount = txtamount.Text
            rs.Update
            rs.Close
            db.Close
            adomembership.Refresh
            dbgridmember.Refresh
                

        Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
        Set rs = db.OpenRecordset("Select *from membership_fee")
                If rs.RecordCount = 0 Then
                frmmembercash.lbltrans.Caption = "00000001"
                Else
                rs.MoveLast
                frmmembercash.lbltrans.Caption = Format(rs!invoice_num + 1, "00000000")
                End If
                
        frmmembercash.Show
        frmmembercash.Caption = "Membership Fee"
        frmmembercash.cmdprint.Caption = "OK"
        frmmembercash.lblamount = txtamount.Text
        frmmembercash.lblidno.Caption = txtidno.Text
        frmmembercash.lblborrower.Caption = txtLname.Text + ", " + txtFname.Text + " " + txtMname.Text
       
       
        
        

'clear entries
        entries

'disable entries
        disable

        cmdsave.Enabled = False
        cmdcancel.Enabled = False

        cmdadd.Enabled = True
        cmdEdit.Enabled = True
        cmdClose.Enabled = True


        Else
        MsgBox ("Members are currently registered"), vbCritical, "Existing Members"
        End If
End If




End Sub

Private Sub cmdUpdate_Click()

Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
Set rs = db.OpenRecordset("Select *from customers where custlname + custfname + custmname + Format(custbday)  = '" & txtLname + txtFname + txtMname + Format(DTBdate) & "'")
rs.Edit
rs!custnum = txtidno.Text
rs!custlname = txtLname.Text
rs!custfname = txtFname.Text
rs!custmname = txtMname.Text
rs!custaddress = txtAdd.Text
rs!custage = lblage.Caption
rs!custcity = txtcity.Text
rs!custstate = cbostate.Text
rs!custbday = DTBdate
rs!custreligion = cboreligion.Text
rs!custnationality = cbonationality.Text
rs!custsex = cbosex.Text
rs!custstatus = cbostatus.Text
rs!custzipcode = txtzip.Text
rs!custnamerelative = txtnamerelative.Text
rs!custtelrelative = txttelrelative.Text
rs!custaddrelative = txtaddrelative.Text
rs!custmobile = txtcmobile.Text
rs!custtelephone = txtctelephone.Text
rs!custvalidid = txtval.Text
rs!custproof = txtproof.Text
rs!custaddrelative = txtaddrelative.Text
rs!custregdate = lbldatereg.Caption
rs!custvalidid = txtval.Text
rs!custnname = txtnname.Text
rs!custamount = txtamount.Text

rs.Update
rs.Close
db.Close
adomembership.Refresh
dbgridmember.Refresh



txtidno.Enabled = False

'clear entries
entries

'disable entries
disable

cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdUpdate.Enabled = False


cmdadd.Enabled = True
cmdEdit.Enabled = True
cmdClose.Enabled = True



MsgBox ("Successfully Updated"), vbInformation, "Successfull"

End Sub


Private Sub Command1_Click()
        
             
CommonDialog1.InitDir = "C:\My Documents"
CommonDialog1.Filter = "JPEG image|*.jpg|GIF image|*.gif|BITMAP image|*.bmp|Icon image|*.ico|Cursor image|*.cur|One Touch image|*.one"
 CommonDialog1.ShowOpen
 strImgN = CommonDialog1.FileName
 txtPictureName.Text = CommonDialog1.FileTitle
 imagepic.Picture = LoadPicture(CommonDialog1.FileName)
      
    
    
End Sub

Private Sub Command2_Click()
imagepic.Picture = LoadPicture("")
End Sub

Private Sub dbgridmember_Click()
lblstatus.Caption = adomembership.Recordset("custlname")

End Sub

Private Sub DTBdate_Change()

If DTBdate > Date Then
MsgBox ("Invalid Birthday, Please Select Again"), vbInformation, "Processing Timeout"
DTBdate.Year = Format(Date, "yyyy")
Else
lblage = Int(DateDiff("d", DTBdate, Now()) / 365.255)
DTBdate.CalendarBackColor = &HFFFFFF
DTBdate.CalendarForeColor = &H80000012
End If

End Sub


Private Sub DTBdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cboreligion.SetFocus
End If

End Sub

Private Sub DTBdate_LostFocus()
If DTBdate > Date Then
MsgBox ("Invalid Birthday, Please Select Again"), vbInformation, "Processing Timeout"
DTBdate.Year = Format(Date, "yyyy")
Else
lblage = Int(DateDiff("d", DTBdate, Now()) / 365.255)
DTBdate.CalendarBackColor = &HFFFFFF
DTBdate.CalendarForeColor = &H80000012
End If
End Sub

Private Sub Form_Load()
'mdifrmmain.acc.Enabled = False


opstudentid.Value = False
opdriveid.Value = False
opsssid.Value = False
opothersid.Value = False
Opelectric.Value = False
Opwater.Value = False
Optel.Value = False
Opothersproof.Value = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdifrmmain.lblstatus.Caption = "Membership Form....... Fill-Up the form clearly and accurate dont leave any blank fields."
End Sub

Private Sub opdriveid_Click()
If opdriveid.Enabled = True Then
txtval.Text = "Drivers License ID"
txtval.ForeColor = &H8000000E

End If
End Sub

Private Sub Opelectric_Click()
If Opelectric.Enabled = True Then
txtproof.Text = "Electric Bill"
txtproof.ForeColor = &H8000000E
End If
End Sub

Private Sub opothersid_Click()
If opstudentid.Enabled = True Then
txtval.ForeColor = &H80000012
txtval.Text = ""
txtval.Locked = False



End If
End Sub

Private Sub Opothersproof_Click()
If Opothersproof.Enabled = True Then
txtproof.ForeColor = &H80000012
txtproof.Text = ""
txtproof.Locked = False

End If

End Sub

Private Sub opsssid_Click()
If opsssid.Enabled = True Then
txtval.Text = "SSS ID"
txtval.ForeColor = &H8000000E

End If
End Sub

Private Sub opstudentid_Click()
If opstudentid.Enabled = True Then
txtval.Text = "Student ID"
txtval.ForeColor = &H8000000E

End If

End Sub


Private Sub Optel_Click()
If Optel.Enabled = True Then
txtproof.Text = "Telphone Bill"
txtproof.ForeColor = &H8000000E
End If
End Sub

Private Sub Opwater_Click()
If Opwater.Enabled = True Then
txtproof.Text = "Water Bill"
txtproof.ForeColor = &H8000000E
End If

End Sub

Private Sub txtAdd_GotFocus()
txtAdd.BackColor = &HFFFFC0
txtAdd.ForeColor = &HFF&

End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtcity.SetFocus
End If

End Sub

Private Sub txtAdd_LostFocus()
txtAdd.Text = StrConv(txtAdd, vbProperCase)
txtAdd.BackColor = &HFFFFFF
txtAdd.ForeColor = &H80000012

End Sub


Private Sub txtcontact_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Exit Sub
End If
If Not IsNumeric(Chr(KeyAscii)) Then
KeyAscii = 0
End If
End Sub

Private Sub txtaddrelative_GotFocus()
txtaddrelative.BackColor = &HFFFFC0
txtaddrelative.ForeColor = &HFF&

End Sub

Private Sub txtaddrelative_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) Then
KeyAscii = 0
End If

End Sub

Private Sub txtaddrelative_LostFocus()
txtaddrelative.Text = StrConv(txtaddrelative, vbProperCase)
txtaddrelative.BackColor = &HFFFFFF
txtaddrelative.ForeColor = &H80000012
End Sub

Private Sub txtamount_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
KeyAscii = 0
End If
End Sub

Private Sub txtcity_GotFocus()
txtcity.BackColor = &HFFFFC0
txtcity.ForeColor = &HFF&

End Sub

Private Sub txtcity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cbostate.SetFocus
Else
If IsNumeric(Chr(KeyAscii)) Then
KeyAscii = 0

End If

End If

End Sub

Private Sub txtcity_LostFocus()
txtcity.Text = StrConv(txtcity, vbProperCase)
txtcity.BackColor = &HFFFFFF
txtcity.ForeColor = &H80000012

End Sub

Private Sub txtcmobile_GotFocus()
txtcmobile.BackColor = &HFFFFC0
txtcmobile.ForeColor = &HFF&


End Sub

Private Sub txtcmobile_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
KeyAscii = 0
End If

End Sub

Private Sub txtcmobile_LostFocus()

txtcmobile.BackColor = &HFFFFFF
txtcmobile.ForeColor = &H80000012

End Sub

Private Sub txtctelephone_GotFocus()
'txtctelephone.DataFormat = True

txtctelephone.BackColor = &HFFFFC0
txtctelephone.ForeColor = &HFF&

End Sub

Private Sub txtctelephone_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
KeyAscii = 0
End If

End Sub

Private Sub txtctelephone_LostFocus()

txtctelephone.BackColor = &HFFFFFF
txtctelephone.ForeColor = &H80000012

End Sub

Private Sub txtFname_GotFocus()
txtFname.BackColor = &HFFFFC0
txtFname.ForeColor = &HFF&
End Sub

Private Sub txtFname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtMname.SetFocus
Else
If IsNumeric(Chr(KeyAscii)) Then
KeyAscii = 0

End If
End If

End Sub

Private Sub txtFname_LostFocus()
txtFname.Text = StrConv(txtFname, vbProperCase)
txtFname.BackColor = &HFFFFFF
txtFname.ForeColor = &H80000012

End Sub


Private Sub txtIdNo_Change()
If txtidno.Text = "" Then
entries
End If
End Sub

Private Sub txtidno_KeyPress(KeyAscii As Integer)
txtidno.Text = Format(txtidno.Text, "00000")
End If
End Sub


Private Sub txtLname_GotFocus()

txtLname.BackColor = &HFFFFC0
txtLname.ForeColor = &HFF&

End Sub

Private Sub txtLname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFname.SetFocus
Else
If IsNumeric(Chr(KeyAscii)) Then
KeyAscii = 0

End If
End If

End Sub

Private Sub txtLname_LostFocus()
txtLname.Text = StrConv(txtLname, vbProperCase)
txtLname.BackColor = &HFFFFFF
txtLname.ForeColor = &H80000012

End Sub

Private Sub txtMname_GotFocus()
txtMname.BackColor = &HFFFFC0
txtMname.ForeColor = &HFF&
End Sub

Private Sub txtMname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAdd.SetFocus
Else
If IsNumeric(Chr(KeyAscii)) Then
KeyAscii = 0

End If
End If

End Sub

Private Sub txtMname_LostFocus()
txtMname.Text = StrConv(txtMname, vbProperCase)
txtMname.BackColor = &HFFFFFF
txtMname.ForeColor = &H80000012
End Sub


Private Sub txtmom_LostFocus()
txtmom.Text = StrConv(txtmom, vbProperCase)
End Sub


Private Sub unit_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Exit Sub
End If
If Not IsNumeric(Chr(KeyAscii)) Then
KeyAscii = 0
End If
End Sub



Sub id()

txtidno.Text = Format(txtidno.Text, "00000")

End Sub


Private Sub txtnamerelative_GotFocus()
txtnamerelative.BackColor = &HFFFFC0
txtnamerelative.ForeColor = &HFF&

End Sub

Private Sub txtnamerelative_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) Then
KeyAscii = 0

End If

End Sub

Private Sub txtnamerelative_LostFocus()
txtnamerelative.Text = StrConv(txtnamerelative, vbProperCase)
txtnamerelative.BackColor = &HFFFFFF
txtnamerelative.ForeColor = &H80000012

End Sub

Private Sub txtnationality_GotFocus()
txtnationality.BackColor = &HFFFFC0
txtnationality.ForeColor = &HFF&

End Sub

Private Sub txtnationality_LostFocus()
txtnationality.Text = StrConv(txtnationality, vbProperCase)
txtnationality.BackColor = &HFFFFFF
txtnationality.ForeColor = &H80000012

End Sub

Private Sub txtreligion_GotFocus()
txtreligion.BackColor = &HFFFFC0
txtreligion.ForeColor = &HFF&

End Sub

Private Sub txtreligion_LostFocus()
txtreligion.Text = StrConv(txtreligion, vbProperCase)
txtreligion.BackColor = &HFFFFFF
txtreligion.ForeColor = &H80000012

End Sub

Private Sub txtstate_GotFocus()
txtstate.BackColor = &HFFFFC0
txtstate.ForeColor = &HFF&

End Sub

Private Sub txtstate_LostFocus()
txtstate.Text = StrConv(txtstate, vbProperCase)
txtstate.BackColor = &HFFFFFF
txtstate.ForeColor = &H80000012

End Sub

Private Sub txttelrelative_GotFocus()
txttelrelative.BackColor = &HFFFFC0
txttelrelative.ForeColor = &HFF&


End Sub

Private Sub txttelrelative_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
KeyAscii = 0
End If

End Sub

Private Sub txttelrelative_LostFocus()
txttelrelative.BackColor = &HFFFFFF
txttelrelative.ForeColor = &H80000012
End Sub

Private Sub txtzip_GotFocus()
txtzip.BackColor = &HFFFFC0
txtzip.ForeColor = &HFF&

End Sub

Private Sub txtzip_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
KeyAscii = 0
End If

End Sub

Private Sub txtzip_LostFocus()

txtzip.BackColor = &HFFFFFF
txtzip.ForeColor = &H80000012

End Sub


'------------------------ sub functions ---------------------------------'

Sub enable()

txtidno.Enabled = True
txtLname.Enabled = True
txtFname.Enabled = True
txtMname.Enabled = True
txtAdd.Enabled = True
txtcity.Enabled = True
cbostate.Enabled = True
DTBdate.Enabled = True
cboreligion.Enabled = True
cbonationality.Enabled = True
cbosex.Enabled = True
cbostatus.Enabled = True
txtzip.Enabled = True
txtnamerelative.Enabled = True
txttelrelative.Enabled = True
txtaddrelative.Enabled = True
txtcmobile.Enabled = True
txtctelephone.Enabled = True



End Sub



Sub disable()

txtLname.Enabled = False
txtFname.Enabled = False
txtMname.Enabled = False
txtAdd.Enabled = False
txtcity.Enabled = False
cbostate.Enabled = False
DTBdate.Enabled = False
cboreligion.Enabled = False
cbonationality.Enabled = False
cbosex.Enabled = False
cbostatus.Enabled = False
txtzip.Enabled = False
txtnamerelative.Enabled = False
txttelrelative.Enabled = False
txtaddrelative.Enabled = False
txtcmobile.Enabled = False
txtctelephone.Enabled = False
txtval.Locked = False
txtproof.Locked = False

End Sub







Sub op1()

Select Case txtval.Text
Case Is = "Student ID"
opstudentid.Value = True
txtval.Text = ""
Case Is = "Drivers License ID"
opdriveid.Value = True
txtval.Text = ""
Case Is = "SSS ID"
opsssid.Value = True
txtval.Text = ""
Case Is = txtval.Text
opothersid.Value = True
End Select
End Sub

Sub op2()

Select Case txtproof.Text
Case Is = "Electric Bill"
Opelectric.Value = True
txtproof.Text = ""
Case Is = "Water Bill"
Opwater.Value = True
txtproof.Text = ""
Case Is = "Telphone Bill"
Optel.Value = True
txtproof.Text = ""
Case Is = txtproof.Text
Opothersproof.Value = True

End Select

End Sub

Sub enable_option()
opstudentid.Enabled = True
opdriveid.Enabled = True
opsssid.Enabled = True
opothersid.Enabled = True
Opwater.Enabled = True
Opelectric.Enabled = True
Optel.Enabled = True
Opothersproof.Enabled = True



End Sub


Sub entries()

txtidno.Text = ""
txtLname.Text = ""
txtFname.Text = ""
txtMname.Text = ""
txtAdd.Text = ""
lblage.Caption = ""
txtcity.Text = ""
cbostate.Text = ""

lbldatereg.Caption = ""
cboreligion.Text = ""
cbonationality.Text = ""
cbosex.Text = ""
cbostatus.Text = ""
txtzip.Text = ""
txtnamerelative.Text = ""
txttelrelative.Text = ""
txtaddrelative.Text = ""
txtcmobile.Text = ""
txtctelephone.Text = ""
txtval.Text = ""
txtproof.Text = ""
txtamount.Text = ""
txtnname.Text = ""
txtval.Text = ""

opstudentid.Value = False
opdriveid.Value = False
opsssid.Value = False
opothersid.Value = False
Opwater.Value = False
Opelectric.Value = False
Optel.Value = False
Opothersproof.Value = False

End Sub





Sub edz()
  On Error Resume Next
  With frmaddmembership
    If .txtPictureName.Text = "" Then
      Call nopicture
      GoTo edz:
    Else
edz:
      Call Images
      Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb ")
      Set rs = db.OpenRecordset("Select *from customers")
        rs!picfilename = .txtPictureName
        rs.Fields("imagepic").AppendChunk BImg
      
    End If
  End With
End Sub

'code used if borrower has no picture. found on frmaddmembership.frm
Sub nopicture()
  With frmaddmembership
    .cd2.InitDir = App.Path & "\image"
    .cd2.FileName = App.Path & "\image\temp.jpg"
      If .cd2.FileName <> "" Then
        strImgN = .cd2.FileName
        .txtPictureName.Text = "No Picture"
        .imagestatus.Picture = LoadPicture(.cd2.FileName)
      End If
  End With
End Sub

'code for image
Sub Images()
  On Error Resume Next
  Dim IntNum As Integer
  IntNum = FreeFile
  Open strImgN For Binary As #IntNum
  ReDim BImg(FileLen(strImgN))
  Get #IntNum, , BImg
  Close #1
End Sub

'code for loading image on frmmembership.frm
Sub LoadImages()
  On Error Resume Next
  Dim ImgS As Long
  Dim OS As Long
  Dim TmpPic As String
  Const conCS = 100
  TmpPic = App.Path & "\Images\tmpPic.bmp"
    If Len(Dir(TmpPic)) > 0 Then
      Kill TmpPic
    End If
      Dim F As Integer
      F = FreeFile
      Open App.Path & "\Images\tmpPic.bmp" For Binary As #F
      ImgS = rs.Fields("imagepic").ActualSize
      Do While OS < ImgS
        BImg() = rs _
        ("imagepic").GetChunk(conCS)
        Put #F, , BImg
        OS = OS + conCS
      Loop
        Close #F
        frmmembership.imagestatus.Picture = LoadPicture(App.Path & "\Images\tmpPic.bmp")
        Kill App.Path & "\Images\tmpPic.bmp"
End Sub

