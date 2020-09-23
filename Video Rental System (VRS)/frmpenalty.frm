VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmpenalty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Rate"
   ClientHeight    =   2565
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmpenalty.frx":0000
   ScaleHeight     =   2565
   ScaleWidth      =   2565
   Begin VB.TextBox txtpenalty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtpen 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   1965
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   930
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2175
      Top             =   2955
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
      RecordSource    =   "penalty"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmpenalty.frx":17F802
      Height          =   480
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   847
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "penalty"
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
      SplitCount      =   1
      BeginProperty Split0 
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Curent Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   1920
      Left            =   105
      Top             =   135
      Width           =   2295
   End
End
Attribute VB_Name = "frmpenalty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
Unload Me

End Sub

Private Sub cmdsave_Click()

Select Case Me.Caption
Case Is = "Penalty Rate"
    If txtpen.Text = "" Then
    MsgBox ("Enter New Amount"), vbCritical, "Missing Data"
    Else
    Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
    Set rs = db.OpenRecordset("Select *from penalty where penalty = '" & txtpen & "'")
        rs.Edit
        rs!penalty = txtpenalty.Text
        rs.Update
        rs.Close
        db.Close
        Adodc1.Refresh
        DataGrid1.Refresh
        txtpenalty.Text = ""
            Unload Me
            MsgBox ("Successfuly Updated"), vbInformation, "Success"
            
    End If

Case Is = "Rental Limits"
    If txtpen.Text = "" Then
    MsgBox ("Enter New Limits"), vbCritical, "Missing Data"
    Else
    Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
    Set rs = db.OpenRecordset("Select *from penalty where rent_limits = '" & txtpen & "'")
        rs.Edit
        rs!rent_limits = txtpenalty.Text
        rs.Update
        rs.Close
        db.Close
        Adodc1.Refresh
        DataGrid1.Refresh
        txtpenalty.Text = ""
        Unload Me
        MsgBox ("Successfuly Updated"), vbInformation, "Success"
    End If

Case Is = "Membership Fee"
    
    If txtpen.Text = "" Then
    MsgBox ("Enter New Limits"), vbCritical, "Missing Data"
    Else
    Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
    Set rs = db.OpenRecordset("Select *from penalty where mem_fee = '" & txtpen & "'")
        rs.Edit
        rs!mem_fee = txtpenalty.Text
        rs.Update
        rs.Close
        db.Close
        Adodc1.Refresh
        DataGrid1.Refresh
        txtpenalty.Text = ""
        Unload Me
        MsgBox ("Successfuly Updated"), vbInformation, "Success"
    End If



End Select

End Sub

Private Sub Form_Load()
    
    Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
    Set rs = db.OpenRecordset("Select *from penalty")
        
Select Case frmpenalty.Caption
Case Is = "Penalty Rate"
            txtpen.Text = rs!penalty

Case Is = "Rental Limits"
            txtpen.Text = rs!rent_limits

Case Is = "Membership Fee"
    txtpen.Text = rs!mem_fee
    
    
End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdifrmmain.lblstatus.Caption = "Penaly Rate....... Change the New Amount for penalty"
End Sub

Private Sub txtpenalty_GotFocus()
Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
    Set rs = db.OpenRecordset("Select *from penalty")
        
Select Case frmpenalty.Caption
Case Is = "Penalty Rate"
            txtpen.Text = rs!penalty

Case Is = "Rental Limits"
            Label1.Caption = "Current Limits"
            Label2.Caption = "New Limits"
            txtpen.Text = rs!rent_limits

Case Is = "Membership Fee"
        Label1.Caption = "Current Fee"
        Label2.Caption = "New Fee"
        txtpen.Text = rs!mem_fee
        

End Select
End Sub

Private Sub txtpenalty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdsave_Click
End If

 If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 13) Then
        KeyAscii = 0
    End If
    
End Sub
