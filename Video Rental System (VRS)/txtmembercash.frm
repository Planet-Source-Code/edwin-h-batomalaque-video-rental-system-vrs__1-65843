VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmmembercash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VRS Cash"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "txtmembercash.frx":0000
   ScaleHeight     =   3480
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblamount 
      Alignment       =   2  'Center
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
      Height          =   420
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   930
      Width           =   2010
   End
   Begin TabDlg.SSTab sstabidno 
      Height          =   435
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   767
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   706
      TabCaption(0)   =   " ID #"
      TabPicture(0)   =   "txtmembercash.frx":DDD0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblidno"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Label lblidno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1110
         TabIndex        =   8
         Top             =   150
         Width           =   45
      End
   End
   Begin VB.CommandButton cmddiscount 
      Caption         =   "Change"
      Height          =   435
      Left            =   2400
      TabIndex        =   4
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox txtcash 
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
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   2130
   End
   Begin VB.CommandButton cmdprint 
      Height          =   450
      Left            =   2325
      TabIndex        =   3
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label lbltrans 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFF00&
      Height          =   270
      Left            =   720
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trans #"
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
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   675
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      Height          =   615
      Left            =   120
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label lblchange 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   2160
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   5
      Top             =   1755
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   600
      Left            =   120
      Top             =   1575
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lblborrower 
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblitemscount 
      Height          =   345
      Left            =   1440
      TabIndex        =   12
      Top             =   2430
      Width           =   735
   End
End
Attribute VB_Name = "frmmembercash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmddiscount_Click()
lblamount.Locked = False
lblamount.SetFocus
End Sub

Private Sub cmddiscount_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 13) Then
KeyAscii = 0
End If
End Sub

Private Sub cmdprint_Click()

If txtcash.Text = "" Then
            MsgBox "Enter Cash Amount", vbInformation, "Cashier"
            txtcash.SetFocus
            SendKeys hl
            Exit Sub
        Else
                txtcash.Text = Format(txtcash.Text, "#,###0.00")
                lblchange.Caption = Val(txtcash.Text) - Val(lblamount.Text)
                lblchange.Caption = Format(lblchange.Caption, "#,###0.00")
                    If Val(txtcash.Text) < Val(lblamount.Text) Then
                        lblchange.Caption = ""
                        MsgBox "Warning! Insufficient Cash Amount", vbCritical, "Cashier"
                        txtcash.Text = ""
                        txtcash.SetFocus
                        SendKeys hl
                        Exit Sub
                    Else
                    cmdprint.SetFocus
                    End If
        End If
        


Prompt$ = "Do you want to print the receipts?"
    reply = MsgBox(Prompt$, vbOKCancel, "Processing Receipts")
    If reply = vbOK Then
    
Select Case frmmembercash.Caption
Case Is = "Membership Fee"
            If txtcash.Text = "" Then
                MsgBox "Enter Cash Amount", vbInformation, "Cashier"
                txtcash.SetFocus
                SendKeys hl
                Exit Sub
            Else
                    txtcash.Text = Format(txtcash.Text, "#,###0.00")
                    lblchange.Caption = Val(txtcash.Text) - Val(lblamount.Text)
                    lblchange.Caption = Format(lblchange.Caption, "#,###0.00")
                        If Val(txtcash.Text) < Val(lblamount.Text) Then
                            lblchange.Caption = ""
                            MsgBox "Warning! Insufficient Cash Amount", vbCritical, "Cashier"
                            txtcash.Text = ""
                            txtcash.SetFocus
                            SendKeys hl
                            Exit Sub
                        Else
                            
                            Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
                            Set rs = db.OpenRecordset("Select *from membership_fee where invoice_num  ='" & lblnum & "'")
                                     rs.AddNew
                                        rs!custnum = lblidno.Caption
                                        rs!invoice_num = lbltrans.Caption
                                        rs!custname = lblborrower.Caption
                                        rs!payments = lblamount.Text
                                        rs!day_paid = Date
                                        rs!cashier = mdifrmmain.StatusBar1.Panels(2)
                                        rs.Update
                                        rs.Close
                                        db.Close
                                        frmrent.adorented.Refresh
                                        frmrent.dbrented.Refresh
                                        frmrent.txtinvoice.Text = ""
                                        receipts.Refresh
                                        receipts.Show
                                        With receipts.Sections("Section2").Controls
                                            .item("lblid").Caption = "Membership Fee"
                                            .item("lbltitle").Caption = ""
                                            .item("lblformat").Caption = ""
                                            .item("lbldue").Caption = ""
                                            .item("items").Caption = ""
                                            .item("lblamount").Caption = ""
                                            End With
                                        
                                        
                            clear_entries
                            Unload Me
                            
                        End If
            End If

Case Is = "Rental Fees"
            If txtcash.Text = "" Then
                MsgBox "Enter Cash Amount", vbInformation, "Cashier"
                txtcash.SetFocus
                SendKeys hl
                Exit Sub
            Else
                    txtcash.Text = Format(txtcash.Text, "#,###0.00")
                    lblchange.Caption = Val(txtcash.Text) - Val(lblamount.Text)
                    lblchange.Caption = Format(lblchange.Caption, "#,###0.00")
                        If Val(txtcash.Text) < Val(lblamount.Text) Then
                            lblchange.Caption = ""
                            MsgBox "Warning! Insufficient Cash Amount", vbCritical, "Cashier"
                            txtcash.Text = ""
                            txtcash.SetFocus
                            SendKeys hl
                            Exit Sub
                        Else
                        
                            Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
                            Set rs = db.OpenRecordset("Select *from rented where invoice_num  ='" & txtinvoice & "'")
                                     rs.AddNew
                                        rs!custnum = frmrent.lblidno.Caption
                                        rs!invoice_num = frmrent.txtinvoice.Text
                                        rs!custname = frmrent.lblborrower.Caption
                                        rs!payments = frmrent.txtrunningamount.Text
                                        rs!tot_items = frmrent.lblitemsno.Caption
                                        rs!day_paid = Date
                                        rs!cashier = mdifrmmain.StatusBar1.Panels(2)
                                        rs.Update
                                        rs.Close
                                        db.Close
                                        frmrent.adorented.Refresh
                                        frmrent.dbrented.Refresh
                                        frmrent.txtinvoice.Text = ""
                                        
                                      receipts.Refresh
                                      receipts.Show
                                        
                                                                                  
                                          
                                        
                            clear_entries
                            Unload Me
                            frmrent.cmdnew.Enabled = True
                            frmrent.cmdprint.Enabled = False
                            frmrent.cmdreceipt.Enabled = False
                            frmrent.lblborrower.Caption = ""
                            frmrent.lblidno.Caption = ""
                            frmrent.txtrunningamount.Text = ""
                            frmrent.lblitemsno.Caption = ""
                            Call frmrent.ado
                            frmrent.cmdprint.Enabled = False
                            frmrent.cmdcan.Enabled = False
                            frmrent.cmdcancel.Enabled = False
                            frmrent.txtitemid.Enabled = False
                            frmrent.txttitle.Enabled = False
                            frmrent.txtstatus.Enabled = False
                            frmrent.txtrunningamount.Enabled = False
                            frmrent.cmdrent.Enabled = False
                            frmrent.cmdover.Enabled = False
                            
                            
                            
                            
                         
                        End If
            End If
         
          
Case Is = "Penalty Fee"
            If txtcash.Text = "" Then
                MsgBox "Enter Cash Amount", vbInformation, "Cashier"
                txtcash.SetFocus
                SendKeys hl
                Exit Sub
            Else
                    txtcash.Text = Format(txtcash.Text, "#,###0.00")
                    lblchange.Caption = Val(txtcash.Text) - Val(lblamount.Text)
                    lblchange.Caption = Format(lblchange.Caption, "#,###0.00")
                        If Val(txtcash.Text) < Val(lblamount.Text) Then
                            lblchange.Caption = ""
                            MsgBox "Warning! Insufficient Cash Amount", vbCritical, "Cashier"
                            txtcash.Text = ""
                            txtcash.SetFocus
                            SendKeys hl
                            Exit Sub
                        Else
                         
                         Set db = OpenDatabase(App.Path & "\VRS Database\vrs.mdb")
                            Set rs = db.OpenRecordset("Select *from membership_fee where invoice_num  ='" & txtinvoice & "'")
                                     rs.AddNew
                                        rs!custnum = frmreturn.lblidno.Caption
                                        rs!invoice_num = lbltrans.Caption
                                        rs!custname = frmreturn.lblborrower.Caption
                                        rs!payments = frmreturn.lbltot_penalty.Caption
                                        rs!day_paid = Date
                                        rs!cashier = mdifrmmain.StatusBar1.Panels(2)
                                        rs.Update
                                        rs.Close
                                        db.Close
                                        
                         receipts.Refresh
                         receipts.Show
                         With receipts.Sections("Section2").Controls
                         .item("lblamount").Caption = ""
                            End With
                                            
                                           
                                            
                         Unload Me
                         clear_entries
                         frmreturn.clear
                            
                        End If
            End If
  
     
End Select

Else
Unload Me
clear_entries
If frmreturn.Visible = True Then
 frmreturn.clear
ElseIf frmrent.Visible = True Then
frmrent.cmdnew.Enabled = True
                            frmrent.cmdprint.Enabled = False
                            frmrent.cmdreceipt.Enabled = False
                            frmrent.lblborrower.Caption = ""
                            frmrent.lblidno.Caption = ""
                            frmrent.txtrunningamount.Text = ""
                            frmrent.lblitemsno.Caption = ""
                            Call frmrent.ado
                            frmrent.cmdprint.Enabled = False
                            frmrent.cmdcan.Enabled = False
                            frmrent.cmdcancel.Enabled = True
                            frmrent.txtitemid.Enabled = False
                            frmrent.txttitle.Enabled = False
                            frmrent.txtstatus.Enabled = False
                            frmrent.txtrunningamount.Enabled = False
                            frmrent.cmdrent.Enabled = False
                            frmrent.cmdover.Enabled = False
Else
End If
End If


mdifrmmain.Enabled = True

End Sub


Private Sub Form_Load()
lbltrans.Caption = frmrent.txtinvoice.Text

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdifrmmain.lblstatus.Caption = "Enter your cash amount correctly....."
End Sub


Private Sub Form_Terminate()
mdifrmmain.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
End Sub

Private Sub lblamount_KeyPress(KeyAscii As Integer)

If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 13) Then
KeyAscii = 0
End If

End Sub

Private Sub txtcash_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 13) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then

    If txtcash.Text = "" Then
            MsgBox "Enter Cash Amount", vbInformation, "Cashier"
            txtcash.SetFocus
            SendKeys hl
            Exit Sub
        Else
                txtcash.Text = Format(txtcash.Text, "#,###0.00")
                lblchange.Caption = Val(txtcash.Text) - Val(lblamount.Text)
                lblchange.Caption = Format(lblchange.Caption, "#,###0.00")
                    If Val(txtcash.Text) < Val(lblamount.Text) Then
                        lblchange.Caption = ""
                        MsgBox "Warning! Insufficient Cash Amount", vbCritical, "Cashier"
                        txtcash.Text = ""
                        txtcash.SetFocus
                        SendKeys hl
                        Exit Sub
                    Else
                    cmdprint.SetFocus
                    End If
        End If
    
    
End If

End Sub


Sub clear_entries()
lblamount = ""
lblidno.Caption = ""
txtcash.Text = ""
lblchange.Caption = ""


End Sub


Private Sub txtitemid_Click()

End Sub
