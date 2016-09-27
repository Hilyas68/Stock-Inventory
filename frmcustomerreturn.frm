VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmcustomerreturn 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customer Return"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdfindbysalesret 
      Caption         =   "FIND"
      Height          =   375
      Left            =   5160
      TabIndex        =   40
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdfindbydate 
      Caption         =   "FIND"
      Height          =   375
      Left            =   5160
      TabIndex        =   38
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   14
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&DELETE"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   11
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "C&ANCEL"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   13
      Top             =   9120
      Width           =   1815
   End
   Begin VB.TextBox txtreference 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9360
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txttotalAmount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox cboarticle 
      Height          =   315
      ItemData        =   "frmcustomerreturn.frx":0000
      Left            =   9360
      List            =   "frmcustomerreturn.frx":0002
      TabIndex        =   5
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtsalesret 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.ComboBox cbocustomername 
      Height          =   315
      ItemData        =   "frmcustomerreturn.frx":0004
      Left            =   2760
      List            =   "frmcustomerreturn.frx":0006
      TabIndex        =   2
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox TXTID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtdate 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   55443457
      CurrentDate     =   39091
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   840
      TabIndex        =   9
      Top             =   4080
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Year"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Customer Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sales Return No"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Quantity"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Article"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Total Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Reference DR/PS"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   4095
      Left            =   720
      Top             =   3960
      Width           =   13935
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      Height          =   3735
      Left            =   960
      TabIndex        =   43
      Top             =   4200
      Width           =   13575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3255
      Left            =   720
      Top             =   600
      Width           =   13935
   End
   Begin VB.Label lblyear 
      Height          =   495
      Left            =   6000
      TabIndex        =   42
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   5280
      TabIndex        =   41
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label38 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   5280
      TabIndex        =   39
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label26 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   840
      TabIndex        =   37
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label27 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   2760
      TabIndex        =   36
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label28 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   8520
      TabIndex        =   35
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label37 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   4680
      TabIndex        =   34
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label42 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   6600
      TabIndex        =   33
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   9480
      TabIndex        =   32
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference DR/RS"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   31
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   9480
      TabIndex        =   30
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   29
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   9480
      TabIndex        =   28
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   27
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   9480
      TabIndex        =   26
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Article"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   25
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   23
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Return No"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   21
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   20
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label41 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frmcustomerreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim toggle As Integer
Dim ssql  As String
Dim x As Variant

Private Sub cboarticle_Change()
Set rs = conn.Execute("select * from incoming_delivery_table where description='" & cboarticle.Text & "'")
If Not rs.EOF Then
txtamount.Text = rs!unit_price
txttotalAmount.Text = rs!total_price
txtquantity.Text = rs!quantity
Call txtenabled(True, False, True, False, True, False, False, True)
End If
End Sub

Private Sub cboarticle_Click()
Set rs = conn.Execute("select * from incoming_delivery_table where description='" & cboarticle.Text & "'")
If Not rs.EOF Then
txtamount.Text = rs!unit_price
txttotalAmount.Text = rs!total_price
txtquantity.Text = rs!quantity
Call txtenabled(True, False, True, False, True, False, False, True)
End If
End Sub

Private Sub cbocustomername_Change()
On Error Resume Next
cboarticle.Clear
Set rs = conn.Execute("select * from incoming_delivery_table where customer_name='" & cbocustomername.Text & "' and year_today='" & lblyear.Caption & "'")
Do While Not rs.EOF
With cboarticle
.AddItem rs!Description
End With
rs.MoveNext
Loop
End Sub

Private Sub cbocustomername_Click()
cboarticle.Clear
Set rs = conn.Execute("select * from incoming_delivery_table where customer_name='" & cbocustomername.Text & "' and year_today='" & lblyear.Caption & "'")
Do While Not rs.EOF
With cboarticle
.AddItem rs!Description
End With
rs.MoveNext
Loop
End Sub

Private Sub cmdadd_Click()
Call txtenabled(True, True, True, True, True, True, True, True)
Call CMDENABLED(False, False, True, True)
End Sub

Private Sub cmdcancel_Click()
lblyear.Caption = Format(Now, "yyyy")
Call txtenabled(True, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call listall
Call txtclear
End Sub

Private Sub CMDCLOSE_Click()
lblyear.Caption = Format(Now, "yyyy")
Call txtenabled(True, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call listall
frmcustomerreturn.Hide
End Sub

Private Sub cmddelete_Click()
x = MsgBox("Are you sure you want to delete this/these item(s)?", vbYesNo + vbCritical, "Confirmation")
If x = vbYes Then
ssql = "delete * from customer_return_table where id=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been deleted!", vbExclamation, "Confirmation"
Call txtenabled(True, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub cmdfindbydate_Click()
Set rs = conn.Execute("select * from customer_return_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
TXTID.Text = rs!ID
txtsalesret.Text = rs!Sales_Ret_No
txtquantity.Text = rs!qty
txtamount.Text = rs!amount
txttotalAmount.Text = rs!total_amount
txtreference.Text = rs!reference_dr_ps
Call txtenabled(True, True, True, False, False, False, False, False)
Call CMDENABLED(False, True, False, True)
Call listbydate
Else
MsgBox "No Result found!", vbExclamation, "Confirmation"
Call txtenabled(True, True, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
End If

End Sub

Private Sub CMDFINDbydate_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Set rs = conn.Execute("select * from customer_return_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cbocustomername.Text = rs!customer_name
End If
Set rs = conn.Execute("select * from customer_return_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cboarticle.Text = rs!article
End If
End Sub

Private Sub cmdfindbysalesret_Click()
Set rs = conn.Execute("select * from customer_return_table where sales_ret_no='" & txtsalesret.Text & "'")
If Not rs.EOF Then
TXTID.Text = rs!ID
txtsalesret.Text = rs!Sales_Ret_No
txtquantity.Text = rs!qty
txtamount.Text = rs!amount
txttotalAmount.Text = rs!total_amount
txtreference.Text = rs!reference_dr_ps
Call txtenabled(True, True, True, False, False, False, False, False)
Call CMDENABLED(False, True, False, True)
Call listbysalesret
Else
MsgBox "No Result found!", vbExclamation, "Confirmation"
Call txtenabled(True, True, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
End If

End Sub

Private Sub cmdfindbysalesret_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Set rs = conn.Execute("select * from customer_return_table where sales_ret_no='" & txtsalesret.Text & "'")
If Not rs.EOF Then
cbocustomername.Text = rs!customer_name
End If
Set rs = conn.Execute("select * from customer_return_table where sales_ret_no='" & txtsalesret.Text & "'")
If Not rs.EOF Then
cboarticle.Text = rs!article
End If
End Sub

Private Sub cmdsave_Click()
ssql = "insert into customer_return_table(year_today,dated_on,customer_name,sales_ret_no,qty,article,amount,total_amount,reference_dr_ps)values('" & lblyear.Caption & "','" & dtdate.Value & _
"','" & cbocustomername.Text & "','" & txtsalesret.Text & "','" & txtquantity.Text & "','" & cboarticle.Text & "','" & txtamount.Text & "','" & txttotalAmount.Text & "','" & txtreference.Text & "')"
conn.Execute (ssql)
MsgBox "New Record has been saved", vbInformation, "Confirmation"
Call listall
Call txtclear
Call txtenabled(True, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
End Sub

Private Sub Form_Load()
On Error Resume Next
lblyear.Caption = Format(Now, "yyyy")
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
Call txtenabled(True, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call listall
Call txtclear

Set rs = conn.Execute("SELECT * FROM CUSTOMER_TABLE ORDER BY CUSTOMER_NAME")
Do While Not rs.EOF
With cbocustomername
.AddItem rs!customer_name
End With
rs.MoveNext
Loop

Set rs = conn.Execute("SELECT * FROM product_table ORDER BY article")
Do While Not rs.EOF
With cboarticle
.AddItem rs!article
End With
rs.MoveNext
Loop
End Sub

Private Sub txtenabled(xdate, xcustomername, xsalesret, xquantity, xarticle, xamount, xtotalamount, xreference)
dtdate.Enabled = xdate
cbocustomername.Enabled = xcustomername
txtsalesret.Enabled = xsalesret
txtquantity.Enabled = xquantity
cboarticle.Enabled = xarticle
txtamount.Enabled = xamount
txttotalAmount.Enabled = xtotalamount
txtreference.Enabled = xreference
End Sub

Private Sub CMDENABLED(XADD, XDELETE, XSAVE, XCANCEL)
cmdadd.Enabled = XADD
cmddelete.Enabled = XDELETE
cmdsave.Enabled = XSAVE
cmdcancel.Enabled = XCANCEL
End Sub

Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from customer_return_table  where year_today ='" & lblyear.Caption & "' ORDER BY sales_ret_no")
Do While Not rs.EOF
    Set x = ListView1.ListItems.Add(, , rs.Fields(0))
  x.SubItems(1) = rs.Fields!year_today
        x.SubItems(2) = rs.Fields!Dated_on
        x.SubItems(3) = rs.Fields!customer_name
        x.SubItems(4) = rs.Fields!Sales_Ret_No
        x.SubItems(5) = rs.Fields!qty
        x.SubItems(6) = rs.Fields!article
        x.SubItems(7) = rs.Fields!amount
        x.SubItems(8) = rs.Fields!total_amount
        x.SubItems(9) = rs.Fields!reference_dr_ps
    rs.MoveNext
Loop
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim col
col = ListView1.SelectedItem.Index
TXTID.Text = ListView1.ListItems.Item(col).Text
lblyear.Caption = ListView1.ListItems.Item(col).SubItems(1)
dtdate.Value = ListView1.ListItems.Item(col).SubItems(2)
cbocustomername.Text = ListView1.ListItems.Item(col).SubItems(3)
txtsalesret.Text = ListView1.ListItems.Item(col).SubItems(4)
txtquantity.Text = ListView1.ListItems.Item(col).SubItems(5)
cboarticle.Text = ListView1.ListItems.Item(col).SubItems(6)
txtamount.Text = ListView1.ListItems.Item(col).SubItems(7)
txttotalAmount.Text = ListView1.ListItems.Item(col).SubItems(8)
txtreference.Text = ListView1.ListItems.Item(col).SubItems(9)
Call CMDENABLED(False, True, False, True)
Call txtenabled(True, False, True, False, False, False, False, False)
End Sub
Private Sub txtclear()
cbocustomername.Text = ""
txtsalesret.Text = ""
txtquantity.Text = ""
cboarticle.Text = ""
txtamount.Text = ""
txttotalAmount.Text = ""
txtreference.Text = ""
End Sub

Private Sub listbydate()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from customer_return_table where dated_on ='" & dtdate.Value & "'ORDER BY sales_ret_no")
Do While Not rs.EOF
    Set x = ListView1.ListItems.Add(, , rs.Fields(0))
        x.SubItems(1) = rs.Fields!year_today
        x.SubItems(2) = rs.Fields!Dated_on
        x.SubItems(3) = rs.Fields!customer_name
        x.SubItems(4) = rs.Fields!Sales_Ret_No
        x.SubItems(5) = rs.Fields!qty
        x.SubItems(6) = rs.Fields!article
        x.SubItems(7) = rs.Fields!amount
        x.SubItems(8) = rs.Fields!total_amount
        x.SubItems(9) = rs.Fields!reference_dr_ps
    rs.MoveNext
Loop
End Sub
Private Sub listbysalesret()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from customer_return_table where sales_ret_no ='" & txtsalesret.Text & "' ORDER BY sales_ret_no")
Do While Not rs.EOF
    Set x = ListView1.ListItems.Add(, , rs.Fields(0))
         x.SubItems(1) = rs.Fields!year_today
        x.SubItems(2) = rs.Fields!Dated_on
        x.SubItems(3) = rs.Fields!customer_name
        x.SubItems(4) = rs.Fields!Sales_Ret_No
        x.SubItems(5) = rs.Fields!qty
        x.SubItems(6) = rs.Fields!article
        x.SubItems(7) = rs.Fields!amount
        x.SubItems(8) = rs.Fields!total_amount
        x.SubItems(9) = rs.Fields!reference_dr_ps
    rs.MoveNext
Loop
End Sub
