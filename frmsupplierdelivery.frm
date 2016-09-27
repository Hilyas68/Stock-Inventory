VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmsupplierdelivery 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Supplier Delivery"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&CANCEL"
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
      Left            =   6120
      TabIndex        =   46
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdfindbydrnumber 
      Caption         =   "FIND"
      Height          =   375
      Left            =   4560
      TabIndex        =   26
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton CMDSAVE 
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
      Left            =   2280
      TabIndex        =   25
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton CMDFINDbydate 
      Caption         =   "FIND"
      Height          =   375
      Left            =   4560
      TabIndex        =   24
      Top             =   1200
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
      Left            =   8040
      TabIndex        =   13
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
      Left            =   4200
      TabIndex        =   12
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
      Left            =   360
      TabIndex        =   11
      Top             =   9120
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4095
      Left            =   480
      TabIndex        =   10
      Top             =   3840
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7223
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
      NumItems        =   11
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
         Text            =   "Supplier Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DR Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Article"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cost Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Total Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Checked By"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox txttotalamount 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ComboBox cbocheckedby 
      Height          =   315
      Left            =   8160
      TabIndex        =   9
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txtcostprice 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ComboBox cboarticle 
      Height          =   315
      Left            =   8160
      TabIndex        =   6
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtunit 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtdrnumber 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ComboBox cbosuppliername 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker dtdate 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "01\09\2007"
      Format          =   58327041
      CurrentDate     =   39091
   End
   Begin VB.TextBox TXTID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtarticleqty 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtarticleid 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtcriticallevel 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblongoing 
      Caption         =   "on going"
      Height          =   375
      Left            =   5400
      TabIndex        =   51
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblneedtopurchase 
      Caption         =   "need to order"
      Height          =   375
      Left            =   3480
      TabIndex        =   50
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3135
      Left            =   360
      Top             =   360
      Width           =   14535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   4455
      Left            =   360
      Top             =   3720
      Width           =   14535
   End
   Begin VB.Label lblyear 
      Height          =   375
      Left            =   3840
      TabIndex        =   48
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label28 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   6240
      TabIndex        =   47
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label27 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   45
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label26 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   44
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Height          =   3975
      Left            =   600
      TabIndex        =   43
      Top             =   4080
      Width           =   14175
   End
   Begin VB.Label Label24 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   8160
      TabIndex        =   42
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label23 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   4320
      TabIndex        =   41
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label22 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   2400
      TabIndex        =   40
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label21 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   480
      TabIndex        =   39
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   8280
      TabIndex        =   38
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   8280
      TabIndex        =   37
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   8280
      TabIndex        =   36
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   8280
      TabIndex        =   35
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   8280
      TabIndex        =   34
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   31
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   30
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Checked By"
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
      TabIndex        =   23
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
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
      Left            =   6480
      TabIndex        =   22
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Price"
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
      TabIndex        =   21
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
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
      Left            =   6480
      TabIndex        =   20
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
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
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
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
      Left            =   480
      TabIndex        =   18
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DR Number"
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
      Left            =   480
      TabIndex        =   17
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name"
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
      Left            =   480
      TabIndex        =   16
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
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
      Left            =   480
      TabIndex        =   15
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmsupplierdelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ssql As String
Dim toggle As Integer
Dim X As Variant
Dim COMPUTE As Double

Private Sub cboarticle_Change()
Set rs = conn.Execute("SELECT * FROM PRODUCT_TABLE WHERE ARTICLE='" & cboarticle.Text & "'")
If Not rs.EOF Then
txtcostprice.Text = rs!cost_price
txtarticleqty.Text = rs!qty
txtarticleid.Text = rs!ID
txtcriticallevel.Text = rs!critical_level
End If
X = Val(txtquantity.Text) * Val(txtcostprice.Text)
txttotalAmount.Text = X
End Sub

Private Sub cboarticle_Click()
Set rs = conn.Execute("SELECT * FROM PRODUCT_TABLE WHERE ARTICLE='" & cboarticle.Text & "'")
If Not rs.EOF Then
txtcostprice.Text = rs!cost_price
txtarticleqty.Text = rs!qty
txtarticleid.Text = rs!ID
txtcriticallevel.Text = rs!critical_level
End If
X = Val(txtquantity.Text) * Val(txtcostprice.Text)
txttotalAmount.Text = X
End Sub

Private Sub cmdadd_Click()
Call txtenabled(True, True, True, True, True, True, True)
Call CMDENABLED(False, False, True, True)
Call txtclear
End Sub

Private Sub cmdcancel_Click()
lblyear.Caption = Format(Now, "yyyy")
Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub CMDCLOSE_Click()
frmsupplierdelivery.Hide
Call listall
lblyear.Caption = Format(Now, "yyyy")
Call txtclear
Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(True, False, False, False)
End Sub

Private Sub cmddelete_Click()
X = MsgBox("Are you sure you want to delete this/these item(s)?", vbYesNo + vbCritical, "Confirmation")
If X = vbYes Then
ssql = "delete * from supplier_delivery_table where id=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been deleted!", vbExclamation, "Confirmation"
Call listall
Call txtclear
Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(True, False, False, False)
End If
End Sub

Private Sub cmdfindbydate_Click()
Set rs = conn.Execute("select * from supplier_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
TXTID.Text = rs!ID
dtdate.Value = rs!Dated_on
cbocheckedby.Text = rs!checked_by
cbosuppliername.Text = rs!supplier_name
txtdrnumber.Text = rs!dr_number
txtquantity.Text = rs!qty
txtunit.Text = rs!unit
cboarticle.Text = rs!article
txtcostprice.Text = rs!cost_price

Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(False, True, False, True)
Call listbydate
Else
MsgBox "No Result found!", vbExclamation, "Confirmation"
Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub CMDFINDbydate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set rs = conn.Execute("select * from supplier_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cbocheckedby.Text = rs!checked_by
End If
Set rs = conn.Execute("select * from supplier_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cbosuppliername.Text = rs!supplier_name
End If
Set rs = conn.Execute("select * from supplier_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cboarticle.Text = rs!article
End If
End Sub

Private Sub cmdfindbydrnumber_Click()
Set rs = conn.Execute("select * from supplier_delivery_table where dr_number='" & txtdrnumber.Text & "'")
If Not rs.EOF Then
TXTID.Text = rs!ID
dtdate.Value = rs!Dated_on

txtdrnumber.Text = rs!dr_number
txtquantity.Text = rs!qty
txtunit.Text = rs!unit
txtcostprice.Text = rs!cost_price

Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(False, True, False, True)
Call listbydrnumber
Else
MsgBox "No Result found!", vbExclamation, "Confirmation"
Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub cmdfindbydrnumber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set rs = conn.Execute("select * from supplier_delivery_table where dr_number='" & txtdrnumber.Text & "'")
If Not rs.EOF Then
cbocheckedby.Text = rs!checked_by
End If
Set rs = conn.Execute("select * from supplier_delivery_table where dr_number='" & txtdrnumber.Text & "'")
If Not rs.EOF Then
cbosuppliername.Text = rs!supplier_name
End If
Set rs = conn.Execute("select * from supplier_delivery_table where dr_number='" & txtdrnumber.Text & "'")
If Not rs.EOF Then
cboarticle.Text = rs!article
End If
End Sub

Private Sub cmdsave_Click()
ssql = "insert into supplier_delivery_table(year_today,dated_on,supplier_name,dr_number,qty,unit,article,cost_price,total_amount,checked_by) values('" & lblyear.Caption & "','" & dtdate.Value & _
"','" & cbosuppliername.Text & "','" & txtdrnumber.Text & "','" & txtquantity.Text & "','" & txtunit & "','" & cboarticle.Text & "','" & txtcostprice.Text & _
"','" & txttotalAmount.Text & "','" & cbocheckedby.Text & "')"
conn.Execute (ssql)
MsgBox "New Record has been saved", vbInformation, "Confirmation"

COMPUTE = Val(txtquantity.Text) + Val(txtarticleqty.Text)
txtarticleqty.Text = COMPUTE
ssql = "update product_table set qty='" & txtarticleqty.Text & "' where id=" & txtarticleid.Text & ""
conn.Execute (ssql)
MsgBox "Product Quantity has been updated", vbInformation, "Confirmation"

If Val(txtarticleqty.Text) < Val(txtcriticallevel.Text) Then
ssql = "update product_table set status='" & lblneedtopurchase.Caption & "' where id=" & txtarticleid.Text & ""
conn.Execute (ssql)
Else
ssql = "update product_table set status='" & lblongoing.Caption & "' where id=" & txtarticleid.Text & ""
conn.Execute (ssql)
End If

Call listall
Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
End Sub


Private Sub Form_Load()
On Error Resume Next
lblyear.Caption = Format(Now, "yyyy")
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
Call txtenabled(True, False, True, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
cbosuppliername.Clear
Set rs = conn.Execute("select * from supplier_table order by supplier_name")
Do While Not rs.EOF
With cbosuppliername
.AddItem rs!supplier_name
End With
rs.MoveNext
Loop

cboarticle.Clear
Set rs = conn.Execute("select * from product_table order by article")
Do While Not rs.EOF
With cboarticle
.AddItem rs!article
End With
rs.MoveNext
Loop

cbocheckedby.Clear
Set rs = conn.Execute("select * from EMPLOYEES_TABLE order by EMPLOYEES_NAME")
Do While Not rs.EOF
With cbocheckedby
.AddItem rs!EMPLOYEES_NAME
End With
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
cbosuppliername.Text = ListView1.ListItems.Item(col).SubItems(3)
txtdrnumber.Text = ListView1.ListItems.Item(col).SubItems(4)
txtquantity.Text = ListView1.ListItems.Item(col).SubItems(5)
txtunit.Text = ListView1.ListItems.Item(col).SubItems(6)
cboarticle.Text = ListView1.ListItems.Item(col).SubItems(7)
txtcostprice.Text = ListView1.ListItems.Item(col).SubItems(8)
txttotalAmount.Text = ListView1.ListItems.Item(col).SubItems(9)
cbocheckedby.Text = ListView1.ListItems.Item(col).SubItems(10)
Call CMDENABLED(False, True, False, True)
Call txtenabled(True, False, True, False, False, False, False)
End Sub

Private Sub txtquantity_Change()
X = Val(txtquantity.Text) * Val(txtcostprice.Text)
txttotalAmount.Text = X
End Sub

Private Sub txtenabled(xdate, XSUPPLIERNAME, XDRNUMBER, xquantity, XUNIT, xarticle, XCHECKEDBY)
dtdate.Enabled = xdate
cbosuppliername.Enabled = XSUPPLIERNAME
txtdrnumber.Enabled = XDRNUMBER
txtquantity.Enabled = xquantity
txtunit.Enabled = XUNIT
cboarticle.Enabled = xarticle
cbocheckedby.Enabled = XCHECKEDBY
End Sub
Private Sub CMDENABLED(XADD, XDELETE, XSAVE, XCANCEL)
cmdadd.Enabled = XADD
cmddelete.Enabled = XDELETE
cmdsave.Enabled = XSAVE
cmdcancel.Enabled = XCANCEL
End Sub

Private Sub txtclear()
cbosuppliername.Text = ""
txtdrnumber.Text = ""
txtquantity.Text = ""
txtunit.Text = ""
cboarticle.Text = ""
cbocheckedby.Text = ""
txtcostprice.Text = ""
txttotalAmount.Text = ""
End Sub
Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from supplier_delivery_table where year_today='" & lblyear.Caption & "' order BY dr_number")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
        X.SubItems(1) = rs.Fields!year_today
        X.SubItems(2) = rs.Fields!Dated_on
        X.SubItems(3) = rs.Fields!supplier_name
        X.SubItems(4) = rs.Fields!dr_number
        X.SubItems(5) = rs.Fields!qty
        X.SubItems(6) = rs.Fields!unit
        X.SubItems(7) = rs.Fields!article
        X.SubItems(8) = rs.Fields!cost_price
        X.SubItems(9) = rs.Fields!total_amount
        X.SubItems(10) = rs.Fields!checked_by
    rs.MoveNext
Loop
End Sub
Private Sub listbydate()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from supplier_delivery_table  where dated_on='" & dtdate.Value & "' order by dr_number")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
        X.SubItems(1) = rs.Fields!year_today
        X.SubItems(2) = rs.Fields!Dated_on
        X.SubItems(3) = rs.Fields!supplier_name
        X.SubItems(4) = rs.Fields!dr_number
        X.SubItems(5) = rs.Fields!qty
        X.SubItems(6) = rs.Fields!unit
        X.SubItems(7) = rs.Fields!article
        X.SubItems(8) = rs.Fields!cost_price
        X.SubItems(9) = rs.Fields!total_amount
        X.SubItems(10) = rs.Fields!checked_by
    rs.MoveNext
Loop
End Sub
Private Sub listbydrnumber()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from supplier_delivery_table  where dr_number='" & txtdrnumber.Text & "' order by dr_number")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
        X.SubItems(1) = rs.Fields!year_today
        X.SubItems(2) = rs.Fields!Dated_on
        X.SubItems(3) = rs.Fields!supplier_name
        X.SubItems(4) = rs.Fields!dr_number
        X.SubItems(5) = rs.Fields!qty
        X.SubItems(6) = rs.Fields!unit
        X.SubItems(7) = rs.Fields!article
        X.SubItems(8) = rs.Fields!cost_price
        X.SubItems(9) = rs.Fields!total_amount
        X.SubItems(10) = rs.Fields!checked_by
    rs.MoveNext
Loop
End Sub
