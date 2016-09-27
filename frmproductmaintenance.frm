VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmproductmaintenance 
   BackColor       =   &H0080C0FF&
   Caption         =   "Product Maintenance"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbosuppliername 
      Height          =   315
      Left            =   2640
      TabIndex        =   50
      Top             =   1440
      Width           =   4335
   End
   Begin VB.CommandButton cmdaddchoices 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdremovechoice 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox TXTID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   26
      Top             =   720
      Width           =   2055
   End
   Begin VB.ComboBox txtarticle 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   3000
      Width           =   4335
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
      Left            =   480
      TabIndex        =   10
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton CMDEDIT 
      Caption         =   "&EDIT"
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
      Left            =   2400
      TabIndex        =   11
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton CMDCLOSE 
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
      Left            =   12840
      TabIndex        =   16
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton CMDCANCEL 
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
      Left            =   10920
      TabIndex        =   15
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "&FIND"
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
      Left            =   9000
      TabIndex        =   14
      Top             =   8880
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
      Left            =   6240
      TabIndex        =   13
      Top             =   8880
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
      Left            =   4320
      TabIndex        =   12
      Top             =   8880
      Width           =   1815
   End
   Begin VB.TextBox txtcriticallevel 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9240
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtqty 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9240
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtsellingprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtcostprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtmktgprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox cboclass 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   840
      TabIndex        =   9
      Top             =   3960
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   6376
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
         Text            =   "Supplier_Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Classification"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Code"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Article"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Mktg Price"
         Object.Width           =   2549
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cost Price"
         Object.Width           =   2549
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Selling Price"
         Object.Width           =   2549
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Qty"
         Object.Width           =   2549
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Critical Level"
         Object.Width           =   2549
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Status"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtidclass 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   29
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblongoing 
      Caption         =   "on going"
      Height          =   375
      Left            =   2280
      TabIndex        =   51
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   3975
      Left            =   600
      Top             =   3840
      Width           =   14175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   3375
      Left            =   600
      Top             =   360
      Width           =   14175
   End
   Begin VB.Label Label30 
      BackColor       =   &H00004080&
      Height          =   3615
      Left            =   960
      TabIndex        =   49
      Top             =   4080
      Width           =   13575
   End
   Begin VB.Label Label29 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   5760
      TabIndex        =   48
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label28 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   5040
      TabIndex        =   47
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label27 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9360
      TabIndex        =   46
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label26 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9360
      TabIndex        =   45
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label25 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9360
      TabIndex        =   44
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label24 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9360
      TabIndex        =   43
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label23 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9360
      TabIndex        =   42
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label22 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   2760
      TabIndex        =   41
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label Label21 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   2760
      TabIndex        =   40
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label20 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   2760
      TabIndex        =   39
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label19 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   2760
      TabIndex        =   38
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label18 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   2760
      TabIndex        =   37
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   12960
      TabIndex        =   36
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   11040
      TabIndex        =   35
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   9120
      TabIndex        =   34
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   6360
      TabIndex        =   33
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   4440
      TabIndex        =   32
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   2520
      TabIndex        =   31
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   600
      TabIndex        =   30
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label10 
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
      Left            =   840
      TabIndex        =   25
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Critical Level"
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
      Left            =   7440
      TabIndex        =   24
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Left            =   7440
      TabIndex        =   23
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Price"
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
      Left            =   7440
      TabIndex        =   22
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label6 
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
      Height          =   495
      Left            =   7440
      TabIndex        =   21
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mktg Price"
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
      Left            =   7440
      TabIndex        =   20
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Height          =   495
      Left            =   840
      TabIndex        =   19
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   840
      TabIndex        =   18
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Classification"
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
      Left            =   840
      TabIndex        =   17
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "frmproductmaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ssql As String
Dim toggle As Integer
Dim X As Variant

Private Sub cboclass_Change()
Set rs = conn.Execute("select * from classification_table where classification='" & cboclass.Text & "'")
If Not rs.EOF Then
txtidclass.Text = rs!ID
End If
End Sub

Private Sub cboclass_Click()
Set rs = conn.Execute("select * from classification_table where classification='" & cboclass.Text & "'")
If Not rs.EOF Then
txtidclass.Text = rs!ID
End If
End Sub

Private Sub cmdadd_Click()
toggle = 0
txtarticle.SetFocus
Call CMDENABLED(False, False, False, True, True)
Call txtenabled(True, True, True, True, True, True, True, True, True)
Call txtclear
End Sub

Private Sub cmdaddchoices_Click()
If cboclass.Text = "" Then
MsgBox "You must write a new choice first!", vbExclamation, "Confirmation"
cboclass.Enabled = True
cboclass.SetFocus
Else
ssql = "insert into classificATION_TABLE(classification) values ('" & cboclass.Text & "')"
conn.Execute (ssql)
Call loadcombo
End If
End Sub

Private Sub cmdcancel_Click()
Call txtenabled(False, False, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub CMDCLOSE_Click()
frmproductmaintenance.Hide
Call txtenabled(False, False, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub cmddelete_Click()
X = MsgBox("Are you sure you want to delete this/these item(s)?", vbYesNo + vbCritical, "Confirmation")
If X = vbYes Then
ssql = "delete * from product_table where id=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been deleted!", vbExclamation, "Confirmation"
Call txtenabled(False, False, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call listall
Call txtclear
End If

txtarticle.Clear
Set rs = conn.Execute("select * from product_table order by article")
Do While Not rs.EOF
With txtarticle
.AddItem rs!article
End With
rs.MoveNext
Loop
Call loadcombo
End Sub

Private Sub CMDEDIT_Click()
toggle = 1
Call txtenabled(True, True, True, True, True, True, True, True, True)
Call CMDENABLED(False, False, False, True, True)
End Sub

Private Sub cmdfind_Click()

Set rs = conn.Execute("select * from product_table where article='" & txtarticle.Text & "'")
If Not rs.EOF Then
TXTID = rs!ID
cbosuppliername.Text = rs!supplier_name
txtcode.Text = rs!code
txtarticle.Text = rs!article
txtmktgprice.Text = rs!mktg_price
txtcostprice.Text = rs!cost_price
txtsellingprice.Text = rs!selling_price
txtqty.Text = rs!qty
txtcriticallevel.Text = rs!critical_level
Call CMDENABLED(False, True, True, False, False)
Call txtenabled(False, False, False, True, False, False, False, False, False)
Else
MsgBox "No Record Found", vbExclamation, "Confirmation"
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(False, False, False, True, False, False, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub cmdremovechoice_Click()
On Error Resume Next
ssql = "delete * from classification_table where id=" & txtidclass.Text & ""
conn.Execute (ssql)
Call loadcombo
End Sub

Private Sub cmdsave_Click()
Select Case toggle
Case 0:
ssql = "insert into product_table(supplier_name,classification,code,article,mktg_price,cost_price,selling_price,qty,critical_level,status) values ('" & cbosuppliername.Text & _
"','" & cboclass.Text & "','" & txtcode.Text & "','" & txtarticle.Text & "','" & txtmktgprice.Text & "','" & txtcostprice.Text & "','" & txtsellingprice.Text & _
"', '" & txtqty.Text & "','" & txtcriticallevel.Text & "','" & "on going" & "')"
conn.Execute (ssql)
MsgBox "New Record has been saved", vbInformation, "Confirmation"

Case 1:
ssql = "update product_table set supplier_name ='" & cbosuppliername.Text & "', classification='" & cboclass.Text & "', code='" & txtcode.Text & _
"',article='" & txtarticle.Text & "',mktg_price='" & txtmktgprice.Text & "',cost_price='" & txtcostprice.Text & "',selling_price='" & txtsellingprice.Text & "',qty='" & txtqty.Text & _
"',critical_level='" & txtcriticallevel.Text & "' where id=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been updated", vbInformation, "Confirmation"
End Select
Call txtenabled(False, False, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall

txtarticle.Clear
Set rs = conn.Execute("select * from product_table order by article")
Do While Not rs.EOF
With txtarticle
.AddItem rs!article
End With
rs.MoveNext
Loop
Call loadcombo

End Sub



Private Sub Form_Load()
On Error Resume Next
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
Call txtenabled(False, False, False, True, False, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
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

txtarticle.Clear
Set rs = conn.Execute("select * from product_table order by article")
Do While Not rs.EOF
With txtarticle
.AddItem rs!article
End With
rs.MoveNext
Loop
Call loadcombo
End Sub
Private Sub CMDENABLED(XADD, xedit, XDELETE, XSAVE, XCANCEL)
cmdadd.Enabled = XADD
CMDEDIT.Enabled = xedit
cmddelete.Enabled = XDELETE
cmdsave.Enabled = XSAVE
cmdcancel.Enabled = XCANCEL
End Sub
Private Sub txtenabled(XSUPPLIERNAME, XCLASSIFICATION, XCODE, xarticle, XMKTGPRICE, XCOSTPRICE, XSELLINGPRICE, XQTY, Xcriticallevel)
cbosuppliername.Enabled = XSUPPLIERNAME
cboclass.Enabled = XCLASSIFICATION
txtcode.Enabled = XCODE
txtarticle.Enabled = xarticle
txtmktgprice.Enabled = XMKTGPRICE
txtcostprice.Enabled = XCOSTPRICE
txtsellingprice.Enabled = XSELLINGPRICE
txtqty.Enabled = XQTY
txtcriticallevel.Enabled = Xcriticallevel
End Sub
Private Sub txtclear()
cbosuppliername.Text = ""
cboclass.Text = ""
txtcode.Text = ""
txtarticle.Text = ""
txtmktgprice.Text = ""
txtcostprice.Text = ""
txtsellingprice.Text = ""
txtqty.Text = ""
txtcriticallevel.Text = ""
End Sub

Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from product_table ORDER BY article")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
        X.SubItems(1) = rs.Fields!supplier_name
        X.SubItems(2) = rs.Fields!classification
        X.SubItems(3) = rs.Fields!code
        X.SubItems(4) = rs.Fields!article
        X.SubItems(5) = rs.Fields!mktg_price
        X.SubItems(6) = rs.Fields!cost_price
        X.SubItems(7) = rs.Fields!selling_price
        X.SubItems(8) = rs.Fields!qty
        X.SubItems(9) = rs.Fields!critical_level
        X.SubItems(10) = rs.Fields!Status
    rs.MoveNext
Loop
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim col
col = ListView1.SelectedItem.Index
TXTID.Text = ListView1.ListItems.Item(col).Text
cbosuppliername.Text = ListView1.ListItems.Item(col).SubItems(1)
cboclass.Text = ListView1.ListItems.Item(col).SubItems(2)
txtcode.Text = ListView1.ListItems.Item(col).SubItems(3)
txtarticle.Text = ListView1.ListItems.Item(col).SubItems(4)
txtmktgprice.Text = ListView1.ListItems.Item(col).SubItems(5)
txtcostprice.Text = ListView1.ListItems.Item(col).SubItems(6)
txtsellingprice.Text = ListView1.ListItems.Item(col).SubItems(7)
txtqty.Text = ListView1.ListItems.Item(col).SubItems(8)
txtcriticallevel.Text = ListView1.ListItems.Item(col).SubItems(9)
lblongoing.Caption = ListView1.ListItems.Item(col).SubItems(10)
Call CMDENABLED(False, True, True, False, True)
Call txtenabled(False, False, False, True, False, False, False, False, False)
End Sub

Private Sub txtarticle_Change()
Set rs = conn.Execute("select * from product_table where article='" & txtarticle.Text & "'")
If Not rs.EOF Then
cbosuppliername.Text = rs!supplier_name
cboclass.Text = rs!classification
End If
End Sub

Private Sub txtarticle_Click()
Set rs = conn.Execute("select * from product_table where article='" & txtarticle.Text & "'")
If Not rs.EOF Then
cbosuppliername.Text = rs!supplier_name
cboclass.Text = rs!classification
End If
End Sub

Private Sub loadcombo()
cboclass.Clear
Set rs = conn.Execute("select * from classification_table order by classification")
Do While Not rs.EOF
With cboclass
.AddItem rs!classification
End With
rs.MoveNext
Loop
End Sub

Private Sub txtcriticallevel_Change()
If Val(txtqty.Text) < Val(txtcriticallevel.Text) Then
MsgBox "Critical level is greater than the quantity required", vbInformation, "Confirmation"
txtcriticallevel.Text = ""
End If
End Sub
