VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsuppliermaintenance 
   BackColor       =   &H0080C0FF&
   Caption         =   "Supplier Maintenance"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txtsuppliername 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   4335
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
      Left            =   7560
      TabIndex        =   21
      Top             =   2160
      Width           =   495
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
      Left            =   6960
      TabIndex        =   20
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox TXTID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   1080
      Width           =   2175
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
      Left            =   4440
      TabIndex        =   9
      Top             =   9240
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
      Left            =   6360
      TabIndex        =   10
      Top             =   9240
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
      Left            =   8880
      TabIndex        =   11
      Top             =   9240
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
      Left            =   10800
      TabIndex        =   12
      Top             =   9240
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
      Left            =   12720
      TabIndex        =   13
      Top             =   9240
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
      Left            =   2520
      TabIndex        =   8
      Top             =   9240
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
      Left            =   600
      TabIndex        =   7
      Top             =   9240
      Width           =   1815
   End
   Begin VB.TextBox txtcontactnumber 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10320
      TabIndex        =   5
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox txtcontactperson 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10320
      TabIndex        =   4
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2640
      Width           =   4335
   End
   Begin VB.ComboBox cboclassification 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   8493
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Supplier Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product Line"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Address"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contact Person"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contact Number"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox txtidclass 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   5175
      Left            =   240
      Top             =   3840
      Width           =   14655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   3135
      Left            =   240
      Top             =   600
      Width           =   14655
   End
   Begin VB.Label Label22 
      BackColor       =   &H00004080&
      Height          =   4815
      Left            =   600
      TabIndex        =   38
      Top             =   4080
      Width           =   14175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00004080&
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
      Left            =   7680
      TabIndex        =   37
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label20 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label Label19 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   2520
      TabIndex        =   35
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   2520
      TabIndex        =   34
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Height          =   855
      Left            =   2520
      TabIndex        =   33
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   10440
      TabIndex        =   32
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label14 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   10440
      TabIndex        =   31
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00004080&
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
      Left            =   7080
      TabIndex        =   30
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   12840
      TabIndex        =   29
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   10920
      TabIndex        =   28
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   9000
      TabIndex        =   27
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   6480
      TabIndex        =   26
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   4560
      TabIndex        =   25
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   2640
      TabIndex        =   24
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   720
      TabIndex        =   23
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Label Label6 
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
      Height          =   495
      Left            =   360
      TabIndex        =   19
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
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
      Left            =   8280
      TabIndex        =   17
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
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
      Left            =   8280
      TabIndex        =   16
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Line"
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
      Left            =   360
      TabIndex        =   14
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "frmsuppliermaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim toggle As Integer

Private Sub cboclassification_Change()
Set rs = conn.Execute("select * from productline_tbl where product_line='" & cboclassification.Text & "'")
If Not rs.EOF Then
txtidclass.Text = rs!ID
End If
End Sub

Private Sub cboclassification_Click()
Set rs = conn.Execute("select * from productline_tbl where product_line='" & cboclassification.Text & "'")
If Not rs.EOF Then
txtidclass.Text = rs!ID
End If
End Sub

Private Sub cmdadd_Click()
toggle = 0
txtsuppliername.SetFocus
Call txtenabled(True, True, True, True, True)
Call CMDENABLED(False, False, False, True, True)
Call txtclear
End Sub

Private Sub cmdaddchoices_Click()
If cboclassification.Text = "" Then
MsgBox "You must write a new choice first!", vbExclamation, "Confirmation"
cboclassification.Enabled = True
cboclassification.SetFocus
Else
ssql = "insert into productline_tbl(product_line) values ('" & cboclassification.Text & "')"
conn.Execute (ssql)
Call comboload
End If
End Sub

Private Sub cmdcancel_Click()
Call txtenabled(True, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub cmdclose_Click()
Call txtenabled(True, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
frmsuppliermaintenance.Hide
End Sub

Private Sub cmddelete_Click()
X = MsgBox("Are you sure you want to delete this/these item(s)?", vbYesNo + vbCritical, "Confirmation")
If X = vbYes Then
ssql = "delete * from supplier_table where id=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been deleted!", vbExclamation, "Confirmation"
Call txtenabled(True, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
End If
Call comboload
End Sub

Private Sub CMDEDIT_Click()
toggle = 1
Call txtenabled(True, True, True, True, True)
Call CMDENABLED(False, False, False, True, True)
End Sub

Private Sub cmdfind_Click()
Set rs = conn.Execute("select * from supplier_table where supplier_name='" & txtsuppliername.Text & "'")
If Not rs.EOF Then
TXTID = rs!ID
txtsuppliername.Text = rs!supplier_name
txtaddress.Text = rs!ADDRESS
txtcontactperson.Text = rs!contact_PERSON
txtcontactnumber.Text = rs!CONTACT_NUMBER
Call CMDENABLED(False, True, True, False, False)
Call txtenabled(True, False, False, False, False)
Else
MsgBox "No Record Found", vbExclamation, "Confirmation"
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub cmdremovechoice_Click()
On Error Resume Next
ssql = "delete * from productline_tbl where id=" & txtidclass.Text & ""
conn.Execute (ssql)
Call comboload
cboclassification.Text = ""
End Sub

Private Sub cmdsave_Click()
Select Case toggle
Case 0:
ssql = "insert into supplier_table(supplier_name,product_line,address,contact_person,contact_number)values ('" & txtsuppliername.Text & _
"','" & cboclassification.Text & "','" & txtaddress.Text & "','" & txtcontactperson.Text & "','" & txtcontactnumber.Text & "')"
conn.Execute (ssql)
MsgBox "New Record has been saved", vbInformation, "Confirmation"

Case 1:
ssql = "update supplier_table set supplier_name='" & txtsuppliername.Text & "',product_line='" & cboclassification.Text & "',address='" & txtaddress.Text & _
"',contact_person='" & txtcontactperson.Text & "',contact_number ='" & txtcontactnumber.Text & "' WHERE ID=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been updated", vbInformation, "Confirmation"
End Select
Call txtenabled(True, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
Call comboload
End Sub

Private Sub Form_Load()
On Error Resume Next
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
Call txtenabled(True, False, False, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
Call comboload


End Sub
Private Sub txtenabled(XSUPPLIERNAME, XCLASSIFICATION, XADDRESS, XCONTACTPERSON, XCONTACTNUMBER)
txtsuppliername.Enabled = XSUPPLIERNAME
cboclassification.Enabled = XCLASSIFICATION
txtaddress.Enabled = XADDRESS
txtcontactperson.Enabled = XCONTACTPERSON
txtcontactnumber.Enabled = XCONTACTNUMBER
End Sub

Private Sub CMDENABLED(XADD, xedit, XDELETE, XSAVE, XCANCEL)
cmdadd.Enabled = XADD
CMDEDIT.Enabled = xedit
cmddelete.Enabled = XDELETE
cmdsave.Enabled = XSAVE
CMDCANCEL.Enabled = XCANCEL
End Sub
Private Sub txtclear()
txtsuppliername.Text = ""
cboclassification.Text = ""
txtaddress.Text = ""
txtcontactperson.Text = ""
txtcontactnumber.Text = ""
End Sub

Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from SUPPLIER_TABLE ORDER BY SUPPLIER_NAME")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
        X.SubItems(1) = rs.Fields!supplier_name
        X.SubItems(2) = rs.Fields!product_line
        X.SubItems(3) = rs.Fields!ADDRESS
        X.SubItems(4) = rs.Fields!contact_PERSON
        X.SubItems(5) = rs.Fields!CONTACT_NUMBER
    rs.MoveNext
Loop
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim col
col = ListView1.SelectedItem.Index
TXTID.Text = ListView1.ListItems.Item(col).Text
txtsuppliername.Text = ListView1.ListItems.Item(col).SubItems(1)
cboclassification.Text = ListView1.ListItems.Item(col).SubItems(2)
txtaddress.Text = ListView1.ListItems.Item(col).SubItems(3)
txtcontactperson.Text = ListView1.ListItems.Item(col).SubItems(4)
txtcontactnumber.Text = ListView1.ListItems.Item(col).SubItems(5)
Call CMDENABLED(False, True, True, False, True)
Call txtenabled(True, False, False, False, False)
End Sub

Private Sub comboload()
cboclassification.Clear
Set rs = conn.Execute("select * from productline_tbl order by product_line")
Do While Not rs.EOF
With cboclassification
.AddItem rs!product_line
End With
rs.MoveNext
Loop
txtsuppliername.Clear
Set rs = conn.Execute("select * from supplier_table order by supplier_name")
Do While Not rs.EOF
With txtsuppliername
.AddItem rs!supplier_name
End With
rs.MoveNext
Loop
End Sub

Private Sub txtsuppliername_Change()
On Error Resume Next
Set rs = conn.Execute("select * from supplier_table where supplier_name='" & txtsuppliername.Text & "'")
If Not rs.EOF Then
cboclassification.Text = rs!product_line
End If
End Sub

Private Sub txtsuppliername_Click()
Set rs = conn.Execute("select * from supplier_table where supplier_name='" & txtsuppliername.Text & "'")
If Not rs.EOF Then
cboclassification.Text = rs!product_line
End If
End Sub
