VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcustomermaintenance 
   BackColor       =   &H0080C0FF&
   Caption         =   "Customer Maintenance"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   14130
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txtcustomername 
      Height          =   315
      ItemData        =   "frmcustomermaintenance.frx":0000
      Left            =   2640
      List            =   "frmcustomermaintenance.frx":0002
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.TextBox TXTID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   720
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
      Left            =   6360
      TabIndex        =   11
      Top             =   9120
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
      Left            =   8640
      TabIndex        =   12
      Top             =   9120
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
      Left            =   10560
      TabIndex        =   13
      Top             =   9120
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
      Left            =   12480
      TabIndex        =   14
      Top             =   9120
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
      TabIndex        =   9
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00C0C0FF&
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
      TabIndex        =   8
      Top             =   9120
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   840
      TabIndex        =   7
      Top             =   3840
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8281
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Address"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contact Person"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contact Number"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Credit Limit"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Terms"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox txtterms 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox txtcreditlimit 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox txtcontactperson 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   4335
   End
   Begin VB.TextBox txtcontactnumber 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9360
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      FillColor       =   &H00004080&
      Height          =   5175
      Left            =   600
      Top             =   3720
      Width           =   13575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   3015
      Left            =   600
      Top             =   480
      Width           =   13575
   End
   Begin VB.Label Label22 
      BackColor       =   &H00004080&
      Height          =   4695
      Left            =   1080
      TabIndex        =   36
      Top             =   3960
      Width           =   12855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9480
      TabIndex        =   35
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Label20 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9480
      TabIndex        =   34
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label19 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9480
      TabIndex        =   33
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label18 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   2760
      TabIndex        =   32
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Height          =   855
      Left            =   2760
      TabIndex        =   31
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   2760
      TabIndex        =   30
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label15 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   12600
      TabIndex        =   28
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   10680
      TabIndex        =   27
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   8760
      TabIndex        =   26
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   6480
      TabIndex        =   25
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   4560
      TabIndex        =   24
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   2640
      TabIndex        =   23
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   720
      TabIndex        =   22
      Top             =   9240
      Width           =   1815
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
      Left            =   840
      TabIndex        =   21
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit"
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
      Left            =   7560
      TabIndex        =   19
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Terms"
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
      Left            =   7560
      TabIndex        =   18
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Left            =   840
      TabIndex        =   17
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label3 
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
      Left            =   840
      TabIndex        =   16
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      Left            =   7560
      TabIndex        =   15
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "frmcustomermaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ssql As String
Dim toggle As Integer
Dim x As Variant
Private Sub cmdadd_Click()
toggle = 0
Call CMDENABLED(False, False, False, True, True)
Call txtenabled(True, True, True, True, True, True)
txtcustomername.SetFocus
Call txtclear
End Sub

Private Sub cmdbrowse_Click()
frmcustomernames.Show
End Sub

Private Sub cmdcancel_Click()
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub CMDCLOSE_Click()
frmcustomermaintenance.Hide
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub cmddelete_Click()
x = MsgBox("Are you sure you want to delete this/these item(s)?", vbYesNo + vbCritical, "Confirmation")
If x = vbYes Then
ssql = "delete * from customer_table where ID=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been deleted!", vbExclamation, "Confirmation"
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False, False, False)
Call txtclear
Call listall
End If

txtcustomername.Clear
Set rs = conn.Execute("select * from customer_table order by customer_name")
Do While Not rs.EOF
With txtcustomername
.AddItem rs!customer_name
End With
rs.MoveNext
Loop
End Sub

Private Sub CMDEDIT_Click()
toggle = 1
Call CMDENABLED(False, False, False, True, True)
Call txtenabled(True, True, True, True, True, True)
End Sub

Private Sub cmdfind_Click()
Set rs = conn.Execute("select * from customer_table where customer_name='" & txtcustomername.Text & "'")
If Not rs.EOF Then
TXTID = rs!ID
txtcustomername.Text = rs!customer_name
txtaddress.Text = rs!ADDRESS
txtcontactperson.Text = rs!contact_PERSON
txtcontactnumber.Text = rs!CONTACT_NUMBER
txtcreditlimit.Text = rs!CREDIT_LIMIT
txtterms.Text = rs!TERMS
Call CMDENABLED(False, True, True, False, False)
Call txtenabled(True, False, False, False, False, False)
Else
MsgBox "No Record Found", vbExclamation, "Confirmation"
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub cmdsave_Click()
Select Case toggle
Case 0:
ssql = "insert into Customer_table(customer_name,address,contact_person,contact_number,credit_limit,terms) values ('" & txtcustomername.Text & _
"', '" & txtaddress.Text & "','" & txtcontactperson.Text & "','" & txtcontactnumber.Text & "','" & txtcreditlimit.Text & "','" & txtterms.Text & "')"
conn.Execute (ssql)
MsgBox "New Record has been saved", vbInformation, "Confirmation"

Case 1:
ssql = "update customer_table set customer_name='" & txtcustomername.Text & "',address='" & txtaddress.Text & "',contact_person='" & txtcontactperson.Text & _
"',contact_number='" & txtcontactnumber.Text & "',credit_limit='" & txtcreditlimit.Text & "',terms='" & txtterms.Text & "' where ID=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been updated", vbInformation, "Confirmation"
End Select
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False, False, False)
Call txtclear
Call listall

txtcustomername.Clear
Set rs = conn.Execute("select * from customer_table order by customer_name")
Do While Not rs.EOF
With txtcustomername
.AddItem rs!customer_name
End With
rs.MoveNext
Loop

End Sub

Private Sub Form_Load()
On Error Resume Next
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False, False, False)
Call txtclear
Call listall

txtcustomername.Clear
Set rs = conn.Execute("select * from customer_table order by customer_name")
Do While Not rs.EOF
With txtcustomername
.AddItem rs!customer_name
End With
rs.MoveNext
Loop

End Sub

Private Sub CMDENABLED(XADD, xedit, XDELETE, XSAVE, XCANCEL)
cmdadd.Enabled = XADD
CMDEDIT.Enabled = xedit
cmddelete.Enabled = XDELETE
cmdsave.Enabled = XSAVE
cmdcancel.Enabled = XCANCEL
End Sub
Private Sub txtenabled(xcustomername, XADDRESS, XCONTACTPERSON, XCONTACTNUMBER _
, xcreditlimit, xterms)
txtcustomername.Enabled = xcustomername
txtaddress.Enabled = XADDRESS
txtcontactperson.Enabled = XCONTACTPERSON
txtcontactnumber.Enabled = XCONTACTNUMBER
txtcreditlimit.Enabled = xcreditlimit
txtterms.Enabled = xterms
End Sub
Private Sub txtclear()
txtcustomername.Text = ""
txtaddress.Text = ""
txtcontactperson.Text = ""
txtcontactnumber.Text = ""
txtcreditlimit.Text = ""
txtterms.Text = ""
End Sub
Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from customer_table ORDER BY CUSTOMER_NAME")
Do While Not rs.EOF
    Set x = ListView1.ListItems.Add(, , rs.Fields(0))
        x.SubItems(1) = rs.Fields!customer_name
        x.SubItems(2) = rs.Fields!ADDRESS
        x.SubItems(3) = rs.Fields!contact_PERSON
        x.SubItems(4) = rs.Fields!CONTACT_NUMBER
        x.SubItems(5) = rs.Fields!CREDIT_LIMIT
        x.SubItems(6) = rs.Fields!TERMS
    rs.MoveNext
Loop
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim col
col = ListView1.SelectedItem.Index
TXTID.Text = ListView1.ListItems.Item(col).Text
txtcustomername.Text = ListView1.ListItems.Item(col).SubItems(1)
txtaddress.Text = ListView1.ListItems.Item(col).SubItems(2)
txtcontactperson.Text = ListView1.ListItems.Item(col).SubItems(3)
txtcontactnumber.Text = ListView1.ListItems.Item(col).SubItems(4)
txtcreditlimit.Text = ListView1.ListItems.Item(col).SubItems(5)
txtterms.Text = ListView1.ListItems.Item(col).SubItems(6)
Call CMDENABLED(False, True, True, False, True)
Call txtenabled(True, False, False, False, False, False)
End Sub

