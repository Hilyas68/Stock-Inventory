VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmemployeesmaintenance 
   BackColor       =   &H0080C0FF&
   Caption         =   "Employee Maintenance"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdadddep 
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
      Left            =   11520
      TabIndex        =   37
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmddeldep 
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
      Left            =   12240
      TabIndex        =   36
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox txtdepartment 
      Height          =   315
      Left            =   9120
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtstarted 
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   2040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   55443457
      CurrentDate     =   39090
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
      Left            =   6120
      TabIndex        =   19
      Top             =   2040
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
      Left            =   5400
      TabIndex        =   18
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox txtemployeename 
      Height          =   315
      ItemData        =   "frmemployeesmaintenance.frx":0000
      Left            =   2880
      List            =   "frmemployeesmaintenance.frx":0002
      TabIndex        =   1
      Top             =   1560
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
      Left            =   840
      TabIndex        =   6
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
      Left            =   2760
      TabIndex        =   7
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
      Left            =   12360
      TabIndex        =   12
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
      Left            =   10440
      TabIndex        =   11
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
      Left            =   8520
      TabIndex        =   10
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
      Left            =   6600
      TabIndex        =   9
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
      Left            =   4680
      TabIndex        =   8
      Top             =   8880
      Width           =   1815
   End
   Begin VB.ComboBox cboposition 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   960
      TabIndex        =   5
      Top             =   3360
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8705
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employee_Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Position"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Department"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date_Started"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox txtidclass 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtdepid 
      Enabled         =   0   'False
      Height          =   615
      Left            =   9120
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label22 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   11640
      TabIndex        =   39
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label21 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   12360
      TabIndex        =   38
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   5535
      Left            =   840
      Top             =   3120
      Width           =   13455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   2175
      Left            =   840
      Top             =   720
      Width           =   13455
   End
   Begin VB.Label Label20 
      BackColor       =   &H00004080&
      Height          =   4815
      Left            =   1080
      TabIndex        =   35
      Top             =   3600
      Width           =   13095
   End
   Begin VB.Label Label19 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   6240
      TabIndex        =   34
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label18 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   5520
      TabIndex        =   33
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   9240
      TabIndex        =   32
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   9240
      TabIndex        =   31
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label14 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   3000
      TabIndex        =   29
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   28
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   12480
      TabIndex        =   27
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   10560
      TabIndex        =   26
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   8640
      TabIndex        =   25
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   6720
      TabIndex        =   24
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   4800
      TabIndex        =   23
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   2880
      TabIndex        =   22
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   960
      TabIndex        =   21
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Started"
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
      Left            =   7680
      TabIndex        =   15
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   7680
      TabIndex        =   14
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      TabIndex        =   13
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employees Name"
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
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "frmemployeesmaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ssql As String
Dim toggle As Integer
Dim x As Variant

Private Sub cboposition_Change()
Set rs = conn.Execute("select * from position_table where position='" & cboposition.Text & "'")
If Not rs.EOF Then
txtidclass.Text = rs!ID
End If
End Sub

Private Sub cboposition_Click()
Set rs = conn.Execute("select * from position_table where position='" & cboposition.Text & "'")
If Not rs.EOF Then
txtidclass.Text = rs!ID
End If

End Sub


Private Sub cmdadd_Click()
toggle = 0
txtemployeename.SetFocus
Call txtenabled(True, True, True, True)
Call CMDENABLED(False, False, False, True, True)
Call txtclear
End Sub

Private Sub cmdaddchoices_Click()
If cboposition.Text = "" Then
MsgBox "You must write a new choice first!", vbExclamation, "Confirmation"
cboposition.Enabled = True
cboposition.SetFocus
Else
ssql = "insert into position_table(position) values ('" & cboposition.Text & "')"
conn.Execute (ssql)
End If
cboposition.Clear
Set rs = conn.Execute("select * from POSITION_TABLE ORDER BY POSITION")
Do While Not rs.EOF
With cboposition
.AddItem rs!Position
End With
rs.MoveNext
Loop

End Sub

Private Sub cmdadddep_Click()
If txtdepartment.Text = "" Then
MsgBox "You must write a new choice first!", vbExclamation, "Confirmation"
txtdepartment.Text = ""
txtdepartment.Enabled = True
Else
ssql = "insert into department_table(department)values('" & txtdepartment.Text & "')"
conn.Execute (ssql)
txtdepartment.Text = ""
Call adddep
End If
End Sub

Private Sub cmdcancel_Click()
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub CMDCLOSE_Click()
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False)
Call txtclear
Call listall
frmemployeesmaintenance.Hide
End Sub

Private Sub cmddeldep_Click()
On Error Resume Next
ssql = "delete * from department_table where id=" & txtdepid.Text & ""
conn.Execute (ssql)
txtdepartment.Text = ""
Call adddep
End Sub

Private Sub cmddelete_Click()
x = MsgBox("Are you sure you want to delete this/these item(s)?", vbYesNo + vbCritical, "Confirmation")
If x = vbYes Then
ssql = "delete * from employees_table where id=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been deleted!", vbExclamation, "Confirmation"
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False)
Call txtclear
Call listall
End If

txtemployeename.Clear
Set rs = conn.Execute("select * from employees_table ORDER BY EMPLOYEES_NAME")
Do While Not rs.EOF
With txtemployeename
.AddItem rs!EMPLOYEES_NAME
End With
rs.MoveNext
Loop
End Sub

Private Sub CMDEDIT_Click()
toggle = 1
Call txtenabled(True, True, True, True)
Call CMDENABLED(False, False, False, True, True)
Call listall
End Sub

Private Sub cmdfind_Click()
Set rs = conn.Execute("select * from EMPLOYEEs_TABLE where EMPLOYEES_name='" & txtemployeename.Text & "'")
If Not rs.EOF Then
TXTID = rs!ID
txtemployeename.Text = rs!EMPLOYEES_NAME
txtdepartment.Text = rs!DEPARTMENT

Call CMDENABLED(False, True, True, False, False)
Call txtenabled(True, False, False, False)
Else
MsgBox "No Record Found", vbExclamation, "Confirmation"
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub cmdfind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Set rs = conn.Execute("select * from EMPLOYEEs_TABLE where EMPLOYEES_name='" & txtemployeename.Text & "'")
If Not rs.EOF Then
dtstarted.Value = rs!Date_STARTED
End If
End Sub

Private Sub cmdremovechoice_Click()
On Error Resume Next
ssql = "delete * from position_table where id=" & txtidclass.Text & ""
conn.Execute (ssql)
cboposition.Text = ""

cboposition.Clear
Set rs = conn.Execute("select * from position_table ORDER BY position")
Do While Not rs.EOF
With cboposition
.AddItem rs!Position
End With
rs.MoveNext
Loop
End Sub

Private Sub cmdsave_Click()
Select Case toggle
Case 0:
ssql = "insert into employees_table(employees_name,position,department,date_started)values('" & txtemployeename.Text & _
"','" & cboposition & "','" & txtdepartment.Text & "','" & dtstarted.Value & "')"
conn.Execute (ssql)
MsgBox "New Record has been saved", vbInformation, "Confirmation"

Case 1:
ssql = "update employees_table set employees_name='" & txtemployeename.Text & "',position='" & cboposition.Text & _
"',department='" & txtdepartment.Text & "',date_started='" & dtstarted.Value & "' where id=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been updated", vbInformation, ""
End Select

txtemployeename.Clear
Set rs = conn.Execute("select * from employees_table ORDER BY EMPLOYEES_NAME")
Do While Not rs.EOF
With txtemployeename
.AddItem rs!EMPLOYEES_NAME
End With
rs.MoveNext
Loop


Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub Form_Load()
On Error Resume Next
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False, False)
Call txtclear
Call listall

txtemployeename.Clear
Set rs = conn.Execute("select * from employees_table ORDER BY EMPLOYEES_NAME")
Do While Not rs.EOF
With txtemployeename
.AddItem rs!EMPLOYEES_NAME
End With
rs.MoveNext
Loop

cboposition.Clear
Set rs = conn.Execute("select * from POSITION_TABLE ORDER BY POSITION")
Do While Not rs.EOF
With cboposition
.AddItem rs!Position
End With
rs.MoveNext
Loop

Call adddep

End Sub
Private Sub CMDENABLED(XADD, xedit, XDELETE, XSAVE, XCANCEL)
cmdadd.Enabled = XADD
CMDEDIT.Enabled = xedit
cmddelete.Enabled = XDELETE
cmdsave.Enabled = XSAVE
cmdcancel.Enabled = XCANCEL
End Sub
Private Sub txtenabled(xemployeename, xposition, xdepartment, xdtstarted)
txtemployeename.Enabled = xemployeename
cboposition.Enabled = xposition
txtdepartment.Enabled = xdepartment
dtstarted.Enabled = xdtstarted
End Sub
Private Sub txtclear()
txtemployeename.Text = ""
cboposition.Text = ""
txtdepartment.Text = ""
End Sub
Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from employees_table ORDER BY employees_NAME")
Do While Not rs.EOF
    Set x = ListView1.ListItems.Add(, , rs.Fields(0))
        x.SubItems(1) = rs.Fields!EMPLOYEES_NAME
        x.SubItems(2) = rs.Fields!Position
        x.SubItems(3) = rs.Fields!DEPARTMENT
        x.SubItems(4) = rs.Fields!Date_STARTED
    rs.MoveNext
Loop
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim col
col = ListView1.SelectedItem.Index
TXTID.Text = ListView1.ListItems.Item(col).Text
txtemployeename.Text = ListView1.ListItems.Item(col).SubItems(1)
cboposition.Text = ListView1.ListItems.Item(col).SubItems(2)
txtdepartment.Text = ListView1.ListItems.Item(col).SubItems(3)
dtstarted.Value = ListView1.ListItems.Item(col).SubItems(4)
Call CMDENABLED(False, True, True, False, True)
Call txtenabled(True, False, False, False)
End Sub

Private Sub txtdepartment_Change()
Set rs = conn.Execute("select * from department_table where department='" & txtdepartment.Text & "'")
If Not rs.EOF Then
txtdepid.Text = rs!ID
End If
End Sub

Private Sub txtdepartment_Click()
Set rs = conn.Execute("select * from department_table where department='" & txtdepartment.Text & "'")
If Not rs.EOF Then
txtdepid.Text = rs!ID
End If
End Sub

Private Sub txtemployeename_Change()
Set rs = conn.Execute("select * from EMPLOYEEs_TABLE where EMPLOYEES_name='" & txtemployeename.Text & "'")
If Not rs.EOF Then
cboposition.Text = rs!Position
End If
End Sub

Private Sub txtemployeename_Click()
Set rs = conn.Execute("select * from EMPLOYEEs_TABLE where EMPLOYEES_name='" & txtemployeename.Text & "'")
If Not rs.EOF Then
cboposition.Text = rs!Position
End If
End Sub
Private Sub adddep()
txtdepartment.Clear
Set rs = conn.Execute("select * from DEPARTMENT_TABLE ORDER BY DEPARTMENT")
Do While Not rs.EOF
With txtdepartment
.AddItem rs!DEPARTMENT
End With
rs.MoveNext
Loop
End Sub
