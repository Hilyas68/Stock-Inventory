VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsecuritymaintenance 
   BackColor       =   &H0080C0FF&
   Caption         =   "Security Maintenance"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      Top             =   8520
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
      Top             =   8520
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
      Left            =   12600
      TabIndex        =   12
      Top             =   8520
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
      Left            =   10680
      TabIndex        =   11
      Top             =   8520
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
      Left            =   8760
      TabIndex        =   10
      Top             =   8520
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
      Top             =   8520
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
      Top             =   8520
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   2520
      TabIndex        =   5
      Top             =   3120
      Width           =   9975
      _ExtentX        =   17595
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Position"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.ComboBox cboposition 
      Height          =   315
      ItemData        =   "frmsecuritymaintenance.frx":0000
      Left            =   7560
      List            =   "frmsecuritymaintenance.frx":000A
      TabIndex        =   4
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtusername 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtuserid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   5295
      Left            =   2400
      Top             =   2880
      Width           =   10335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      Height          =   1695
      Left            =   2400
      Top             =   960
      Width           =   10215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00004080&
      Height          =   255
      Left            =   7680
      TabIndex        =   27
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   7680
      TabIndex        =   26
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00004080&
      Height          =   4815
      Left            =   2640
      TabIndex        =   23
      Top             =   3240
      Width           =   9975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   12720
      TabIndex        =   22
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   10800
      TabIndex        =   21
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   8880
      TabIndex        =   20
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   6720
      TabIndex        =   19
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   4800
      TabIndex        =   18
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   2880
      TabIndex        =   17
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00004080&
      Height          =   615
      Left            =   960
      TabIndex        =   16
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Left            =   6480
      TabIndex        =   15
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   6480
      TabIndex        =   14
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   2520
      TabIndex        =   13
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frmsecuritymaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim toggle As Integer
Dim ssql As String
Dim X As Variant
Private Sub cmdadd_Click()
toggle = 0
txtusername.SetFocus
Call txtenabled(True, True, True)
Call CMDENABLED(False, False, False, True, True)
Call txtclear
End Sub

Private Sub cmdcancel_Click()
Call txtenabled(True, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub cmdclose_Click()
Call txtenabled(True, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
frmsecuritymaintenance.Hide
End Sub

Private Sub cmddelete_Click()
X = MsgBox("Are you sure you want to delete this/these item(s)?", vbYesNo + vbCritical, "Confirmation")
If X = vbYes Then
ssql = "delete * from security_tbl where user_id=" & txtuserid.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been deleted!", vbExclamation, "Confirmation"
Call txtenabled(True, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub CMDEDIT_Click()
toggle = 1
Call txtenabled(True, True, True)
Call CMDENABLED(False, False, False, True, True)
End Sub

Private Sub cmdfind_Click()
Set rs = conn.Execute("select * from security_tbl where user_name='" & txtusername.Text & "'")
If Not rs.EOF Then
txtuserid = rs!user_id
txtusername.Text = rs!user_name
txtpassword.Text = rs!pASSWORD
cboposition.Text = rs!Position
Call CMDENABLED(False, True, True, False, False)
Call txtenabled(True, False, False)
Else
MsgBox "No Record Found", vbExclamation, "Confirmation"
Call CMDENABLED(True, False, False, False, False)
Call txtenabled(True, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub cmdsave_Click()
Select Case toggle
Case 0:
ssql = "INSERT INTO SECURITY_TBL(USER_NAME,PASSWORD,POSITION)VALUES('" & txtusername.Text & "','" & txtpassword.Text & "','" & cboposition.Text & "')"
conn.Execute (ssql)
MsgBox "Existing Record has been saved", vbInformation, "Confirmation"

Case 1:
ssql = "update security_tbl set user_name='" & txtusername.Text & "',password='" & txtpassword.Text & "',position='" & cboposition.Text & "' where user_id=" & txtuserid.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been updated", vbInformation, "Confirmation"
End Select
Call txtenabled(True, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub Form_Load()
On Error Resume Next
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
Call txtenabled(True, False, False)
Call CMDENABLED(True, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub CMDENABLED(XADD, xedit, XDELETE, XSAVE, XCANCEL)
cmdadd.Enabled = XADD
CMDEDIT.Enabled = xedit
cmddelete.Enabled = XDELETE
cmdsave.Enabled = XSAVE
CMDCANCEL.Enabled = XCANCEL
End Sub

Private Sub txtenabled(XUSERNAME, XPASSWORD, xposition)
txtusername.Enabled = XUSERNAME
txtpassword.Enabled = XPASSWORD
cboposition.Enabled = xposition
End Sub
Private Sub txtclear()
txtusername.Text = ""
txtpassword.Text = ""
cboposition.Text = ""
End Sub
Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from security_tbl ORDER BY user_name")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
        X.SubItems(1) = rs.Fields!user_name
        X.SubItems(2) = rs.Fields!pASSWORD
        X.SubItems(3) = rs.Fields!Position
    rs.MoveNext
Loop
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim col
col = ListView1.SelectedItem.Index
txtuserid.Text = ListView1.ListItems.Item(col).Text
txtusername.Text = ListView1.ListItems.Item(col).SubItems(1)
txtpassword.Text = ListView1.ListItems.Item(col).SubItems(2)
cboposition.Text = ListView1.ListItems.Item(col).SubItems(3)
Call CMDENABLED(False, True, True, False, True)
Call txtenabled(True, False, False)
End Sub
