VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmstockonhand 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Stock on Hand"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   9255
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "ON HAND"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   19
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdfindall 
      Caption         =   "CRITICAL"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   16
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton CMDCLEAR 
      Caption         =   "&CLEAR"
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
      TabIndex        =   14
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton CMDCLOSE 
      Caption         =   "C&LOSE"
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
      Left            =   8280
      TabIndex        =   12
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIND"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton CMDFINDbyinvoiceno 
      Caption         =   "FIND"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   1800
      TabIndex        =   5
      Top             =   2280
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9340
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
         Text            =   "Code"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Article"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Stock on hand"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Critical Level"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.ComboBox cboarticle 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
   End
   Begin VB.ComboBox cboclass 
      Height          =   315
      Left            =   3600
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblongoing 
      Caption         =   "on going"
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404000&
      Height          =   855
      Left            =   3720
      TabIndex        =   20
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Label lblnone 
      Height          =   495
      Left            =   2520
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404000&
      Height          =   855
      Left            =   1800
      TabIndex        =   17
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404000&
      Height          =   615
      Left            =   6480
      TabIndex        =   15
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackColor       =   &H00404000&
      Height          =   615
      Left            =   8400
      TabIndex        =   13
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404000&
      BorderWidth     =   3
      Height          =   5655
      Left            =   1680
      Top             =   2160
      Width           =   11295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404000&
      BorderWidth     =   3
      Height          =   1455
      Left            =   1680
      Top             =   600
      Width           =   11295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      Caption         =   "Label3"
      Height          =   5295
      Left            =   1920
      TabIndex        =   11
      Top             =   2400
      Width           =   10935
   End
   Begin VB.Label lblneedtopurchase 
      Caption         =   "need to order"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label38 
      BackColor       =   &H00404000&
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   960
      Width           =   735
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label22 
      BackColor       =   &H00404000&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "frmstockonhand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cboclass_Click()
cboarticle.Clear
Set rs = conn.Execute("select * from product_table where classification='" & cboclass.Text & "'")
Do While Not rs.EOF
With cboarticle
.AddItem rs!article
End With
rs.MoveNext
Loop

End Sub

Private Sub CMDCLEAR_Click()
cboclass.Text = ""
cboarticle.Text = ""
Call listnone
End Sub

Private Sub CMDCLOSE_Click()
cboarticle.Text = ""
cboclass.Text = ""
frmstockonhand.Hide
Call listnone
End Sub

Private Sub cmdfindall_Click()
Call LISTALLARTICLE
End Sub

Private Sub CMDFINDbyinvoiceno_Click()
Call listall
End Sub

Private Sub Command1_Click()
Call listbyarticle
End Sub

Private Sub Command2_Click()
Call LISTALLPRODUCTS
End Sub

Private Sub Form_Load()
On Error Resume Next
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG

Set rs = conn.Execute("SELECT * FROM CLASSIFICATION_TABLE")
Do While Not rs.EOF
With cboclass
.AddItem rs!classification
End With
rs.MoveNext
Loop

Set rs = conn.Execute("SELECT * FROM product_TABLE")
Do While Not rs.EOF
With cboarticle
.AddItem rs!article
End With
rs.MoveNext
Loop
End Sub
Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from product_table  where status ='" & lblneedtopurchase.Caption & "' and classification='" & cboclass.Text & "' ORDER BY code")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(3))
        X.SubItems(1) = rs.Fields!article
        X.SubItems(2) = rs.Fields!qty
        X.SubItems(3) = rs.Fields!critical_level
    rs.MoveNext
Loop
End Sub
Private Sub listbyarticle()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from product_table  where status ='" & lblneedtopurchase.Caption & "' and article='" & cboarticle.Text & "' ORDER BY article")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(3))
        X.SubItems(1) = rs.Fields!article
        X.SubItems(2) = rs.Fields!qty
        X.SubItems(3) = rs.Fields!critical_level
    rs.MoveNext
Loop
End Sub
Private Sub LISTALLARTICLE()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from product_table  where status ='" & lblneedtopurchase.Caption & "' order by code ")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(3))
        X.SubItems(1) = rs.Fields!article
        X.SubItems(2) = rs.Fields!qty
        X.SubItems(3) = rs.Fields!critical_level
    rs.MoveNext
Loop
End Sub
Private Sub listnone()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from product_table  where status ='" & lblnone.Caption & "'")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(3))
        X.SubItems(1) = rs.Fields!article
        X.SubItems(2) = rs.Fields!qty
        X.SubItems(3) = rs.Fields!critical_level
    rs.MoveNext
Loop
End Sub

Private Sub LISTALLPRODUCTS()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from product_table WHERE STATUS ='" & lblongoing.Caption & "' order by code ")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(3))
        X.SubItems(1) = rs.Fields!article
        X.SubItems(2) = rs.Fields!qty
        X.SubItems(3) = rs.Fields!critical_level
    rs.MoveNext
Loop
End Sub

