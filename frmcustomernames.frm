VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcustomernames 
   Caption         =   "Customer Names"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer Name"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmcustomernames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
connstring = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstring

frmcustomernames.ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from customer_table ORDER BY CUSTOMER_NAME")
Do While Not rs.EOF
    Set X = frmcustomernames.ListView1.ListItems.Add(, , rs.Fields(0))
        X.SubItems(1) = rs.Fields!CUSTOMER_NAME
    rs.MoveNext
Loop
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim col
col = ListView1.SelectedItem.Index
frmcustomermaintenance.TXTID.Text = ListView1.ListItems.Item(col).Text
frmcustomermaintenance.txtcustomername.Text = ListView1.ListItems.Item(col).SubItems(1)
frmcustomermaintenance.txtaddress.Text = ListView1.ListItems.Item(col).SubItems(2)
frmcustomermaintenance.txtcontactperson.Text = ListView1.ListItems.Item(col).SubItems(3)
frmcustomermaintenance.txtcontactnumber.Text = ListView1.ListItems.Item(col).SubItems(4)
frmcustomermaintenance.txtcreditlimit.Text = ListView1.ListItems.Item(col).SubItems(5)
frmcustomermaintenance.txtterms.Text = ListView1.ListItems.Item(col).SubItems(6)
End Sub
