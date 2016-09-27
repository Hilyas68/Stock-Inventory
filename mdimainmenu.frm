VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdimainmenu 
   BackColor       =   &H8000000C&
   Caption         =   "STOCK INVENTORY MANAGER"
   ClientHeight    =   5700
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9165
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdimainmenu.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   1320
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5325
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "&TRANSACTION"
      Begin VB.Menu mnuInvoice 
         Caption         =   "Issue Invoice"
      End
      Begin VB.Menu mnusuppliercategory 
         Caption         =   "Supp&lier"
         Begin VB.Menu mnusupplierdelivery 
            Caption         =   "Suppli&er Deliveries"
         End
      End
      Begin VB.Menu mnucustomercategory 
         Caption         =   "Cus&tomer"
         Begin VB.Menu mnucustomerdelivery 
            Caption         =   "Incoming Deliveries"
         End
         Begin VB.Menu mnucustomerreturn 
            Caption         =   "Customer Return"
         End
      End
   End
   Begin VB.Menu mnuinquiry 
      Caption         =   "&INQUIRY"
      Begin VB.Menu mnustockonhand 
         Caption         =   "Stock on Hand"
      End
   End
   Begin VB.Menu mnumaintenance 
      Caption         =   "&MAINTENANCE"
      Begin VB.Menu mnucustomer 
         Caption         =   "C&ustomer"
      End
      Begin VB.Menu mnuproduct 
         Caption         =   "&Product"
      End
      Begin VB.Menu mnusupplier 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu mnuemployee 
         Caption         =   "&Employee"
      End
      Begin VB.Menu mnusecurity 
         Caption         =   "&Security"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&REPORT"
      Begin VB.Menu mnucustomerreport 
         Caption         =   "C&ustomer"
      End
      Begin VB.Menu mnuproductreport 
         Caption         =   "&Product"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&ABOUT"
      Begin VB.Menu mnusysteminformation 
         Caption         =   "Sy&stem Information"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "&EXIT"
   End
End
Attribute VB_Name = "mdimainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset



Private Sub MDIForm_Load()
On Error Resume Next
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
StatusBar1.Panels(1) = "Date: " + Format(Now, "MMMM-DD-YYYY")
End Sub

Private Sub mnucustomer_Click()
On Error Resume Next
frmcustomermaintenance.Show
End Sub

Private Sub mnucustomerdelivery_Click()
frmincomingdeliveries.Show
End Sub

Private Sub mnucustomerreport_Click()
On Error Resume Next
rs.Close
rs.Open "select * from customer_table", conn
Set DataReport1.DataSource = rs
DataReport1.Show
End Sub

Private Sub mnucustomerreturn_Click()
On Error Resume Next
frmcustomerreturn.Show
End Sub

Private Sub mnuemployee_Click()
On Error Resume Next
frmemployeesmaintenance.Show
End Sub

Private Sub mnuexit_Click()
x = MsgBox("Log out from the system?", vbYesNo + vbInformation, "Log Out")
If x = vbYes Then
''End
    mdimainmenu.Hide
    frmsecurity.Combo1.Text = ""
    frmsecurity.txtpassword.Text = ""
    frmsecurity.Show
End If
End Sub

Private Sub mnuInvoice_Click()
invoice.Show
End Sub

Private Sub mnuproduct_Click()

frmproductmaintenance.Show
End Sub

Private Sub mnuproductreport_Click()
On Error Resume Next
rs.Close
rs.Open "select * from product_table", conn
Set DataReport2.DataSource = rs
DataReport2.Show
End Sub

Private Sub mnusecurity_Click()
On Error Resume Next
frmsecuritymaintenance.Show
End Sub

Private Sub mnustockonhand_Click()
On Error Resume Next
frmstockonhand.Show
End Sub

Private Sub mnusupplier_Click()
On Error Resume Next
frmsuppliermaintenance.Show
End Sub

Private Sub mnusupplierdelivery_Click()
On Error Resume Next
frmsupplierdelivery.Show
End Sub

Private Sub mnusystemdeveloper_Click()
On Error Resume Next
frmsystemdeveloper.Show
End Sub

Private Sub mnusysteminformation_Click()
On Error Resume Next
frmsysteminformation.Show
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(2) = "Time: " + Format(Now, "HH:MM:SS AMPM")
End Sub
