VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmincomingdeliveries 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Incoming Deliveries"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
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
      TabIndex        =   75
      Top             =   9120
      Width           =   1815
   End
   Begin VB.ComboBox cbocustomername 
      Height          =   315
      ItemData        =   "frmincomingdeliveries.frx":0000
      Left            =   2160
      List            =   "frmincomingdeliveries.frx":0002
      TabIndex        =   3
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CommandButton cmdfindbydate 
      Caption         =   "FIND"
      Height          =   375
      Left            =   4560
      TabIndex        =   69
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton CMDFINDbyinvoiceno 
      Caption         =   "FIND"
      Height          =   375
      Left            =   4560
      TabIndex        =   67
      Top             =   1200
      Width           =   735
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
      TabIndex        =   16
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
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
      Left            =   10440
      TabIndex        =   59
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command5 
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
      Left            =   11040
      TabIndex        =   58
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command4 
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
      Left            =   10440
      TabIndex        =   55
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command3 
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
      Left            =   11040
      TabIndex        =   54
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10440
      TabIndex        =   51
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
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
      Left            =   11040
      TabIndex        =   50
      Top             =   360
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
      Left            =   10440
      TabIndex        =   47
      Top             =   960
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
      Left            =   11040
      TabIndex        =   46
      Top             =   960
      Width           =   495
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
      TabIndex        =   15
      Top             =   9120
      Width           =   1815
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
      TabIndex        =   17
      Top             =   9120
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   480
      TabIndex        =   13
      Top             =   4680
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   5741
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Year"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Customer Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Serial No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Brand Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Color"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Fabric"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Article"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Total Price"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txttotalprice 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtunitprice 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   3240
      Width           =   2175
   End
   Begin VB.ComboBox cboarticle 
      Height          =   315
      ItemData        =   "frmincomingdeliveries.frx":0004
      Left            =   8040
      List            =   "frmincomingdeliveries.frx":0006
      TabIndex        =   10
      Top             =   2760
      Width           =   4215
   End
   Begin VB.ComboBox cbofabric 
      Height          =   315
      ItemData        =   "frmincomingdeliveries.frx":0008
      Left            =   8040
      List            =   "frmincomingdeliveries.frx":000A
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox cbocolor 
      Height          =   315
      ItemData        =   "frmincomingdeliveries.frx":000C
      Left            =   8040
      List            =   "frmincomingdeliveries.frx":000E
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ComboBox cbotype 
      Height          =   315
      ItemData        =   "frmincomingdeliveries.frx":0010
      Left            =   8040
      List            =   "frmincomingdeliveries.frx":0012
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox cbobrandname 
      Height          =   315
      ItemData        =   "frmincomingdeliveries.frx":0014
      Left            =   8040
      List            =   "frmincomingdeliveries.frx":0016
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtinvoiceno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox TXTID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtdate 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39091
   End
   Begin VB.TextBox txtidbrand 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   62
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtidtype 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      TabIndex        =   63
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtidcolor 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   64
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtidfabric 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   65
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TXTPRODID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   72
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TXTPRODQTY 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   71
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtcritical 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   78
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TXTZERO 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      TabIndex        =   80
      Text            =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblneedtopurchase 
      Caption         =   "need to order"
      Height          =   375
      Left            =   3720
      TabIndex        =   79
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   4215
      Left            =   360
      Top             =   240
      Width           =   14535
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   3615
      Left            =   360
      Top             =   4560
      Width           =   14535
   End
   Begin VB.Label lblyear 
      Height          =   495
      Left            =   2160
      TabIndex        =   77
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label42 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   6600
      TabIndex        =   76
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label41 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   74
      Top             =   2520
      Width           =   4335
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
      Left            =   480
      TabIndex        =   73
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label39 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   70
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label38 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   68
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label37 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   4680
      TabIndex        =   66
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label36 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   10560
      TabIndex        =   61
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label35 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   11160
      TabIndex        =   60
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label34 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   10560
      TabIndex        =   57
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   11160
      TabIndex        =   56
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label32 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   10560
      TabIndex        =   53
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label31 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   11160
      TabIndex        =   52
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label30 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   10560
      TabIndex        =   49
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label29 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   11160
      TabIndex        =   48
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label28 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   8520
      TabIndex        =   45
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label27 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   2760
      TabIndex        =   44
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label26 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   840
      TabIndex        =   43
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Height          =   3255
      Left            =   600
      TabIndex        =   42
      Top             =   4800
      Width           =   14175
   End
   Begin VB.Label Label24 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      TabIndex        =   41
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label23 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      TabIndex        =   40
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label22 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   8160
      TabIndex        =   39
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label21 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   8160
      TabIndex        =   38
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   8160
      TabIndex        =   37
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   8160
      TabIndex        =   36
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   35
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   8160
      TabIndex        =   34
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   31
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   30
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price"
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
      Left            =   6720
      TabIndex        =   29
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
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
      Left            =   6720
      TabIndex        =   28
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Article "
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
      Left            =   6720
      TabIndex        =   27
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Fabric"
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
      Left            =   6720
      TabIndex        =   26
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
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
      Left            =   6720
      TabIndex        =   25
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   6720
      TabIndex        =   24
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
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
      Left            =   6720
      TabIndex        =   23
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No"
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
      TabIndex        =   21
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
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
      Left            =   480
      TabIndex        =   20
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   19
      Top             =   720
      Width           =   1815
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
      Left            =   480
      TabIndex        =   18
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "frmincomingdeliveries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim COMPUTE As Double

Private Sub cboarticle_Change()
Set rs = conn.Execute("select * from product_table where article='" & cboarticle.Text & "' AND QTY <> '" & TXTZERO.Text & "'")
If Not rs.EOF Then
cboarticle.Text = rs!article
txtserialno.Text = rs!code
txtunitprice.Text = rs!selling_price
TXTPRODQTY.Text = rs!qty
TXTPRODID.Text = rs!ID
txtcritical.Text = rs!critical_level
Call COMPUTEAMOUNT
End If
End Sub

Private Sub cboarticle_Click()
Set rs = conn.Execute("select * from product_table where article='" & cboarticle.Text & "' AND QTY <> '" & TXTZERO.Text & "'")
If Not rs.EOF Then
cboarticle.Text = rs!article
txtserialno.Text = rs!code
txtunitprice.Text = rs!selling_price
TXTPRODQTY.Text = rs!qty
TXTPRODID.Text = rs!ID
txtcritical.Text = rs!critical_level
Call COMPUTEAMOUNT
End If
End Sub

Private Sub cbobrandname_Change()
Set rs = conn.Execute("select * from brand_name_table where brand_name='" & cbobrandname.Text & "'")
If Not rs.EOF Then
txtidbrand.Text = rs!ID
End If

End Sub

Private Sub cbobrandname_Click()
Set rs = conn.Execute("select * from brand_name_table where brand_name='" & cbobrandname.Text & "'")
If Not rs.EOF Then
txtidbrand.Text = rs!ID
End If
End Sub

Private Sub cbocolor_Change()
Set rs = conn.Execute("select * from color_table where color='" & cbocolor.Text & "'")
If Not rs.EOF Then
txtidcolor.Text = rs!ID
End If
End Sub

Private Sub cbocolor_Click()
Set rs = conn.Execute("select * from color_table where color='" & cbocolor.Text & "'")
If Not rs.EOF Then
txtidcolor.Text = rs!ID
End If
End Sub



Private Sub cbofabric_Change()
Set rs = conn.Execute("select * from fabric_table where fabric='" & cbofabric.Text & "'")
If Not rs.EOF Then
txtidfabric.Text = rs!ID
End If
End Sub

Private Sub cbofabric_Click()
Set rs = conn.Execute("select * from fabric_table where fabric='" & cbofabric.Text & "'")
If Not rs.EOF Then
txtidfabric.Text = rs!ID
End If
End Sub


Private Sub cbotype_Change()
Set rs = conn.Execute("select * from type_table where type='" & cbotype.Text & "'")
If Not rs.EOF Then
txtidtype.Text = rs!ID
End If
End Sub

Private Sub cbotype_Click()
Set rs = conn.Execute("select * from type_table where type='" & cbotype.Text & "'")
If Not rs.EOF Then
txtidtype.Text = rs!ID
End If
End Sub

Private Sub cmdadd_Click()
Call txtenabled(True, True, True, True, True, True, True, True, True, True, True, True)
Call txtclear
Call CMDENABLED(False, False, True, True)
End Sub

Private Sub cmdaddchoices_Click()
If cbotype.Text = "" Then
MsgBox "You must write a new choice first!", vbExclamation, "Confirmation"
cbotype.Text = ""
cbotype.Enabled = True
Else
ssql = "insert into type_table(type)values('" & cbotype.Text & "')"
conn.Execute (ssql)
cbotype.Text = ""
Call cboaddtype
End If
End Sub

Private Sub cmdcancel_Click()
lblyear.Caption = Format(Now, "yyyy")
Call listall
Call cboaddchoices
Call cboaddtype
Call cboaddcolor
Call cboaddfabric
Call txtenabled(True, True, False, False, False, False, False, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear

End Sub

Private Sub CMDCLOSE_Click()
lblyear.Caption = Format(Now, "yyyy")
frmincomingdeliveries.Hide
Call listall
Call txtclear
Call txtenabled(True, True, False, False, False, False, False, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
End Sub

Private Sub cmddelete_Click()
X = MsgBox("Are you sure you want to delete this/these item(s)?", vbYesNo + vbCritical, "Confirmation")
If X = vbYes Then
ssql = "delete * from incoming_delivery_table where id=" & TXTID.Text & ""
conn.Execute (ssql)
MsgBox "Existing Record has been deleted!", vbExclamation, "Confirmation"
Call listall
Call txtclear
Call txtenabled(True, True, False, False, False, False, False, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
End If
End Sub

Private Sub cmdfindbydate_Click()

Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cboarticle.Text = rs!Description
End If

Set rs = conn.Execute("select * from incoming_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
TXTID.Text = rs!ID
lblyear.Caption = rs!year_today
dtdate.Value = rs!Dated_on
txtserialno.Text = rs!serial_no
txtquantity.Text = rs!quantity
txtunitprice.Text = rs!unit_price
txttotalprice.Text = rs!total_price

Call txtenabled(True, True, False, False, False, False, False, False, False, False, False, False)
Call CMDENABLED(False, True, False, True)
Call listbydate
Else
MsgBox "No Result found!", vbExclamation, "Confirmation"
Call txtenabled(True, True, True, False, False, False, False, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
End If
End Sub

Private Sub CMDFINDbydate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cbobrandname.Text = rs!brand_name
End If
Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cbotype.Text = rs!Type
End If
Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cbocolor.Text = rs!Color
End If
Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where dated_on='" & dtdate.Value & "'")
If Not rs.EOF Then
cbofabric.Text = rs!fabric
End If
End Sub

Private Sub CMDFINDbyinvoiceno_Click()

Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where invoice_no='" & txtinvoiceno.Text & "'")
If Not rs.EOF Then
cboarticle.Text = rs!Description
End If


Set rs = conn.Execute("select * from incoming_delivery_table where invoice_no='" & txtinvoiceno.Text & "'")
If Not rs.EOF Then
txtinvoiceno.Text = rs!invoice_no
TXTID.Text = rs!ID
lblyear.Caption = rs!year_today
dtdate.Value = rs!Dated_on
txtserialno.Text = rs!serial_no
txtquantity.Text = rs!quantity
txtunitprice.Text = rs!unit_price
txttotalprice.Text = rs!total_price
Call txtenabled(True, True, False, False, False, False, False, False, False, False, False, False)
Call CMDENABLED(False, True, False, True)
Call listbyinvoiceno
Else
MsgBox "No Result found!", vbExclamation, "Confirmation"
Call txtenabled(True, True, True, False, False, False, False, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear
Call listall
End If
End Sub
Private Sub CMDFINDbyinvoiceno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where invoice_no='" & txtinvoiceno.Text & "'")
If Not rs.EOF Then
cbobrandname.Text = rs!brand_name
End If
Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where invoice_no='" & txtinvoiceno.Text & "'")
If Not rs.EOF Then
cbotype.Text = rs!Type
End If
Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where invoice_no='" & txtinvoiceno.Text & "'")
If Not rs.EOF Then
cbocolor.Text = rs!Color
End If
Set rs = conn.Execute("SELECT * FROM incoming_delivery_table where invoice_no='" & txtinvoiceno.Text & "'")
If Not rs.EOF Then
cbofabric.Text = rs!fabric
End If
End Sub

Private Sub cmdremovechoice_Click()
On Error Resume Next
ssql = "delete * from type_table where id=" & txtidtype.Text & ""
conn.Execute (ssql)
cbotype.Text = ""
Call cboaddtype
End Sub

Private Sub cmdsave_Click()
ssql = "INSERT INTO INCOMING_DELIVERY_TABLE(INVOICE_NO,year_today,DATED_ON, CUSTOMER_NAME, SERIAL_NO,QUANTITY,BRAND_NAME,TYPE,COLOR,FABRIC,description,UNIT_PRICE,TOTAL_PRICE)VALUES ('" & txtinvoiceno.Text & _
"','" & lblyear.Caption & "','" & dtdate.Value & "','" & cbocustomername.Text & "','" & txtserialno.Text & "','" & txtquantity.Text & "','" & cbobrandname.Text & "','" & cbotype.Text & "','" & cbocolor.Text & "','" & cbofabric.Text & _
"','" & cboarticle.Text & "','" & txtunitprice.Text & "','" & txttotalprice.Text & "')"
conn.Execute (ssql)
MsgBox "New Record has been saved", vbInformation, "Confirmation"

COMPUTE = Val(TXTPRODQTY.Text) - Val(txtquantity.Text)
TXTPRODQTY.Text = COMPUTE


ssql = "UPDATE PRODUCT_TABLE SET QTY='" & TXTPRODQTY.Text & "' WHERE ID=" & TXTPRODID.Text & ""
conn.Execute (ssql)

If Val(TXTPRODQTY.Text) < Val(txtcritical.Text) Then
MsgBox "You need to purchase additional items!", vbCritical, "Critical"
ssql = "update product_table set status='" & lblneedtopurchase.Caption & "' where id=" & TXTPRODID.Text & ""
conn.Execute (ssql)
End If


Call CMDENABLED(True, False, False, False)
Call txtenabled(True, True, False, False, False, False, False, False, False, False, False, False)
Call txtclear
Call listall
End Sub

Private Sub Command1_Click()
On Error Resume Next
ssql = "delete * from brand_name_table where id=" & txtidbrand.Text & ""
conn.Execute (ssql)
cbobrandname.Text = ""
Call cboaddchoices
End Sub

Private Sub Command2_Click()
If cbobrandname.Text = "" Then
MsgBox "You must write a new choice first!", vbExclamation, "Confirmation"
cbobrandname.Text = ""
cbobrandname.Enabled = True
Else
ssql = "insert into brand_name_table(brand_name)values('" & cbobrandname.Text & "')"
conn.Execute (ssql)
cbobrandname.Text = ""
Call cboaddchoices
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
ssql = "delete * from color_table where id=" & txtidcolor.Text & ""
conn.Execute (ssql)
cbocolor.Text = ""
Call cboaddcolor
End Sub

Private Sub Command4_Click()
If cbocolor.Text = "" Then
MsgBox "You must write a new choice first!", vbExclamation, "Confirmation"
cbocolor.Text = ""
cbocolor.Enabled = True
Else
ssql = "insert into color_table(color)values('" & cbocolor.Text & "')"
conn.Execute (ssql)
cbocolor.Text = ""
Call cboaddcolor
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
ssql = "delete * from fabric_table where id=" & txtidfabric.Text & ""
conn.Execute (ssql)
cbofabric.Text = ""
Call cboaddfabric
End Sub

Private Sub Command6_Click()
If cbofabric.Text = "" Then
MsgBox "You must write a new choice first!", vbExclamation, "Confirmation"
cbofabric.Text = ""
cbofabric.Enabled = True
Else
ssql = "insert into fabric_table(fabric)values('" & cbofabric.Text & "')"
conn.Execute (ssql)
cbofabric.Text = ""
Call cboaddfabric
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
lblyear.Caption = Format(Now, "yyyy")
connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
conn.Open connstrinG
Call listall
Call cboaddchoices
Call cboaddtype
Call cboaddcolor
Call cboaddfabric
Call txtenabled(True, True, False, False, False, False, False, False, False, False, False, False)
Call CMDENABLED(True, False, False, False)
Call txtclear

cboarticle.Clear
Set rs = conn.Execute("select * from product_table  WHERE QTY <> '" & TXTZERO.Text & "'ORDER by article")
Do While Not rs.EOF
With cboarticle
.AddItem rs!article
End With
rs.MoveNext
Loop

Set rs = conn.Execute("select * from customer_table order by customer_name")
Do While Not rs.EOF
With cbocustomername
.AddItem rs!customer_name
End With
rs.MoveNext
Loop
End Sub
Private Sub listall()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from incoming_delivery_table where year_today='" & lblyear.Caption & "' order BY invoice_no")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
        X.SubItems(1) = rs.Fields!invoice_no
        X.SubItems(2) = rs.Fields!year_today
        X.SubItems(3) = rs.Fields!Dated_on
        X.SubItems(4) = rs.Fields!customer_name
        X.SubItems(5) = rs.Fields!serial_no
        X.SubItems(6) = rs.Fields!quantity
        X.SubItems(7) = rs.Fields!brand_name
        X.SubItems(8) = rs.Fields!Type
        X.SubItems(9) = rs.Fields!Color
        X.SubItems(10) = rs.Fields!fabric
        X.SubItems(11) = rs.Fields!Description
        X.SubItems(12) = rs.Fields!unit_price
        X.SubItems(13) = rs.Fields!total_price
    rs.MoveNext
Loop
End Sub
Private Sub cboaddchoices()
cbobrandname.Clear
Set rs = conn.Execute("select * from brand_name_table order by brand_name")
Do While Not rs.EOF
With cbobrandname
.AddItem rs!brand_name
End With
rs.MoveNext
Loop
End Sub
Private Sub cboaddtype()
cbotype.Clear
Set rs = conn.Execute("select * from type_table order by type")
Do While Not rs.EOF
With cbotype
.AddItem rs!Type
End With
rs.MoveNext
Loop
End Sub
Private Sub cboaddcolor()
cbocolor.Clear
Set rs = conn.Execute("select * from color_table order by color")
Do While Not rs.EOF
With cbocolor
.AddItem rs!Color
End With
rs.MoveNext
Loop
End Sub
Private Sub cboaddfabric()
cbofabric.Clear
Set rs = conn.Execute("select * from fabric_table order by fabric")
Do While Not rs.EOF
With cbofabric
.AddItem rs!fabric
End With
rs.MoveNext
Loop
End Sub
Private Sub txtenabled(xinvoiceno, xdate, xserialno, xquantity, xbrandname, xtype, xcolor, xfabric, xarticle _
, xuntiprice, xtotalprice, xcustomername)
txtinvoiceno.Enabled = xinvoiceno
dtdate.Enabled = xdate
txtquantity.Enabled = xquantity
cbobrandname.Enabled = xbrandname
cbotype.Enabled = xtype
cbocolor.Enabled = xcolor
cbofabric.Enabled = xfabric
cboarticle.Enabled = xarticle
cbocustomername.Enabled = xcustomername

End Sub
Private Sub txtclear()
txtinvoiceno.Text = ""
txtserialno.Text = ""
txtquantity.Text = ""
cbobrandname.Text = ""
cbotype.Text = ""
cbocolor.Text = ""
cbofabric.Text = ""
cboarticle.Text = ""
txtunitprice.Text = ""
txttotalprice.Text = ""
cbocustomername.Text = ""
End Sub
Private Sub CMDENABLED(XADD, XDELETE, XSAVE, XCANCEL)
cmdadd.Enabled = XADD
cmddelete.Enabled = XDELETE
cmdsave.Enabled = XSAVE
cmdcancel.Enabled = XCANCEL
End Sub
Private Sub COMPUTEAMOUNT()
COMPUTE = Val(txtquantity.Text) * Val(txtunitprice.Text)
txttotalprice.Text = COMPUTE

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim col
col = ListView1.SelectedItem.Index
TXTID.Text = ListView1.ListItems.Item(col).Text
txtinvoiceno.Text = ListView1.ListItems.Item(col).SubItems(1)
lblyear.Caption = ListView1.ListItems.Item(col).SubItems(2)
dtdate.Value = ListView1.ListItems.Item(col).SubItems(3)
cbocustomername.Text = ListView1.ListItems.Item(col).SubItems(4)
txtserialno.Text = ListView1.ListItems.Item(col).SubItems(5)
txtquantity.Text = ListView1.ListItems.Item(col).SubItems(6)
cbobrandname.Text = ListView1.ListItems.Item(col).SubItems(7)
cbotype.Text = ListView1.ListItems.Item(col).SubItems(8)
cbocolor.Text = ListView1.ListItems.Item(col).SubItems(9)
cbofabric.Text = ListView1.ListItems.Item(col).SubItems(10)
cboarticle.Text = ListView1.ListItems.Item(col).SubItems(11)
txtunitprice.Text = ListView1.ListItems.Item(col).SubItems(12)
txttotalprice.Text = ListView1.ListItems.Item(col).SubItems(13)
Call CMDENABLED(False, True, False, True)
Call txtenabled(True, False, False, False, False, False, False, False, False, False, False, False)
End Sub

Private Sub txtquantity_Change()
Call COMPUTEAMOUNT
End Sub
Private Sub listbyinvoiceno()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from incoming_delivery_table where invoice_no='" & txtinvoiceno.Text & "' ORDER BY invoice_no")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
    X.SubItems(1) = rs.Fields!invoice_no
        X.SubItems(2) = rs.Fields!year_today
        X.SubItems(3) = rs.Fields!Dated_on
        X.SubItems(4) = rs.Fields!customer_name
        X.SubItems(5) = rs.Fields!serial_no
        X.SubItems(6) = rs.Fields!quantity
        X.SubItems(7) = rs.Fields!brand_name
        X.SubItems(8) = rs.Fields!Type
        X.SubItems(9) = rs.Fields!Color
        X.SubItems(10) = rs.Fields!fabric
        X.SubItems(11) = rs.Fields!Description
        X.SubItems(12) = rs.Fields!unit_price
        X.SubItems(13) = rs.Fields!total_price
    rs.MoveNext
Loop
End Sub
Private Sub listbydate()
ListView1.ListItems.Clear
Set rs = conn.Execute("SELECT * from incoming_delivery_table where dated_on='" & dtdate.Value & "' ORDER BY invoice_no")
Do While Not rs.EOF
    Set X = ListView1.ListItems.Add(, , rs.Fields(0))
       X.SubItems(1) = rs.Fields!invoice_no
        X.SubItems(2) = rs.Fields!year_today
        X.SubItems(3) = rs.Fields!Dated_on
        X.SubItems(4) = rs.Fields!customer_name
        X.SubItems(5) = rs.Fields!serial_no
        X.SubItems(6) = rs.Fields!quantity
        X.SubItems(7) = rs.Fields!brand_name
        X.SubItems(8) = rs.Fields!Type
        X.SubItems(9) = rs.Fields!Color
        X.SubItems(10) = rs.Fields!fabric
        X.SubItems(11) = rs.Fields!Description
        X.SubItems(12) = rs.Fields!unit_price
        X.SubItems(13) = rs.Fields!total_price
    rs.MoveNext
Loop
End Sub
