VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form invoice 
   Caption         =   "INVOICE"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTot 
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton NEW_FIX 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   17
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   11280
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   10575
      Begin VB.TextBox TxtQty 
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtDesc 
         Height          =   615
         Index           =   0
         Left            =   4320
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtAmt 
         Height          =   495
         Index           =   0
         Left            =   8280
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label LblQty 
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblDesc 
         Caption         =   "Description"
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblAmt 
         Caption         =   "Amount"
         Height          =   255
         Index           =   0
         Left            =   7440
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   1920
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98697217
         CurrentDate     =   42405
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   " Invoice Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Invoice date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   8160
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mmsql As String
Dim connstrinG As String
Dim cn As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim objIE As SHDocVw.InternetExplorer
Dim OrgBox As HTMLInputElement
Dim URL As String
Dim strHTML As String
Dim bNumToWordInit As Boolean
Dim strNumberWords(19) As String
Dim strNumberGroups(10) As String
Dim strNumberTenWords(2 To 9) As String
Dim strNumberDecimalWords(20 To 90) As String
Dim db_name, db_server, db_port, db_user, db_pass, constr As String
Dim msql As String
Dim you As Date
Dim ie As Object
Dim i As Integer
Dim j As Integer
Option Explicit

Private Sub Command1_Click()
Dim new_index As Integer
new_index = Frame2.UBound + 1
    
    Load Frame2(new_index)
    Frame2(new_index).Move Frame2(0).Left, Frame2(new_index - 1).Top + Frame2(0).Height + 120
    Frame2(new_index).Visible = True


    Load LblQty(new_index)
    Set LblQty(new_index).Container = Frame2(new_index)
    LblQty(new_index).Visible = True

    Load TxtQty(new_index)
    Set TxtQty(new_index).Container = Frame2(new_index)
    TxtQty(new_index).Visible = True
    TxtQty(new_index) = ""
    
    Load LblDesc(new_index)
    Set LblDesc(new_index).Container = Frame2(new_index)
    LblDesc(new_index).Visible = True

    Load txtDesc(new_index)
    Set txtDesc(new_index).Container = Frame2(new_index)
    txtDesc(new_index).Visible = True
    txtDesc(new_index) = ""

    Load LblAmt(new_index)
    Set LblAmt(new_index).Container = Frame2(new_index)
    LblAmt(new_index).Visible = True

    Load txtAmt(new_index)
    Set txtAmt(new_index).Container = Frame2(new_index)
    txtAmt(new_index).Visible = True
    txtAmt(new_index) = ""
    
Me.Height = Frame2(new_index).Top + Frame2(new_index).Height + 120 + (Me.Height - Me.ScaleHeight)


    
End Sub

Private Sub Form_Load()
Label3 = Now()
 'Text2.Text = RandomNumber(1000000, 1)
 'db_name = "micmac_cerm_lat"
    ''db_server = "200.0.0.114"
    'db_server = "localhost"
    'db_port = ""    'default port is 3306
    'db_user = "root"
    'db_pass = ""

'Set cn = New ADODB.Connection
  '' cn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=" & db_server & ";" _
                        & "DATABASE=" & db_name & ";" _
                        & "UID=" & db_user & ";PWD=" & db_pass & "; OPTION=3"

'cn.Open

connstrinG = "driver={Microsoft access driver (*.mdb)};dbq=" & App.Path & "/inventory.mdb"
cn.Open connstrinG


msql = "SELECT * FROM Invoice"
If rs1.State = adStateOpen Then
Set rs1 = Nothing
End If
rs1.Open msql, cn, adOpenDynamic, adLockOptimistic

Dim ms As String
ms = Len(rs1!InvoiceNum)
rs1!InvoiceNum = rs1!InvoiceNum + 1

If Len(rs1!InvoiceNum) = 1 Then rs1!InvoiceNum = "000000" & rs1!InvoiceNum
If Len(rs1!InvoiceNum) = 2 Then rs1!InvoiceNum = "00000" & rs1!InvoiceNum
If Len(rs1!InvoiceNum) = 3 Then rs1!InvoiceNum = "0000" & rs1!InvoiceNum
If Len(rs1!InvoiceNum) = 4 Then rs1!InvoiceNum = "000" & rs1!InvoiceNum
If Len(rs1!InvoiceNum) = 5 Then rs1!InvoiceNum = "00" & rs1!InvoiceNum
If Len(rs1!InvoiceNum) = 6 Then rs1!InvoiceNum = "0" & rs1!InvoiceNum

rs1.Update
Text3.Text = rs1!InvoiceNum
Command1.Enabled = True

End Sub

Private Sub NEW_FIX_Click()

''strHTML = "<HTML><HEAD><TITLE>Details</TITLE></HEAD><BODY></BODY></HTML>"
Dim rep As String
Label3 = Replace(Now(), ":", "_")
rep = Replace(Label3, " ", "_")
you = DTPicker1.Value

Open "C:\AddaxInvoice\addaxInvoice_" & rep & ".html" For Output As #1
'Print #1, "<!DOCTYPE html>"
Print #1, "<HTML>"
Print #1, "<HEAD>"
Print #1, "<TITLE>Details</TITLE>"
Print #1, "</HEAD>"
Print #1, "<BODY style=""border:1px solid #acacac; width:45em; margin:auto; font-family:Andalus;"">"

Print #1, "<div class=container style=""width:45em"">"

Print #1, "<div>"

Print #1, "<h1 style=""text-align:center; border:1px solid #aaa; background-color:#CC0000; color:#fff"">INVOICE</h1>"
Print #1, "</div>"

Print #1, "<div>"
        Print #1, "<table style=""border:1px solid #aaa; width:45em;"">"
'Print #1, "<tr>"
       'Print #1, "<td><strong>Coop Name:</strong></td>"
        'Print #1, "<td> <strong>Branch office:</strong></td>"
        'Print #1, "<td><strong> Tel:</strong></td>"
        'Print #1, "<td><strong> Email:</strong></td>"
      'Print #1, "</tr>"
     
     Print #1, "<tr>"
        Print #1, "<td><strong> Addax Staff Cooperative Society Ltd</strong></td>"
      Print #1, "</tr>"
       Print #1, "<tr>"
         Print #1, "<td><strong> 32, Ozumba Mbadiwe Street,Victoria Isalnd, Lagos.</strong></td>"
      Print #1, "</tr>"
       Print #1, "<tr>"
       Print #1, "<td><strong>Phone: 08087182057</strong></td>"
      Print #1, "</tr>"
       Print #1, "<tr>"
         Print #1, "<td><strong>Email: addax.coop@addaxpetroleum.com</strong></td>"
      Print #1, "</tr>"
   Print #1, "</table>"
   Print #1, "</div>"
   
   Print #1, "<br/>"
  ' Print #1, "<br/>"
   'Print #1, "<br/>"
      
      'Print #1, "<div style=""border:1px solid #acacac;  max-width:35em"">"
       ' Print #1, "<strong>Billed to: </strong>" & Text2.Text & "<br/>"
        'Print #1, "<strong>Address:</strong>" & Text1.Text & ""
      'Print #1, "</div>"
      
    Print #1, "<div>"
      Print #1, "<table style=""border:1px solid #aaa; width:45em; max-height:10em"">"
        Print #1, "<tr>"
         Print #1, "<td rowspan=""2"" style=""word-wrap:break-word; max-width:25em""><strong>Billed to:</strong> " & Text2.Text & "<br/>" & Text1.Text & "</td>"
         'Print #1, "<td colspan=""2""; style=""color:red;""><h2 style=""padding-left:100px;""></h2></td>"
          Print #1, "<td""><strong></strong></td>"
          Print #1, "<td""><strong></strong></td>"
          Print #1, "<td><strong></strong></td>"
          Print #1, "<td""><strong></strong></td>"
           Print #1, "<td><strong>Invoice Num: " & Text3.Text & "<br/>" & " Invoice Date: " & you & "<br/>" & " Date: " & Now() & "<br/>" & "</strong></td>"
           'Print #1, "<td><strong>Invoice Date: " & you & "</strong></td>"
           'Print #1, "<td><strong>Date: " & Now() & "</strong></td>"
        Print #1, "</tr>"
        
        'Print #1, "<tr>"
           ' Print #1, "<td><strong></strong></td>"
            'Print #1, "<td><strong></strong></td>"
           ' Print #1, "<td><strong></strong></td>"
            'Print #1, "<td><strong>Invoice Date: " & you & "</strong></td>"
        'Print #1, "</tr>"
        
         'Print #1, "<tr>"
           ' Print #1, "<td><strong></strong></td>"
           ' Print #1, "<td><strong></strong></td>"
           ' Print #1, "<td><strong></strong></td>"
           ' Print #1, "<td><strong>Invoice Num: " & Text3.Text & "</strong></td>"
        'Print #1, "</tr>"
        
       ' Print #1, "<tr>"
          '  Print #1, "<td><strong></strong></td>"
            'Print #1, "<td><strong></strong></td>"
            'Print #1, "<td><strong></strong></td>"
            'Print #1, "<td><strong></strong></td>"
            'Print #1, "<td><strong>Date: " & Now() & "</strong></td>"
      'Print #1, "</tr>"
      
      Print #1, "</table>"
      
        Print #1, "<br />"
     Print #1, "<div>"
       Print #1, "<table style=""border:1px solid #acacac; width:45em;>"
          Print #1, "<tr style=""border:1px solid #acacac"">"
            Print #1, "<td style=""border:1px solid #acacac""><strong>Qty.</strong> </td>"
            Print #1, "<td style=""border:1px solid #acacac""><strong>Description</strong></td>"
            Print #1, "<td style=""border:1px solid #acacac""><strong>Amount</strong></td>"
            Print #1, "<td style=""border:1px solid #acacac""><strong>Total</strong></td>"
            Print #1, "</tr>"
          
          For i = 0 To Frame2.UBound
          Dim tot As String
          Dim sumtot As Double
          Dim amttot As Double
         Dim varAmt As Double
         Dim lDecimalPos As Long
         varAmt = txtAmt(i)
         tot = TxtQty(i).Text * varAmt
         sumtot = Val(sumtot) + Val(tot)
         amttot = Val(amttot) + Val(txtAmt(i).Text)
         
           
          
            Print #1, "<tr>"
            Print #1, "<td style=""border:1px solid #acacac"">" & TxtQty(i).Text & "</td>"
            Print #1, "<td style=""border:1px solid #acacac"">" & txtDesc(i).Text & "</td>"
            Print #1, "<td style=""border:1px solid #acacac"">" & FormatNumber(txtAmt(i).Text, 2, False, True, True) & "</td>"
            Print #1, "<td style=""border:1px solid #acacac"">" & FormatNumber(tot, 2, False, True, True) & "</td>"
            
        Print #1, "</tr>"
        Next
    
        
         Print #1, "<tr>"
            Print #1, "<td style=""border:1px solid #acacac""></td>"
            Print #1, "<td style=""border:1px solid #acacac""></td>"
            Print #1, "<td style=""border:1px solid #acacac""></td>"
            Print #1, "<td style=""border:1px solid #acacac""></td>"
            
        Print #1, "</tr>"
        
        Print #1, "<tr>"
            Print #1, "<td style=""border:1px solid #acacac""></td>"
            Print #1, "<td><strong>Total</strong></td>"
            Print #1, "<td style=""border:1px solid #acacac"">" & FormatNumber(tot, 2, False, True, True) & "</td>"
            Print #1, "<td style=""border:1px solid #acacac"">" & FormatNumber(sumtot, 2, False, True, True) & "</td>"
            
        Print #1, "</tr>"
         Print #1, "</table>"
         
         Print #1, "<br />"
         Print #1, "<br />"

         Print #1, "<p><strong>Amount in words:</strong> " & NumberToWords(sumtot) & " Naira" & " </p>"
       Print #1, "</div>"
       
       Print #1, "<br />"
    
    Print #1, "<div>"
Print #1, "<p> <strong>Authorise Signature:</strong> ___________________<strong>Date:</strong>_________________</p>"
Print #1, "</div>"

Print #1, "</div>" 'container div

    Print #1, "<br />"
    Print #1, "<div>"
        Print #1, "<footer style=""text-align:center; border:1px solid #aaa; background-color:#CC0000; color:#fff"">powered by Sekat &copy; 2016  </footer>"
    Print #1, "</div>"

Print #1, "</BODY>"
Print #1, "</HTML>"

Close #1

Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.Navigate "C:\AddaxInvoice\addaxInvoice_" & rep & ".html"


 MsgBox "saved successfully"
End Sub

Private Sub InitNumToWords()

'array for values 0 to 19
strNumberWords(0) = "zero"
strNumberWords(1) = "one"
strNumberWords(2) = "two"
strNumberWords(3) = "three"
strNumberWords(4) = "four"
strNumberWords(5) = "five"
strNumberWords(6) = "six"
strNumberWords(7) = "seven"
strNumberWords(8) = "eight"
strNumberWords(9) = "nine"
strNumberWords(10) = "ten"
strNumberWords(11) = "eleven"
strNumberWords(12) = "twelve"
strNumberWords(13) = "thirteen"
strNumberWords(14) = "fourteen"
strNumberWords(15) = "fifteen"
strNumberWords(16) = "sixteen"
strNumberWords(17) = "seventeen"
strNumberWords(18) = "eightteen"
strNumberWords(19) = "nineteen"

'ArrayOfDecimal
'strNumberDecimalWords(10) = "ten"
strNumberDecimalWords(20) = "twenty"
strNumberDecimalWords(30) = "thirty"
strNumberDecimalWords(40) = "forty"
strNumberDecimalWords(50) = "fifty"
strNumberDecimalWords(60) = "sixty"
strNumberDecimalWords(70) = "seventy"
strNumberDecimalWords(80) = "eighty"
strNumberDecimalWords(90) = "ninty"

'array for 10's digit
strNumberTenWords(2) = "twenty"
strNumberTenWords(3) = "thirty"
strNumberTenWords(4) = "forty"
strNumberTenWords(5) = "fifty"
strNumberTenWords(6) = "sixty"
strNumberTenWords(7) = "seventy"
strNumberTenWords(8) = "eighty"
strNumberTenWords(9) = "ninety"


'array for number groups
strNumberGroups(1) = "thousand"
strNumberGroups(2) = "million"
strNumberGroups(3) = "billion"
strNumberGroups(4) = "trillion"

'set flag
bNumToWordInit = True

End Sub


Public Function NumberToWords(ByVal vNumber, Optional bMoney As Boolean = False) As String

Dim strTemp As String
Dim strChar As String
Dim strWhole As String
Dim strDecimal As String

Dim lNumberGroupCount As Long
Dim lDecimalPos As Long
Dim loop1 As Long
Dim dTemp As Double


'intialize the arrays (if not yet done)
If Not bNumToWordInit Then
    InitNumToWords
End If

'make sure it's a valid number
If Not IsNumeric(vNumber) Then
    NumberToWords = "Invalid Number"
    Exit Function
End If

If Abs(Val(vNumber)) >= 999999999999999# Then
    NumberToWords = "Number too big"
    Exit Function
End If

strTemp = CStr(vNumber)

'clean up non-numerics
strTemp = Replace(strTemp, "$", "")
strTemp = Replace(strTemp, ",", "")
strTemp = Replace(strTemp, " ", "")

'convert '(number)' to '-number'
If Left$(strTemp, 1) = "(" And Right$(strTemp, 1) = ")" Then
    strTemp = "-" & Mid$(strTemp, 2, Len(strTemp) - 2)
End If

'find the decimal
lDecimalPos = InStr(1, strTemp, ".")

'if there is a decimal
If lDecimalPos > 0 Then
    'get integer part
    strWhole = Left$(strTemp, lDecimalPos - 1)
    
    'get the fractional part
    strDecimal = Right$(strTemp, Len(strTemp) - lDecimalPos)
    
    If strDecimal = "" Then strDecimal = "0"
    
    'if optional money param is true
    If bMoney Then
        'handle >2 digit decimal
        If Len(strDecimal) > 2 Then
            strDecimal = CStr(CInt(Val("." & strDecimal) * 100))
        'handle <2 digit decimal
        ElseIf Len(strDecimal) < 2 Then
            strDecimal = Left(strDecimal & "00", 2)
        End If
    End If
Else 'otherwise
    If bMoney Then
        strDecimal = "0"
    End If
    strWhole = strTemp
End If

    
vNumber = Val(strWhole)

'handle negatives
If vNumber < 0 Then
    NumberToWords = "negative"
    vNumber = Abs(vNumber)
End If

'if the number is at least 1
If vNumber > 0 Then

    'get count of three digit number groups (log base 1000)
    lNumberGroupCount = Int(Log(CDbl(vNumber)) / Log(1000))
    
    'if the number has more that the "hundreds" group
    If lNumberGroupCount > 0 Then
        'get the hundres value of the current group
        dTemp = vNumber / (1000 ^ lNumberGroupCount)
        dTemp = Int(dTemp)
        
        'build the output by recursively calling this function and
        'getting the Group word from the array
        NumberToWords = Trim$(NumberToWords(dTemp)) & " " & strNumberGroups(lNumberGroupCount)
        'if the remainder is more than 0
        If vNumber - (dTemp * 1000 ^ lNumberGroupCount) > 0 Then
               NumberToWords = NumberToWords & " " & _
                 NumberToWords(vNumber - (dTemp * 1000 ^ lNumberGroupCount))
        End If
    Else
    
        'if the number is at least 100
        If vNumber > 99 Then
        
            'get the number word for the hundreds digit
            NumberToWords = Trim$(NumberToWords & _
               " " & strNumberWords(Int(vNumber / 100)) & " hundred ")
            
            'subtract from the number
            vNumber = vNumber Mod 100 '- 100 * Int(vNumber / 100)
            
            If vNumber > 0 Then
            NumberToWords = NumberToWords & "  and"
            End If
        End If
    
        'if the remaining value is at least 20
        If vNumber > 19 Then
            
            'append the the number word for the 10's digit
            NumberToWords = Trim$(NumberToWords & _
               " " & strNumberTenWords(Int(vNumber / 10)))
            
            'subtract from the number
            vNumber = vNumber Mod 10 '- 10 * Int(vNumber / 10)
            
            'if the remainder is at least 1
            If vNumber > 0 Then
                'append the the number word for the 1's digit
                NumberToWords = Trim$(NumberToWords & " " & strNumberWords(vNumber))
            End If
            
            'If vNumber = strDecimal Then NumberToWords = NumberToWords & " kobo"
            
        Else ' otherwise (less than 20)
            'if the remainder is at least 1
            If vNumber > 0 Then
                'append the number word for less than 20
                NumberToWords = Trim$(NumberToWords & " " & strNumberWords(vNumber))
            End If
            
            If vNumber > 0 Then
                'append the number word for less than 20
                NumberToWords = Trim$(NumberToWords & " " & strNumberWords(vNumber))
            End If
        End If
    End If
Else 'otherwise (less than 1 i.e. 0)
    NumberToWords = "zero"
End If

'if optional Money parameter is true
If bMoney Then
    'format as money
    NumberToWords = Trim$(NumberToWords & " dollars and")
    NumberToWords = NumberToWords & " " & NumberToWords(strDecimal)
    NumberToWords = Trim$(NumberToWords & " cents")
Else
    'if there is a decimal portion
    ''strDecimal = hassan
    
    If strDecimal <> "" Then
    
    If Len(strDecimal) = 1 Then
    strDecimal = strDecimal & "0"
    NumberToWords = Trim$(NumberToWords & " Naira " & "and")
    NumberToWords = NumberToWords & " " & strNumberDecimalWords(Val(strDecimal))
            NumberToWords = NumberToWords & " kobo"
    Else
    
        'append the word point
        NumberToWords = Trim$(NumberToWords & " Naira " & "and")
        'build the decimal portion
        'For loop1 = 1 To Len(strDecimal)
            'strChar = Mid$(strDecimal, loop1, 1)
            NumberToWords = NumberToWords & " " & strNumberWords(Val(strDecimal))
            NumberToWords = NumberToWords & " kobo"
       ' Next 'loop1
        End If
    End If
End If
End Function


Private Sub txtAmt_Change(Index As Integer)
txtTot = 0
Dim i As Integer
For i = 0 To Frame2.UBound
txtTot = Val(txtTot) + Val(txtAmt(i))
Next i
End Sub



