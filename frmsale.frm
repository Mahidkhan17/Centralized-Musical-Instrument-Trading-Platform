VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmsale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CENTRALIZED MUSICAL INSTRUMENT TRADING PLATFORM"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmsale.frx":0000
   ScaleHeight     =   5535
   ScaleWidth      =   9705
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Details "
      Enabled         =   0   'False
      Height          =   3495
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   3615
      Begin VB.TextBox txtquantity 
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   24
         Top             =   2520
         Width           =   1350
      End
      Begin VB.TextBox txtitemno 
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
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   1350
      End
      Begin VB.TextBox txtmname 
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
         Left            =   1320
         TabIndex        =   18
         Top             =   840
         Width           =   2070
      End
      Begin VB.TextBox txtprice 
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
         Left            =   1320
         TabIndex        =   17
         Top             =   2040
         Width           =   1350
      End
      Begin VB.TextBox txtcategory 
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
         Left            =   1320
         TabIndex        =   16
         Top             =   1440
         Width           =   2070
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity :"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item No.:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Category :"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Price :"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
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
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0080FFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdsell 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sell"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   3495
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
      Begin VB.TextBox txttotal 
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   14
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtqty 
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
         Left            =   1320
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtpay 
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
         Left            =   1320
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker dtpbill 
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   119799809
         CurrentDate     =   45678
      End
      Begin VB.TextBox txtbillno 
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
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
         TabIndex        =   13
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity :"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay :"
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
         TabIndex        =   7
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
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
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No. :"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Integer
Private Sub cmdback_Click()
    Unload Me
End Sub

Private Sub cmdsell_Click()
    If Val(txtpay.Text) <> Val(txttotal.Text) Then
        MsgBox " The payment must be Same as that of Price"
        ElseIf txtname.Text = "" Then
            MsgBox " Write the Name First"
        Else
            str = "select * from Billing order by BillNo"
            rs.Open str, cn, 1, 3
            rs.AddNew
            rs.Fields("BillNo") = txtbillno.Text
            rs.Fields("BillDate") = dtpbill.Value
            rs.Fields("cname") = txtname.Text
            rs.Fields("Payment") = txtpay.Text
            rs.Fields("ItemNo") = txtitemno.Text
            rs.Fields("Category") = txtcategory.Text
            rs.Fields("ItemName") = txtmname.Text
            rs.Fields("Quantity") = txtqty.Text
            rs.Update
            rs.Close
            MsgBox " Musical Instrument Sold", vbInformation
            str = " select * from ItemEntry where ItemNo=" & c
            rs.Open str, cn, 1, 3
            rs.Fields("Quantity") = rs.Fields("Quantity") - txtqty.Text
            rs.Update
            rs.Close
            Unload Me
    End If
End Sub

Private Sub Form_Load()
dtpbill.Value = Date
If rs.State = 1 Then
    rs.Close
    Set rs = Nothing
End If
    
    str = " select * from ItemEntry where ItemNo=" & c
    rs.Open str, cn, 1, 3
    txtitemno.Text = rs.Fields("ItemNO")
    txtcategory.Text = rs.Fields("Category")
    txtmname.Text = rs.Fields("ItemName")
    txtprice.Text = rs.Fields("Selling")
    txtquantity.Text = rs.Fields("Quantity")
    q = rs.Fields("Quantity")
    rs.Close
    
    str = "select * from Billing order by BillNo"
    rs.Open str, cn, 1, 3
    If rs.EOF Then
        txtbillno.Text = 1
    Else
        rs.MoveLast
        txtbillno.Text = rs.Fields("BillNo") + 1
    End If
    rs.Close
    
End Sub

Private Sub txtpay_KeyPress(KeyAscii As Integer)
Call CHECKNUM(KeyAscii)
End Sub

Private Sub txtqty_Change()
If Val(txtqty.Text) > q Then
    MsgBox " Required Quantity is not present"
Else
txttotal.Text = Val(txtprice.Text) * Val(txtqty.Text)
End If
End Sub


