VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmaddstock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CENTRALIZED MUSICAL INSTRUMENT TRADING PLATFORM"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmaddstock.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   10860
   Begin MSComCtl2.DTPicker dtpbill 
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   119799809
      CurrentDate     =   45678
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdphoto 
      BackColor       =   &H0080FFFF&
      Caption         =   "Photo"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H0080FFFF&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtselling 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   12
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox txtqty 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   10
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox txtcost 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   8
      Top             =   3240
      Width           =   3015
   End
   Begin VB.ComboBox cmbcategory 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txtno 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Item No.:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image imgmusical 
      Height          =   5055
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Price :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Price :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item No.:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Stock"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmaddstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdphoto_Click()
    CommonDialog1.Filter = "File(*.jpg)|*.jpg|File (*.bmp)|*.bmp"
    CommonDialog1.DefaultExt = "jpg"
    CommonDialog1.DialogTitle = "Select File"
    CommonDialog1.ShowOpen
    a = CommonDialog1.FileName
    imgmusical.Picture = LoadPicture(a)
End Sub

Private Sub cmdsave_Click()
    If cmdsave.Caption = "New" Then
        cmdsave.Caption = "Save"
        Call nextnum
    Else
        str = "select * from ItemEntry"
        rs.Open str, cn, 1, 3
        rs.AddNew
        rs.Fields("BillDate") = dtpbill.Value
        rs.Fields("ItemNo") = txtno.Text
        rs.Fields("Category") = cmbcategory.Text
        rs.Fields("ItemName") = txtname.Text
        rs.Fields("CostPrice") = txtcost.Text
        rs.Fields("Quantity") = txtqty.Text
        rs.Fields("Selling") = txtselling.Text
        rs.Fields("Photo") = a
        rs.Update
        rs.Close
        MsgBox " New Item of category " & cmbcategory.Text & "  is purchased"
        Unload Me
        
    End If
End Sub

Private Sub nextnum()
 str = "select distinct category from Category"
    rs.Open str, cn, 1, 3
    cmbcategory.Clear
    While Not rs.EOF
        cmbcategory.AddItem (rs.Fields("category"))
        rs.MoveNext
    Wend
    rs.Close
    str = "select * from ItemEntry order by ItemNo"
    rs.Open str, cn, 1, 3
    If rs.EOF Then
        txtno.Text = 1
    Else
        rs.MoveLast
        txtno.Text = rs.Fields("itemno") + 1
    End If
    rs.Close
    
End Sub

Private Sub Form_Load()
    dtpbill.Value = Date
End Sub



