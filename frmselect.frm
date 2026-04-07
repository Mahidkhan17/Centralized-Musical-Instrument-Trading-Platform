VERSION 5.00
Begin VB.Form frmselect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CENTRALIZED MUSICAL INSTRUMENT TRADING PLATFORM"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmselect.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   12450
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
      Left            =   2520
      TabIndex        =   19
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Details "
      Enabled         =   0   'False
      Height          =   3375
      Left            =   8760
      TabIndex        =   9
      Top             =   2280
      Width           =   3615
      Begin VB.TextBox txtqty 
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
         TabIndex        =   20
         Top             =   2640
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
         TabIndex        =   13
         Top             =   1440
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
         TabIndex        =   12
         Top             =   2040
         Width           =   1350
      End
      Begin VB.TextBox txtname 
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
         TabIndex        =   11
         Top             =   840
         Width           =   2070
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
         TabIndex        =   10
         Top             =   360
         Width           =   1350
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
         TabIndex        =   21
         Top             =   2640
         Width           =   1095
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
         TabIndex        =   17
         Top             =   840
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
         TabIndex        =   16
         Top             =   2040
         Width           =   855
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
         TabIndex        =   15
         Top             =   1440
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
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Height          =   975
      Left            =   8760
      TabIndex        =   4
      Top             =   5760
      Width           =   3615
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H0080FFFF&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdselect 
         BackColor       =   &H0080FFFF&
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8535
      Begin VB.Image imgmusical 
         Height          =   6375
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   7800
      Width           =   8535
      Begin VB.CommandButton cmdlast 
         BackColor       =   &H0080FFFF&
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdfirst 
         BackColor       =   &H0080FFFF&
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H0080FFFF&
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080FFFF&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbcategory_Click()
    str = "select * from ItemEntry where Quantity>0 and Category='" & cmbcategory.Text & "'"
    rs.Open str, cn, 1, 3
    Call showdata
End Sub

Private Sub cmdback_Click()
    Unload Me
    rs.Close
End Sub

Private Sub cmdfirst_Click()
    rs.MoveFirst
    Call showdata
End Sub

Private Sub cmdlast_Click()
    rs.MoveLast
    Call showdata
End Sub

Private Sub cmdNext_Click()
    rs.MoveNext
       If rs.EOF = True Then
        rs.MoveFirst
    End If
    Call showdata
End Sub

Private Sub cmdPrevious_Click()
    rs.MovePrevious
    If rs.BOF = True Then
        rs.MoveLast
    End If
Call showdata

End Sub

Private Sub cmdselect_Click()
    c = Val(txtitemno.Text)
    frmsale.Show
    Unload Me
    
End Sub

Private Sub Form_Load()
    str = "select * from Category"
    rs.Open str, cn, 1, 3
    While Not rs.EOF
        cmbcategory.AddItem (rs.Fields("category"))
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Function showdata()
    txtitemno.Text = rs.Fields("ItemNO")
    txtcategory.Text = rs.Fields("Category")
    txtname.Text = rs.Fields("ItemName")
    txtprice.Text = rs.Fields("Selling")
    imgmusical.Picture = LoadPicture(rs.Fields("Photo"))
    txtqty.Text = rs.Fields("Quantity")
End Function
