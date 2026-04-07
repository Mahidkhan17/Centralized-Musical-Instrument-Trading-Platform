VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmviewstock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CENTRALIZED MUSICAL INSTRUMENT TRADING PLATFORM"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmviewstock.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   11280
   Begin VB.Frame framed 
      BackColor       =   &H0080FFFF&
      Caption         =   "Duration"
      Height          =   855
      Left            =   5640
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdview 
         BackColor       =   &H0080FF80&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dof 
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   127533057
         CurrentDate     =   45678
      End
      Begin MSComCtl2.DTPicker dot 
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   127533057
         CurrentDate     =   45678
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
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
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
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
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   5415
      Begin VB.OptionButton optmonth 
         BackColor       =   &H0080FFFF&
         Caption         =   "Monthly"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optyear 
         BackColor       =   &H0080FFFF&
         Caption         =   "Yearly"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optall 
         BackColor       =   &H0080FFFF&
         Caption         =   "All"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optduration 
         BackColor       =   &H0080FFFF&
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0080FF80&
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txttotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1320
      TabIndex        =   6
      Top             =   7800
      Width           =   870
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmviewstock.frx":5E8B3
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8916
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ItemNo"
         Caption         =   "Item No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Category"
         Caption         =   "Category"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ItemName"
         Caption         =   "Item Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CostPrice"
         Caption         =   "Cost Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Quantity"
         Caption         =   "Quantity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Selling"
         Caption         =   "Selling"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "BillDate"
         Caption         =   "Bill Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbitem 
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
      Left            =   7920
      TabIndex        =   2
      Top             =   720
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
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4560
      Top             =   8160
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BSCIT\Musical\Musical.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BSCIT\Musical\Musical.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select  * from ItemEntry"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label6 
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
      Left            =   480
      TabIndex        =   7
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock View"
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
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   720
      Width           =   1455
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
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmviewstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbcategory_Click()
     str = "select * from ItemEntry where Category='" & cmbcategory.Text & "'"
     rs.Open str, cn, 1, 3
     While Not rs.EOF
        cmbitem.AddItem (rs.Fields("ItemName"))
        rs.MoveNext
     Wend
     rs.Close
End Sub

Private Sub cmbitem_Click()
     str = "select * from ItemEntry where Category='" & cmbcategory.Text & "'and ItemName='" & cmbitem.Text & "'"
     Adodc1.RecordSource = str
       Adodc1.Refresh
     txttotal.Text = Adodc1.Recordset.RecordCount

End Sub

Private Sub cmdback_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    str = "select * from ItemEntry"
    rs.Open str, cn, 1, 3
    While Not rs.EOF
        cmbcategory.AddItem (rs.Fields("Category"))
        rs.MoveNext
    Wend
    rs.Close
     str = "select * from ItemEntry"
     Adodc1.RecordSource = str
     Adodc1.Refresh
     txttotal.Text = Adodc1.Recordset.RecordCount
End Sub
Private Sub cmdview_Click()
    str = "Select * From ItemEntry where BillDate between #" & dof.Value & "# " & "and" & " #" & dot.Value & "#"
     Adodc1.RecordSource = str
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    txttotal.Text = Adodc1.Recordset.RecordCount
    
End Sub
Private Sub optall_Click()
str = "SELECT * FROM ItemEntry ORDER BY ItemNo"
 Adodc1.RecordSource = str
 Adodc1.Refresh
 Set DataGrid1.DataSource = Adodc1
    txttotal.Text = Adodc1.Recordset.RecordCount
 
End Sub

Private Sub optduration_Click()
    framed.Visible = True
End Sub

Private Sub optmonth_Click()
    str = "SELECT * FROM ItemEntry WHERE MONTH(BillDate)  = " _
        & Month(Date) & " AND YEAR(BillDate)  = " & Year(Date) & " ORDER BY ItemNo"

    Adodc1.RecordSource = str
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    txttotal.Text = Adodc1.Recordset.RecordCount
    
End Sub

Private Sub optyear_Click()
    str = "SELECT * FROM ItemEntry WHERE YEAR(BillDate)  = " _
           & Year(Date) & " ORDER BY ItemNo"
         
    Adodc1.RecordSource = str
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    txttotal.Text = Adodc1.Recordset.RecordCount
         
End Sub
