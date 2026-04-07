VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CENTRALIZED MUSICAL INSTRUMENT TRADING PLATFORM"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmlogin.frx":0000
   ScaleHeight     =   5535
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtpassword 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "@"
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtuser 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdlogin_Click()
    If txtuser.Text = "a" And txtpassword.Text = "b" Then
        Unload Me
        frmmain.Show
    Else
        MsgBox " Enter Correct Details"
    End If
End Sub
