VERSION 5.00
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   "CENTRALIZED MUSICAL INSTRUMENT TRADING PLATFORM"
   ClientHeight    =   6915
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9360
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmmain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnucreate 
      Caption         =   "Create"
      Begin VB.Menu mnucategory 
         Caption         =   "Category"
      End
   End
   Begin VB.Menu mnuentry 
      Caption         =   "Entry"
      Begin VB.Menu mnunewstock 
         Caption         =   "New Stock"
      End
   End
   Begin VB.Menu mnuselect 
      Caption         =   "Select"
      Begin VB.Menu mnuitem 
         Caption         =   "Item"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnustock 
         Caption         =   "Stock"
      End
      Begin VB.Menu mnusale 
         Caption         =   "Sale"
      End
      Begin VB.Menu mnupurchase 
         Caption         =   "Purchase"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnucategory_Click()
    frmnew.Show
End Sub

Private Sub mnuexit_Click()
    a = MsgBox(" Do You Want to Close?", vbQuestion + vbYesNo)
    If a = vbYes Then
        End
    End If
End Sub

Private Sub mnuitem_Click()
    frmselect.Show
End Sub

Private Sub mnunewstock_Click()
    frmaddstock.Show
End Sub

Private Sub mnupurchase_Click()
    frmviewstock.Show
    frmviewstock.Frame1.Visible = True
    frmviewstock.cmbcategory.Visible = False
    frmviewstock.cmbitem.Visible = False
    frmviewstock.Label1.Visible = False
    frmviewstock.Label5.Visible = False
    frmviewstock.Label2.Caption = "Purchase View"
End Sub

Private Sub mnusale_Click()
    frmviewsale.Show
End Sub

Private Sub mnustock_Click()
    frmviewstock.Show
End Sub
