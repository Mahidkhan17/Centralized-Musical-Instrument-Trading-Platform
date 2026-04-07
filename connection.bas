Attribute VB_Name = "connection"
Public cn As New ADODB.connection
Public rs As New ADODB.Recordset
Public str As String
Public c As Integer
Sub main()
str = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" & App.Path & "\Musical.mdb; persist Security Info=False;"
         cn.Open str
        frmsplash.Show

End Sub
Function CHECKTEXT(K As Integer)
Select Case K
        Case 65 To 90, 97 To 122, 8, 32
                 K = K
        Case Else
                 K = 0
End Select
CHECKTEXT = K
End Function

Function CHECKNUM(K As Integer)
Select Case K
        Case 48 To 57, 8
                 K = K
        Case Else
                 K = 0
End Select
CHECKNUM = K
End Function
    
    
    
  
   

