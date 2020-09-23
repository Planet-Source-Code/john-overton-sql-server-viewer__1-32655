Attribute VB_Name = "Module1"
Option Explicit

Public strServer As String
Public strDatabase As String
Public strUID As String
Public strPWD As String
Public cnOpen As Boolean
Public cn As ADODB.Connection
Public cmd As ADODB.Command
Public Click As Boolean
Public DGError




Public Sub Open_cn()
On Error GoTo errhandler
 Set cmd = New ADODB.Command
    Set cn = New ADODB.Connection
  
    With cn
        .Provider = "MSDASQL;DRIVER={SQL Server};SERVER=" & strServer & ";trusted_connection=no;user id= " & strUID & ";password=" & strPWD & " ;database=" & strDatabase & ";"
        .Open
        
    End With
    cnOpen = True
    
    
 Exit Sub
errhandler:
 Call MsgBox("Connection Error.", vbOKOnly, "Error")
 Form1.sb.Panels(1).Text = "Disconnected"
 strServer = ""

 
 cnOpen = False

End Sub
Public Sub Close_cn()
     If cnOpen = True Then
     
        cn.Close
        Set cn = Nothing
        
         cnOpen = False
    End If
      
End Sub
