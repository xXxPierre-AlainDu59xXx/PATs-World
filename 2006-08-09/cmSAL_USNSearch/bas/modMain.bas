Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()

  Dim cSearch As New CUSNSearch
  Call cSearch.Execute(Command$)
  
End Sub


