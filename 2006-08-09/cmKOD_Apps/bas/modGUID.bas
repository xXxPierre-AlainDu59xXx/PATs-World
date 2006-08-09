Attribute VB_Name = "modCreateGUID"
' declare structs
Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

' decalres API functions
Declare Function CoCreateGuid Lib "ole32.dll" (GUID As GUID) As Long
Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long

'
' CreateGUID
'
' Purpose       : creates GUID (e.g. for creating unique filenames)
' Parameters    : none
' Return value  : string (GUID)
'
' Comment       : project public function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036, 2003-09-29
'
'
Function CreateGUID() As String
    
  ' locals
  Dim pguid As GUID
  Dim szGUID As String
  Dim lRet As Long
        
  ' create GUID
  lRet = CoCreateGuid(pguid)
  szGUID = String$(38, 0)
  StringFromGUID2 pguid, StrPtr(szGUID), Len(szGUID) + 1
  
  ' remove leading and trailing curly brackets
  CreateGUID = CStr(Mid(szGUID, 2, 36))

End Function

