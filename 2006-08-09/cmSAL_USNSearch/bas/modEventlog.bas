Attribute VB_Name = "modEventLog"
Option Explicit

Private Declare Function RegisterEventSource Lib "advapi32" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32" (ByVal hEventLog As Long) As Boolean
Private Declare Function ReportEvent Lib "advapi32" Alias "ReportEventA" _
                        (ByVal hEventLog As Long, _
                        ByVal wType As Long, _
                        ByVal wCategory As Long, _
                        ByVal dwEventID As Long, _
                        ByVal lpUserSid As Long, _
                        ByVal wNumStrings As Long, _
                        ByVal dwDataSize As Long, _
                        lpStrings As Any, _
                        lpRawData As Any) As Long

'   Event Type Constants
Public Const EVENTLOG_SUCCESS = &H0            'Success event
Public Const EVENTLOG_ERROR_TYPE = &H1         'Error event
Public Const EVENTLOG_WARNING_TYPE = &H2       'Warning event
Public Const EVENTLOG_INFORMATION_TYPE = &H4   'Information event
Public Const EVENTLOG_AUDIT_SUCCESS = &H8      'Success audit event
Public Const EVENTLOG_AUDIT_FAILURE = &H10     'Failure audit event

Public Sub LogEvent(strSource As String, dEventType As Long, ByRef strError As String, Optional iEventID As Integer = 26, Optional iCategory As Integer = 26)
  
    Dim hEvt As Long
  
  hEvt = RegisterEventSource(vbNullChar, strSource)
  If hEvt Then
    Call ReportEvent(hEvt, dEventType, iCategory, iEventID, 0, 1, 0, strError, vbNullChar)
    DeregisterEventSource (hEvt)
  End If

End Sub

