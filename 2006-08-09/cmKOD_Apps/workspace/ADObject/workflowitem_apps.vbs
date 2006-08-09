'
' apps.vbs
'
' Purpose : contains all KOD applications (target system SAL wrapper)
'
' Copyright © 2004-2005 by econet AG
'

Option Explicit

'
' adobject_exec
'
' Purpose       : vbs wrapper for KOD app "cmKOD_apps.ADObject"
' Parameters    : none
' Return value  : none
'
' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-07-21
'
Public Function adobject_exec()

    ' identify
    Dim strThisMethod
    Dim oKODApps

  ' activate error handler
  on error resume next

  ' identify
  strThisMethod = "adobject_exec"

  ' echo
  Call Echo(strThisMethod, "")

  set oKODApps = CreateObject("cmKOD_apps.ADObject")
  if not IsObject(oKODApps) then
    Call HandleException(cmErr, true, 100027, strThisMethod, "failed to create new instance of 'cmKOD_apps.ADObject'", cmThisScript)
  else
    Call oKODApps.Exec(cmData, cmConfigData, cmErr)
  end if

  ' final error check
  Call HandleException(cmErr, true, 100028, strThisMethod, TXT_ERR_UNEXPECTED, "")

end function

'
' syncsvcdsp_exec
'
' Purpose       : vbs wrapper for KOD app "cmKOD_apps.SyncSvcDsp"
' Parameters    : none
' Return value  : none
'
' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-16
'
Public Function syncsvcdsp_exec()

    ' identify
    Dim strThisMethod
    Dim oKODApps

  ' activate error handler
  on error resume next

  ' identify
  strThisMethod = "syncsvcdsp_exec"

  ' echo
  Call Echo(strThisMethod, "")

  set oKODApps = CreateObject("cmKOD_apps.SyncSvcDsp")
  if not IsObject(oKODApps) then
    Call HandleException(cmErr, true, -999, strThisMethod, "failed to create new instance of 'cmKOD_apps.SyncSvcDsp'", cmThisScript)
  else
    Call oKODApps.Exec(cmData, cmConfigData, cmErr)
  end if

  ' final error check
  Call HandleException(cmErr, true, -999, strThisMethod, TXT_ERR_UNEXPECTED, "")

end function
