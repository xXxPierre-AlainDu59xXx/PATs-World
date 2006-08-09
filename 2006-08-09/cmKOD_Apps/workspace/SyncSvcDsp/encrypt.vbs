Option Explicit

  Dim strPath
  Dim oDoc

  Dim oKODApps

set oKODApps = CreateObject("cmKOD_Apps.SyncSvcDsp")

strPath = "C:\VSS\Islands\KOD\cmKOD_Apps\workspace\SyncSvcDsp\"

With oKODApps

  Call .Encrypt(strPath & "wse_config.xml")

End With

msgbox "encrypted!"