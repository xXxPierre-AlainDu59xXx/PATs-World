Option Explicit

  Dim strPath
  Dim oCmData, oCmConfigData, oCmErr
  Dim oDoc

  Dim oKODApps

set oKODApps = CreateObject("cmKOD_Apps.SyncSvcDsp")

Set oCmErr = CreateObject("CMEGG_XML.Errors")
Call oCmErr.Reset()

Set oCmData = CreateObject("MSXML2.DOMDocument.4.0")
oCmData.async = False
oCmData.setProperty "SelectionLanguage", "XPath"

Set oCmConfigData = CreateObject("MSXML2.DOMDocument.4.0")
oCmConfigData.async = False
oCmConfigData.setProperty "SelectionLanguage", "XPath"

strPath = "C:\VSS\Islands\KOD\cmKOD_Apps\workspace\SyncSvcDsp\"

Call oCmConfigData.load(strPath & "wse_config.xml")
Call oCmData.load(strPath & "prepared_cmdata_call.xml")

With oKODApps

  Call .Exec(oCmData, oCmConfigData, oCmErr)

  set oDoc = oCmErr.GetDOMDoc
  Call oDoc.save(strPath & "out\oCmErr_call.xml")

  Call oCmData.save(strPath & "out\oCmData_call.xml")

End With

msgbox "finished!"