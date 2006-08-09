Attribute VB_Name = "modXML"
Option Explicit

'
' XSLTTransform
'
' Purpose       :
' Parameters    :
' Return value  :
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036, 2005-07-11
'
Sub XSLTTransform(ByRef poXMLObj As Object, ByVal strXSLTFN As String, ByRef poCmErr As Object)
  
    ' some declarations
    Dim oDOMXSL As MSXML2.FreeThreadedDOMDocument40
    Dim oDOMXML As MSXML2.FreeThreadedDOMDocument40
    Dim oXSLT As New MSXML2.XSLTemplate40
    Dim oXSLProc As MSXML2.IXSLProcessor
  
  ' instantiate objects
  Set oDOMXSL = CreateObject("MSXML2.FreeThreadedDOMDocument.4.0")
  oDOMXSL.async = False
  oDOMXSL.setProperty "SelectionLanguage", "XPath"
  
  Set oDOMXML = CreateObject("MSXML2.FreeThreadedDOMDocument.4.0")
  oDOMXML.async = False
  oDOMXML.setProperty "SelectionLanguage", "XPath"
  
  Set oXSLT = CreateObject("MSXML2.XSLTemplate.4.0")

  ' activate error handler
  On Error GoTo handle_err
  
  ' load XML string
  If CBool(oDOMXML.loadXML(poXMLObj.xml)) Then

    ' dont parse cause entities cannot be resolved
    oDOMXSL.validateOnParse = False
    oDOMXSL.async = False ' never ever forget it !

    ' load XSL file from harddisk
    If CBool(oDOMXSL.Load(strXSLTFN)) Then
      
      Set oXSLT.stylesheet = oDOMXSL
      Set oXSLProc = oXSLT.createProcessor()
      oXSLProc.input = oDOMXML
      
      ' transform data
      oXSLProc.Transform
      If Not CBool(poXMLObj.loadXML(oXSLProc.output)) Then
        Call HandleException(poCmErr, 100023, App.Title, "failed to load transformation result")
      End If
    Else
      Call HandleXMLException(oDOMXSL.parseError, poCmErr, 100024, App.Title, "failed to load xsl file '" & strXSLTFN & "'")
    End If
  Else
    Call HandleXMLException(oDOMXML.parseError, poCmErr, 100025, App.Title, "failed to load xml")
  End If
  
  Exit Sub
  
handle_err:
  
  Call HandleException(poCmErr, 100026, App.Title, "transformation failed")
  Exit Sub

End Sub

