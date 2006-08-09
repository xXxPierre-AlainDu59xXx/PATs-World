VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' ADObject
'
' Purpose :
'
' Copyright © 2001-2005 by econet AG
'

Option Explicit

  ' MSXML2 document
  Private m_oXMLCfg As Object ' MSXML2.DOMDocument40
  Private m_oXMLObj As Object ' MSXML2.DOMDocument40

  ' ref's to MSXML2 documents
  Private m_oCmData As Object ' MSXML2.DOMDocument40
  Private m_oCmConfigData As Object ' MSXML2.DOMDocument40
  Private m_oCmErr As Object ' CMEGG_XMLLib.Errors

  ' internal store
  Private m_strCfgID As String
  Private m_strXSLTFN As String
  Private m_oNConfig As Object ' MSXML2.IXMLDOMNode

'
' Exec
'
' Purpose       : to avoid excessive details, see documentation
'
' Parameters    : ByVal pCmData As MSXML2.DOMDocument40       : reference to context 'cmData'
'                 ByVal pCmConfigData As MSXML2.DOMDocument40 : reference to context 'cmConfigData'
'                 ByVal pCmErr As CMEGG_XMLLib.Errors         : reference to context 'cmErr'
'
' Return value  : none

' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-07-08
'
' Public Function (ByVal pCmData As Object, ByVal pCmConfigData As Object, ByVal pCmErr As Object) As Boolean
Public Sub Exec(ByVal pCmData As MSXML2.DOMDocument40, ByVal pCmConfigData As MSXML2.DOMDocument40, ByVal pCmErr As CMEGG_XMLLib.Errors)

    ' MSXML2 interfaces
    Dim oNode As Object ' MSXML2.IXMLDOMNode
    Dim oAttr As Object ' MSXML2.IXMLDOMAttribute
  
    ' locals
    Dim strTemp As String
  
    ' control booleans
    Dim bExecXSLT As Boolean
    Dim bExecPlugIn As Boolean
  
  ' activate error handler
  On Error GoTo handle_err

  ' store interfaces
  Set m_oCmData = pCmData
  Set m_oCmConfigData = pCmConfigData
  Set m_oCmErr = pCmErr
  
  ' load configuration : mandatory : 'app_adobj_details'
  ' if the attribute xmlid is available, try to resolve the filepath the standard way
  ' if not the configuration is located as subnodes of the element 'app_adobj_details'
  Set oNode = m_oCmConfigData.documentElement.selectSingleNode("./app_adobj_details")
  If Not oNode Is Nothing Then
    Set oAttr = oNode.Attributes.getNamedItem("xmlid")
    If Not oAttr Is Nothing Then
      If oAttr.Text <> "" Then
        strTemp = GetXMLFN(oAttr.Text)
        If strTemp <> "" Then
          If Not CBool(m_oXMLCfg.Load(strTemp)) Then
            Call HandleException(m_oCmErr, 100018, App.Title, "failed to load external xml configuration file")
          End If
        Else
          Call HandleException(m_oCmErr, 100019, App.Title, "external xml configuration not available")
        End If
      Else
        Call m_oXMLCfg.loadXML(oNode.xml)
      End If
    Else
      Call m_oXMLCfg.loadXML(oNode.xml)
    End If
  Else
    Call HandleException(m_oCmErr, 100020, App.Title, "no configuration available")
  End If

  ' if any error occurred, bail out
  If m_oCmErr.Count > 0 Then
    Exit Sub
  End If
  
  '
  ' ok. lets see what kind of execution we're facing
  '

  ' defaults
  m_strCfgID = "default"

  Set oNode = m_oCmData.documentElement.selectSingleNode("./cmrequest/cfgid")
  If Not oNode Is Nothing Then m_strCfgID = Trim(oNode.Text)

  Set m_oNConfig = m_oXMLCfg.documentElement.selectSingleNode("./*[@id='" & m_strCfgID & "']")
  If m_oNConfig Is Nothing Then
    Call HandleException(m_oCmErr, 100021, App.Title, "no configuration available ('" & m_strCfgID & "')")
    Exit Sub
  End If
  
  '
  ' processing handling
  '
  
  ' defaults
  bExecXSLT = False
  bExecPlugIn = False
  
  Set oAttr = m_oNConfig.Attributes.getNamedItem("xslid")
  If Not oAttr Is Nothing Then
    
    ' do some verification
    Call VerifyXSLTHandling
    If m_oCmErr.Count = 0 Then bExecXSLT = True
    ' everythings prepared. now access the AD, bind the object and get all required properties
    Call BindObjectAndEnrichXML
    
    ' store resulting xml
    If Not m_oXMLObj.documentElement Is Nothing Then
      Set oNode = m_oCmData.createElement("adobjecttmp")
      Call m_oCmData.documentElement.appendChild(oNode)
      Call oNode.appendChild(m_oXMLObj.documentElement.cloneNode(True)) ' m_oXMLObj.childNodes.Item(0))
    End If
  
  End If
  
  If m_oCmErr.Count = 0 Then
    
    Set oAttr = m_oNConfig.Attributes.getNamedItem("progid")
    If Not oAttr Is Nothing Then
      bExecPlugIn = True
      Call ExecutePlugIn
    End If
  
    If bExecXSLT And m_oCmErr.Count = 0 Then
      ' ok. now we will transform the enriched XML
      Call XSLTTransform(m_oXMLObj, m_strXSLTFN, m_oCmErr)
    End If
  
  End If
  
  ' valid configuration ?
  If Not bExecPlugIn And Not bExecXSLT Then
    Call HandleException(m_oCmErr, 100030, App.Title, "neither plugin nor xslt transformation executed")
  End If
  
  '
  ' if there is no xml result and no error pending, create error information
  '
  If m_oCmErr.Count = 0 And m_oXMLObj.documentElement Is Nothing Then
    Call HandleException(m_oCmErr, 100022, App.Title, "no xml result available")
  End If
  
  ' append response xml only if no error occurred
  If m_oCmErr.Count = 0 Then
    Call m_oCmData.documentElement.selectSingleNode("./cmresponse").appendChild(m_oXMLObj.documentElement)
  Else
    If Not m_oXMLObj.documentElement Is Nothing Then
      Call m_oCmData.documentElement.appendChild(m_oXMLObj.documentElement)
    End If
  End If
  
  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100017, App.Title, TXT_ERR_UNEXPECTED)
  Exit Sub

End Sub

'
' VerifyXSLTHandling
'
' Purpose       : verify xml parameters, verify xslt, verify and load xml template
' Parameters    : none
' Return value  : none
'
' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-07-20
'
Private Sub VerifyXSLTHandling()
    
    ' locals
    Dim strXSLTID As String
    Dim lCountNodes As Long
    
    ' MSXML2 interfaces
    Dim oNode As Object ' MSXML2.IXMLDOMNode

  Set oNode = m_oCmData.documentElement.selectSingleNode("./cmrequest/objectguid")
  If oNode Is Nothing Then Call HandleException(m_oCmErr, 100010, App.Title, "parameter 'objectguid' is missing")
  
  Set oNode = m_oCmData.documentElement.selectSingleNode("./cmrequest/domainfqdn")
  If oNode Is Nothing Then Call HandleException(m_oCmErr, 100011, App.Title, "parameter 'domainfqdn' is missing")
  
  Set oNode = m_oNConfig.selectSingleNode("./template")
  If oNode Is Nothing Then
    Call HandleException(m_oCmErr, 100012, App.Title, "configuration parameter 'template' is missing")
    Exit Sub
  End If
  
  lCountNodes = CLng(m_oNConfig.selectSingleNode("./template").childNodes.length)
  
  Select Case lCountNodes
    
    Case 0
      Call HandleException(m_oCmErr, 100013, App.Title, "configuration template contains no xml")
    
    Case Is > 1
      Call HandleException(m_oCmErr, 100014, App.Title, "configuration template contains more than one childnode")
  
  End Select
 
  ' bail out in case of errors
  If m_oCmErr.Count > 0 Then Exit Sub
  
  If Not CBool(m_oXMLObj.loadXML(m_oNConfig.selectSingleNode("./template").childNodes.Item(0).xml)) Then
    Call HandleException(m_oCmErr, 100015, App.Title, "failed to load configuration template")
    Exit Sub
  End If
  
  ' get xslt identifier (text or number)
  strXSLTID = m_oNConfig.Attributes.getNamedItem("xslid").Text
  m_strXSLTFN = LCase(Trim(GetXSLTFN(strXSLTID)))
  If m_strXSLTFN = "" Then
    Call HandleException(m_oCmErr, 100016, App.Title, "xslid '" & strXSLTID & "' not found")
    Exit Sub
  End If

End Sub

'
' BindObjectAndEnrichXML
'
' Purpose       : bind ad object and fill provided xml structure with property values
' Parameters    : none
' Return value  : none
'
' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-07-11
'
Private Sub BindObjectAndEnrichXML()

    ' MSXML2 interfaces
    Dim oAttr As Object ' MSXML2.IXMLDOMAttribute
    Dim oNode As Object ' MSXML2.IXMLDOMNode
    Dim oNL As Object ' MSXML2.IXMLDOMNodeList
    Dim oNodeNew As Object ' MSXML2.IXMLDOMNode
    Dim oCDS As Object ' MSXML2.IXMLDOMCDATASection
  
    ' key values
    Dim strObjectGUID As String
    Dim strDomainFQDN As String
    Dim strDCFQDN As String
    
    ' AD property
    Dim strObjCN As String
    
    Dim strPropName As String
    Dim strPropSyntax As String
    Dim bPropIsMultiValued As Boolean
    Dim lPropAdsType As Long
    Dim lPropVTDataType As Long
    Dim strPropTypeName As String
    Dim varPropValue As Variant
    
    Dim lPropMode As Long
    
    Dim varValue As Variant
    Dim varValueList As Variant
    
    ' cmSAL_AD interfaces
    Dim objADSIWrapper As Object
    
    ' ADSI interfaces
    Dim oADObj As IADs
    Dim oScheme As IADs
    Dim oUserScheme As IADsClass
    Dim oProp As IADsProperty
    Dim oPropSyntax As IADsSyntax
  
    Dim oPropList As IADsPropertyList
    Dim oPropEntry As IADsPropertyEntry
    Dim oPropVal As IADsPropertyValue

    ' generic
    Dim lPropCount As Long
    Dim i As Long
    Dim bExec As Boolean
  
  ' get parameters (already verified)
  strObjectGUID = Trim(m_oCmData.documentElement.selectSingleNode("./cmrequest/objectguid").Text)
  strDomainFQDN = Trim(m_oCmData.documentElement.selectSingleNode("./cmrequest/domainfqdn").Text)
  
  ' cmrequest is check first, if no dc fqdn was specified it will be searched in the wse_config xml
  Set oNode = m_oCmData.documentElement.selectSingleNode("./cmrequest/dcfqdn")
  If Not oNode Is Nothing Then
    strDCFQDN = oNode.Text
  Else
    Set oNode = m_oNConfig.selectSingleNode("./dcfqdn")
    If Not oNode Is Nothing Then
      strDCFQDN = oNode.Text
    End If
  End If
  If strDCFQDN = "*" Or strDCFQDN = "" Then
    strDCFQDN = ""
  End If

  ' resolve domain controller ?

  If strDCFQDN = "" Then
  
    ' try to create required instances
    On Error GoTo handle_err_salad
    Set objADSIWrapper = CreateObject("cmSAL_AD.ADSIWrapper")
    
    ' activate generic error handler again
    On Error GoTo handle_err_getanydc
    
    With objADSIWrapper
      .DomainName = strDomainFQDN
      Call .GetAnyDC
      strDCFQDN = .ServerName
    End With

  End If
  
  On Error GoTo handle_err_bind

  ' bind object
  Set oADObj = GetObject("LDAP://" & strDCFQDN & "/<GUID=" & strObjectGUID & ">")
  Set oUserScheme = GetObject(oADObj.schema)
  Set oScheme = GetObject(oUserScheme.Parent)
  
  ' for informational purposes only : the common name
  strObjCN = oADObj.Name
  
  ' get property cache
  Set oPropList = oADObj
  oPropList.GetInfo
  lPropCount = oPropList.PropertyCount
  
  On Error GoTo handle_err

  ' iterate thru all elements
  Set oNL = m_oXMLObj.selectNodes("//properties/property")
  For Each oNode In oNL
    
    ' defaults
    bExec = False
    lPropMode = 0

    ' how to read the attribute ?
    Set oAttr = oNode.Attributes.getNamedItem("mode")
    If Not oAttr Is Nothing Then
      lPropMode = CLng(oAttr.Text)
    End If
    
    ' get property name
    strPropName = oNode.selectSingleNode("./name").Text
  
    ' bind required interfaces
    On Error Resume Next

    If lPropMode = 1 Then

      ' try to get property
      varPropValue = CallByName(oADObj, strPropName, VbGet)
      
      Select Case Err.Number
        
        Case 0
          bExec = True
        
        Case 438
            ' property could not be bound
            Set oNodeNew = m_oXMLObj.createElement("value")
            Call oNode.appendChild(oNodeNew)
            oNodeNew.Text = "null"
        
        Case Else
          Call HandleException(m_oCmErr, 100034, App.Title, "failed to access property '" & strPropName & "' (object: '" & "LDAP://" & strDCFQDN & "/<GUID=" & strObjectGUID & ">" & "')")
      
      End Select
  
      On Error GoTo handle_err
  
      If bExec Then
  
        ' get details about the bound property
        bPropIsMultiValued = CBool(IsArray(varPropValue))
        strPropTypeName = TypeName(varPropValue)
        
        ' store those information for later usage
        Set oAttr = m_oXMLObj.createAttribute("multivalued")
        Call oNode.Attributes.setNamedItem(oAttr)
        oAttr.Text = CLng(bPropIsMultiValued)
        
        Set oAttr = m_oXMLObj.createAttribute("typename")
        Call oNode.Attributes.setNamedItem(oAttr)
        oAttr.Text = strPropTypeName
  
        ' remove all value elements (if there are any)
        Do While Not oNode.selectSingleNode("./value") Is Nothing
          Set oNodeNew = oNode.selectSingleNode("./value")
          Call oNodeNew.parentNode.removeChild(oNodeNew)
        Loop
  
        ' unsupported values:
        
        ' Empty : Uninitialized
        ' Null : No valid data
        ' Object : Generic object
        ' Unknown : Unknown object type
        ' Nothing : Object variable that doesn't yet refer to an object instance Error
        
        If LCase(strPropTypeName) <> "empty" And _
           LCase(strPropTypeName) <> "null" And _
           LCase(strPropTypeName) <> "object" And _
           LCase(strPropTypeName) <> "unknown" And _
           LCase(strPropTypeName) <> "nothing" Then
        
          ' switch error handler
          On Error GoTo handle_err_getvalue_mode1
          
          ' get all values
          If bPropIsMultiValued Then
            
            For Each varValue In varPropValue
            
              Set oNodeNew = m_oXMLObj.createElement("value")
              Call oNode.appendChild(oNodeNew)
              ' store it as cdata section
              Set oCDS = m_oXMLObj.createCDATASection(varValue)
              Call oNodeNew.appendChild(oCDS)
            
            Next
          
          Else
            
            Set oNodeNew = m_oXMLObj.createElement("value")
            Call oNode.appendChild(oNodeNew)
            ' store it as cdata section
            Set oCDS = m_oXMLObj.createCDATASection(varPropValue)
            Call oNodeNew.appendChild(oCDS)
          
          End If
          
          ' back to default error handler
          On Error GoTo handle_err
          
        Else
          
          Call HandleException(m_oCmErr, 100035, App.Title, "unsupported typename. property='" & strPropName & "' typename='" & strPropTypeName & "'")
          Set oNodeNew = m_oXMLObj.createElement("unsupported")
          Call oNode.appendChild(oNodeNew)
        
        End If
      
      End If ' If bExec Then

    ElseIf lPropMode = 0 Then
    
      Set oProp = oScheme.GetObject("Property", strPropName)
      Set oPropSyntax = GetObject(oScheme.ADsPath & "/" & oProp.Syntax)
      
      Select Case Err.Number
        Case 0
          bExec = True
        Case -2147463168
          Call HandleException(m_oCmErr, 100032, App.Title, "failed to find property '" & strPropName & "' in schema (object: '" & "LDAP://" & strDCFQDN & "/<GUID=" & strObjectGUID & ">" & "')")
        Case Else
          Call HandleException(m_oCmErr, 100033, App.Title, "failed to access property '" & strPropName & "' in schema (object: '" & "LDAP://" & strDCFQDN & "/<GUID=" & strObjectGUID & ">" & "')")
      End Select
  
      On Error GoTo handle_err
      
      If bExec Then
      
        ' get details about the bound property
        bPropIsMultiValued = CBool(oProp.MultiValued)
        lPropVTDataType = oPropSyntax.OleAutoDataType
        strPropSyntax = oProp.Syntax
          
        ' store those information for later usage
        Set oAttr = m_oXMLObj.createAttribute("multivalued")
        Call oNode.Attributes.setNamedItem(oAttr)
        oAttr.Text = CLng(bPropIsMultiValued)
        
        Set oAttr = m_oXMLObj.createAttribute("syntax")
        Call oNode.Attributes.setNamedItem(oAttr)
        oAttr.Text = strPropSyntax
    
        Set oAttr = m_oXMLObj.createAttribute("oleautodatatype")
        Call oNode.Attributes.setNamedItem(oAttr)
        oAttr.Text = GetAsText_VT_GENERIC_ENUM(lPropVTDataType)
    
        ' remove all value elements (if there are any)
        Do While Not oNode.selectSingleNode("./value") Is Nothing
          Set oNodeNew = oNode.selectSingleNode("./value")
          Call oNodeNew.parentNode.removeChild(oNodeNew)
        Loop
        
        ' try to find property in cacht to retrieve the appropriate adstype
        On Error Resume Next
        lPropAdsType = -1
        Set oPropEntry = oPropList.Item(strPropName)
        Select Case Err.Number
          
          Case 0
            
            ' re-activate common handler
            On Error GoTo handle_err
            
            lPropAdsType = oPropEntry.ADsType
            
            Set oAttr = m_oXMLObj.createAttribute("adstype")
            Call oNode.Attributes.setNamedItem(oAttr)
            oAttr.Text = GetAsText_ADSTYPE_ENUM(lPropAdsType)
        
            ' ok. which adstypes do we support ? the following of course! :-)
            If lPropAdsType = ADSTYPE_DN_STRING Or _
               lPropAdsType = ADSTYPE_CASE_EXACT_STRING Or _
               lPropAdsType = ADSTYPE_CASE_IGNORE_STRING Or _
               lPropAdsType = ADSTYPE_PRINTABLE_STRING Or _
               lPropAdsType = ADSTYPE_NUMERIC_STRING Or _
               lPropAdsType = ADSTYPE_BOOLEAN Or _
               lPropAdsType = ADSTYPE_INTEGER Or _
               lPropAdsType = ADSTYPE_UTC_TIME Then
                  
              ' prepare
              Select Case lPropAdsType
                Case ADSTYPE_DN_STRING
                  strPropTypeName = "DNString"
                Case ADSTYPE_CASE_EXACT_STRING
                  strPropTypeName = "CaseExactString"
                Case ADSTYPE_CASE_IGNORE_STRING
                  strPropTypeName = "CaseIgnoreString"
                Case ADSTYPE_PRINTABLE_STRING
                  strPropTypeName = "PrintableString"
                Case ADSTYPE_NUMERIC_STRING
                  strPropTypeName = "NumericString"
                Case ADSTYPE_BOOLEAN
                  strPropTypeName = "Boolean"
                Case ADSTYPE_INTEGER
                  strPropTypeName = "Integer"
                Case ADSTYPE_UTC_TIME
                  strPropTypeName = "UTCTime"
              End Select
              
              ' switch error handler
              On Error GoTo handle_err_getvalue_mode0
              
              ' get all values
              For Each varValue In oPropEntry.Values
                Set oPropVal = varValue
                Set oNodeNew = m_oXMLObj.createElement("value")
                Call oNode.appendChild(oNodeNew)
                ' read single property value and store it as cdata section
                varPropValue = CallByName(oPropVal, strPropTypeName, VbGet)
                Set oCDS = m_oXMLObj.createCDATASection(varPropValue)
                Call oNodeNew.appendChild(oCDS)
              Next varValue
    
              ' back to default error handler
              On Error GoTo handle_err
            
            Else
    
              Call HandleException(m_oCmErr, 100004, App.Title, "unsupported property. name='" & strPropName & "' syntax='" & strPropSyntax & "' adstype='" & GetAsText_ADSTYPE_ENUM(lPropAdsType) & "' oleautodatatype='" & GetAsText_VT_GENERIC_ENUM(lPropVTDataType) & "'")
              Set oNodeNew = m_oXMLObj.createElement("unsupported")
              Call oNode.appendChild(oNodeNew)
            
            End If
          
          Case &H8000500D
            
            ' re-activate common handler
            On Error GoTo handle_err
            
            ' property could not be found in cache
            Set oNodeNew = m_oXMLObj.createElement("value")
            Call oNode.appendChild(oNodeNew)
            oNodeNew.Text = "null"
          
          Case Else
            
            ' unexpected exception
            Call HandleException(m_oCmErr, 100005, App.Title, "unexpected error while finding property '" & strPropName & "' in property cache of '" & strObjCN & "' ('" & strObjectGUID & "')")
            
            ' re-activate common handler
            On Error GoTo handle_err
        
        End Select
      
      End If ' if bExec then ...
  
    End If ' ... elseif lPropmode=0 (default)
  
  Next ' For Each oNode In oNL


  Exit Sub

handle_err:

  Call HandleException(m_oCmErr, 100006, App.Title, TXT_ERR_UNEXPECTED)
  Exit Sub

handle_err_getvalue_mode0:

  Call HandleException(m_oCmErr, 100031, App.Title, "failed to read values of property '" & strPropName & "' (type='" & strPropTypeName & "') (syntax='" & strPropSyntax & "' adstype='" & GetAsText_ADSTYPE_ENUM(lPropAdsType) & "' oleautodatatype='" & GetAsText_VT_GENERIC_ENUM(lPropVTDataType) & "')")
  Exit Sub

handle_err_getvalue_mode1:

  Call HandleException(m_oCmErr, 100036, App.Title, "failed to read values of property '" & strPropName & "' (typename='" & TypeName(varPropValue) & "'")
  Exit Sub

handle_err_salad:
  
  Call HandleException(m_oCmErr, 100007, App.Title, "Installation failure : 'cmSAL_AD.DLL' is missing or not properly registered")
  Exit Sub

handle_err_bind:

  Call HandleException(m_oCmErr, 100008, App.Title, "failed to bind '" & "LDAP://" & strDCFQDN & "/<GUID=" & strObjectGUID & ">" & "' ")
  Exit Sub

handle_err_getanydc:

  Call HandleException(m_oCmErr, 100009, App.Title, "failed to bind to domain '" & strDomainFQDN & "' ")
  Exit Sub

End Sub

'
' ExecutePlugIn
'
' Purpose       : execute COM plugin
' Parameters    : none
' Return value  : none
'
' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-07-12
'
Private Sub ExecutePlugIn()

    ' locals
    Dim oObj As Object
    Dim strProgID As String
    Dim strMethod As String
    Dim lIndex As Long
    Dim lResult As Long
    Dim strTemp As String
    
    ' MSXML2 interfaces
    Dim oAttr As Object ' MSXML2.IXMLDOMAttribute
    Dim oNode As Object ' MSXML2.IXMLDOMNode
  
  ' activate error handler
  On Error GoTo handle_err
  
  Set oNode = m_oNConfig
  If Not oNode Is Nothing Then
    
    Set oAttr = oNode.Attributes.getNamedItem("progid")
    If Not oAttr Is Nothing Then strProgID = oAttr.Text
  
    Set oAttr = oNode.Attributes.getNamedItem("method")
    If Not oAttr Is Nothing Then strMethod = oAttr.Text
  
  End If
  
  If strProgID = "" Or strMethod = "" Then
    Call HandleException(m_oCmErr, 100029, App.Title, "plugin configuraton failure : no progid or method specified")
    Exit Sub
  End If
  
  On Error GoTo handle_err_1
  
  Set oObj = CreateObject(strProgID)
  
  On Error GoTo handle_err_2
  
  strTemp = CallByName(oObj, strMethod, VbMethod, m_oCmData, m_oCmConfigData, m_oCmErr)
  If Not CBool(m_oXMLObj.loadXML(strTemp)) Then
    Call HandleException(m_oCmErr, 100000, App.Title, "failed to load the xml-string result of '" & strProgID & "." & strMethod & "'")
  End If
  
  Exit Sub

handle_err:

  Call HandleException(m_oCmErr, 100001, App.Title, TXT_ERR_UNEXPECTED)
  Exit Sub

handle_err_1:

  Call HandleException(m_oCmErr, 100002, App.Title, "failed to create new instance of '" & strProgID & "'")
  Exit Sub

handle_err_2:

  Call HandleException(m_oCmErr, 100003, App.Title, "failed to call function '" & strMethod & "'")
  Exit Sub

End Sub


'
' GetXMLFN
'
' Purpose       : get xml-filename by id
' Parameters    : strID = element identifier (wse_config.xml)
' Return value  : filename
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036, 2005-07-11
'
Private Function GetXMLFN(ByVal strID As String) As String
  GetXMLFN = GetConfigFN(strID, "xmls")
End Function

'
' GetXSLTFN
'
' Purpose       : get xsl-filename by id
' Parameters    : strID = element identifier (wse_config.xml)
' Return value  : filename
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036, 2005-07-11
'
Private Function GetXSLTFN(ByVal strID As String) As String
  GetXSLTFN = GetConfigFN(strID, "xsls")
End Function

'
' GetConfigFN
'
' Purpose       : get filename by id and root-name
' Parameters    : strID = element identifier (wse_config.xml)
'               : strName = root element name
' Return value  : filename
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036, 2005-07-11
'
Private Function GetConfigFN(ByVal strID As String, ByVal strName As String) As String

    ' locals
    Dim strFN As String
    Dim oAttr As Object ' MSXML2.IXMLDOMAttribute
    Dim oNode As Object ' MSXML2.IXMLDOMNode
  
  ' default
  GetConfigFN = ""
  
  Set oNode = m_oCmConfigData.documentElement.selectSingleNode("./" & LCase(strName) & "/*[@id='" & strID & "']")
  If Not oNode Is Nothing Then
    Set oAttr = oNode.Attributes.getNamedItem("userelpath")
    If Not oAttr Is Nothing Then
      If CBool(oAttr.Text) Then
        strFN = m_oCmData.documentElement.selectSingleNode("./control/rootpath").Text
      End If
    End If
    GetConfigFN = strFN & oNode.Text
  End If

End Function

Private Sub Class_Initialize()

  Set m_oXMLCfg = CreateObject("msxml2.DOMDocument.4.0")
  m_oXMLCfg.async = False
  m_oXMLCfg.setProperty "SelectionLanguage", "XPath"

  Set m_oXMLObj = CreateObject("msxml2.DOMDocument.4.0")
  m_oXMLObj.async = False
  m_oXMLObj.setProperty "SelectionLanguage", "XPath"

End Sub

'
' tr036 2005-07-21
' >> code snippets from development time
' >> since there is no code-snippet library i will store it here for later usage
'
        
        ' special handling : if the property exists in the ad-scheme but is not set for the bound object
        ' an error will be thrown. this kind of error has to be trapped and the value will equal "null"
'        On Error Resume Next
'        varValueList = oADObj.GetEx(strPropName)
'        Select Case Err.Number
'
'          Case 0
'            On Error GoTo handle_err
'            For Each varValue In varValueList
'              Set oNodeNew = m_oXMLTpl.createElement("value")
'              Call oNode.appendChild(oNodeNew)
'              oNodeNew.Text = varValue
'            Next
'          Case &H8000500D
'            Set oNodeNew = m_oXMLTpl.createElement("value")
'            Call oNode.appendChild(oNodeNew)
'            oNodeNew.Text = "null"
'
'          Case Else
'            Call HandleException(m_oCmErr, , App.Title, "failed to bind property '" & strPropName & "' for '" & strObjCN & "' ('" & strObjectGUID & "')")
'
'        End Select
'
'        On Error GoTo handle_err


' VT_ARRAY | VT_UI4
' ADSTYPE_OCTET_STRING

' ADSTYPE_UTC_TIME, ADSTYPE_LARGE_INTEGER


'    Case ADSTYPE_LARGE_INTEGER
'
'Dim liLarge As LargeInteger
'strPropName = strPropName & ""
'Set liLarge = oPropVal.LargeInteger
'Call LargeInteger_To_Time(liLarge, varValue)
'oNodeNew.Text = varValue


'pwdlastset
'On Error Resume Next
'
'Const E_ADS_PROPERTY_NOT_FOUND = &H8000500D      ' LINE 3
'
'Set objUser = GetObject("LDAP://CN=myerken,OU=management,DC=fabrikam,DC=com")
'dtmValue = objUser.PasswordLastChanged           ' LINE 6
'
'If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
'    WScript.Echo "The password has never been set."
'    WScript.Quit
'Else
'    WScript.Echo "The password was last set on " & _
'                 DateValue(dtmValue) & " at " & TimeValue(dtmValue)
'End If

' maxpwdage
'Const ONE_HUNDRED_NANOSECOND = 0.0000001    ' .000000100 is equal to 10^-7
'Const SECONDS_IN_DAY = 86400
'
'Set objDomain = GetObject("LDAP://DC=fabrikam,DC=com")     ' LINE 4
'Set objMaxPwdAge = objDomain.Get("maxPwdAge")              ' LINE 5
'
'If objMaxPwdAge.LowPart = 0 Then
'  WScript.Echo "The Maximum Password Age is set to 0 in the " & _
'               "domain. Therefore, the password does not expire."
'  WScript.Quit
'Else
'  dblMaxPwdNano = Abs(objMaxPwdAge.HighPart * 2 ^ 32 + objMaxPwdAge.LowPart)
'  dblMaxPwdSecs = dblMaxPwdNano * ONE_HUNDRED_NANOSECOND   ' LINE 13
'  dblMaxPwdDays = Int(dblMaxPwdSecs / SECONDS_IN_DAY)      ' LINE 14
'  WScript.Echo "Maximum password age: " & dblMaxPwdDays & " days"
'End If



'    Select Case lVTDataType
'
'    '
'    '    Dim oDomain As IADs
'    '    Dim oLargeInt As LargeInteger
'    '    Set oDomain = GetObject("LDAP://DC=fabrikam,DC=com")
'    '    Set oLargeInt = oDomain.Get("creationTime")
'    '    Debug.Print oLargeInt.HighPart
'    '    Debug.Print oLargeInt.LowPart
'    '    strTemp = "&H" + CStr(Hex(oLargeInt.HighPart)) + CStr(Hex(oLargeInt.LowPart))
'    '    Debug.Print strTemp
'
'
'    'Dim oTarget As IADs
'    'Dim Octet(5) As Byte
'    'Dim MultiOctet(2) As Variant
'    'Dim i As Integer, j As Integer
'    '
'    'On Error GoTo Cleanup
'    '
'    '' Set up MultiOctetString.
'    'For i = 0 To 2
'    '    For j = 0 To 5
'    '        Octet(j) = CByte(i * j)
'    '    Next j
'    '    MultiOctet(i) = Octet
'    'Next i
'    '
'    '' Bind to the object and set MultiOctetString.
'    'Set oTarget = GetObject("LDAP://CN=SomeUser,CN=Users,DC=Fabrikam, DC=COM")
'    'oTarget.Put "multiOctetString", MultiOctet
'    'oTarget.SetInfo
'    '
'    'Dim GetOctet As Variant
'    'Dim Temp As Variant
'    '
'    '' Read back and print MultiOctetString.
'    'GetOctet = oTarget.Get("multiOctetString")
'    'For i = LBound(GetOctet) To UBound(GetOctet)
'    '    Temp = GetOctet(i)
'    '    For j = LBound(Temp) To UBound(Temp)
'    '        Debug.Print Temp(j)
'    '    Next j
'    '    Debug.Print "----"
'    'Next i
'    '
'    'Exit Sub
'    '
'    'Cleanup:
'    '   MsgBox ("An error has occurred. " & Err.Number)
'    '   Set oTarget = Nothing
'
'' unklar ob wirs supporten
'
'' ADSTYPE_DN_WITH_STRING : see pdf or ad schema description
'' ADSTYPE_OBJECT_CLASS
'
'' schwer zu supporten in VB :
'' INTEGER8 : erst .HighPart dann .LowPart ["&H" + CStr(Hex(varValue.HighPart)) + CStr(Hex(varValue.LowPart))]
'
'      ' VT_DISPATCH
'
'      Case VT_BOOL, VT_BSTR, VT_CY, VT_DATE, _
'           VT_EMPTY, VT_ERROR, _
'           VT_I2, VT_I4, VT_R4, VT_R8, VT_UI1
'
''      Case ADSTYPE_INTEGER, ADSTYPE_BOOLEAN, _
'           ADSTYPE_DN_STRING, ADSTYPE_CASE_EXACT_STRING, ADSTYPE_CASE_IGNORE_STRING, _
'           ADSTYPE_PRINTABLE_STRING, ADSTYPE_NUMERIC_STRING, _
'           ADSTYPE_OCTET_STRING, _
'           ADSTYPE_DN_WITH_BINARY
'
'
'        ' special handling : if the property exists in the ad-scheme but is not set for the bound object
'        ' an error will be thrown. this kind of error has to be trapped and the value will equal "null"
'        On Error Resume Next
'        varValueList = oADObj.GetEx(strPropName)
'        Select Case Err.Number
'
'          Case 0
'            On Error GoTo handle_err
'            For Each varValue In varValueList
'              Set oNodeNew = m_oXMLTpl.createElement("value")
'              Call oNode.appendChild(oNodeNew)
'              oNodeNew.Text = varValue
'            Next
'
'
'
''For Each v In propEntry.Values
''Set propVal = v
''If propEntry.ADsType = 10 Then
''  Set liLArge = propVal.LargeInteger
''  Call LargeInteger_To_Time(liLArge, varValue)
''  Debug.Print varValue
''End If
'''    Debug.Print propVal.ADsType & " " &     Debug.Print GetAsText_ADSTYPE_ENUM(propVal.ADsType)
''Next
''Next
''i = 1
'
'          Case &H8000500D
'            Set oNodeNew = m_oXMLTpl.createElement("value")
'            Call oNode.appendChild(oNodeNew)
'            oNodeNew.Text = "null"
'
'          Case Else
'            Call HandleException(m_oCmErr, , App.Title, "failed to bind property '" & strPropName & "' for '" & strObjCN & "' ('" & strObjectGUID & "')")
'
'        End Select
'
'        On Error GoTo handle_err
'
'      Case Else
'
'        Call HandleException(m_oCmErr, , App.Title, "unsupported property. name='" & strPropName & "' syntax='" & strPropSyntax & "' adstype='" & GetAsText_ADSTYPE_ENUM(lPropAdsType) & "' oleautodatatype='" & GetAsText_VT_GENERIC_ENUM(lVTDataType) & "'")
'
'        Set oNodeNew = m_oXMLTpl.createElement("unsupported")
'        Call oNode.appendChild(oNodeNew)
'
'    End Select


'' VB Sample: Enumerating properties in the property list
'Dim propList As IADsPropertyList
'Dim propEntry As IADsPropertyEntry
'Dim propVal As IADsPropertyValue
'Dim count As Long, i As Long, v As Variant
'
'Set propList = GetObject("LDAP://" & strDCFQDN & "/<GUID=" & strObjectGUID & ">")
'propList.GetInfo
'
'count = propList.PropertyCount
''Debug.Print "No of Property Found: " & count
'
'Dim liLArge As LargeInteger
'
'Set propEntry = propList.GetPropertyItem(strPropName, ADSTYPE_CASE_IGNORE_STRING)
'Set propEntry = propList.Item(strPropName)
'
'
'
'For i = 0 To count - 1 'Each item in property list has a property entry
'  Set propEntry = propList.Item(i)
'  Debug.Print propEntry.Name & " " & propEntry.ADsType & " " & GetAsText_ADSTYPE_ENUM(propEntry.ADsType)
'  'Each value in property entry has property values
'
'
'    Set oProp = oScheme.GetObject("Property", propEntry.Name)
'    Set oPropSyntax = GetObject(oScheme.ADsPath & "/" & oProp.Syntax)
'    Debug.Print "strPropSyntax = " & oProp.Syntax
'    Debug.Print "lVTDataType = " & oPropSyntax.OleAutoDataType
'    Debug.Print "textual lVTDataType = " & GetAsText_VT_GENERIC_ENUM(oPropSyntax.OleAutoDataType)
'
'
'  For Each v In propEntry.Values
'    Set propVal = v
'
'' If LCase(propEntry.Name) = "pwdlastset" Then
'If propEntry.ADsType = 10 Then
'  Set liLArge = propVal.LargeInteger
'  Call LargeInteger_To_Time(liLArge, varValue)
'  Debug.Print varValue
'End If
'
''    Debug.Print propVal.ADsType
''    Debug.Print GetAsText_ADSTYPE_ENUM(propVal.ADsType)
'  Next
'Next
'i = 1

'Dim strTemp As String
'Dim x As IADsUser
'Dim cls As IADsClass, v As Variant, op As Variant
'
'Dim l As ISchema
'
'Set x = GetObject("LDAP://CN=Timm Ralf,OU=MUC-users,DC=central,DC=econet,DC=de")
'Set x = GetObject("LDAP://muc018/<GUID=36039ca6749fbd4eb2fbc973abb5c6ef>")
'
'' 36039ca6749fbd4eb2fbc973abb5c6ef
'
'strTemp = x.Name
'strTemp = x.Class
'
'Debug.Print x.GUID
'
'' Get optional property values of this object.
'Set cls = GetObject(x.schema)
'
'Dim scc As IADs
'Dim prc As IADsProperty
'
'Set scc = GetObject(cls.Parent)
'
'Set prc = scc.GetObject("Property", "Owner")
'Debug.Print "Attribute: " & prc.Name
'Debug.Print "Syntax: " & prc.Syntax
'Debug.Print "MultiValued: " & prc.MultiValued
'
'Set prc = scc.GetObject("Property", "LogonHours")
'Debug.Print "Attribute: " & prc.Name
'Debug.Print "Syntax: " & prc.Syntax
'Debug.Print "MultiValued: " & prc.MultiValued
'
'Dim sy As IADs
'Set sy = GetObject(scc.ADsPath & "/" & prc.Syntax)
'Debug.Print "Automation data types: " & sy.OleAutoDataType

'For Each op In cls.MandatoryProperties
'  v = x.Get(op)
'  strTemp = "'" & op & "' = '" & v & "'"
'  Debug.Print strTemp
'Next

