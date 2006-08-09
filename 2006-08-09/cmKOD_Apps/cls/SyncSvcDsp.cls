VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SyncSvcDsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' SyncSvcDsp
'
' Purpose : sync-service dispatcher.
'
' Copyright © 2001-2005 by econet AG
'

Option Explicit

  ' MSXML2 documents and interfaces
  Private m_oXMLCfg As Object ' MSXML2.DOMDocument40
  Private m_oXMLRS As Object ' MSXML2.DOMDocument40
  Private m_oXMLRq As Object ' MSXML2.DOMDocument40
  
  Private m_oNCfg As Object ' MSXML2.IXMLDOMNode

  ' BOR extensions
  Private m_oReq As Object
  
  ' MAI de/en-cryption
'  Private m_oCrypt As cmMAI_Crypt.BlowFish
  
  ' internal keystore
  Private m_cKeyStore As New KeyStore
 
  ' ref's to MSXML2 documents
  Private m_oCmData As Object ' MSXML2.DOMDocument40
  Private m_oCmConfigData As Object ' MSXML2.DOMDocument40
  Private m_oCmErr As Object ' CMEGG_XMLLib.Errors

  ' private members
  Private m_strConnectionString As String
  Private m_strTablename As String
  
  ' request related info
  Dim m_strInstance As String
  Dim m_strLogin As String
  Dim m_strPW As String
  Dim m_strSession As String
  
  ' status enumeration
  Private Enum eStatus
    status_open = 0
    status_success = 1
    status_locked = 2
    status_err_exec_failed = 3
    status_err_still_locked = 4
    status_pending = 5
    status_skipped = 6
  End Enum

  ' textual constants
  Private Const TXT_APP = "econet cMatrix : Sync Service Dispatcher"

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
' Responsible   : tr036 2005-08-16
'
' Public Sub Exec(ByVal pCmData As Object, ByVal pCmConfigData As Object, ByVal pCmErr As Object)
Public Sub Exec(ByVal pCmData As MSXML2.DOMDocument40, ByVal pCmConfigData As MSXML2.DOMDocument40, ByVal pCmErr As CMEGG_XMLLib.Errors)

    ' MSXML2 interfaces
    Dim oNode As Object ' MSXML2.IXMLDOMNode
    Dim oAttr As Object ' MSXML2.IXMLDOMAttribute
    
    ' locals
    Dim strTemp As String
    Dim lCmd As Long
    Dim strXPathCfg As String
  
    Dim lNumberOfRuns As Long
    Dim varValue As Variant
  
    Dim strFilterSource As String
    Dim strSQLGet As String
    Dim bDecrease As Boolean
  
  ' activate error handler
  On Error GoTo handle_err

  ' store interfaces
  Set m_oCmData = pCmData
  Set m_oCmConfigData = pCmConfigData
  Set m_oCmErr = pCmErr

  ' load configuration : mandatory : 'app_adobj_details'
  ' if the attribute xmlid is available, try to resolve the filepath the standard way
  ' if not the configuration is located as subnodes of the root configuration element
  Set oNode = m_oCmConfigData.documentElement.selectSingleNode("./app_syncsvcdsp_details")
  If Not oNode Is Nothing Then
    Set oAttr = oNode.Attributes.getNamedItem("xmlid")
    If Not oAttr Is Nothing Then
      If oAttr.Text <> "" Then
        strTemp = GetXMLFN(oAttr.Text)
        If strTemp <> "" Then
          If Not CBool(m_oXMLCfg.Load(strTemp)) Then
            Call HandleException(m_oCmErr, 100100, App.Title, "failed to load external xml configuration file")
          End If
        Else
          Call HandleException(m_oCmErr, 100101, App.Title, "external xml configuration not available")
        End If
      Else
        Call m_oXMLCfg.loadXML(oNode.xml)
      End If
    Else
      Call m_oXMLCfg.loadXML(oNode.xml)
    End If
  Else
    Call HandleException(m_oCmErr, 100102, App.Title, "no configuration available")
  End If

  If m_oCmErr.Count > 0 Then
    Call LogEvent(TXT_APP, EVENTLOG_ERROR_TYPE, "Sync Service Dispatcher : '" & m_oCmErr.GetXML & "'", 1, 0)
    Exit Sub
  End If

  ' determine which configuration to execute
  lCmd = -1
  Set oNode = m_oCmData.documentElement.selectSingleNode("./cmrequest")
  If Not oNode Is Nothing Then lCmd = CLng(oNode.Attributes.getNamedItem("cmd").Text)
  strXPathCfg = "./app_syncsvcdsp_details/cfg[@id='" & lCmd & "']"
  Set m_oNCfg = m_oXMLCfg.selectSingleNode(strXPathCfg)
  If m_oNCfg Is Nothing Then
    Call HandleException(m_oCmErr, 100103, App.Title, "no configuration for command '" & lCmd & "' available")
  End If

  If m_oCmErr.Count > 0 Then
    Call LogEvent(TXT_APP, EVENTLOG_ERROR_TYPE, "Sync Service Dispatcher : '" & m_oCmErr.GetXML & "'", 1, 0)
    Exit Sub
  End If
  
  ' heyho, lets go!
  
  ' first of all decrypt all encrypted data
  Call DecryptCfg
  
  ' defaults
  strFilterSource = ""
  m_strConnectionString = ""
  lNumberOfRuns = 0
  m_strSession = ""
  bDecrease = True
  
  ' read parameters from XML
  
  Set oNode = m_oNCfg.selectSingleNode("./number")
  If Not oNode Is Nothing Then lNumberOfRuns = CLng(oNode.Text)
  If lNumberOfRuns = -1 Then
    bDecrease = False
    lNumberOfRuns = 1
  End If
  
  Set oNode = m_oNCfg.selectSingleNode("./db")
  If Not oNode Is Nothing Then m_strConnectionString = oNode.Text
  
  Set oNode = m_oNCfg.selectSingleNode("./tablename")
  If Not oNode Is Nothing Then m_strTablename = oNode.Text
  
  Set oNode = m_oNCfg.selectSingleNode("./sql_a")
  If Not oNode Is Nothing Then strSQLGet = oNode.Text
  
  Set oNode = m_oNCfg.selectSingleNode("./source")
  If Not oNode Is Nothing Then strFilterSource = oNode.Text
  If strFilterSource = "*" Then strFilterSource = "%"
  
  ' soap request data
  
  ' activate error handler
  On Error GoTo handle_err
  
  Set m_oReq = CreateObject("cmBOR_Extensions.CMRequest")
  
  ' prepare CMRequest class for upcoming system calls
  With m_oReq
    
    Set oNode = m_oNCfg.selectSingleNode("./request/endpointurl")
    If Not oNode Is Nothing Then varValue = CStr(oNode.Text)
    .EndPointUrl = varValue
    
    Set oNode = m_oNCfg.selectSingleNode("./request/winhttpauthscheme")
    If Not oNode Is Nothing Then varValue = CLng(oNode.Text)
    .WinHTTPAuthScheme = CLng(varValue)
    
    Set oNode = m_oNCfg.selectSingleNode("./request/basicauthlogin")
    If Not oNode Is Nothing Then varValue = CStr(oNode.Text)
    .Login = varValue
    
    Set oNode = m_oNCfg.selectSingleNode("./request/basicauthpassword")
    If Not oNode Is Nothing Then varValue = CStr(oNode.Text)
    .Password = varValue
  
  End With
  
  Set oNode = m_oNCfg.selectSingleNode("./request/instance")
  If Not oNode Is Nothing Then m_strInstance = oNode.Text
  Set oNode = m_oNCfg.selectSingleNode("./request/login")
  If Not oNode Is Nothing Then m_strLogin = oNode.Text
  Set oNode = m_oNCfg.selectSingleNode("./request/password")
  If Not oNode Is Nothing Then m_strPW = oNode.Text
  
  ' prepare SQL for retrieving data
  strSQLGet = Replace(strSQLGet, "%SOURCE%", strFilterSource)
  strSQLGet = Replace(strSQLGet, "%TABLE%", m_strTablename)
  
  Do While (lNumberOfRuns > 0) And (m_oCmErr.Count = 0)
  
    ' default
    Call m_oXMLRS.loadXML("<none/>")
    
    If Not CBool(SQLExec(strSQLGet)) Then Exit Do
    If LCase(m_oXMLRS.documentElement.nodeName) = "none" Then Exit Do
    If m_oXMLRS.documentElement.selectNodes("./rs:data/z:row").Length = 0 Then Exit Do
    
    ' ok. lets continue.
    Call RSAnalyse
    
    ' any soap request ?
    If m_oXMLRS.documentElement.selectNodes("./rs:data/z:row").Length > 0 Then
    
      If m_strSession = "" Then
'        Call WSLogon
      End If
      
      If m_oCmErr.Count = 0 Then
'        Call RSExecuteRequest
      End If
    
    End If
    
    ' update all recordsets
    Call RSUpdate
    
    If bDecrease Then
      ' decrease counter
      lNumberOfRuns = lNumberOfRuns - 1
    End If
  
  Loop
  
  ' log-off from webservice if session is available
  If m_strSession <> "" Then
    Call WSLogoff
  End If
  
  ' create eventlog entry
  If m_oCmErr.Count > 0 Then
    Call LogEvent(TXT_APP, EVENTLOG_ERROR_TYPE, "Sync Service Dispatcher : '" & m_oCmErr.GetXML & "'", 1, 0)
    Exit Sub
  End If
  
  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100104, App.Title, TXT_ERR_UNEXPECTED)
  Call LogEvent(TXT_APP, EVENTLOG_ERROR_TYPE, "Sync Service Dispatcher : '" & m_oCmErr.GetXML & "'", 1, 0)

End Sub

'
' WSLogon
'
' Purpose       : log-on to webservice
' Parameters    : none
' Return value  : none
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-18
'
Private Sub WSLogon()

    ' MSXML2 interfaces
    Dim oNode As Object ' MSXML2.IXMLDOMNode
  
  ' activate error handler
  On Error GoTo handle_err
  
  '
  ' execute log-on SOAP request
  '
  With m_oReq
    
    .ClearSession
    .ClearJob
    
    Call m_oXMLRq.loadXML(m_oNCfg.selectSingleNode("./request/logon/cmrequest").xml)
    
    ' set instance
    m_oXMLRq.documentElement.Attributes.getNamedItem("instance").Text = m_strInstance
    
    ' set login
    m_oXMLRq.documentElement.selectSingleNode("./login").Text = m_strLogin
    
    ' set password
    m_oXMLRq.documentElement.selectSingleNode("./password").Text = m_strPW
    
    ' set xml dom document for execution
    .RequestXMLDoc = m_oXMLRq
    
    ' execute call
    Call .Execute
        
    If .ResponseCountErrors = 0 Then
      Set oNode = .ResponseXMLDoc.documentElement.selectSingleNode("session")
      If Not oNode Is Nothing Then m_strSession = oNode.Text
    Else
      Call m_oCmErr.SetXML(.ResponseErrors)
    End If
  
  End With

  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100105, App.Title, "failed to logon to '" & m_oReq.EndPointUrl & "' ('" & m_strInstance & "')")
  Exit Sub

End Sub

'
' WSExec
'
' Purpose       : execute webservice call
' Parameters    : none
' Return value  : none
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-18
'
Private Sub WSExec()

    ' MSXML2 interfaces
    Dim oNode As Object ' MSXML2.IXMLDOMNode
  
  ' activate error handler
  On Error GoTo handle_err
  
  '
  ' execute log-on SOAP request
  '
  With m_oReq
    
    .ClearJob
    
    ' set instance & session
    m_oXMLRq.documentElement.Attributes.getNamedItem("instance").Text = m_strInstance
    m_oXMLRq.documentElement.Attributes.getNamedItem("session").Text = m_strSession
    
    ' set xml dom document for execution
    .RequestXMLDoc = m_oXMLRq
    
    ' execute call
    Call .Execute
  
  End With

  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100106, App.Title, "failed to call webservice '" & m_oReq.EndPointUrl & "' ('" & m_strInstance & "')")
  Exit Sub

End Sub

'
' WSLogoff
'
' Purpose       : log-off from webservice
' Parameters    : none
' Return value  : none
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-18
'
Private Sub WSLogoff()

    ' MSXML2 interfaces
    Dim oNode As Object ' MSXML2.IXMLDOMNode
  
  ' activate error handler
  On Error GoTo handle_err
  
  '
  ' execute log-on SOAP request
  '
  With m_oReq
        
    .ClearJob
    
    Call m_oXMLRq.loadXML(m_oNCfg.selectSingleNode("./request/logoff/cmrequest").xml)
    
    ' set instance & session
    m_oXMLRq.documentElement.Attributes.getNamedItem("instance").Text = m_strInstance
    m_oXMLRq.documentElement.Attributes.getNamedItem("session").Text = m_strSession
    m_oXMLRq.documentElement.selectSingleNode("./session").Text = m_strSession
    
    ' set xml dom document for execution
    .RequestXMLDoc = m_oXMLRq
    
    ' execute call
    Call .Execute
        
    If .ResponseCountErrors = 0 Then
      m_strSession = ""
    Else
      Call m_oCmErr.Set(.ResponseErrors)
    End If
  
  End With

  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100107, App.Title, "failed to logoff from '" & m_oReq.EndPointUrl & "' ('" & m_strInstance & "') : '" & m_strSession & "'")
  Exit Sub

End Sub

'
' RSAnalyse
'
' Purpose       : analyse and prepare recordset for upcoming execution
' Parameters    : none
' Return value  : none
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-18
'
Private Sub RSAnalyse()

    ' MSXML2 interfaces
    Dim oNL As MSXML2.IXMLDOMNodeList
    Dim oNodeNew As Object ' MSXML2.IXMLDOMNode
    Dim oNode As Object ' MSXML2.IXMLDOMNode
    Dim oAttr As Object ' MSXML2.IXMLDOMAttribute
    Dim oNodeStatus As Object ' MSXML2.IXMLDOMNode
 
  ' activate error handler
  On Error GoTo handle_err
  
  Set oNL = m_oXMLRS.documentElement.selectNodes("./rs:data/z:row[@activity='1']")
  If oNL.Length > 1 Then
    Call HandleException(m_oCmErr, 100108, App.Title, "imported data inconsistent (more than one activity=1 (add)) for objectGUID '" & m_oXMLRS.documentElement.selectSingleNode("./rs:data/z:row").Attributes.getNamedItem("objectGUID").Text & "'")
  End If
  
  Set oNL = m_oXMLRS.documentElement.selectNodes("./rs:data/z:row[@activity='3']")
  If oNL.Length > 1 Then
    Call HandleException(m_oCmErr, 100109, App.Title, "imported data inconsistent (more than one activity=3 (del)) for objectGUID '" & m_oXMLRS.documentElement.selectSingleNode("./rs:data/z:row").Attributes.getNamedItem("objectGUID").Text & "'")
  End If
  
  ' if any error occurred, bail out
  If m_oCmErr.Count > 0 Then
    Exit Sub
  End If
    
    Dim lCount As Long
    Dim lCountRecords As Long
    Dim lActivity As Long
    
    Dim bAddOnly As Boolean
    Dim bDelOnly As Boolean
    Dim bSkipAll As Boolean
    Dim bAlreadyUpdated As Boolean
  
  bAddOnly = False
  bDelOnly = False
  bSkipAll = False
  bAlreadyUpdated = False
  
  Set oNL = m_oXMLRS.documentElement.selectNodes("./rs:data/z:row")
  lCountRecords = oNL.Length
  
  ' evaluate first record
  Set oNode = oNL.Item(0)
  lActivity = CLng(oNode.Attributes.getNamedItem("activity").Text)
  If lActivity = 1 Then
    bAddOnly = True
  End If

  ' evaluate last record
  Set oNode = oNL.Item(lCountRecords - 1)
  lActivity = CLng(oNode.Attributes.getNamedItem("activity").Text)
  If lActivity = 3 Then
    bDelOnly = True
  End If

  ' if delete and add are altogether we can skip the whole object. it does not exist anymore.
  If bDelOnly And bAddOnly Then bSkipAll = True

  lCount = 0
  For Each oNode In oNL

    Set oNodeStatus = m_oXMLRS.createElement("status")
    Call oNode.appendChild(oNodeStatus)
    oNodeStatus.Text = status_skipped

    Set oNodeNew = m_oXMLRS.createElement("info")
    Call oNode.appendChild(oNodeNew)

    Set oAttr = m_oXMLRS.createAttribute("noexec")
    Call oNode.Attributes.setNamedItem(oAttr)
    oAttr.Text = 0

    ' activity 2 only (updates)
    If Not bAddOnly And Not bDelOnly Then

      lActivity = CLng(oNode.Attributes.getNamedItem("activity").Text)
      If lActivity = 2 And oAttr.Text = 0 Then
        If Not bAlreadyUpdated Then
          bAlreadyUpdated = True
        Else
          oAttr.Text = 1
        End If
      End If

    Else

      ' activity 1 (add) or activity 3 (del)
      If bSkipAll Then
        ' none of the records left
        oAttr.Text = 1
      Else
        ' the first record only
        If bAddOnly And lCount = 0 Then
          bSkipAll = True
        Else
          ' the last record only
          If bDelOnly And lCount < lCountRecords - 1 Then oAttr.Text = 1
        End If
      End If
    
    End If
    
    ' increase internal counter
    lCount = lCount + 1

  Next

  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100110, App.Title, TXT_ERR_UNEXPECTED)
  Exit Sub

End Sub

'
' RSExecuteRequest
'
' Purpose       : execute webservice call
' Parameters    : none
' Return value  : none
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-18
'
Private Sub RSExecuteRequest()

    ' MSXML2 interfaces
    Dim oNL As MSXML2.IXMLDOMNodeList
    Dim oNodeNew As Object ' MSXML2.IXMLDOMNode
    Dim oNode As Object ' MSXML2.IXMLDOMNode
    Dim oAttr As Object ' MSXML2.IXMLDOMAttribute
    ' node to info structure
    Dim oNInfo As Object ' MSXML2.IXMLDOMNode
  
  ' activate error handler
  On Error GoTo handle_err
  
  ' since we expect exact one webservice call this nodelist is waste of time and code
  ' but who knows ... so we leave it as it is
  Set oNL = m_oXMLRS.documentElement.selectNodes("./rs:data/z:row[@noexec='0']")
  For Each oNode In oNL
    
    Call m_oXMLRq.loadXML(m_oNCfg.selectSingleNode("./request/call/cmrequest").xml)
    
    For Each oAttr In oNode.Attributes
      Set oNodeNew = m_oXMLRq.createElement(LCase(oAttr.Name))
      Call m_oXMLRq.documentElement.appendChild(oNodeNew)
      oNodeNew.Text = oAttr.Text
    Next
    
    ' remove interfering elements
    Set oNInfo = m_oXMLRq.documentElement.selectSingleNode("./id")
    If Not oNInfo Is Nothing Then Call oNInfo.parentNode.removeChild(oNInfo)
    Set oNInfo = m_oXMLRq.documentElement.selectSingleNode("./status")
    If Not oNInfo Is Nothing Then Call oNInfo.parentNode.removeChild(oNInfo)
    Set oNInfo = m_oXMLRq.documentElement.selectSingleNode("./noexec")
    If Not oNInfo Is Nothing Then Call oNInfo.parentNode.removeChild(oNInfo)
    Set oNInfo = m_oXMLRq.documentElement.selectSingleNode("./info")
    If Not oNInfo Is Nothing Then Call oNInfo.parentNode.removeChild(oNInfo)
    
    ' try to get get base command from configuration
    Set oNInfo = m_oNCfg.selectSingleNode("./request/call/basecmd")
    If Not oNInfo Is Nothing Then
      m_oXMLRq.documentElement.Attributes.getNamedItem("cmd").Text = CLng(oNInfo.Text) + CLng(m_oXMLRq.documentElement.selectSingleNode("./activity").Text)
    End If
    
    ' prepare node to info
    Set oNInfo = oNode.selectSingleNode("./info")
    
    ' execute SOAP request
    Call WSExec
    
    ' set state for upcoming sql record update
    If m_oReq.ResponseCountErrors > 0 Or m_oCmErr.Count > 0 Then
      oNode.selectSingleNode("./status").Text = status_err_exec_failed
    Else
      oNode.selectSingleNode("./status").Text = status_success
    End If
  
    ' enrich info element with interesting data :-)
    
    ' on which server did we execute the request
    Set oNodeNew = m_oXMLRS.createElement("endpointurl")
    Call oNInfo.appendChild(oNodeNew)
    oNodeNew.Text = m_oReq.EndPointUrl
  
    ' the instance
    Set oNodeNew = m_oXMLRS.createElement("instance")
    Call oNInfo.appendChild(oNodeNew)
    oNodeNew.Text = m_strInstance
    
    ' jobGUID fpr later request identification
    Set oNodeNew = m_oXMLRS.createElement("job")
    Call oNInfo.appendChild(oNodeNew)
    oNodeNew.Text = m_oReq.Job
    
    ' append errors structure
    If m_oReq.ResponseCountErrors > 0 Then
      Call oNInfo.appendChild(m_oReq.ResponseErrorsObj.GetDOMDoc.documentElement.cloneNode(True))
    End If
    If m_oCmErr.Count > 0 Then
      Call oNInfo.appendChild(m_oCmErr.GetDOMDoc.documentElement.cloneNode(True))
    End If
  
  Next

  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100111, App.Title, TXT_ERR_UNEXPECTED)
  Exit Sub

End Sub

'
' RSUpdate
'
' Purpose       : update recordsets (status + info-xml)
' Parameters    : none
' Return value  : none
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-18
'
Private Sub RSUpdate()

    ' MSXML2 interfaces
    Dim oNL As MSXML2.IXMLDOMNodeList
    Dim oNodeSub As Object ' MSXML2.IXMLDOMNode
    Dim oNode As Object ' MSXML2.IXMLDOMNode
    Dim oAttr As Object ' MSXML2.IXMLDOMAttribute
  
    Dim strSQLUpd As String
    Dim strSQLTmp As String
    Dim strTmp As String
  
  ' activate error handler
  On Error GoTo handle_err
  
  Set oNode = m_oNCfg.selectSingleNode("./sql_b")
  If Not oNode Is Nothing Then strSQLUpd = oNode.Text
  strSQLUpd = Replace(strSQLUpd, "%TABLE%", m_strTablename)
  
  Set oNL = m_oXMLRS.documentElement.selectNodes("./rs:data/z:row")
  For Each oNode In oNL
    
    strSQLTmp = Replace(strSQLUpd, "%ID%", oNode.Attributes.getNamedItem("id").Text)
    
    ' no session indicates that logon failed. reset!
    If m_strSession = "" Then
      strSQLTmp = Replace(strSQLTmp, "%INFO%", "null")
      strSQLTmp = Replace(strSQLTmp, "%STATUS%", "null")
    Else

      Set oNodeSub = oNode.selectSingleNode("./info")
      If Not oNodeSub Is Nothing Then
        strTmp = oNodeSub.xml
      Else
        strTmp = "<info/>"
      End If
      strSQLTmp = Replace(strSQLTmp, "%INFO%", "'" & Replace(strTmp, "'", "''") & "'")
      
      Set oNodeSub = oNode.selectSingleNode("./status")
      If Not oNodeSub Is Nothing Then
        strTmp = oNodeSub.Text
      Else
        strTmp = CStr(status_skipped)
      End If
      strSQLTmp = Replace(strSQLTmp, "%STATUS%", strTmp)
    
    End If
    
    Call SQLExec(strSQLTmp)
  
  Next

  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100112, App.Title, TXT_ERR_UNEXPECTED)
  Exit Sub

End Sub

'
' ExecSQL
'
' Purpose       : execute SQL command text
'
' Parameters    : strConnectionString As String : connection string
'                 strSQL As String              : sql command text
'
' Return value  : none

' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08
'
Private Function SQLExec(strSQL As String) As Boolean

    ' ADO objects
    Dim oCmd As ADODB.Command
    Dim oRs As ADODB.Recordset
    Dim oStream As ADODB.Stream
    Dim oParam As ADODB.Parameter
    Dim oCon As ADODB.Connection

    ' locals
    Dim lAffected As Long
    Dim varTemp As Variant, i As Long
    Dim varRetVal As Variant
  
  ' activate error handler
  On Error GoTo handle_err

  ' open connection
  Set oCon = New ADODB.Connection
  oCon.ConnectionString = m_strConnectionString
  oCon.CommandTimeout = 600
  oCon.CursorLocation = adUseClient
  Call oCon.open
  
  Set oCmd = CreateObject("ADODB.Command")

  ' set connection and SQL query text
  Set oCmd.ActiveConnection = oCon
  oCmd.CommandType = adCmdText
  oCmd.CommandText = strSQL

  ' execute
  Set oRs = oCmd.Execute(lAffected)

  ' step thru command parameters
  For Each oParam In oCmd.Parameters
    Select Case oParam.Direction
      ' set return value, if available
      Case adParamReturnValue
        varRetVal = oParam.Value
      Case adParamInput
      Case adParamOutput
      Case adParamInputOutput
      Case Else
    End Select
  Next

  ' any recordset available
  If Not oRs Is Nothing Then

    ' open ? adStateOpen:=1 adStateClosed:=0
    If oRs.State = adStateOpen Or oRs.State <> adStateClosed Then

      ' convert recordset to XML
      Set oStream = CreateObject("ADODB.Stream")
      Call oRs.save(oStream, adPersistXML)

      ' set return value
      Call m_oXMLRS.loadXML(CStr(oStream.ReadText(adReadAll)))

      ' close open recordset
      oRs.Close

    End If ' if oRs.State = adStateOpen or oRs.State <> adStateClosed then

  End If ' if not oRs is nothing then

  ' close connection
  Call oCon.Close

  SQLExec = True
  
  Exit Function

handle_err:

  SQLExec = False
  
  varTemp = ""
  If oCon.Errors.Count > 0 Then
    i = 0
    While i < oCon.Errors.Count
      varTemp = varTemp & "Source='" & oCon.Errors.Item(i).Source & "' Description = '" & oCon.Errors.Item(i).Description & "' SQLState='" & oCon.Errors.Item(i).SQLState & "' NativeError='" & oCon.Errors.Item(i).NativeError & "' Number='" & oCon.Errors.Item(i).Number & "' ;"
      i = i + 1
    Wend
  End If
    
  Call HandleException(m_oCmErr, 100113, App.Title, "database access failure : '" & varTemp & "'")

End Function

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

'
' DecryptCfg
'
' Purpose       : decrypt all marked elements
' Parameters    : none
' Return value  : none
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-19
'
Private Sub DecryptCfg()

    Dim oNodeList As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMAttribute

    Dim strTemp As String
    Dim bResult As Boolean
    Dim strContent As String

  ' activate error handler
  On Error GoTo handle_err
  
  Set oNodeList = m_oXMLCfg.selectNodes(".//*[@*[name()='enc']]")
  For Each oNode In oNodeList
    Set oAttr = oNode.Attributes.getNamedItem("enc")
    If CLng(oAttr.Text) = 1 Then
      
      oAttr.Text = 0
      
      strTemp = oNode.Text
      bResult = m_cKeyStore.DecryptText(KN_ASE, strTemp, strContent)
      If Not bResult Then
        strContent = ""
        Call HandleException(m_oCmErr, 100115, App.Title, "failed to decrypt configuration")
      End If
      oNode.Text = strContent
    
    End If
  Next

  ' thats it
  Exit Sub

handle_err:

  ' oops.
  Call HandleException(m_oCmErr, 100114, App.Title, "failed while decrypting secured information")

End Sub

'
' Encrypt
'
' Purpose       : encrypt all marked elements
' Parameters    : File : full path and filename (xml)
' Return value  : none
'
' Comment       : internal function
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-08-19
'
Public Sub Encrypt(ByVal File As String)

    Dim oNodeList As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMAttribute
    
    Dim strEn As String
    Dim strTemp As String
    Dim bResult As Boolean
  
  Dim oXML As MSXML2.DOMDocument40
  Set oXML = CreateObject("MSXML2.DOMDocument.4.0")
  oXML.async = False
  If oXML.Load(File) = False Then Exit Sub
  
  Set oNode = oXML.documentElement.selectSingleNode("./app_syncsvcdsp_details")
  If Not oNode Is Nothing Then
    
    Set oNodeList = oNode.selectNodes(".//*[@*[name()='enc']]")
    For Each oNode In oNodeList
      Set oAttr = oNode.Attributes.getNamedItem("enc")
      If CLng(oAttr.Text) = -1 Then
        
        oAttr.Text = 1
        
        strTemp = oNode.Text
        bResult = m_cKeyStore.EncryptText(KN_ASE, strTemp, strEn)
        If Not bResult Then
          Call m_cKeyStore.ErrorsXMLObj.GetDOMDoc.save(File & ".errors")
          Call Err.Raise(vbObjectError, "encrypt", "encryption failed.")
        End If
        oNode.Text = strEn
      
      End If
    Next
    
    Call oXML.save(File)
  
  End If

End Sub

Private Sub Class_Initialize()

  Set m_oXMLCfg = CreateObject("msxml2.DOMDocument.4.0")
  m_oXMLCfg.async = False
  m_oXMLCfg.setProperty "SelectionLanguage", "XPath"

  Set m_oXMLRq = CreateObject("msxml2.DOMDocument.4.0")
  m_oXMLRq.async = False
  m_oXMLRq.setProperty "SelectionLanguage", "XPath"
  
  Set m_oXMLRS = CreateObject("MSXML2.DOMDocument.4.0")
  m_oXMLRS.async = False
  m_oXMLRS.setProperty "SelectionLanguage", "XPath"
  m_oXMLRS.setProperty "SelectionNamespaces", "xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'"

End Sub

