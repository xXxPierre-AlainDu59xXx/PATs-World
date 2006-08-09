Attribute VB_Name = "modErr"
Option Explicit

Global Const TXT_ERR_UNEXPECTED = "Unexpected exception - please contact your system administrator"

'
' HandleException
'
' Purpose       : handles errors and generates error information elements
'
' Parameters    : ByRef poErr           : reference to error management object
'                 ByVal lCmCode         : cmCode
'                 ByVal strSource       : Source
'                 ByVal strDetails      : Details
'
' Return value  : none
'
' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-04-21
'
Public Sub HandleException(ByRef poErr As CMEGG_XMLLib.Errors, ByVal lCmCode, ByVal strSource, ByVal strDetails)

  With poErr

    Call .Add(lCmCode, strSource, strDetails)

    ' if this flag is set we assume that the internal err-object is initialized
    If Err.Number <> 0 Then

      ' take the internal err-object's data
      .Nr = Err.Number
      .Description = Err.Description
      .Src = Err.Source
      .Fn = ""

      ' do not forget to clean up
      Call Err.Clear

    Else

      ' no exception - no information !
      .Nr = -1
      .Description = ""
      .Src = ""
      .Fn = ""

    End If

  End With

  ' thats it. bye.

End Sub

'
' HandleXMLException
'
' Purpose       : handles errors and generates error information elements
'
' Parameters    : ByRef poParseError    : reference to IParseError (MSXMLDOM)
'                 ByRef poErr           : reference to error management object
'                 ByVal lCmCode         : cmCode
'                 ByVal strSource       : Source
'                 ByVal strDetails      : Details
'
' Return value  : none
'
' Comment       :
' Side Effects  :
' Prerequisites :
'
' Responsible   : tr036 2005-07-13
'
Public Sub HandleXMLException(ByRef poParseError As MSXML2.IXMLDOMParseError, ByRef poErr As CMEGG_XMLLib.Errors, ByVal lCmCode, ByVal strSource, ByVal strDetails)

  With poErr

    Call .Add(lCmCode, strSource, strDetails)

    .Nr = poParseError.errorCode
    .Description = Replace(Replace("['" & poParseError.reason & "'] [line='" & poParseError.Line & "'] [srcText='" & poParseError.srcText & "']", vbLf, " "), vbCr, " ")
    .Src = ""
    .Fn = poParseError.url
  
  End With

  ' thats it. bye.

End Sub

