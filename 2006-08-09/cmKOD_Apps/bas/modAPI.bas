Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, _
   lpSystemTime As SYSTEMTIME) As Long

Public Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" _
    (lpSystemTime As SYSTEMTIME, _
    dbTime As Double) As Long

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


' see MSDN ADSTYPEENUM enumeration

Public Enum ADSTYPE_ENUM
  
  ADSTYPE_INVALID = 0
  ADSTYPE_DN_STRING = 1
  ADSTYPE_CASE_EXACT_STRING = 2
  ADSTYPE_CASE_IGNORE_STRING = 3
  ADSTYPE_PRINTABLE_STRING = 4
  ADSTYPE_NUMERIC_STRING = 5
  ADSTYPE_BOOLEAN = 6
  ADSTYPE_INTEGER = 7
  ADSTYPE_OCTET_STRING = 8
  ADSTYPE_UTC_TIME = 9
  ADSTYPE_LARGE_INTEGER = 10
  ADSTYPE_PROV_SPECIFIC = 11
  ADSTYPE_OBJECT_CLASS = 12
  ADSTYPE_CASEIGNORE_LIST = 13
  ADSTYPE_OCTET_LIST = 14
  ADSTYPE_PATH = 15
  ADSTYPE_POSTALADDRESS = 16
  ADSTYPE_TIMESTAMP = 17
  ADSTYPE_BACKLINK = 18
  ADSTYPE_TYPEDNAME = 19
  ADSTYPE_HOLD = 20
  ADSTYPE_NETADDRESS = 21
  ADSTYPE_REPLICAPOINTER = 22
  ADSTYPE_FAXNUMBER = 23
  ADSTYPE_EMAIL = 24
  ADSTYPE_NT_SECURITY_DESCRIPTOR = 25
  ADSTYPE_UNKNOWN = 26
  ADSTYPE_DN_WITH_BINARY = 27
  ADSTYPE_DN_WITH_STRING = 28

End Enum

Public Enum VT_GENERIC_ENUM

  VT_EMPTY = 0
  VT_NULL = 1
  VT_I2 = 2
  VT_I4 = 3
  VT_R4 = 4
  VT_R8 = 5
  VT_CY = 6
  VT_DATE = 7
  VT_BSTR = 8
  VT_DISPATCH = 9
  VT_ERROR = 10
  VT_BOOL = 11
  VT_VARIANT = 12
  VT_UNKNOWN = 13
  VT_DECIMAL = 14
  VT_I1 = 16
  VT_UI1 = 17
  VT_UI2 = 18
  VT_UI4 = 19
  VT_I8 = 20
  VT_UI8 = 21
  VT_INT = 22
  VT_UINT = 23
  VT_VOID = 24
  VT_HRESULT = 25
  VT_PTR = 26
  VT_SAFEARRAY = 27
  VT_CARRAY = 28
  VT_USERDEFINED = 29
  VT_LPSTR = 30
  VT_LPWSTR = 31
  VT_RECORD = 36
  VT_INT_PTR = 37
  VT_UINT_PTR = 38
  VT_FILETIME = 64
  VT_BLOB = 65
  VT_STREAM = 66
  VT_STORAGE = 67
  VT_STREAMED_OBJECT = 68
  VT_STORED_OBJECT = 69
  VT_BLOB_OBJECT = 70
  VT_CF = 71
  VT_CLSID = 72
  VT_VERSIONED_STREAM = 73
  VT_BSTR_BLOB = &HFFF
  VT_VECTOR = &H1000
  VT_ARRAY = &H2000
  VT_BYREF = &H4000
  VT_RESERVED = &H8000
  VT_ILLEGAL = &HFFFF
  VT_ILLEGALMASKED = &HFFF
  VT_TYPEMASK = &HFFF

End Enum

Public Function GetAsText_ADSTYPE_ENUM(ByVal lValue As Long) As String

  Select Case lValue
  
  Case ADSTYPE_INVALID
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_INVALID"
  Case ADSTYPE_DN_STRING
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_DN_STRING"
  Case ADSTYPE_CASE_EXACT_STRING
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_CASE_EXACT_STRING"
  Case ADSTYPE_CASE_IGNORE_STRING
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_CASE_IGNORE_STRING"
  Case ADSTYPE_PRINTABLE_STRING
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_PRINTABLE_STRING"
  Case ADSTYPE_NUMERIC_STRING
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_NUMERIC_STRING"
  Case ADSTYPE_BOOLEAN
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_BOOLEAN"
  Case ADSTYPE_INTEGER
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_INTEGER"
  Case ADSTYPE_OCTET_STRING
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_OCTET_STRING"
  Case ADSTYPE_UTC_TIME
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_UTC_TIME"
  Case ADSTYPE_LARGE_INTEGER
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_LARGE_INTEGER"
  Case ADSTYPE_PROV_SPECIFIC
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_PROV_SPECIFIC"
  Case ADSTYPE_OBJECT_CLASS
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_OBJECT_CLASS"
  Case ADSTYPE_CASEIGNORE_LIST
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_CASEIGNORE_LIST"
  Case ADSTYPE_OCTET_LIST
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_OCTET_LIST"
  Case ADSTYPE_PATH
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_PATH"
  Case ADSTYPE_POSTALADDRESS
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_POSTALADDRESS"
  Case ADSTYPE_TIMESTAMP
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_TIMESTAMP"
  Case ADSTYPE_BACKLINK
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_BACKLINK"
  Case ADSTYPE_TYPEDNAME
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_TYPEDNAME"
  Case ADSTYPE_HOLD
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_HOLD"
  Case ADSTYPE_NETADDRESS
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_NETADDRESS"
  Case ADSTYPE_REPLICAPOINTER
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_REPLICAPOINTER"
  Case ADSTYPE_FAXNUMBER
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_FAXNUMBER"
  Case ADSTYPE_EMAIL
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_EMAIL"
  Case ADSTYPE_NT_SECURITY_DESCRIPTOR
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_NT_SECURITY_DESCRIPTOR"
  Case ADSTYPE_UNKNOWN
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_UNKNOWN"
  Case ADSTYPE_DN_WITH_BINARY
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_DN_WITH_BINARY"
  Case ADSTYPE_DN_WITH_STRING
    GetAsText_ADSTYPE_ENUM = "ADSTYPE_DN_WITH_STRING"
  Case Else
    GetAsText_ADSTYPE_ENUM = "*no*mapping*unknown*type*"
  End Select

End Function


'
'
'





Public Function GetAsText_VT_GENERIC_ENUM(ByVal lValue As Long) As String

  Select Case lValue
  
  Case VT_EMPTY
    GetAsText_VT_GENERIC_ENUM = "VT_EMPTY"
  Case VT_NULL
    GetAsText_VT_GENERIC_ENUM = "VT_NULL"
  Case VT_I2
    GetAsText_VT_GENERIC_ENUM = "VT_I2"
  Case VT_I4
    GetAsText_VT_GENERIC_ENUM = "VT_I4"
  Case VT_R4
    GetAsText_VT_GENERIC_ENUM = "VT_R4"
  Case VT_R8
    GetAsText_VT_GENERIC_ENUM = "VT_R8"
  Case VT_CY
    GetAsText_VT_GENERIC_ENUM = "VT_CY"
  Case VT_DATE
    GetAsText_VT_GENERIC_ENUM = "VT_DATE"
  Case VT_BSTR
    GetAsText_VT_GENERIC_ENUM = "VT_BSTR"
  Case VT_DISPATCH
    GetAsText_VT_GENERIC_ENUM = "VT_DISPATCH"
  Case VT_ERROR
    GetAsText_VT_GENERIC_ENUM = "VT_ERROR"
  Case VT_BOOL
    GetAsText_VT_GENERIC_ENUM = "VT_BOOL"
  Case VT_VARIANT
    GetAsText_VT_GENERIC_ENUM = "VT_VARIANT"
  Case VT_UNKNOWN
    GetAsText_VT_GENERIC_ENUM = "VT_UNKNOWN"
  Case VT_DECIMAL
    GetAsText_VT_GENERIC_ENUM = "VT_DECIMAL"
  Case VT_I1
    GetAsText_VT_GENERIC_ENUM = "VT_I1"
  Case VT_UI1
    GetAsText_VT_GENERIC_ENUM = "VT_UI1"
  Case VT_UI2
    GetAsText_VT_GENERIC_ENUM = "VT_UI2"
  Case VT_UI4
    GetAsText_VT_GENERIC_ENUM = "VT_UI4"
  Case VT_I8
    GetAsText_VT_GENERIC_ENUM = "VT_I8"
  Case VT_UI8
    GetAsText_VT_GENERIC_ENUM = "VT_UI8"
  Case VT_INT
    GetAsText_VT_GENERIC_ENUM = "VT_INT"
  Case VT_UINT
    GetAsText_VT_GENERIC_ENUM = "VT_UINT"
  Case VT_VOID
    GetAsText_VT_GENERIC_ENUM = "VT_VOID"
  Case VT_HRESULT
    GetAsText_VT_GENERIC_ENUM = "VT_HRESULT"
  Case VT_PTR
    GetAsText_VT_GENERIC_ENUM = "VT_PTR"
  Case VT_SAFEARRAY
    GetAsText_VT_GENERIC_ENUM = "VT_SAFEARRAY"
  Case VT_CARRAY
    GetAsText_VT_GENERIC_ENUM = "VT_CARRAY"
  Case VT_USERDEFINED
    GetAsText_VT_GENERIC_ENUM = "VT_USERDEFINED"
  Case VT_LPSTR
    GetAsText_VT_GENERIC_ENUM = "VT_LPSTR"
  Case VT_LPWSTR
    GetAsText_VT_GENERIC_ENUM = "VT_LPWSTR"
  Case VT_RECORD
    GetAsText_VT_GENERIC_ENUM = "VT_RECORD"
  Case VT_INT_PTR
    GetAsText_VT_GENERIC_ENUM = "VT_INT_PTR"
  Case VT_UINT_PTR
    GetAsText_VT_GENERIC_ENUM = "VT_UINT_PTR"
  Case VT_FILETIME
    GetAsText_VT_GENERIC_ENUM = "VT_FILETIME"
  Case VT_BLOB
    GetAsText_VT_GENERIC_ENUM = "VT_BLOB"
  Case VT_STREAM
    GetAsText_VT_GENERIC_ENUM = "VT_STREAM"
  Case VT_STORAGE
    GetAsText_VT_GENERIC_ENUM = "VT_STORAGE"
  Case VT_STREAMED_OBJECT
    GetAsText_VT_GENERIC_ENUM = "VT_STREAMED_OBJECT"
  Case VT_STREAMED_OBJECT
    GetAsText_VT_GENERIC_ENUM = "VT_STORED_OBJECT"
  Case VT_BLOB_OBJECT
    GetAsText_VT_GENERIC_ENUM = "VT_BLOB_OBJECT"
  Case VT_CF
    GetAsText_VT_GENERIC_ENUM = "VT_CF"
  Case VT_CLSID
    GetAsText_VT_GENERIC_ENUM = "VT_CLSID"
  Case VT_VERSIONED_STREAM
    GetAsText_VT_GENERIC_ENUM = "VT_VERSIONED_STREAM"
  Case VT_BSTR_BLOB
    GetAsText_VT_GENERIC_ENUM = "VT_BSTR_BLOB"
  Case VT_VECTOR
    GetAsText_VT_GENERIC_ENUM = "VT_VECTOR"
  Case VT_ARRAY
    GetAsText_VT_GENERIC_ENUM = "VT_ARRAY"
  Case VT_BYREF
    GetAsText_VT_GENERIC_ENUM = "VT_BYREF"
  Case VT_RESERVED
    GetAsText_VT_GENERIC_ENUM = "VT_RESERVED"
  Case VT_ILLEGAL
    GetAsText_VT_GENERIC_ENUM = "VT_ILLEGAL"
  Case VT_ILLEGALMASKED
    GetAsText_VT_GENERIC_ENUM = "VT_ILLEGALMASKED"
  Case VT_TYPEMASK
    GetAsText_VT_GENERIC_ENUM = "VT_TYPEMASK"
  
  End Select

End Function

' This function will convert the ADSI datatype LargeInteger to a Variant time
' value in Greenwich Mean Time (GMT).
Function LargeInteger_To_Time(oLargeInt As LargeInteger, vTime As Variant) As Boolean
    On Error Resume Next
    Dim pFileTime As FILETIME
    Dim pSysTime As SYSTEMTIME
    Dim dbTime As Double
    Dim lResult As Long
    
    If (oLargeInt.HighPart = 0 And oLargeInt.LowPart = 0) Then
        vTime = 0
        LargeInteger_To_Time = True
        Exit Function
    End If
    
    If (oLargeInt.LowPart = -1) Then
        vTime = -1
        LargeInteger_To_Time = True
        Exit Function
    End If
    
    pFileTime.dwHighDateTime = oLargeInt.HighPart
    pFileTime.dwLowDateTime = oLargeInt.LowPart
    
    ' Convert the FileTime to System time.
    lResult = FileTimeToSystemTime(pFileTime, pSysTime)
    If lResult = 0 Then
        LargeInteger_To_Time = False
        Debug.Print "FileTimeToSystemTime: " + Err.Number + "  - " + Err.Description
        Exit Function
    End If
    
    ' Convert System Time to a Double.
    lResult = SystemTimeToVariantTime(pSysTime, dbTime)
    If lResult = 0 Then
        LargeInteger_To_Time = False
        Debug.Print "SystemTimeToVariantTime: " + Err.Number + "  - " + Err.Description
        Exit Function
    End If
    
    ' Place the double in the variant.
    vTime = CDate(dbTime)
    LargeInteger_To_Time = True

End Function

