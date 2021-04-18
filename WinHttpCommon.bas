Attribute VB_Name = "WinHttpCommon"
Option Explicit

' flags for WinHttpOpen():
Public Const WINHTTP_FLAG_SYNC = &H0          ' session is synchronous - use this
Public Const WINHTTP_FLAG_ASYNC = &H10000000  ' session is asynchronous - not implemented

' codes from GetLastError
' see: https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-
Public Const ERROR_SUCCESS = 0
Public Const ERROR_INVALID_FUNCTION = 1
Public Const ERROR_INVALID_HANDLE = 6
Public Const ERROR_NOT_ENOUGH_MEMORY = 8

' flags for status check
Public Const WINHTTP_QUERY_STATUS_CODE = 19  ' special: part of status line
Public Const WINHTTP_QUERY_FLAG_NUMBER = &H20000000   ' bit flag to get result as number

' WinHttpQueryHeaders constants for code readability
Public Const WINHTTP_HEADER_NAME_BY_INDEX = 0
Public Const WINHTTP_NO_OUTPUT_BUFFER = 0
Public Const WINHTTP_NO_HEADER_INDEX = 0
Public Const WINHTTP_NO_REQUEST_DATA = 0

' a few HTTP Response Status Codes
Public Const HTTP_STATUS_CONTINUE = 100           ' OK to continue with request
Public Const HTTP_STATUS_SWITCH_PROTOCOLS = 101   ' server has switched protocols in upgrade header
Public Const HTTP_STATUS_OK = 200                 ' request completed

Public Const WINHTTP_WEB_SOCKET_BUFFER_TYPE = 0  ' text, not specified in original sample code

Public Const INTERNET_DEFAULT_PORT = 0           ' use the protocol-specific default
Public Const INTERNET_DEFAULT_HTTP_PORT = 80     ' use the HTTP default
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443   ' use the HTTPS default

Public Const WINHTTP_ACCESS_TYPE_DEFAULT_PROXY = 0
Public Const WINHTTP_ACCESS_TYPE_NO_PROXY = 1
Public Const WINHTTP_ACCESS_TYPE_NAMED_PROXY = 3
Public Const WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY = 4

Public Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET = 114
Public Const WINHTTP_NO_ADDITIONAL_HEADERS = 0

' buffer types
Public Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE = 0
Public Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE = 1
Public Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE = 2
Public Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE = 3
Public Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE = 4

' status codes
Public Const WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS = 1000
Public Const WINHTTP_WEB_SOCKET_ENDPOINT_TERMINATED_CLOSE_STATUS = 1001
Public Const WINHTTP_WEB_SOCKET_PROTOCOL_ERROR_CLOSE_STATUS = 1002
Public Const WINHTTP_WEB_SOCKET_INVALID_DATA_TYPE_CLOSE_STATUS = 1003
Public Const WINHTTP_WEB_SOCKET_EMPTY_CLOSE_STATUS = 1005
Public Const WINHTTP_WEB_SOCKET_ABORTED_CLOSE_STATUS = 1006
Public Const WINHTTP_WEB_SOCKET_INVALID_PAYLOAD_CLOSE_STATUS = 1007
Public Const WINHTTP_WEB_SOCKET_POLICY_VIOLATION_CLOSE_STATUS = 1008
Public Const WINHTTP_WEB_SOCKET_MESSAGE_TOO_BIG_CLOSE_STATUS = 1009
Public Const WINHTTP_WEB_SOCKET_UNSUPPORTED_EXTENSIONS_CLOSE_STATUS = 1010
Public Const WINHTTP_WEB_SOCKET_SERVER_ERROR_CLOSE_STATUS = 1011
Public Const WINHTTP_WEB_SOCKET_SECURE_HANDSHAKE_ERROR_CLOSE_STATUS = 1015

' ====================================================
' API functions
' ====================================================
' DO Use StrPtr to pass strings,
' DO NOT append Chr(0) to strings before passing them
' ====================================================

Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Public Declare PtrSafe Function WinHttpOpen Lib "winhttp" ( _
   ByVal pszAgentW As LongPtr, _
   ByVal dwAccessType As Long, _
   ByVal pszProxyW As LongPtr, _
   ByVal pszProxyBypassW As LongPtr, _
   ByVal dwFlags As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpConnect Lib "winhttp" ( _
   ByVal hSession As LongPtr, _
   ByVal pswzServerName As LongPtr, _
   ByVal nServerPort As Long, _
   ByVal dwReserved As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp" ( _
   ByVal hConnect As LongPtr, _
   ByVal pwszVerb As LongPtr, _
   ByVal pwszObjectName As LongPtr, _
   ByVal pwszVersion As LongPtr, _
   ByVal pwszReferrer As LongPtr, _
   ByVal ppwszAcceptTypes As LongPtr, _
   ByVal dwFlags As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpSetOption Lib "winhttp" ( _
   ByVal hInternet As LongPtr, _
   ByVal dwOption As Long, _
   ByVal lpBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpSendRequest Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpszHeaders As LongPtr, _
   ByVal dwHeadersLength As Long, _
   ByVal lpOptional As LongPtr, _
   ByVal dwOptionalLength As Long, _
   ByVal dwTotalLength As Long, _
   ByVal dwContext As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpReserved As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal pContext As LongPtr _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp" ( _
   ByVal hRequest As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketSend Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal eBufferType As Long, _
   ByVal pvBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketReceive Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef pvBuffer As Any, _
   ByVal dwBufferLength As Long, _
   ByRef pdwBytesRead As LongPtr, _
   ByRef peBufferType As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketClose Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal usStatus As Integer, _
   ByVal pvReason As LongPtr, _
   ByVal dwReasonLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketQueryCloseStatus Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef usStatus As Integer, _
   ByRef pvReason As Any, _
   ByVal dwReasonLength As Long, _
   ByRef pdwReasonLengthConsumed As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpQueryHeaders Lib "winhttp" ( _
  ByVal hRequest As LongPtr, _
  ByVal dwInfoLevel As Long, _
  ByVal pwszName As LongPtr, _
  ByRef lpBuffer As Long, _
  ByRef lpdwBufferLength As Long, _
  ByRef lpdwIndex As Long _
   ) As Long
