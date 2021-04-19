Attribute VB_Name = "Main"
Option Explicit

' VBA version of original sample code, pulled from:
' https://github.com/microsoft/Windows-classic-samples/blob/master/Samples/WinhttpWebsocket/cpp/WinhttpWebsocket.cpp
' Note: synchronous communication - fails if server does not reply
' ===================================================
' =  Echo server thanks to: https://websocket.org/  =
' ===================================================

Sub Main()

Dim dwError As Long
Dim fStatus As Long   ' in some cases, boolean in disguise, zero => false
Dim dwStatusCode As Long
Dim sizeStatusCode As Long  ' for HTTP result request
sizeStatusCode = 4 ' four bytes for long ' for HTTP result request
dwError = ERROR_SUCCESS
fStatus = False
Dim hSessionHandle As Long
Dim hConnectionHandle As Long
Dim hRequestHandle As Long
Dim hWebSocketHandle As Long
hSessionHandle = 0
hConnectionHandle = 0
hRequestHandle = 0
hWebSocketHandle = 0

' HTTP User-Agent Header
Dim AgentHeader As String
AgentHeader = "Websocket Sample"

' Create session handle
hSessionHandle = WinHttpOpen(StrPtr(AgentHeader), _
      WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, 0, 0, 0)
If hSessionHandle = 0 Then
  dwError = GetLastError
  GoTo quit
End If

' Server Name
Dim ServerName As String
ServerName = "echo.websocket.org"

' Server Port
Dim Port As Long
Port = INTERNET_DEFAULT_HTTP_PORT ' 80

' Create connection handle
hConnectionHandle = WinHttpConnect(hSessionHandle, StrPtr(ServerName), Port, 0)
If hConnectionHandle = 0 Then
  dwError = GetLastError
  GoTo quit
End If

' Request Method & Path
Dim method As String, Path As String
method = "GET" ' always
Path = "/" ' echo.websocket.org

' Create request handle - use 0 for null pointer to empty strings: Version, Referrer, AcceptTypes
hRequestHandle = WinHttpOpenRequest(hConnectionHandle, StrPtr(method), StrPtr(Path), 0, 0, 0, 0)
If hRequestHandle = 0 Then
  dwError = GetLastError
  GoTo quit
End If

' Request client protocol upgrade from http to websocket, returns true if success
fStatus = WinHttpSetOption(hRequestHandle, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0)
If (fStatus = 0) Then
  dwError = GetLastError
  GoTo quit
End If

' Perform websocket handshake by sending the upgrade request to server
' --------------------------------------------------------------------
' Application may specify additional headers if needed.
' --------------------------------------------------------------------
' Not in original code...
' Each header except the last must be terminated by a carriage return/line feed (vbCrLf).
' Uses an odd API feature: pass string length as -1, and API measures string
' Note: This is where websocket (internal, RFC "subprotocol") protocol is set
' --------------------------------------------------------------------
Dim HeaderText As String
Dim HeaderTextLength As Long
HeaderText = ""
HeaderText = HeaderText & "Host: " & ServerName & vbCrLf   ' may be redundant or unnecessary
HeaderText = HeaderText & "Sec-WebSocket-Version: 13" & vbCrLf  ' 8 or 13, may be redundant or unnecessary
HeaderText = HeaderText & "Sec-Websocket-Protocol: echo-protocol" & vbCrLf  ' subprotocol
' setup for API call, trim any trailing vbCrLf
If (Right(HeaderText, 2) = vbCrLf) Then
  HeaderText = Left(HeaderText, Len(HeaderText) - 2)
End If

If Len(HeaderText) > 0 Then ' let the API figure it out
  HeaderTextLength = -1
  fStatus = WinHttpSendRequest(hRequestHandle, StrPtr(HeaderText), _
               HeaderTextLength, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
Else  ' call without adding headers
  fStatus = WinHttpSendRequest(hRequestHandle, WINHTTP_NO_ADDITIONAL_HEADERS, _
               0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
End If
If (fStatus = 0) Then
  dwError = GetLastError
  MsgBox "quitting with dwError: " & dwError
  GoTo quit
End If

' Receive server reply
fStatus = WinHttpReceiveResponse(hRequestHandle, 0)
If (fStatus = 0) Then
  dwError = GetLastError
  GoTo quit
End If

' See if the HTTP Response confirms the upgrade, with HTTP status code 101.
fStatus = WinHttpQueryHeaders(hRequestHandle, _
    (WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER), _
    WINHTTP_HEADER_NAME_BY_INDEX, _
    dwStatusCode, sizeStatusCode, WINHTTP_NO_HEADER_INDEX)
If (fStatus = 0) Then
  dwError = GetLastError
  GoTo quit
End If
If dwStatusCode <> 101 Then
  Debug.Print "Code needs to be 101, ending..."
  dwError = 0
  GoTo quit
End If

' finally, get handle to websocket
hWebSocketHandle = WinHttpWebSocketCompleteUpgrade(hRequestHandle, 0)
If hWebSocketHandle = 0 Then
  dwError = GetLastError
  GoTo quit
End If

' The request handle is not needed anymore. From now on we will use the websocket handle.
WinHttpCloseHandle (hRequestHandle)
hRequestHandle = 0
Debug.Print "Succesfully upgraded to websocket protocol at: " & ServerName & ":" & Port & Path

'first message!
Dim Message As String
Message = "Hello World!"
Dim cdwMessageLength As Long
cdwMessageLength = 2 * Len(Message) ' two bytes per character

' Send and receive data on the websocket protocol.
dwError = WinHttpWebSocketSend(hWebSocketHandle, _
             WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE, StrPtr(Message), cdwMessageLength)
If (dwError <> ERROR_SUCCESS) Then
  dwError = GetLastError
  GoTo quit
End If
Debug.Print "Sent message to the server: " & Message

Dim rgbBuffer(1023) As Byte
Dim dwBufferLength As Long
' get correct value in case of change to buffer size
dwBufferLength = (UBound(rgbBuffer) - LBound(rgbBuffer) + 1)

Dim pbCurrentBufferPointer As Long
pbCurrentBufferPointer = VarPtr(rgbBuffer(0))
Dim dwBytesTransferred As Long
dwBytesTransferred = 0
Dim eBufferType As Long
eBufferType = 0
Dim dwCloseReasonLength As Long
dwCloseReasonLength = 0
Dim usStatus As Integer
usStatus = 0

' loop if we get a message fragment packet
' added a failsafe counter to original code - probably unnecessary
Dim MessageComplete As Boolean, count As Long
MessageComplete = False
count = 0
Do Until (MessageComplete Or (count > 40))
  If (dwBufferLength = 0) Then
    dwError = ERROR_NOT_ENOUGH_MEMORY
    GoTo quit
  End If
  count = count + 1
  Debug.Print "Check receive buffer: " & count
  ' receive - will hang if no server response
  dwError = WinHttpWebSocketReceive(hWebSocketHandle, rgbBuffer(0), dwBufferLength, dwBytesTransferred, eBufferType)
  If (dwError <> ERROR_SUCCESS) Then
    dwError = GetLastError
    GoTo quit
  End If
' If we receive just part of the message restart the receive operation.
  pbCurrentBufferPointer = pbCurrentBufferPointer + dwBytesTransferred
  dwBufferLength = dwBufferLength - dwBytesTransferred
  If Not (eBufferType = WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE) Then
    MessageComplete = True
  End If
Loop

' Expected server to just echo single binary message - complain if different
If (eBufferType <> WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE) Then
  Debug.Print "Unexpected buffer type"
  dwError = 87 ' ERROR_INVALID_PARAMETER
  GoTo quit
End If

' convert buffer into string (crudely, just ignore zeros)
Dim j As Long, strBuffer As String
strBuffer = ""
For j = LBound(rgbBuffer) To UBound(rgbBuffer)
  If rgbBuffer(j) <> 0 Then
    strBuffer = strBuffer & Chr(rgbBuffer(j))
  End If
Next
Debug.Print "Received message from the server: " & strBuffer

' wrap up
dwError = WinHttpWebSocketClose(hWebSocketHandle, WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, 0, 0)
If (dwError <> ERROR_SUCCESS) Then
  dwError = GetLastError
  GoTo quit
End If

' Check close status returned by the server.
Dim rgbCloseReasonBuffer(123) As Byte
Dim dwCloseReasonBufferLength As Long
' get correct value in case of change to buffer size
dwCloseReasonBufferLength = (UBound(rgbCloseReasonBuffer) - LBound(rgbCloseReasonBuffer) + 1)
dwError = WinHttpWebSocketQueryCloseStatus(hWebSocketHandle, usStatus, _
             rgbCloseReasonBuffer(0), dwCloseReasonBufferLength, dwCloseReasonLength)
If (dwError <> ERROR_SUCCESS) Then
  GoTo quit
End If
' report result
If usStatus = WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS Then
  Debug.Print "Program ended correctly"
Else
  Debug.Print "The server closed the connection with status code: " & usStatus
  Dim i As Long, strCloseReason As String
  strCloseReason = ""
  For i = LBound(rgbCloseReasonBuffer) To UBound(rgbCloseReasonBuffer)
    If rgbCloseReasonBuffer(i) <> 0 Then
      strCloseReason = strCloseReason & Chr(rgbCloseReasonBuffer(i))
    End If
  Next
  If Len(strCloseReason) > 0 Then
    Debug.Print " and reason: " & strCloseReason
   End If
End If

quit:
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hWebSocketHandle <> 0) Then
    WinHttpCloseHandle (hWebSocketHandle)
    hWebSocketHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  If (dwError <> ERROR_SUCCESS) Then
    Select Case dwError
      Case 12029
        Debug.Print "Error: 12029: The connection to the server failed"
      Case 12030
        Debug.Print "Error: 12030: The connection with the server has been reset or terminated"
      Case Else
        Debug.Print "Application failed with error: " & dwError
      End Select
  End If
  
End Sub







