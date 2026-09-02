Attribute VB_Name = "VBA_WinsockAPI_UDP_IPv6_Sample"
Option Explicit

'===============================================================================
' Excel VBA Winsock UDP/IPv6 送信側・受信側サンプル
'-------------------------------------------------------------------------------
' WindowsのWinsock APIをVBAから直接呼び出し、IPv6ループバックアドレス
' 「::1」上でUDPデータグラムを送受信します。外部コマンドは使用しません。
'
' UDPで特に理解してほしい点
'   ・connect/acceptを行わず、データグラム単位で送受信します。
'   ・sendtoの成功は「相手が受信した」ことを保証しません。
'   ・到達確認、順序保証、再送はUDP自身にはありません。
'   ・この教材では応答を返す簡易サーバーを用意し、結果を目で確認します。
'===============================================================================

Private Const AF_INET6 As Long = 23
Private Const SOCK_DGRAM As Long = 2
Private Const IPPROTO_UDP As Long = 17
Private Const IPPROTO_IPV6 As Long = 41
Private Const IPV6_V6ONLY As Long = 27
Private Const SOL_SOCKET As Long = &HFFFF&
' WindowsでSO_REUSEADDRを使うと、使用中ポートへの強制bindや不定な配送を
' 招く場合があります。本教材では共有せず、排他的にポートを確保します。
Private Const SO_EXCLUSIVEADDRUSE As Long = -5
Private Const SO_RCVTIMEO As Long = &H1006&
Private Const SO_SNDTIMEO As Long = &H1005&
Private Const SOCKET_ERROR As Long = -1
Private Const INVALID_SOCKET_VALUE As Long = -1

Private Const DEFAULT_SERVER_ADDRESS As String = "::1"
Private Const DEFAULT_SERVER_PORT As Long = 60053
Private Const RECEIVE_BUFFER_SIZE As Long = 8192
Private Const SOCKET_TIMEOUT_MS As Long = 5000

' 起動マクロが終了した後も別Excelプロセスを確実に保持するための参照です。
' サーバー終了時には読み取り専用ブック、Excelプロセスの順で明示的に解放します。
Private m_serverExcelInstance As Excel.Application
Private m_serverWorkbookInstance As Excel.Workbook

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
#If Win64 Then
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As LongPtr
#End If
    'C言語側はchar配列です。VBAのStringでは構造体サイズが変わるため、
    '1要素1バイトのByte配列でWindows SDKのレイアウトに合わせます。
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
#If Not Win64 Then
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As LongPtr
#End If
End Type

Private Type SOCKADDR_IN6
    sin6_family As Integer
    sin6_port As Integer
    sin6_flowinfo As Long
    sin6_addr(0 To 15) As Byte
    sin6_scope_id As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" ( _
    ByVal requestedVersion As Integer, ByRef data As WSADATA) As Long
Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function socket Lib "ws2_32.dll" ( _
    ByVal addressFamily As Long, ByVal socketType As Long, _
    ByVal protocol As Long) As LongPtr
Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr) As Long
Private Declare PtrSafe Function bind Lib "ws2_32.dll" Alias "bind" ( _
    ByVal socketHandle As LongPtr, ByRef socketAddress As SOCKADDR_IN6, _
    ByVal addressLength As Long) As Long
Private Declare PtrSafe Function sendto Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, ByRef buffer As Any, _
    ByVal bufferLength As Long, ByVal flags As Long, _
    ByRef destinationAddress As SOCKADDR_IN6, ByVal addressLength As Long) As Long
Private Declare PtrSafe Function recvfrom Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, ByRef buffer As Any, _
    ByVal bufferLength As Long, ByVal flags As Long, _
    ByRef sourceAddress As SOCKADDR_IN6, ByRef addressLength As Long) As Long
Private Declare PtrSafe Function setsockopt Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, ByVal level As Long, ByVal optionName As Long, _
    ByRef optionValue As Any, ByVal optionLength As Long) As Long
Private Declare PtrSafe Function htons Lib "ws2_32.dll" ( _
    ByVal hostValue As Integer) As Integer
Private Declare PtrSafe Function ntohs Lib "ws2_32.dll" ( _
    ByVal networkValue As Integer) As Integer
Private Declare PtrSafe Function InetPtonW Lib "ws2_32.dll" ( _
    ByVal addressFamily As Long, ByVal addressText As LongPtr, _
    ByRef binaryAddress As Any) As Long
Private Declare PtrSafe Function InetNtopW Lib "ws2_32.dll" ( _
    ByVal addressFamily As Long, ByRef binaryAddress As Any, _
    ByVal outputText As LongPtr, ByVal outputCharCount As LongPtr) As LongPtr
#End If

'-------------------------------------------------------------------------------
' 初心者向け実行マクロ
'-------------------------------------------------------------------------------

Public Sub StartUdpIPv6ServerInNewExcel()
    'Excel VBA自身の型を使い、Open/Closeの名前付き引数をコンパイル時に確認します。
    Dim errorText As String

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "先にこのブックをマクロ有効ブックとして保存してください。", _
               vbInformation
        Exit Sub
    End If

    If Not m_serverExcelInstance Is Nothing Then
        MsgBox "このブックから起動したサーバー用Excelは、すでに存在します。", _
               vbInformation
        Exit Sub
    End If

    On Error GoTo ERROR_HANDLER

    Set m_serverExcelInstance = CreateObject("Excel.Application")
    m_serverExcelInstance.Visible = True
    Set m_serverWorkbookInstance = m_serverExcelInstance.Workbooks.Open( _
                         Filename:=ThisWorkbook.FullName, _
                         UpdateLinks:=False, _
                         ReadOnly:=True, _
                         IgnoreReadOnlyRecommended:=True)
    m_serverExcelInstance.Run WorkbookMacroReference( _
                        CStr(m_serverWorkbookInstance.Name), "ScheduleUdpIPv6Server")

    MsgBox "IPv6 UDPサーバー用のExcelを起動しました。" & vbCrLf & _
           "サーバー側TraceシートのSERVER_READYを確認してください。", vbInformation
    Exit Sub

ERROR_HANDLER:
    errorText = Err.Description
    '起動途中で失敗した場合は、このマクロが作成した別Excelだけを閉じます。
    On Error Resume Next
    If Not m_serverWorkbookInstance Is Nothing Then _
        m_serverWorkbookInstance.Close SaveChanges:=False
    If Not m_serverExcelInstance Is Nothing Then m_serverExcelInstance.Quit
    Set m_serverWorkbookInstance = Nothing
    Set m_serverExcelInstance = Nothing
    On Error GoTo 0
    MsgBox "サーバー用Excelを起動できませんでした。" & vbCrLf & _
           errorText, vbExclamation
End Sub

Public Sub ScheduleUdpIPv6Server()
    Application.OnTime Now + TimeSerial(0, 0, 1), _
                       WorkbookMacroReference( _
                           ThisWorkbook.Name, "RunUdpIPv6Server")
End Sub

'ネイティブAPIを呼ぶ前に、VBA側の構造体サイズを単独で確認するマクロです。
Public Sub CheckUdpIPv6StructureSizes()
    ValidateWinsockStructureSizes
    MsgBox "構造体サイズは想定どおりです。" & vbCrLf & _
           "SOCKADDR_IN6=28 bytes", vbInformation
End Sub

Public Sub SendHello()
    SendUdpIPv6Command "HELLO", vbNullString
End Sub

Public Sub SendEchoSample()
    SendUdpIPv6Command "ECHO", "こんにちは。IPv6 UDP通信の確認です。"
End Sub

Public Sub StopUdpIPv6Server()
    SendUdpIPv6Command "QUIT", vbNullString
End Sub

' 終了応答を受け取った後、サーバー側TraceのSHUTDOWNを確認してから実行し、
' サーバー用ブックを保存せず閉じて専用Excelプロセスも終了します。
Public Sub CloseUdpIPv6ServerExcel()
    Dim closeError As String

    If m_serverExcelInstance Is Nothing Then Exit Sub

    On Error Resume Next
    If Not m_serverWorkbookInstance Is Nothing Then
        m_serverWorkbookInstance.Close SaveChanges:=False
    End If
    m_serverExcelInstance.Quit
    If Err.Number <> 0 Then closeError = Err.Description
    Set m_serverWorkbookInstance = Nothing
    Set m_serverExcelInstance = Nothing
    On Error GoTo 0

    If Len(closeError) > 0 Then
        MsgBox "サーバー用Excelの終了処理を確認してください。" & vbCrLf & _
               closeError, vbExclamation
    End If
End Sub

'-------------------------------------------------------------------------------
' UDP/IPv6 サーバー
'-------------------------------------------------------------------------------

Public Sub RunUdpIPv6Server()
    Dim winsockData As WSADATA
    Dim serverSocket As LongPtr
    Dim serverAddress As SOCKADDR_IN6
    Dim clientAddress As SOCKADDR_IN6
    Dim clientAddressLength As Long
    Dim returnCode As Long
    Dim optionValue As Long
    Dim winsockStarted As Boolean
    Dim exitRequested As Boolean
    Dim requestText As String
    Dim responseText As String
    Dim fields As Variant
    Dim traceId As String
    Dim commandName As String
    Dim payload As String
    Dim serverTraceId As String
    Dim startedAt As Double

    On Error GoTo ERROR_HANDLER

    serverSocket = INVALID_SOCKET_VALUE
    serverTraceId = NewTraceId("UDP6-SERVER")
    startedAt = Timer

    WriteTrace serverTraceId, "SERVER", "STARTUP", "LOCAL", _
               "Winsock 2.2を初期化します。", "START", 0

    ValidateWinsockStructureSizes
    returnCode = WSAStartup(MakeWord(2, 2), winsockData)
    If returnCode <> 0 Then
        Err.Raise vbObjectError + 2301, "RunUdpIPv6Server", _
                  "WSAStartupに失敗しました。エラーコード=" & CStr(returnCode)
    End If
    winsockStarted = True

    serverSocket = socket(AF_INET6, SOCK_DGRAM, IPPROTO_UDP)
    If serverSocket = INVALID_SOCKET_VALUE Then
        RaiseLastSocketError "IPv6 UDPソケットを作成できませんでした。"
    End If

    optionValue = 1
    returnCode = setsockopt(serverSocket, SOL_SOCKET, SO_EXCLUSIVEADDRUSE, _
                            optionValue, LenB(optionValue))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "SO_EXCLUSIVEADDRUSEを設定できませんでした。"
    End If

    optionValue = 1
    returnCode = setsockopt(serverSocket, IPPROTO_IPV6, IPV6_V6ONLY, _
                            optionValue, LenB(optionValue))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "IPV6_V6ONLYを設定できませんでした。"
    End If

    If Not FillIPv6Address(serverAddress, DEFAULT_SERVER_ADDRESS, DEFAULT_SERVER_PORT) Then
        Err.Raise vbObjectError + 2302, "RunUdpIPv6Server", _
                  "サーバーIPv6アドレスを変換できませんでした。"
    End If

    returnCode = bind(serverSocket, serverAddress, LenB(serverAddress))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "IPv6アドレスとポートのbindに失敗しました。"
    End If

    WriteTrace serverTraceId, "SERVER", "SERVER_READY", "LOCAL", _
               EndpointText(serverAddress) & " でデータグラムを待ちます。", _
               "READY", ElapsedMilliseconds(startedAt)

    Do While Not exitRequested
        clientAddressLength = LenB(clientAddress)

        'recvfromはデータグラムが届くまで待ちます。別Excelプロセスで実行するため、
        '普段操作するクライアント側Excelはこの待機の影響を受けません。
        requestText = ReceiveDatagram(serverSocket, clientAddress, clientAddressLength)
        fields = Split(TrimLineEnding(requestText), vbTab, 3)

        If UBound(fields) < 1 Or Not IsSafeTraceId(CStr(fields(0))) Then
            traceId = serverTraceId
            commandName = "INVALID"
            payload = vbNullString
        Else
            traceId = CStr(fields(0))
            commandName = UCase$(CStr(fields(1)))
            If UBound(fields) >= 2 Then payload = CStr(fields(2)) Else payload = vbNullString
        End If

        WriteTrace traceId, "SERVER", "DATAGRAM_RECEIVED", "IN", _
                   EndpointText(clientAddress) & ", Command=" & commandName & _
                   ", Payload=" & payload, "OK", ElapsedMilliseconds(startedAt)

        Select Case commandName
            Case "HELLO"
                responseText = traceId & vbTab & "OK" & vbTab & _
                               "HELLO from the Excel VBA IPv6 UDP server." & vbLf
            Case "ECHO"
                responseText = traceId & vbTab & "OK" & vbTab & payload & vbLf
            Case "QUIT"
                responseText = traceId & vbTab & "OK" & vbTab & _
                               "The IPv6 UDP server will stop." & vbLf
                exitRequested = True
            Case Else
                responseText = traceId & vbTab & "ERROR" & vbTab & _
                               "Unknown command." & vbLf
        End Select

        SendDatagram serverSocket, responseText, clientAddress
        WriteTrace traceId, "SERVER", "DATAGRAM_SENT", "OUT", _
                   EndpointText(clientAddress) & ", " & TrimLineEnding(responseText), _
                   "SENT", ElapsedMilliseconds(startedAt)
    Loop

    WriteTrace serverTraceId, "SERVER", "SHUTDOWN", "LOCAL", _
               "QUITを受信したため受信ループを終了します。", "COMPLETED", _
               ElapsedMilliseconds(startedAt)

CLEANUP:
    If serverSocket <> INVALID_SOCKET_VALUE Then closesocket serverSocket
    If winsockStarted Then WSACleanup
    Exit Sub

ERROR_HANDLER:
    WriteTrace serverTraceId, "SERVER", "ERROR", "LOCAL", _
               Err.Description, "FAILED", ElapsedMilliseconds(startedAt)
    MsgBox "IPv6 UDPサーバーでエラーが発生しました。" & vbCrLf & _
           Err.Description, vbExclamation
    Resume CLEANUP
End Sub

'-------------------------------------------------------------------------------
' UDP/IPv6 クライアント
'-------------------------------------------------------------------------------

Private Function SendUdpIPv6Command( _
    ByVal commandName As String, _
    ByVal payload As String) As Boolean
    Dim winsockData As WSADATA
    Dim clientSocket As LongPtr
    Dim serverAddress As SOCKADDR_IN6
    Dim responseSource As SOCKADDR_IN6
    Dim responseSourceLength As Long
    Dim returnCode As Long
    Dim winsockStarted As Boolean
    Dim traceId As String
    Dim requestText As String
    Dim responseText As String
    Dim startedAt As Double

    On Error GoTo ERROR_HANDLER

    clientSocket = INVALID_SOCKET_VALUE
    traceId = NewTraceId("UDP6")
    startedAt = Timer
    requestText = traceId & vbTab & UCase$(commandName) & vbTab & payload & vbLf

    WriteTrace traceId, "CLIENT", "STARTUP", "LOCAL", _
               "Winsock 2.2を初期化します。", "START", 0

    ValidateWinsockStructureSizes
    returnCode = WSAStartup(MakeWord(2, 2), winsockData)
    If returnCode <> 0 Then
        Err.Raise vbObjectError + 2311, "SendUdpIPv6Command", _
                  "WSAStartupに失敗しました。エラーコード=" & CStr(returnCode)
    End If
    winsockStarted = True

    clientSocket = socket(AF_INET6, SOCK_DGRAM, IPPROTO_UDP)
    If clientSocket = INVALID_SOCKET_VALUE Then
        RaiseLastSocketError "IPv6 UDPクライアントソケットを作成できませんでした。"
    End If
    SetSocketTimeouts clientSocket, SOCKET_TIMEOUT_MS

    If Not FillIPv6Address(serverAddress, DEFAULT_SERVER_ADDRESS, DEFAULT_SERVER_PORT) Then
        Err.Raise vbObjectError + 2312, "SendUdpIPv6Command", _
                  "送信先IPv6アドレスを変換できませんでした。"
    End If

    SendDatagram clientSocket, requestText, serverAddress
    WriteTrace traceId, "CLIENT", "DATAGRAM_SENT", "OUT", _
               EndpointText(serverAddress) & ", Command=" & UCase$(commandName) & _
               ", Payload=" & payload, "SENT", ElapsedMilliseconds(startedAt)

    'UDPのsendtoが成功しても相手の受信は保証されません。
    'このサンプルではサーバーからの応答を最大5秒待って、到達を確認します。
    responseSourceLength = LenB(responseSource)
    responseText = ReceiveDatagram(clientSocket, responseSource, responseSourceLength)
    ValidateResponseEndpoint responseSource, serverAddress
    ValidateResponseTraceId responseText, traceId

    WriteTrace traceId, "CLIENT", "DATAGRAM_RECEIVED", "IN", _
               EndpointText(responseSource) & ", " & TrimLineEnding(responseText), _
               "COMPLETED", ElapsedMilliseconds(startedAt)

    MsgBox "IPv6 UDPサーバーから応答を受信しました。" & vbCrLf & _
           TrimLineEnding(responseText), vbInformation
    SendUdpIPv6Command = True

CLEANUP:
    If clientSocket <> INVALID_SOCKET_VALUE Then closesocket clientSocket
    If winsockStarted Then WSACleanup
    Exit Function

ERROR_HANDLER:
    WriteTrace traceId, "CLIENT", "ERROR", "LOCAL", _
               Err.Description, "FAILED", ElapsedMilliseconds(startedAt)
    MsgBox "IPv6 UDPクライアントでエラーが発生しました。" & vbCrLf & _
           Err.Description, vbExclamation
    Resume CLEANUP
End Function

'-------------------------------------------------------------------------------
' Winsock補助処理
'-------------------------------------------------------------------------------

Private Sub SendDatagram(ByVal socketHandle As LongPtr, _
                         ByVal value As String, _
                         ByRef destination As SOCKADDR_IN6)
    Dim bytes() As Byte
    Dim byteCount As Long
    Dim sentBytes As Long

    byteCount = StringToUtf8(value, bytes)
    sentBytes = sendto(socketHandle, bytes(0), byteCount, 0, _
                       destination, LenB(destination))
    If sentBytes = SOCKET_ERROR Then
        RaiseLastSocketError "sendtoに失敗しました。"
    ElseIf sentBytes <> byteCount Then
        Err.Raise vbObjectError + 2321, "SendDatagram", _
                  "データグラム全体を送信できませんでした。"
    End If
End Sub

Private Function ReceiveDatagram(ByVal socketHandle As LongPtr, _
                                 ByRef source As SOCKADDR_IN6, _
                                 ByRef sourceLength As Long) As String
    Dim buffer(0 To RECEIVE_BUFFER_SIZE - 1) As Byte
    Dim receivedBytes As Long

    receivedBytes = recvfrom(socketHandle, buffer(0), RECEIVE_BUFFER_SIZE, 0, _
                             source, sourceLength)
    If receivedBytes = SOCKET_ERROR Then
        RaiseLastSocketError "recvfromに失敗またはタイムアウトしました。"
    End If

    ReceiveDatagram = Utf8ToString(buffer, receivedBytes)
End Function

Private Function FillIPv6Address(ByRef result As SOCKADDR_IN6, _
                                 ByVal addressText As String, _
                                 ByVal portNumber As Long) As Boolean
    If portNumber < 0 Or portNumber > 65535 Then
        Err.Raise vbObjectError + 2322, "FillIPv6Address", _
                  "ポート番号は0～65535で指定してください。"
    End If

    result.sin6_family = AF_INET6
    result.sin6_port = htons(ToSignedInteger(portNumber))
    result.sin6_flowinfo = 0
    result.sin6_scope_id = 0
    FillIPv6Address = (InetPtonW(AF_INET6, StrPtr(addressText), _
                       result.sin6_addr(0)) = 1)
End Function

Private Sub SetSocketTimeouts(ByVal socketHandle As LongPtr, ByVal timeoutMs As Long)
    Dim returnCode As Long

    returnCode = setsockopt(socketHandle, SOL_SOCKET, SO_RCVTIMEO, _
                            timeoutMs, LenB(timeoutMs))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "受信タイムアウトを設定できませんでした。"
    End If

    returnCode = setsockopt(socketHandle, SOL_SOCKET, SO_SNDTIMEO, _
                            timeoutMs, LenB(timeoutMs))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "送信タイムアウトを設定できませんでした。"
    End If
End Sub

Private Function EndpointText(ByRef value As SOCKADDR_IN6) As String
    Dim addressBuffer As String
    Dim convertedPointer As LongPtr
    Dim nullPosition As Long

    addressBuffer = String$(46, vbNullChar)
    convertedPointer = InetNtopW(AF_INET6, value.sin6_addr(0), _
                                 StrPtr(addressBuffer), Len(addressBuffer))
    If convertedPointer = 0 Then
        EndpointText = "[IPv6 address conversion failed]"
    Else
        nullPosition = InStr(addressBuffer, vbNullChar)
        If nullPosition > 0 Then addressBuffer = Left$(addressBuffer, nullPosition - 1)
        EndpointText = "[" & addressBuffer & "]:" & _
                       CStr(UnsignedInteger(ntohs(value.sin6_port)))
    End If
End Function

'UDPは接続を確立しないため、recvfromが返した送信元を明示的に確認します。
'別のローカルプロセスが同じクライアントポートへ送ったデータを、正しい応答と
'誤認しないための最小限の検証です。
Private Sub ValidateResponseEndpoint( _
    ByRef actualSource As SOCKADDR_IN6, _
    ByRef expectedSource As SOCKADDR_IN6)

    Dim index As Long

    If actualSource.sin6_family <> expectedSource.sin6_family Or _
       actualSource.sin6_port <> expectedSource.sin6_port Or _
       actualSource.sin6_scope_id <> expectedSource.sin6_scope_id Then
        Err.Raise vbObjectError + 2323, "ValidateResponseEndpoint", _
                  "UDP応答の送信元アドレスまたはポートが期待値と異なります。"
    End If

    For index = 0 To 15
        If actualSource.sin6_addr(index) <> expectedSource.sin6_addr(index) Then
            Err.Raise vbObjectError + 2324, "ValidateResponseEndpoint", _
                      "UDP応答の送信元IPv6アドレスが期待値と異なります。"
        End If
    Next index
End Sub

Private Sub ValidateResponseTraceId( _
    ByVal responseText As String, _
    ByVal expectedTraceId As String)

    Dim fields As Variant

    fields = Split(TrimLineEnding(responseText), vbTab, 3)
    If UBound(fields) < 1 Then
        Err.Raise vbObjectError + 2325, "ValidateResponseTraceId", _
                  "UDP応答の形式が不正です。"
    End If

    If StrComp(CStr(fields(0)), expectedTraceId, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 2326, "ValidateResponseTraceId", _
                  "UDP応答のTraceIdが要求と一致しません。"
    End If
End Sub

Private Function IsSafeTraceId(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String

    If Len(value) = 0 Or Len(value) > 128 Then Exit Function

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If Not ((character >= "0" And character <= "9") Or _
                (character >= "A" And character <= "Z") Or _
                (character >= "a" And character <= "z") Or _
                character = "-") Then Exit Function
    Next index

    IsSafeTraceId = True
End Function

'WSADATAは32bit/64bitでメンバー順とサイズが異なります。
'Windows APIに渡す前に実際のLenBを確認し、範囲外書き込みを防ぎます。
Private Sub ValidateWinsockStructureSizes()
    Dim winsockData As WSADATA
    Dim ipv6Address As SOCKADDR_IN6
    Dim expectedWsaDataBytes As Long

#If Win64 Then
    expectedWsaDataBytes = 408
#Else
    expectedWsaDataBytes = 400
#End If

    If LenB(winsockData) <> expectedWsaDataBytes Then
        Err.Raise vbObjectError + 2327, "ValidateWinsockStructureSizes", _
                  "WSADATAのサイズが想定と異なります。actual=" & _
                  CStr(LenB(winsockData)) & ", expected=" & _
                  CStr(expectedWsaDataBytes)
    End If

    If LenB(ipv6Address) <> 28 Then
        Err.Raise vbObjectError + 2328, "ValidateWinsockStructureSizes", _
                  "SOCKADDR_IN6のサイズが想定と異なります。actual=" & _
                  CStr(LenB(ipv6Address)) & ", expected=28"
    End If
End Sub

Private Function WorkbookMacroReference( _
    ByVal workbookName As String, _
    ByVal macroName As String) As String

    WorkbookMacroReference = "'" & Replace(workbookName, "'", "''") & _
                             "'!" & macroName
End Function

Private Function TrimLineEnding(ByVal value As String) As String
    Do While Len(value) > 0 And _
             (Right$(value, 1) = vbCr Or Right$(value, 1) = vbLf)
        value = Left$(value, Len(value) - 1)
    Loop
    TrimLineEnding = value
End Function

Private Function ToSignedInteger(ByVal unsignedValue As Long) As Integer
    If unsignedValue <= 32767 Then
        ToSignedInteger = CInt(unsignedValue)
    Else
        ToSignedInteger = CInt(unsignedValue - 65536)
    End If
End Function

Private Function UnsignedInteger(ByVal signedValue As Integer) As Long
    UnsignedInteger = signedValue And &HFFFF&
End Function

Private Function MakeWord(ByVal lowByte As Byte, ByVal highByte As Byte) As Integer
    MakeWord = lowByte Or (CLng(highByte) * 256&)
End Function

Private Sub RaiseLastSocketError(ByVal message As String)
    Dim errorCode As Long

    errorCode = WSAGetLastError()
    Err.Raise vbObjectError + 2399, "Winsock", _
              message & " Winsockエラー=" & CStr(errorCode)
End Sub
