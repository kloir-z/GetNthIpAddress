Attribute VB_Name = "Module1"
Option Explicit

Private Type ipBytes
    Octet1 As Byte  ' 最上位バイト
    Octet2 As Byte
    Octet3 As Byte
    Octet4 As Byte  ' 最下位バイト
End Type

' IPアドレス文字列をバイト配列に変換
Private Function ParseIpAddress(ByVal ipAddress As String) As ipBytes
    Dim octets() As String
    octets = Split(ipAddress, ".")
    
    With ParseIpAddress
        .Octet1 = CByte(octets(0))
        .Octet2 = CByte(octets(1))
        .Octet3 = CByte(octets(2))
        .Octet4 = CByte(octets(3))
    End With
End Function

' バイト配列をIPアドレス文字列に変換
Private Function IpBytesToString(ByRef ip As ipBytes) As String
    IpBytesToString = CStr(ip.Octet1) & "." & _
                     CStr(ip.Octet2) & "." & _
                     CStr(ip.Octet3) & "." & _
                     CStr(ip.Octet4)
End Function

' プレフィックス長に基づいてIPアドレスをマスク
Private Function ApplyNetworkMask(ByRef ip As ipBytes, ByVal prefix As Integer) As ipBytes
    Dim remainingBits As Integer
    Dim maskByte As Integer
    
    With ApplyNetworkMask
        ' 最初のオクテット
        remainingBits = WorksheetFunction.Min(8, prefix)
        maskByte = &HFF - (2 ^ (8 - remainingBits) - 1)
        .Octet1 = ip.Octet1 And CByte(maskByte)
        
        ' 2番目のオクテット
        prefix = prefix - remainingBits
        If prefix > 0 Then
            remainingBits = WorksheetFunction.Min(8, prefix)
            maskByte = &HFF - (2 ^ (8 - remainingBits) - 1)
            .Octet2 = ip.Octet2 And CByte(maskByte)
        Else
            .Octet2 = 0
        End If
        
        ' 3番目のオクテット
        prefix = prefix - remainingBits
        If prefix > 0 Then
            remainingBits = WorksheetFunction.Min(8, prefix)
            maskByte = &HFF - (2 ^ (8 - remainingBits) - 1)
            .Octet3 = ip.Octet3 And CByte(maskByte)
        Else
            .Octet3 = 0
        End If
        
        ' 4番目のオクテット
        prefix = prefix - remainingBits
        If prefix > 0 Then
            remainingBits = WorksheetFunction.Min(8, prefix)
            maskByte = &HFF - (2 ^ (8 - remainingBits) - 1)
            .Octet4 = ip.Octet4 And CByte(maskByte)
        Else
            .Octet4 = 0
        End If
    End With
End Function

' n番目のホストアドレスを計算
Private Function AddToIp(ByRef baseIp As ipBytes, ByVal n As Long) As ipBytes
    Dim carry As Long
    Dim result As Long
    
    With AddToIp
        ' 4番目のオクテット
        result = CLng(baseIp.Octet4) + (n Mod 256)
        carry = result \ 256
        .Octet4 = CByte(result Mod 256)
        
        ' 3番目のオクテット
        n = n \ 256
        result = CLng(baseIp.Octet3) + n Mod 256 + carry
        carry = result \ 256
        .Octet3 = CByte(result Mod 256)
        
        ' 2番目のオクテット
        n = n \ 256
        result = CLng(baseIp.Octet2) + n Mod 256 + carry
        carry = result \ 256
        .Octet2 = CByte(result Mod 256)
        
        ' 最初のオクテット
        n = n \ 256
        result = CLng(baseIp.Octet1) + n + carry
        If result > 255 Then
            Err.Raise 6, "AddToIp", "Result exceeds valid IP range"
        End If
        .Octet1 = CByte(result)
    End With
End Function

' CIDR表記を解析
Private Sub ParseCIDR(ByVal cidrNotation As String, ByRef networkAddress As String, ByRef prefix As Integer)
    Dim parts() As String
    parts = Split(cidrNotation, "/")
    
    If UBound(parts) <> 1 Then
        Err.Raise 5, "ParseCIDR", "Invalid CIDR notation"
    End If
    
    networkAddress = parts(0)
    prefix = CInt(parts(1))
End Sub

' メイン関数
Public Function GetNthIpAddress(ByVal cidrNotation As String, ByVal n As Long) As String
    On Error GoTo ErrorHandler
    
    ' CIDR表記の解析
    Dim networkAddress As String
    Dim prefix As Integer
    ParseCIDR cidrNotation, networkAddress, prefix
    
    ' プレフィックス長の検証
    If prefix < 0 Or prefix > 32 Then
        GetNthIpAddress = "Invalid prefix length"
        Exit Function
    End If
    
    ' IPアドレスをバイトに分解
    Dim ipBytes As ipBytes
    ipBytes = ParseIpAddress(networkAddress)
    
    ' ネットワークマスクの適用
    ipBytes = ApplyNetworkMask(ipBytes, prefix)
    
    ' ホスト部のビット数を計算
    Dim hostBits As Integer
    hostBits = 32 - prefix
    
    ' 最大ホスト数を計算
    Dim maxHosts As Long
    Select Case prefix
        Case 32
            ' /32の特殊ケース: 単一のアドレス
            maxHosts = 1
        Case 31
            ' /31の特殊ケース: 2つのアドレスを使用可能
            maxHosts = 2
        Case Else
            If hostBits > 30 Then
                ' 2^31 - 2 would overflow Long
                maxHosts = &H7FFFFFFF
            Else
                ' 通常のケース: ネットワークアドレスとブロードキャストアドレスを除外
                maxHosts = WorksheetFunction.Power(2, hostBits) - 2
            End If
    End Select
    
    ' 範囲チェック
    If n < 1 Or n > maxHosts Then
        GetNthIpAddress = "Invalid host number"
        Exit Function
    End If
    
    ' n番目のアドレスを計算
    Dim resultIp As ipBytes
    Select Case prefix
        Case 32
            ' /32の場合、指定されたアドレスをそのまま返す
            resultIp = ipBytes
        Case 31
            ' /31の場合、n=1で最初のアドレス、n=2で2番目のアドレスを返す
            resultIp = AddToIp(ipBytes, n - 1)
        Case Else
            ' 通常のケース
            resultIp = AddToIp(ipBytes, n)
    End Select
    
    ' 結果を文字列に変換
    GetNthIpAddress = IpBytesToString(resultIp)
    Exit Function
    
ErrorHandler:
    GetNthIpAddress = "Error: " & Err.Description
End Function


' テスト用サブルーチン
Public Sub TestIpCalculator()
    Dim result As String
    
    ' /32のテスト
    result = GetNthIpAddress("192.168.1.1/32", 1)
    Debug.Print "192.168.1.1/32の1番目: " & result  ' 192.168.1.1
    result = GetNthIpAddress("192.168.1.1/32", 2)
    Debug.Print "192.168.1.1/32の2番目: " & result  ' Invalid host number
    
    ' /31のテスト
    result = GetNthIpAddress("192.168.1.0/31", 1)
    Debug.Print "192.168.1.0/31の1番目: " & result  ' 192.168.1.0
    result = GetNthIpAddress("192.168.1.0/31", 2)
    Debug.Print "192.168.1.0/31の2番目: " & result  ' 192.168.1.1
    
    ' 通常のケース
    result = GetNthIpAddress("192.168.1.0/24", 1)
    Debug.Print "192.168.1.0/24の1番目: " & result  ' 192.168.1.1
    result = GetNthIpAddress("192.168.1.0/30", 1)
    Debug.Print "192.168.1.0/30の1番目: " & result  ' 192.168.1.1
End Sub
