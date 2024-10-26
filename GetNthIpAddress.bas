Attribute VB_Name = "Module1"
Option Explicit

Private Type ipBytes
    Octet1 As Byte  ' �ŏ�ʃo�C�g
    Octet2 As Byte
    Octet3 As Byte
    Octet4 As Byte  ' �ŉ��ʃo�C�g
End Type

' IP�A�h���X��������o�C�g�z��ɕϊ�
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

' �o�C�g�z���IP�A�h���X������ɕϊ�
Private Function IpBytesToString(ByRef ip As ipBytes) As String
    IpBytesToString = CStr(ip.Octet1) & "." & _
                     CStr(ip.Octet2) & "." & _
                     CStr(ip.Octet3) & "." & _
                     CStr(ip.Octet4)
End Function

' �v���t�B�b�N�X���Ɋ�Â���IP�A�h���X���}�X�N
Private Function ApplyNetworkMask(ByRef ip As ipBytes, ByVal prefix As Integer) As ipBytes
    Dim remainingBits As Integer
    Dim maskByte As Integer
    
    With ApplyNetworkMask
        ' �ŏ��̃I�N�e�b�g
        remainingBits = WorksheetFunction.Min(8, prefix)
        maskByte = &HFF - (2 ^ (8 - remainingBits) - 1)
        .Octet1 = ip.Octet1 And CByte(maskByte)
        
        ' 2�Ԗڂ̃I�N�e�b�g
        prefix = prefix - remainingBits
        If prefix > 0 Then
            remainingBits = WorksheetFunction.Min(8, prefix)
            maskByte = &HFF - (2 ^ (8 - remainingBits) - 1)
            .Octet2 = ip.Octet2 And CByte(maskByte)
        Else
            .Octet2 = 0
        End If
        
        ' 3�Ԗڂ̃I�N�e�b�g
        prefix = prefix - remainingBits
        If prefix > 0 Then
            remainingBits = WorksheetFunction.Min(8, prefix)
            maskByte = &HFF - (2 ^ (8 - remainingBits) - 1)
            .Octet3 = ip.Octet3 And CByte(maskByte)
        Else
            .Octet3 = 0
        End If
        
        ' 4�Ԗڂ̃I�N�e�b�g
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

' n�Ԗڂ̃z�X�g�A�h���X���v�Z
Private Function AddToIp(ByRef baseIp As ipBytes, ByVal n As Long) As ipBytes
    Dim carry As Long
    Dim result As Long
    
    With AddToIp
        ' 4�Ԗڂ̃I�N�e�b�g
        result = CLng(baseIp.Octet4) + (n Mod 256)
        carry = result \ 256
        .Octet4 = CByte(result Mod 256)
        
        ' 3�Ԗڂ̃I�N�e�b�g
        n = n \ 256
        result = CLng(baseIp.Octet3) + n Mod 256 + carry
        carry = result \ 256
        .Octet3 = CByte(result Mod 256)
        
        ' 2�Ԗڂ̃I�N�e�b�g
        n = n \ 256
        result = CLng(baseIp.Octet2) + n Mod 256 + carry
        carry = result \ 256
        .Octet2 = CByte(result Mod 256)
        
        ' �ŏ��̃I�N�e�b�g
        n = n \ 256
        result = CLng(baseIp.Octet1) + n + carry
        If result > 255 Then
            Err.Raise 6, "AddToIp", "Result exceeds valid IP range"
        End If
        .Octet1 = CByte(result)
    End With
End Function

' CIDR�\�L�����
Private Sub ParseCIDR(ByVal cidrNotation As String, ByRef networkAddress As String, ByRef prefix As Integer)
    Dim parts() As String
    parts = Split(cidrNotation, "/")
    
    If UBound(parts) <> 1 Then
        Err.Raise 5, "ParseCIDR", "Invalid CIDR notation"
    End If
    
    networkAddress = parts(0)
    prefix = CInt(parts(1))
End Sub

' ���C���֐�
Public Function GetNthIpAddress(ByVal cidrNotation As String, ByVal n As Long) As String
    On Error GoTo ErrorHandler
    
    ' CIDR�\�L�̉��
    Dim networkAddress As String
    Dim prefix As Integer
    ParseCIDR cidrNotation, networkAddress, prefix
    
    ' �v���t�B�b�N�X���̌���
    If prefix < 0 Or prefix > 32 Then
        GetNthIpAddress = "Invalid prefix length"
        Exit Function
    End If
    
    ' IP�A�h���X���o�C�g�ɕ���
    Dim ipBytes As ipBytes
    ipBytes = ParseIpAddress(networkAddress)
    
    ' �l�b�g���[�N�}�X�N�̓K�p
    ipBytes = ApplyNetworkMask(ipBytes, prefix)
    
    ' �z�X�g���̃r�b�g�����v�Z
    Dim hostBits As Integer
    hostBits = 32 - prefix
    
    ' �ő�z�X�g�����v�Z
    Dim maxHosts As Long
    Select Case prefix
        Case 32
            ' /32�̓���P�[�X: �P��̃A�h���X
            maxHosts = 1
        Case 31
            ' /31�̓���P�[�X: 2�̃A�h���X���g�p�\
            maxHosts = 2
        Case Else
            If hostBits > 30 Then
                ' 2^31 - 2 would overflow Long
                maxHosts = &H7FFFFFFF
            Else
                ' �ʏ�̃P�[�X: �l�b�g���[�N�A�h���X�ƃu���[�h�L���X�g�A�h���X�����O
                maxHosts = WorksheetFunction.Power(2, hostBits) - 2
            End If
    End Select
    
    ' �͈̓`�F�b�N
    If n < 1 Or n > maxHosts Then
        GetNthIpAddress = "Invalid host number"
        Exit Function
    End If
    
    ' n�Ԗڂ̃A�h���X���v�Z
    Dim resultIp As ipBytes
    Select Case prefix
        Case 32
            ' /32�̏ꍇ�A�w�肳�ꂽ�A�h���X�����̂܂ܕԂ�
            resultIp = ipBytes
        Case 31
            ' /31�̏ꍇ�An=1�ōŏ��̃A�h���X�An=2��2�Ԗڂ̃A�h���X��Ԃ�
            resultIp = AddToIp(ipBytes, n - 1)
        Case Else
            ' �ʏ�̃P�[�X
            resultIp = AddToIp(ipBytes, n)
    End Select
    
    ' ���ʂ𕶎���ɕϊ�
    GetNthIpAddress = IpBytesToString(resultIp)
    Exit Function
    
ErrorHandler:
    GetNthIpAddress = "Error: " & Err.Description
End Function


' �e�X�g�p�T�u���[�`��
Public Sub TestIpCalculator()
    Dim result As String
    
    ' /32�̃e�X�g
    result = GetNthIpAddress("192.168.1.1/32", 1)
    Debug.Print "192.168.1.1/32��1�Ԗ�: " & result  ' 192.168.1.1
    result = GetNthIpAddress("192.168.1.1/32", 2)
    Debug.Print "192.168.1.1/32��2�Ԗ�: " & result  ' Invalid host number
    
    ' /31�̃e�X�g
    result = GetNthIpAddress("192.168.1.0/31", 1)
    Debug.Print "192.168.1.0/31��1�Ԗ�: " & result  ' 192.168.1.0
    result = GetNthIpAddress("192.168.1.0/31", 2)
    Debug.Print "192.168.1.0/31��2�Ԗ�: " & result  ' 192.168.1.1
    
    ' �ʏ�̃P�[�X
    result = GetNthIpAddress("192.168.1.0/24", 1)
    Debug.Print "192.168.1.0/24��1�Ԗ�: " & result  ' 192.168.1.1
    result = GetNthIpAddress("192.168.1.0/30", 1)
    Debug.Print "192.168.1.0/30��1�Ԗ�: " & result  ' 192.168.1.1
End Sub
