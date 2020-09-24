Attribute VB_Name = "HexStringAPI"
' CryptoAPI versions
Option Explicit

Private Const CRYPT_STRING_NOCR As Long = &H80000000
Private Const CRYPT_STRING_NOCRLF As Long = &H40000000
Private Const CRYPT_STRING_HEX As Long = 4&

Private Declare Function CryptBinaryToString Lib "Crypt32" Alias "CryptBinaryToStringW" (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
Private Declare Function CryptStringToBinary Lib "Crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long

Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long

Private LH(0 To 5) As Long, LHP As Long
Private LA() As Long, LP As Long

Public Function BytesToHexString_C1(Bytes() As Byte, Optional ByVal Format As Boolean = True, Optional ByVal NewLine As Boolean = True, Optional ByVal Lowercase As Boolean = True) As String
    Dim F As Long, L As Long, LB As Long, P As Long, UB As Long
    P = Not Not Bytes: Debug.Assert App.hInstance
    If P <> 0 Then
        LB = LBound(Bytes)
        UB = UBound(Bytes) + 1
        If UB > LB Then
            F = CRYPT_STRING_HEX Or (CLng(Not Format Or Not NewLine) And (CRYPT_STRING_NOCR Or CRYPT_STRING_NOCRLF))
            If CryptBinaryToString(Bytes(LB), UB, F, 0, L) <> 0 Then
                P = SysAllocStringLen(0, L - 1)
                PutMem4 VarPtr(BytesToHexString_C1), P
                If CryptBinaryToString(Bytes(LB), UB, F, P, L) <> 0 Then
                    If Not Format Then
                        If InStr(BytesToHexString_C1, vbLf) <> 0 Then
                            BytesToHexString_C1 = Replace(Replace(BytesToHexString_C1, vbLf, vbNullString), " ", vbNullString)
                        Else
                            BytesToHexString_C1 = Replace(BytesToHexString_C1, " ", vbNullString)
                        End If
                    End If
                    If Not Lowercase Then BytesToHexString_C1 = UCase$(BytesToHexString_C1)
                Else
                    BytesToHexString_C1 = vbNullString
                End If
            End If
        End If
    End If
End Function

Public Function HexStringToBytes_C1(Hex As String) As Byte()
    Dim L As Long, LO As Long, P As Long, U As Long

    If LHP = 0 Then
        LH(0) = 1: LH(1) = 4: LH(4) = &H7FFFFFFF
        LHP = VarPtr(LH(0))
        LP = ArrPtr(LA)
    End If

    PutMem4 LP, LHP
    
    L = LenB(Hex) \ 4
    LO = L - 6
    If LO < 0 Then LO = 0
    L = LO + 6
    P = SysAllocStringByteLen(0, LO) - 4
    LO = L
    HexStringToBytes_C1 = vbNullString
    LH(3) = Not Not HexStringToBytes_C1: Debug.Assert App.hInstance
    LA(3) = P
    LA(4) = L
    
    U = CLng(CryptStringToBinary(StrPtr(Hex), Len(Hex), CRYPT_STRING_HEX, P, LO, 0, U) <> 0&)
    Select Case LO And U
    Case L
    Case 0
        HexStringToBytes_C1 = vbNullString
    Case Else
        ReDim Preserve HexStringToBytes_C1(0 To LO - 1)
    End Select

    LH(3) = LP: LA(0) = 0
End Function
