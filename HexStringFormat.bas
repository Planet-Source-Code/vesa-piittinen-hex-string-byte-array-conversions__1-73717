Attribute VB_Name = "HexStringFormat"
' all these functions work with formatted hex strings
' this means strings that contain spaces or line changes are accepted
Option Explicit

Private Const CRYPT_HEX_FORMAT = "00 00 00 00 00 00 00 00  00 00 00 00 00 00 00 00"

Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long

Private LH(0 To 5) As Long, LHP As Long
Private LA() As Long, LP As Long

Private IH(0 To 5) As Long, IHP As Long
Private IA() As Integer, IP As Long

Private LH_(0 To 5) As Long, LHP_ As Long
Private LA_() As Long, LP_ As Long

Private BHex(0 To 511) As Long, BHexI As Boolean

' a very advanced function: allows for many kinds of formatting options and it is very fast too
Public Function BytesToHexString_F1(Bytes() As Byte, Optional Format As String = CRYPT_HEX_FORMAT, Optional Separator As String = vbNewLine, Optional ByVal Lowercase As Boolean = True) As String
    Dim BytesBase As Long, BytesPtr As Long, StringPtr As Long
    Dim C As Long, CH As Long, CL As Long, CS As Long, F() As Long, I As Long, J As Long, L As Long, LF As Long, LS As Long, P As Long
    ' get pointer to safe array header
    BytesPtr = Not Not Bytes: Debug.Assert App.hInstance
    ' valid array
    If BytesPtr <> 0 Then
        ' calculate size
        L = UBound(Bytes) - LBound(Bytes) + 1
        ' valid size
        If L > 0 Then
            ' safe arrays prepared for use?
            If IHP = 0 Then
                ' safe array: Long
                LH(0) = 1: LH(1) = 4: LH(4) = &H3FFFFFFF
                LHP = VarPtr(LH(0))
                LP = ArrPtr(LA)
                ' safe array: Integer
                IH(0) = 1: IH(1) = 2
                IHP = VarPtr(IH(0))
                IP = ArrPtr(IA)
            End If
            ' hex array prepared for use?
            If BHexI = False Then
                For I = 0 To 255
                    ' upper case
                    CH = ((I And &HF0&) \ &H10&) Or &H30&
                    If CH > 57& Then CH = CH + 7&
                    CL = (I And &HF&) Or &H30&
                    If CL > 57& Then CL = CL + 7&
                    BHex(I) = CH Or (CL * &H10000)
                    ' lower case
                    If CH > 64& Then CH = CH Or &H20&
                    If CL > 64& Then CL = CL Or &H20&
                    BHex(I Or 256&) = CH Or (CL * &H10000)
                Next I
                BHexI = True
            End If
            
            ' safe array: Long (first to get better speed)
            PutMem4 LP, LHP
            
            ' fast mode or format mode?
            If InStr(Format, "00") = 0 Then
                ' prepare buffer
                StringPtr = SysAllocStringByteLen(0, L * 4&)
                LH(3) = VarPtr(BytesToHexString_F1): LA(0) = StringPtr
                ' modify byte array to zero base
                If LBound(Bytes) <> 0 Then
                    LH(3) = BytesPtr
                    BytesBase = LA(5)
                    LA(5) = 0
                End If
                ' point long array to string buffer
                LH(3) = StringPtr
                ' convert the bytes
                If Lowercase = False Then
                    For I = 0 To UBound(Bytes): LA(I) = BHex(Bytes(I)): Next
                Else
                    For I = 0 To UBound(Bytes): LA(I) = BHex(Bytes(I) Or 256&): Next
                End If
                ' restore byte array to non-zero base
                If BytesBase <> 0 Then LH(3) = BytesPtr: LA(5) = BytesBase
            Else
                LF = Len(Format)
                LS = Len(Separator)
                ' find out how many bytes we output per line
                ReDim F(0 To LF \ 2 - 1)
                I = 0
                Do
                    Do: I = InStrB(I + 1, Format, "00")
                    Loop Until (I = 0&) Or (I And 1&) = 1&
                    If I <> 0& Then
                        F(C) = (I - 1&)
                        C = C + 1&
                        I = I + 3
                    End If
                Loop Until I = 0&
                ReDim Preserve F(C - 1)
                ' calculate separator & amount of characters after last line
                CL = L - 1
                If LS <> 0& Then CS = LS * (CL \ C)
                If (L Mod C) <> 0& Then CS = CS + F(CL Mod C) \ 2& + 2&
                ' prepare buffer
                StringPtr = SysAllocStringLen(0, LF * (L \ C) + CS)
                LH(3) = VarPtr(BytesToHexString_F1): LA(0) = StringPtr
                ' replicate
                Mid$(BytesToHexString_F1, 1, LF) = Format
                If Len(BytesToHexString_F1) > LF Then
                    Mid$(BytesToHexString_F1, 1 + LF, LS) = Separator
                    Mid$(BytesToHexString_F1, 1 + LF + LS) = BytesToHexString_F1
                End If
                ' modify byte array to zero base
                If LBound(Bytes) <> 0 Then
                    LH(3) = BytesPtr
                    BytesBase = LA(5)
                    LA(5) = 0
                End If
                P = StringPtr
                LS = (LF + LS) * 2&
                ' convert the bytes
                If Lowercase = False Then
                    For I = 0 To UBound(Bytes) - C + 1 Step C
                        For J = 0 To C - 1
                            LH(3) = P + F(J)
                            LA(0) = BHex(Bytes(I + J))
                        Next J
                        P = P + LS
                    Next I
                    If (L Mod C) <> 0& Then
                        For J = 0 To CL Mod C
                            LH(3) = P + F(J)
                            LA(0) = BHex(Bytes(I + J))
                        Next J
                    End If
                Else
                    For I = 0 To UBound(Bytes) - C + 1 Step C
                        For J = 0 To C - 1
                            LH(3) = P + F(J)
                            LA(0) = BHex(Bytes(I + J) Or 256&)
                        Next J
                        P = P + LS
                    Next I
                    If (L Mod C) <> 0& Then
                        For J = 0 To CL Mod C
                            LH(3) = P + F(J)
                            LA(0) = BHex(Bytes(I + J) Or 256&)
                        Next J
                    End If
                End If
                ' restore byte array to non-zero base
                If BytesBase <> 0 Then LH(3) = BytesPtr: LA(5) = BytesBase
            End If
            
            ' safe array: Long
            LH(3) = LP: LA(0) = 0
        End If
    End If
End Function

' a very fast function that also allows for any kind of hex string input
' supports upper & lowercase and any non-hex character pair or lone character is simply ignored
Public Function HexStringToBytes_F1(Hex As String) As Byte()

    Dim B() As Byte, C As Long, CH As Long, CL As Long, H As Long, I As Long, J As Long, L As Long, LB As Long
    
    L = Len(Hex)
    If L > 1 Then
        ' safe arrays prepared for use?
        If IHP = 0 Then
            ' safe array: Long
            LH(0) = 1: LH(1) = 4: LH(4) = &H3FFFFFFF
            LHP = VarPtr(LH(0))
            LP = ArrPtr(LA)
            ' safe array: Integer
            IH(0) = 1: IH(1) = 2
            IHP = VarPtr(IH(0))
            IP = ArrPtr(IA)
        End If
        
        ' safe array: Long (first to get better speed)
        PutMem4 LP, LHP
        ' safe array: Integer
        LH(3) = IP: LA(0) = IHP
        
        ' prepare output byte array
        HexStringToBytes_F1 = vbNullString
        ' get pointer to safe array header for manipulation
        LH(3) = Not Not HexStringToBytes_F1: Debug.Assert App.hInstance
        ' calculate size of BSTR allocation
        LB = (L \ 2) - 6: If LB < 0 Then LB = 0
        ' create a BSTR, it works as our byte array!
        LA(3) = SysAllocStringByteLen(0, LB) - 4: LA(4) = LB + 6
        
        ' access string via Integer array
        IH(3) = StrPtr(Hex): IH(4) = L
        ' set long array to output data (= byte array)
        LH(3) = LA(3)
        
        Do
            ' byte 1
            Do While I + 1 < L
                CH = IA(I)
                Select Case CH
                Case 48 To 57
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = ((CH And Not 48&) * &H10&) Or (CL And Not 48&)
                        C = 1
                        Exit Do
                    Case 65 To 70
                        H = ((CH And Not 48&) * &H10&) Or (CL - 55&)
                        C = 1
                        Exit Do
                    Case 97 To 102
                        H = ((CH And Not 48&) * &H10&) Or (CL - 87&)
                        C = 1
                        Exit Do
                    End Select
                Case 65 To 70
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = ((CH - 55&) * &H10&) Or (CL And Not 48&)
                        C = 1
                        Exit Do
                    Case 65 To 70
                        H = ((CH - 55&) * &H10&) Or (CL - 55&)
                        C = 1
                        Exit Do
                    Case 97 To 102
                        H = ((CH - 55&) * &H10&) Or (CL - 87&)
                        C = 1
                        Exit Do
                    End Select
                Case 97 To 102
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = ((CH - 87&) * &H10&) Or (CL And Not 48&)
                        C = 1
                        Exit Do
                    Case 65 To 70
                        H = ((CH - 87&) * &H10&) Or (CL - 55&)
                        C = 1
                        Exit Do
                    Case 97 To 102
                        H = ((CH - 87&) * &H10&) Or (CL - 87&)
                        C = 1
                        Exit Do
                    End Select
                End Select
                I = I + 1
            Loop
            ' done?
            If I + 2 < L Then I = I + 1 Else Exit Do
            ' byte 2
            Do While I + 1 < L
                CH = IA(I)
                Select Case CH
                Case 48 To 57
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH And Not 48&) * &H1000&) Or ((CL And Not 48&) * &H100&)
                        C = 2
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH And Not 48&) * &H1000&) Or ((CL - 55&) * &H100&)
                        C = 2
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH And Not 48&) * &H1000&) Or ((CL - 87&) * &H100&)
                        C = 2
                        Exit Do
                    End Select
                Case 65 To 70
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH - 55&) * &H1000&) Or ((CL And Not 48&) * &H100&)
                        C = 2
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH - 55&) * &H1000&) Or ((CL - 55&) * &H100&)
                        C = 2
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH - 55&) * &H1000&) Or ((CL - 87&) * &H100&)
                        C = 2
                        Exit Do
                    End Select
                Case 97 To 102
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH - 87&) * &H1000&) Or ((CL And Not 48&) * &H100&)
                        C = 2
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH - 87&) * &H1000&) Or ((CL - 55&) * &H100&)
                        C = 2
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH - 87&) * &H1000&) Or ((CL - 87&) * &H100&)
                        C = 2
                        Exit Do
                    End Select
                End Select
                I = I + 1
            Loop
            ' done?
            If I + 2 < L Then I = I + 1 Else Exit Do
            ' byte 3
            Do While I + 1 < L
                CH = IA(I)
                Select Case CH
                Case 48 To 57
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH And Not 48&) * &H100000) Or ((CL And Not 48&) * &H10000)
                        C = 3
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH And Not 48&) * &H100000) Or ((CL - 55&) * &H10000)
                        C = 3
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH And Not 48&) * &H100000) Or ((CL - 87&) * &H10000)
                        C = 3
                        Exit Do
                    End Select
                Case 65 To 70
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH - 55&) * &H100000) Or ((CL And Not 48&) * &H10000)
                        C = 3
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH - 55&) * &H100000) Or ((CL - 55&) * &H10000)
                        C = 3
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH - 55&) * &H100000) Or ((CL - 87&) * &H10000)
                        C = 3
                        Exit Do
                    End Select
                Case 97 To 102
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH - 87&) * &H100000) Or ((CL And Not 48&) * &H10000)
                        C = 3
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH - 87&) * &H100000) Or ((CL - 55&) * &H10000)
                        C = 3
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH - 87&) * &H100000) Or ((CL - 87&) * &H10000)
                        C = 3
                        Exit Do
                    End Select
                End Select
                I = I + 1
            Loop
            ' done?
            If I + 2 < L Then I = I + 1 Else Exit Do
            ' byte 4
            Do While I + 1 < L
                CH = IA(I)
                Select Case CH
                Case 48 To 55
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH And Not 48&) * &H10000000) Or ((CL And Not 48&) * &H1000000)
                        C = 0
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH And Not 48&) * &H10000000) Or ((CL - 55&) * &H1000000)
                        C = 0
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH And Not 48&) * &H10000000) Or ((CL - 87&) * &H1000000)
                        C = 0
                        Exit Do
                    End Select
                Case 56 To 57
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH And Not 56&) * &H10000000) Or ((CL And Not 48&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH And Not 56&) * &H10000000) Or ((CL - 55&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH And Not 56&) * &H10000000) Or ((CL - 87&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    End Select
                Case 65 To 70
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH - 63&) * &H10000000) Or ((CL And Not 48&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH - 63&) * &H10000000) Or ((CL - 55&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH - 63&) * &H10000000) Or ((CL - 87&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    End Select
                Case 97 To 102
                    I = I + 1
                    CL = IA(I)
                    Select Case CL
                    Case 48 To 57
                        H = H Or ((CH - 95&) * &H10000000) Or ((CL And Not 48&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    Case 65 To 70
                        H = H Or ((CH - 95&) * &H10000000) Or ((CL - 55&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    Case 97 To 102
                        H = H Or ((CH - 95&) * &H10000000) Or ((CL - 87&) * &H1000000) Or &H80000000
                        C = 0
                        Exit Do
                    End Select
                End Select
                I = I + 1
            Loop
            ' write
            If C = 0 Then LA(J) = H: J = J + 1
            ' done?
            If I + 2 < L Then I = I + 1 Else Exit Do
        Loop
        
        ' check for unwritten bytes & avoid buffer overwrite
        Select Case C
            Case 0
            Case 1: LA(J) = (LA(J) And &HFFFFFF00) Or H
            Case 2: LA(J) = (LA(J) And &HFFFF0000) Or H
            Case 3: LA(J) = (LA(J) And &HFF000000) Or H
        End Select
        
        ' calculate final length
        L = J * 4 + C
        Select Case L
            Case LB + 6 ' do nothing!
            Case 0: HexStringToBytes_F1 = vbNullString
            Case Else
                LH(3) = ArrPtr(B)
                LA(0) = Not Not HexStringToBytes_F1: Debug.Assert App.hInstance
                ReDim Preserve B(0 To L - 1)
                LA(0) = 0
        End Select
                
        ' safe array: Integer
        LH(3) = IP: LA(0) = 0
        ' safe array: Long
        LH(3) = LP: LA(0) = 0
    Else
        ' empty array
        HexStringToBytes_F1 = vbNullString
    End If
End Function

Public Function z_broken_HexStringToBytes_F2(Hex As String) As Byte()

    Dim B() As Byte, C As Long, CH As Long, CL As Long, H As Long, I As Long, J As Long, L As Long, LB As Long
    
    L = Len(Hex)
    If L > 1 Then
        ' safe arrays prepared for use?
        If LHP_ = 0 Then
            ' safe array: Long
            LH(0) = 1: LH(1) = 4: LH(4) = &H3FFFFFFF
            LHP = VarPtr(LH(0))
            LP = ArrPtr(LA)
            ' safe array: Long 2
            LH_(0) = 1: LH_(1) = 4: LH_(4) = &H3FFFFFFF
            LHP_ = VarPtr(LH_(0))
            LP_ = ArrPtr(LA_)
        End If
        
        ' safe array: Long (first to get better speed)
        PutMem4 LP, LHP
        ' safe array: Long 2
        LH(3) = LP_: LA(0) = LHP_
        
        ' prepare output byte array
        z_broken_HexStringToBytes_F2 = vbNullString
        ' get pointer to safe array header for manipulation
        LH(3) = Not Not z_broken_HexStringToBytes_F2: Debug.Assert App.hInstance
        ' calculate size of BSTR allocation
        LB = (L \ 2) - 6: If LB < 0 Then LB = 0
        ' create a BSTR, it works as our byte array!
        LA(3) = SysAllocStringByteLen(0, LB) - 4: LA(4) = LB + 6
        
        ' access string via Long array 2
        LH_(3) = StrPtr(Hex)
        ' set long array to output data (= byte array)
        LH(3) = LA(3)

        L = L \ 2

        Do
            ' byte 1
            For I = I To L - 1
                CH = LA_(I)
                If (CH And &HFF80FF80) = 0& Then
                    CL = CH And &H7F&
                    CH = (CH And &H7F0000) \ &H10000
                    If CL > 47& And CL < 58& Then
                        If CH > 47& And CH < 58& Then
                            H = ((CH And Not 48&) * &H10&) Or (CL And Not 48&)
                            Debug.Print I, VBA.Hex$(H)
                            C = 1
                            Exit For
                        ElseIf CH > 64 And CH < 71 Then
                            H = ((CH - 55&) * &H10&) Or (CL)
                            C = 1
                            Exit For
                        End If
                    ElseIf CL > 64 And CL < 71 Then
                        If CH < 10 Then
                            H = (CH * &H10&) Or (CL - 55&)
                            C = 1
                            Exit For
                        ElseIf CH > 64 And CH < 71 Then
                            H = ((CH - 55&) * &H10&) Or (CL - 55&)
                            C = 1
                            Exit For
                        End If
                    End If
                End If
            Next I
            If I + 1 < L Then I = I + 1 Else Exit Do
            ' byte 2
            For I = I To L - 1
                CH = LA_(I)
                If (CH And &HFF80FF80) = 0& Then
                    CL = CH And &H4F&
                    CH = (CH And &H4F0000) \ &H10000
                    If CL < 10 Then
                        If CH < 10 Then
                            H = H Or (CH * &H1000&) Or (CL * &H100&)
                            C = 2
                            Exit For
                        ElseIf CH > 64 And CH < 71 Then
                            H = H Or ((CH - 55&) * &H1000&) Or (CL * &H100&)
                            C = 2
                            Exit For
                        End If
                    ElseIf CL > 64 And CL < 71 Then
                        If CH < 10 Then
                            H = H Or (CH * &H1000&) Or ((CL - 55&) * &H100&)
                            C = 2
                            Exit For
                        ElseIf CH > 64 And CH < 71 Then
                            H = H Or ((CH - 55&) * &H1000&) Or ((CL - 55&) * &H100&)
                            C = 2
                            Exit For
                        End If
                    End If
                End If
            Next I
            If I + 1 < L Then I = I + 1 Else Exit Do
            ' byte 3
            For I = I To L - 1
                CH = LA_(I)
                If (CH And &HFF80FF80) = 0& Then
                    CL = CH And &H4F&
                    CH = (CH And &H4F0000) \ &H10000
                    If CL < 10 Then
                        If CH < 10 Then
                            H = H Or (CH * &H100000) Or (CL * &H10000)
                            C = 3
                            Exit For
                        ElseIf CH > 64 And CH < 71 Then
                            H = H Or ((CH - 55&) * &H100000) Or (CL * &H10000)
                            C = 3
                            Exit For
                        End If
                    ElseIf CL > 64 And CL < 71 Then
                        If CH < 10 Then
                            H = H Or (CH * &H100000) Or ((CL - 55&) * &H10000)
                            C = 3
                            Exit For
                        ElseIf CH > 64 And CH < 71 Then
                            H = H Or ((CH - 55&) * &H100000) Or ((CL - 55&) * &H10000)
                            C = 3
                            Exit For
                        End If
                    End If
                End If
            Next I
            If I + 1 < L Then I = I + 1 Else Exit Do
            ' byte 4
            For I = I To L - 1
                CH = LA_(I)
                If (CH And &HFF80FF80) = 0& Then
                    CL = CH And &H4F&
                    CH = (CH And &H4F0000) \ &H10000
                    If CL < 10 Then
                        If CH < 8 Then
                            H = H Or (CH * &H10000000) Or (CL * &H1000000)
                            C = 0
                            Exit For
                        ElseIf CH < 10 Then
                            H = H Or ((CH And &H7&) * &H10000000) Or (CL * &H1000000) Or &H80000000
                            C = 0
                            Exit For
                        ElseIf CH > 64 And CH < 71 Then
                            H = H Or ((CH - 63&) * &H10000000) Or (CL * &H1000000) Or &H80000000
                            C = 0
                            Exit For
                        End If
                    ElseIf CL > 64 And CL < 71 Then
                        If CH < 8 Then
                            H = H Or (CH * &H10000000) Or ((CL - 55&) * &H1000000)
                            C = 0
                            Exit For
                        ElseIf CH < 10 Then
                            H = H Or ((CH And &H7&) * &H10000000) Or ((CL - 55&) * &H1000000) Or &H80000000
                            C = 0
                            Exit For
                        ElseIf CH > 64 And CH < 71 Then
                            H = H Or ((CH - 63&) * &H10000000) Or ((CL - 55&) * &H1000000) Or &H80000000
                            C = 0
                            Exit For
                        End If
                    End If
                End If
            Next I
            ' write
            If C = 0 Then LA(J) = H: J = J + 1
            ' done?
            If I + 1 < L Then I = I + 1 Else Exit Do
        Loop

        ' check for unwritten bytes & avoid buffer overwrite
        Select Case C
            Case 0
            Case 1: LA(J) = (LA(J) And &HFFFFFF00) Or H
            Case 2: LA(J) = (LA(J) And &HFFFF0000) Or H
            Case 3: LA(J) = (LA(J) And &HFF000000) Or H
        End Select
        
        ' calculate final length
        L = J * 4 + C
        Select Case L
            Case LB + 6 ' do nothing!
            Case 0: z_broken_HexStringToBytes_F2 = vbNullString
            Case Else
                LH(3) = ArrPtr(B)
                LA(0) = Not Not z_broken_HexStringToBytes_F2: Debug.Assert App.hInstance
                ReDim Preserve B(0 To L - 1)
                LA(0) = 0
        End Select

        ' safe array: Long 2
        LH(3) = LP_: LA(0) = 0
        ' safe array: Long
        LH(3) = LP: LA(0) = 0
    Else
        ' empty array
        z_broken_HexStringToBytes_F2 = vbNullString
    End If
End Function
