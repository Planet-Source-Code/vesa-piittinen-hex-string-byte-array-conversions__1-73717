Attribute VB_Name = "HexStringSimple"
' all these functions work with simple hex strings only
' this means the strings must not contain any formatting: no spaces, no linechanges or anything else
Option Explicit

' all declarations here are only used by S1 version of the function!
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long

Private LH(0 To 5) As Long, LHP As Long
Private LA() As Long, LP As Long

Private IH(0 To 5) As Long, IHP As Long
Private IA() As Integer, IP As Long

' creates a "&H##" string by concatenation and then coerces the result to Byte
' result: this function is slow
Public Function HexStringToBytes_H1(Hex As String) As Byte()
    Dim B() As Byte, I As Long
    ReDim B(Len(Hex) \ 2 - 1)
    For I = 1 To Len(Hex) - 1 Step 2
        B((I - 1) \ 2) = "&H" & Mid$(Hex, I, 2)
    Next I
    HexStringToBytes_H1 = B
End Function

' creates a "&H####" string by using a new copy of Hex as a buffer (passed ByVal)
' result: this function is slow, but roughly needs only half the processing time than above
Public Function HexStringToBytes_H2(ByVal Hex As String) As Byte()
    Dim C As Long, I As Long
    ' the first two bytes have to be dealt using string concatenation
    C = "&H" & Left$(Hex, 4)
    ' beautiful math here, must be done to swap the bytes to correct order
    Mid$(Hex, 1, 1) = ChrW$(((C And &HFF&) * &H100&) Or ((C And &HFF00&) \ &H100&))
    For I = 3 To Len(Hex) - 5 Step 4
        ' now, instead of string concatenation we place &H into Hex string
        ' we save time because we do not create a new string
        Mid$(Hex, I, 2) = "&H"
        ' then coerce the result into a Long (here Mid$ creates a new string btw...)
        C = Mid$(Hex, I, 6)
        ' beautiful math here, must be done to swap the bytes to correct order
        Mid$(Hex, ((I + 1) \ 4) + 1, 1) = ChrW$(((C And &HFF&) * &H100&) Or ((C And &HFF00&) \ &H100&))
    Next I
    ' and then the result
    HexStringToBytes_H2 = LeftB$(Hex, Len(Hex) \ 2)
End Function

' original version: http://www.vbforums.com/showthread.php?p=3289346#post3289346
' gets character code by using AscW and then converts the characters to 4-bit parts of a byte
' result: this function is slow
Public Function HexStringToBytes_A1(Hex As String) As Byte()
    Dim B() As Byte, BH As Long, BL As Long, I As Long
    If LenB(Hex) Then
        ' reserve memory for output buffer
        ReDim B(Len(Hex) \ 2 - 1)
        ' jump by every two characters (in this case we happen to use byte positions for greater speed)
        For I = 1 To LenB(Hex) - 3 Step 4
            ' get the character value and decrease by 48
            ' note: each MidB$ creates a new string
            BH = AscW(MidB$(Hex, I, 2)) - 48&
            BL = AscW(MidB$(Hex, I + 2, 2)) - 48&
            ' move old A - F values down even more
            If BH > 9 Then BH = BH - 7
            If BL > 9 Then BL = BL - 7
            ' combine the two 4 bit parts into a single byte
            B(I \ 4) = ((BH And &HF&) * &H10&) Or (BL And &HF&)
        Next I
        ' return the output
        HexStringToBytes_A1 = B
    End If
End Function

' the same as above, but uses the passed Hex string as a buffer (passed ByVal)
' result: this function is slow
Public Function HexStringToBytes_A2(ByVal Hex As String) As Byte()
    Dim BH As Long, BL As Long, I As Long
    If LenB(Hex) Then
        ' jump by every two characters (in this case we happen to use byte positions for greater speed)
        For I = 1 To LenB(Hex) - 3 Step 4
            ' get the character value and decrease by 48
            ' note: each MidB$ creates a new string
            BH = AscW(MidB$(Hex, I, 2)) - 48&
            BL = AscW(MidB$(Hex, I + 2, 2)) - 48&
            ' move old A - F values down even more
            If BH > 9 Then BH = BH - 7
            If BL > 9 Then BL = BL - 7
            ' combine the two 4 bit parts into a single byte
            MidB$(Hex, ((I - 1) \ 4) + 1, 1) = ChrW$(((BH And &HF&) * &H10&) Or (BL And &HF&))
        Next I
        ' return the output
        HexStringToBytes_A2 = LeftB$(Hex, (I - 1) \ 4)
    End If
End Function

' 2011-01-31 "the quite insane version really"
' this function is long, but it does a few tricks to avoid any kind of speed bottlenecks
' result: the current fastest VB6 implementation known
Public Function HexStringToBytes_S1(Hex As String) As Byte()
    Dim C As Long, H As Long, L As Long, LB As Long
    
    ' ignore half byte information
    L = Len(Hex) And Not 1
    ' check length
    If L >= 12 Then
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
        ' safe array: Long
        PutMem4 LP, LHP
        ' safe array: Integer
        LH(3) = IP: LA(0) = IHP
        ' create an empty byte array
        HexStringToBytes_S1 = vbNullString
        ' length of byte array
        LB = (L \ 2)
        ' get pointer to safe array header for manipulation
        LH(3) = Not Not HexStringToBytes_S1: Debug.Assert App.hInstance
        ' create a BSTR, it works as our byte array!
        LA(3) = SysAllocStringByteLen(0, LB - 6) - 4: LA(4) = LB
        
        IH(3) = StrPtr(Hex): IH(4) = L
        ' set long array to output data (= byte array)
        LH(3) = LA(3)
        ' go through 8 hex string characters at a time = 4 bytes = 32-bits
        For L = 0 To UBound(IA) - 7 Step 8
            ' byte 1
            C = IA(L + 1)
            Select Case C
                Case 48 To 57: H = C And Not 48&
                Case 65 To 70: H = C - 55&
                Case Else: H = 0
            End Select
            C = IA(L)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H10&)
                Case 65 To 70: H = H Or ((C - 55&) * &H10&)
            End Select
            ' byte 2
            C = IA(L + 3)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H100&)
                Case 65 To 70: H = H Or ((C - 55&) * &H100&)
            End Select
            C = IA(L + 2)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H1000&)
                Case 65 To 70: H = H Or ((C - 55&) * &H1000&)
            End Select
            ' byte 3
            C = IA(L + 5)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H10000)
                Case 65 To 70: H = H Or ((C - 55&) * &H10000)
            End Select
            C = IA(L + 4)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H100000)
                Case 65 To 70: H = H Or ((C - 55&) * &H100000)
            End Select
            ' byte 4
            C = IA(L + 7)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H1000000)
                Case 65 To 70: H = H Or ((C - 55&) * &H1000000)
            End Select
            C = IA(L + 6)
            Select Case C
                Case 48 To 55: H = H Or ((C And Not 48&) * &H10000000)
                Case 56 To 57: H = H Or ((C And Not 56&) * &H10000000) Or &H80000000
                Case 65 To 70: H = H Or ((C - 63&) * &H10000000) Or &H80000000
            End Select
            ' write
            LA(L \ 8) = H
        Next L
        
        ' memory safety
        Select Case (UBound(IA) + 1) - L
        Case 0 ' we are done!
        Case 2
            ' read
            H = LA(L \ 8) And &HFFFFFF00
            ' byte 1
            C = IA(L + 1)
            Select Case C
                Case 48 To 57: H = H Or (C And Not 48&)
                Case 65 To 70: H = H Or (C - 55&)
            End Select
            C = IA(L)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H10&)
                Case 65 To 70: H = H Or ((C - 55&) * &H10&)
            End Select
            ' write
            LA(L \ 8) = H
        Case 4
            ' read
            H = LA(L \ 8) And &HFFFF0000
            ' byte 1
            C = IA(L + 1)
            Select Case C
                Case 48 To 57: H = H Or (C And Not 48&)
                Case 65 To 70: H = H Or (C - 55&)
            End Select
            C = IA(L)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H10&)
                Case 65 To 70: H = H Or ((C - 55&) * &H10&)
            End Select
            ' byte 2
            C = IA(L + 3)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H100&)
                Case 65 To 70: H = H Or ((C - 55&) * &H100&)
            End Select
            C = IA(L + 2)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H1000&)
                Case 65 To 70: H = H Or ((C - 55&) * &H1000&)
            End Select
            ' write
            LA(L \ 8) = H
        Case 6
            ' read
            H = LA(L \ 8) And &HFF000000
            ' byte 1
            C = IA(L + 1)
            Select Case C
                Case 48 To 57: H = H Or (C And Not 48&)
                Case 65 To 70: H = H Or (C - 55&)
            End Select
            C = IA(L)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H10&)
                Case 65 To 70: H = H Or ((C - 55&) * &H10&)
            End Select
            ' byte 2
            C = IA(L + 3)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H100&)
                Case 65 To 70: H = H Or ((C - 55&) * &H100&)
            End Select
            C = IA(L + 2)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H1000&)
                Case 65 To 70: H = H Or ((C - 55&) * &H1000&)
            End Select
            ' byte 3
            C = IA(L + 5)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H10000)
                Case 65 To 70: H = H Or ((C - 55&) * &H10000)
            End Select
            C = IA(L + 4)
            Select Case C
                Case 48 To 57: H = H Or ((C And Not 48&) * &H100000)
                Case 65 To 70: H = H Or ((C - 55&) * &H100000)
            End Select
            ' write
            LA(L \ 8) = H
        End Select
        ' end safearrays
        LH(3) = IP: LA(0) = 0
        LH(3) = LP: LA(0) = 0
    ElseIf L > 0 Then
        Dim B() As Byte, BL As Byte, BH As Byte
        B = LeftB$(Hex, L \ 2)
        For L = 0 To UBound(B)
            BH = AscB(Mid$(Hex, L + L + 1, 1)) And Not 48
            BL = AscB(Mid$(Hex, L + L + 2, 1)) And Not 48
            If BH < 10 Then BH = BH * 16 Else BH = ((BH - 7) And 15) * 16
            If BL < 10 Then B(L) = BL Or BH Else B(L) = ((BL - 7) And 15) Or BH
        Next L
        HexStringToBytes_S1 = B
    End If
End Function

