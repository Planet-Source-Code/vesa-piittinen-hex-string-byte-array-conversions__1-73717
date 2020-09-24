VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run the code compiled & with optimizations for ""remove integer overflow checks"" & ""remove array bounds checks"""
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   921
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Byte Array to Hex String (other)"
      Height          =   2295
      Left            =   6960
      TabIndex        =   35
      Top             =   2520
      Width           =   6735
      Begin VB.CommandButton cmdOther 
         Caption         =   "BytesToHexString_F1 (silly XML)"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "BytesToHexString_F1 (HTML)"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "BytesToHexString_F1 (none)"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblOther 
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   42
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label lblOther 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   40
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Outputs other kinds of custom formatting."
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label lblOther 
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   37
         Top             =   720
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Formatted Hex String to Byte Array"
      Height          =   2295
      Left            =   120
      TabIndex        =   26
      Top             =   4920
      Width           =   6735
      Begin VB.CommandButton cmdFormatted 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton cmdFormatted 
         Caption         =   "HexStringToBytes_C1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdFormatted 
         Caption         =   "HexStringToBytes_F1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblFormatted 
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   34
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label lblFormatted 
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   32
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label lblFormatted 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   31
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Input: string with formatted hex string, with spaces and line changes."
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Byte Array to Hex String (CryptoAPI formatted)"
      Height          =   2295
      Left            =   6960
      TabIndex        =   19
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdFormat 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton cmdFormat 
         Caption         =   "BytesToHexString_F1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdFormat 
         Caption         =   "BytesToHexString_C1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Output: hex string with CryptoAPI like output."
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label lblFormat 
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   25
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label lblFormat 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   23
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label lblFormat 
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   21
         Top             =   720
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Simple Hex String to Byte Array"
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdSimple 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton cmdSimple 
         Caption         =   "HexStringToBytes_C1"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CommandButton cmdSimple 
         Caption         =   "HexStringToBytes_H1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdSimple 
         Caption         =   "HexStringToBytes_H2"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdSimple 
         Caption         =   "HexStringToBytes_A1"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton cmdSimple 
         Caption         =   "HexStringToBytes_A2"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton cmdSimple 
         Caption         =   "HexStringToBytes_S1"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdSimple 
         Caption         =   "HexStringToBytes_F1"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label lblSimple 
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   18
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label lblSimple 
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   16
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label lblSimple 
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   14
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Input: string with hex data only, no formatting: ""410042004300"" is ok, ""41 00 42 00"" is not."
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label lblSimple 
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   12
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label lblSimple 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   11
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label lblSimple 
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   10
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label lblSimple 
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   9
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label lblSimple 
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   8
         Top             =   2640
         Width           =   3855
      End
   End
   Begin RichTextLib.RichTextBox txtPreview 
      Height          =   2295
      Left            =   6960
      TabIndex        =   0
      Top             =   4920
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Related thread: http://www.vbforums.com/showthread.php?t=639667

' COMPILE! for best results
' File > Make project > Options... > Compile tab > [Advanced optimizations] > [X] Remove Array Bounds Checks & [X] Remove Integer Overflow Checks
' NOTE: this is not recommended for any application that is still in development & testing, you do want to have these overflow checks there
Option Explicit

Private TESTBYTES() As Byte
Private TESTFORMAT As String
Private TESTSIMPLE As String

Private Sub cmdFormat_Click(Index As Integer)
    Dim S As String, T As Double
    Timing = 0
    Select Case Index
        Case 0: S = BytesToHexString_C1(TESTBYTES)
        Case 1: S = BytesToHexString_F1(TESTBYTES)
    End Select
    T = Timing
    lblFormat(Index).Caption = Format$(T * 1000, "0.00000") & " ms (output: " & Len(S) & " chars)"
    txtPreview.Text = S
End Sub

Private Sub cmdFormatted_Click(Index As Integer)
    Dim B() As Byte, T As Double
    Timing = 0
    Select Case Index
        Case 0: B = HexStringToBytes_C1(TESTFORMAT)
        Case 1: B = HexStringToBytes_F1(TESTFORMAT)
        ' just in case of a mistake
        Case Else: B = vbNullString
    End Select
    T = Timing
    lblFormatted(Index).Caption = Format$(T * 1000, "0.00000") & " ms (output: " & (UBound(B) + 1) & " bytes)"
    txtPreview.Text = Replace(StrConv(CStr(B), vbUnicode), vbNullChar, " ")
End Sub

Private Sub cmdOther_Click(Index As Integer)
    Dim S As String, T As Double
    Timing = 0
    Select Case Index
        Case 0: S = BytesToHexString_F1(TESTBYTES, vbNullString, , False)
        Case 1: S = BytesToHexString_F1(TESTBYTES, vbTab & "00000000 00000000 00000000 00000000", "<br>" & vbNewLine, True)
        Case 2
            S = BytesToHexString_F1( _
                TESTBYTES, _
                "<row id=""0x00000000"">" & vbNewLine & _
                vbTab & "<one>00</one>" & vbNewLine & _
                vbTab & "<two>00</two>" & vbNewLine & _
                vbTab & "<three>00</three>" & vbNewLine & _
                vbTab & "<four>00</four>" & vbNewLine & _
                "</row>", _
                vbNewLine, _
                False _
            )
    End Select
    T = Timing
    lblOther(Index).Caption = Format$(T * 1000, "0.00000") & " ms (output: " & Len(S) & " chars)"
    txtPreview.Text = S
End Sub

Private Sub cmdSimple_Click(Index As Integer)
    Dim B() As Byte, T As Double
    Timing = 0
    Select Case Index
        Case 0: B = HexStringToBytes_H1(TESTSIMPLE)
        Case 1: B = HexStringToBytes_H2(TESTSIMPLE)
        Case 2: B = HexStringToBytes_A1(TESTSIMPLE)
        Case 3: B = HexStringToBytes_A2(TESTSIMPLE)
        Case 4: B = HexStringToBytes_S1(TESTSIMPLE)
        Case 5: B = HexStringToBytes_F1(TESTSIMPLE)
        Case 6: B = HexStringToBytes_C1(TESTSIMPLE)
        ' just in case of a mistake
        Case Else: B = vbNullString
    End Select
    T = Timing
    lblSimple(Index).Caption = Format$(T * 1000, "0.00000") & " ms (output: " & (UBound(B) + 1) & " bytes)"
    txtPreview.Text = Replace(StrConv(CStr(B), vbUnicode), vbNullChar, " ")
End Sub

Private Sub Form_Load()
    Dim I As Long
    
    ReDim TESTBYTES(0 To 65535)
    
    For I = 0 To UBound(TESTBYTES)
        TESTBYTES(I) = CByte(I And &HFF&)
    Next I

    TESTSIMPLE = BytesToHexString_C1(TESTBYTES, False, False, False)
    TESTFORMAT = BytesToHexString_C1(TESTBYTES, True, True, True)
    
    Debug.Print TESTFORMAT
End Sub
