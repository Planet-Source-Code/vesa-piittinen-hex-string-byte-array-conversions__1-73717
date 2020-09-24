Attribute VB_Name = "Timings"
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean

Dim m_Time As Double
Dim m_TimeFreq As Double
Dim m_TimeStart As Currency

Public Property Get Timing() As Double
    Dim curTime As Currency
    QueryPerformanceCounter curTime
    Timing = (curTime - m_TimeStart) * m_TimeFreq + m_Time
End Property

Public Property Let Timing(ByVal NewValue As Double)
    Dim curFreq As Currency, curOverhead As Currency
    m_Time = NewValue
    QueryPerformanceFrequency curFreq
    m_TimeFreq = 1 / curFreq
    QueryPerformanceCounter curOverhead
    QueryPerformanceCounter m_TimeStart
    m_TimeStart = m_TimeStart + (m_TimeStart - curOverhead)
End Property
