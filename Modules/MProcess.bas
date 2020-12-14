Attribute VB_Name = "MProcess"
Option Explicit

Private Declare Function GetCurrentThread Lib "kernel32" () _
        As Long
        
Private Declare Function GetCurrentProcess Lib "kernel32" () _
        As Long
        
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal _
        hThread As Long, ByVal eTHREAD_PRIORITY As Long) As Long
        
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal _
        hProcess As Long, ByVal dwPriorityClass As ePRIORITY_CLASS) As Long
        
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal _
        hThread As Long) As Long
        
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal _
        hProcess As Long) As Long
        

Const THREAD_BASE_PRIORITY_IDLE = -15
Const THREAD_BASE_PRIORITY_LOWRT = 15
Const THREAD_BASE_PRIORITY_MIN = -2
Const THREAD_BASE_PRIORITY_MAX = 2


Public Enum eTHREAD_PRIORITY
    THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
    THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
    THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
    THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
    THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
    THREAD_PRIORITY_ERROR_RETURN = &H7FFFFFFF
End Enum

Public Enum ePRIORITY_CLASS
    IDLE_PRIORITY_CLASS = &H40
    NORMAL_PRIORITY_CLASS = &H20
    HIGH_PRIORITY_CLASS = &H80
    REALTIME_PRIORITY_CLASS = &H100
End Enum

Private m_lng_OLD_PRIORITY_CLASS As Long
Private m_lng_OLD_THREAD_PRIORITY As Long

Public Sub ReSetProcessPriority()
Dim hThread As Long
Dim hProcess As Long
    hThread = GetCurrentThread
    hProcess = GetCurrentProcess

    Call SetThreadPriority(hThread, m_lng_OLD_THREAD_PRIORITY)
    Call SetPriorityClass(hProcess, m_lng_OLD_PRIORITY_CLASS)

End Sub

Public Sub SetProcessPriority(Optional ByVal new_THREAD_PRIORITY As eTHREAD_PRIORITY = THREAD_PRIORITY_HIGHEST, Optional ByVal new_PRIORITY_CLASS As ePRIORITY_CLASS = HIGH_PRIORITY_CLASS)
Dim hThread As Long
Dim hProcess As Long
    hThread = GetCurrentThread
    hProcess = GetCurrentProcess
        
    m_lng_OLD_THREAD_PRIORITY = GetThreadPriority(hThread)
    m_lng_OLD_PRIORITY_CLASS = GetPriorityClass(hProcess)
    
    Call SetThreadPriority(hThread, new_THREAD_PRIORITY)
    Call SetPriorityClass(hProcess, new_PRIORITY_CLASS)

End Sub


