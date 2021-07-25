VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum SnapShotFlags
    SnapHeapLists = &H1
    SnapProcesses = &H2
    SnapThreads = &H4
    SnapModules = &H8
    SnapAll = &HF
End Enum

Private m_SnapHandle As Long
Private m_Flags      As Long
Private m_pHeapList  As Long
Private m_pProcess   As Long
Private m_pThreads   As Long
Private m_pModule    As Long

Public Sub NewC(Flags As SnapShotFlags)
    m_Flags = Flags
End Sub
Public Function NextProcess() As Process
    Dim hr As Long
    Set NextProcess = New Process
    If m_pProcess = 0 Then
        m_pProcess = NextProcess.Ptr
        'hr = Process32First(m_SnapHandle, m_pProcess)
    Else
        m_pProcess = NextProcess.Ptr
        'hr = Process32First(m_SnapHandle, m_pProcess)
        If hr = 0 Then
        End If
    End If
End Function
Public Function GetProcessByID(ProcID As Long) As Process

End Function
Public Function GetProcessByName(procname As String) As Process
    m_SnapHandle = CreateToolhelp32Snapshot(m_Flags, 0&)
    
    Call CloseHandle(m_SnapHandle)
End Function



