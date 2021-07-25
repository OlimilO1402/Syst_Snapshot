VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISnapShotEntry
'Public Type THREADENTRY32
'    dwSize             As Long 'The size of the structure, in bytes. Before calling the Thread32First function, set this member to sizeof(THREADENTRY32). If you do not initialize dwSize, Thread32First fails.
'    cntUsage           As Long 'This member is no longer used and is always set to zero
'    th32ThreadID       As Long 'The thread identifier, compatible with the thread identifier returned by the CreateProcess function
'    th32OwnerProcessID As Long 'The identifier of the process that created the thread.
'    tpBasePri          As Long 'The kernel base priority level assigned to the thread. The priority is a number from 0 to 31, with 0 representing the lowest possible thread priority. For more information, see KeQueryPriorityThread.
'    tpDeltaPri         As Long 'This member is no longer used and is always set to zero.
'    dwFlags            As Long 'This member is no longer used and is always set to zero.
'End Type
Private m_ThreadEntry As THREADENTRY32
Private m_Snapshot    As SnapShot

Public Sub NewC(aSnapShot As SnapShot)
    Set m_Snapshot = aSnapShot
End Sub

Public Property Get Ptr() As Long
    Ptr = ISnapShotEntry_Ptr
End Property
Public Property Get Key() As String
    Key = ISnapShotEntry_Key
End Property

Public Property Get ThreadID() As Long
    ThreadID = m_ThreadEntry.th32ThreadID
End Property
Public Property Get OwnerProcessID() As Long
    OwnerProcessID = m_ThreadEntry.th32OwnerProcessID
End Property
Public Property Get BasePriority() As Long
    BasePriority = m_ThreadEntry.th32OwnerProcessID
End Property

Public Function ToString() As String
    Dim s As String
    With m_ThreadEntry
        s = s & "ThreadID:       " & .th32ThreadID & vbCrLf
        s = s & "OwnerProcessID: " & .th32OwnerProcessID & vbCrLf
        s = s & "Base priority:  " & .tpBasePri & vbCrLf
    End With
    ToString = s
End Function

Private Function ISnapShotEntry_ID() As Long
    ISnapShotEntry_ID = m_ThreadEntry.th32ThreadID
End Function
Private Function ISnapShotEntry_Key() As String
    ISnapShotEntry_Key = CStr(m_ThreadEntry.th32ThreadID)
End Function
Private Function ISnapShotEntry_Name() As String
    ISnapShotEntry_Name = CStr(m_ThreadEntry.th32ThreadID)
End Function
Private Function ISnapShotEntry_Ptr() As Long
    m_ThreadEntry.dwSize = LenB(m_ThreadEntry)
    ISnapShotEntry_Ptr = VarPtr(m_ThreadEntry)
End Function

