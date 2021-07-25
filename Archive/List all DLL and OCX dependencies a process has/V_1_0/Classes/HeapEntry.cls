VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HeapEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISnapShotEntry
'Public Type HEAPENTRY32
'    dwSize        As Long ' The size of the structure, in bytes.
'                          ' Before calling the Heap32First function, set this member.
'                          ' If you do not initialize dwSize, Heap32First fails. ' SIZE_T
'    hHandle       As Long ' A handle to the heap block.
'    dwAddress     As Long ' The linear address of the start of the block. ' As ULONG_PTR
'    dwBlockSize   As Long ' The size of the heap block, in bytes ' As SIZE_T
'    dwFlags       As Long ' This member can be one of the following values (see below LF32_...).
'    dwLockCount   As Long ' This member is no longer used and is always set to zero
'    dwReserved    As Long ' Reserved; do not use or alter.
'    th32ProcessID As Long ' The identifier of the process that uses the heap
'    th32HeapID    As Long ' The heap identifier.
'                          ' This is not a handle, and has meaning only to the tool help functions'ULONG_PTR
'End Type
'Public Const LF32_FIXED    As Long = &H1 'The memory block has a fixed (unmovable) location.
'Public Const LF32_FREE     As Long = &H2 'The memory block is not used.
'Public Const LF32_MOVEABLE As Long = &H4 'The memory block location can be moved.
Public Enum HeapFlags
    HeapFixed = &H1
    HeapFree = &H2
    HeapMoveable = &H4
End Enum
Private m_HeapEntry As HEAPENTRY32
Private m_Snapshot  As SnapShot
Private m_Index     As Long

Public Sub NewC(aSnapShot As SnapShot, ByVal Index As Long)
    Set m_Snapshot = aSnapShot
    m_Index = Index
End Sub

Public Property Get Ptr() As Long
    Ptr = ISnapShotEntry_Ptr
End Property
Public Property Get Key() As String
    Key = ISnapShotEntry_Key
End Property
Public Property Get Handle() As Long
    Handle = m_HeapEntry.hHandle
End Property
Public Property Get Address() As Long
    Address = m_HeapEntry.dwAddress
End Property
Public Property Get BlockSize() As Long
    BlockSize = m_HeapEntry.dwBlockSize
End Property
Public Property Get Flags() As HeapFlags
    Flags = m_HeapEntry.dwFlags
End Property
Public Property Get ProcessID() As Long
    ProcessID = m_HeapEntry.th32ProcessID
End Property
Public Property Get HeapID() As Long
    HeapID = m_HeapEntry.th32HeapID
End Property

Public Function ToString() As String
    Dim s As String
    With m_HeapEntry
        s = s & "BlockHandle: " & .hHandle & vbCrLf
        s = s & "Address:     " & .dwAddress & vbCrLf
        s = s & "BlockSize:   " & .dwBlockSize & vbCrLf
        s = s & "Flags:       " & FlagsToString(.dwFlags) & vbCrLf
        s = s & "ProcessID:   " & .th32ProcessID & vbCrLf
        s = s & "HeapID:      " & .th32HeapID & vbCrLf
    End With
    ToString = s
End Function

Private Function FlagsToString(ByVal f As Long) As String
    Dim s As String
    If (f And LF32_FIXED) <> 0 Then _
                                    s = s & IIf(Len(s), " ", "") & "fixed"
    If (f And LF32_FREE) <> 0 Then _
                                    s = s & IIf(Len(s), " ", "") & "free"
    If (f And LF32_MOVEABLE) <> 0 Then _
                                    s = s & IIf(Len(s), " ", "") & "moveable"
    FlagsToString = s
End Function

Private Function ISnapShotEntry_ID() As Long
    ISnapShotEntry_ID = m_HeapEntry.hHandle
End Function
Private Function ISnapShotEntry_Key() As String
    'ISnapShotEntry_Key = CStr(m_HeapEntry.hHandle)
    ISnapShotEntry_Key = m_Index
End Function
Private Function ISnapShotEntry_Name() As String
    ISnapShotEntry_Name = m_HeapEntry.hHandle
End Function

Private Function ISnapShotEntry_Ptr() As Long
    m_HeapEntry.dwSize = LenB(m_HeapEntry)
    ISnapShotEntry_Ptr = VarPtr(m_HeapEntry)
End Function

