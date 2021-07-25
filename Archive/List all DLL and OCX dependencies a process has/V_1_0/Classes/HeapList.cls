VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HeapList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISnapShotEntry

'Public Type HEAPLIST32
'    dwSize        As Long 'The size of the structure, in bytes. Before calling the Heap32ListFirst function, set this member. If you do not initialize dwSize, Heap32ListFirst will fail.
'    th32ProcessID As Long 'The identifier of the process to be examined
'    th32HeapID    As Long 'The heap identifier. This is not a handle, and has meaning only to the tool help functions
'    dwFlags       As Long 'This member can be one of the following values.
'End Type
'Public Const HF32_DEFAULT As Long = 1

Private m_HeapList    As HEAPLIST32
Private m_Snapshot    As SnapShot
Private m_HeapEntries As Collection

Public Sub NewC(aSnapShot As SnapShot)
    Set m_Snapshot = aSnapShot
    'set
End Sub

Public Property Get Ptr() As Long
    Ptr = ISnapShotEntry_Ptr
End Property
Public Property Get Key() As String
    Key = ISnapShotEntry_Key
End Property
Public Sub FetchHeapEntries()
    'Call HeapEntries
    Set m_HeapEntries = HeapEntries
End Sub
Public Function ToString() As String
    Dim s As String
    With m_HeapList
        s = s & "Flags:  " & FlagsToString(.dwFlags) & vbCrLf
        s = s & "HeapID: " & CStr(.th32HeapID) & vbCrLf
        s = s & "ProcessID: " & CStr(.th32ProcessID) & vbCrLf
    End With
    ToString = s
End Function

Private Function FlagsToString(f As Long) As String
    If (f And HF32_DEFAULT) <> 0 Then FlagsToString = "default"
End Function
Public Property Get ProcessID() As Long
    ProcessID = m_HeapList.th32ProcessID
End Property
Public Property Get HeapID() As Long
    HeapID = m_HeapList.th32HeapID
End Property

Public Property Get HeapEntries() As Collection
    If (m_Snapshot.Flags And SnapHeapEntry) = SnapHeapEntry Then
        If m_HeapEntries Is Nothing Then
            'Dim snap As SnapShot: Set snap = New_Snapshot(SnapHeapLists)
            Dim iter As SnapIter: iter = m_Snapshot.GetIterator(Me)
            'MsgBox m_Snapshot.Flags
            'iter = snap.GetIterator(Me)
            Dim col As New Collection
            'Dim ent As HeapEntry ': Set ent = MSnapIter.NextHeapEntry(iter)
            'Do While Not ent Is Nothing
            While MSnapIter.HasNextHeapEntry(iter)
                'Set ent = iter.NNext
                'If ent.HeapID = Me.HeapID Then 'm_HeapList.th32HeapID Then
                'If iter.NNext.ID = Me.HeapID Then
                '    col.Add ent, ent.Key
                   col.Add iter.NNext, iter.NNext.Key
                'End If
                'Set ent = MSnapIter.NextHeapEntry(iter)
            Wend
            'Loop
            Set m_HeapEntries = col
        End If
    End If
    Set HeapEntries = m_HeapEntries
End Property

Private Function ISnapShotEntry_ID() As Long
    ISnapShotEntry_ID = m_HeapList.th32HeapID
End Function
Private Function ISnapShotEntry_Key() As String
    ISnapShotEntry_Key = CStr(m_HeapList.th32HeapID)
End Function
Private Function ISnapShotEntry_Name() As String
    ISnapShotEntry_Name = CStr(m_HeapList.th32HeapID)
End Function
Private Function ISnapShotEntry_Ptr() As Long
    m_HeapList.dwSize = LenB(m_HeapList)
    ISnapShotEntry_Ptr = VarPtr(m_HeapList)
End Function
