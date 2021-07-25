Attribute VB_Name = "MSnapIter"
Option Explicit

Public Type SnapIter
    Snapshot  As Snapshot
    ProcessID As Long 'HeapList  As HeapList
    HeapID    As Long
    hSnapshot As Long
    hResult   As Long
    Count     As Long
    bFirst    As Boolean
    NNext     As ISnapShotEntry
End Type

Public Function New_SnapIter(aSnapShot As Snapshot, Optional aHeapList As HeapList) As SnapIter
    With New_SnapIter
        Set .Snapshot = aSnapShot
        'Set .HeapList = aHeapList
        If Not aHeapList Is Nothing Then
            .ProcessID = aHeapList.ProcessID
            .HeapID = aHeapList.HeapID
        End If
        .hSnapshot = .Snapshot.Handle
        .bFirst = True
    End With
End Function

Public Function HasNextProcessEntry(this As SnapIter) As Boolean 'ISnapShotEntry 'ProcessEntry
    With this
        If .bFirst Then
            If .hResult = 0 Then
                Set .NNext = MNew.ProcessEntry(.Snapshot)
                .hResult = Process32First(.hSnapshot, ByVal .NNext.Ptr)
                .bFirst = False
                .Count = .Count + 1
                HasNextProcessEntry = .hResult <> 0 'True
            End If
        Else
            If .hResult <> 0 Then
                Set .NNext = MNew.ProcessEntry(.Snapshot)
                .hResult = Process32Next(.hSnapshot, ByVal .NNext.Ptr)
                .Count = .Count + 1
                HasNextProcessEntry = .hResult <> 0 'True
            End If
        End If
    End With
End Function

Public Function HasNextModuleEntry(this As SnapIter) As Boolean 'ISnapShotEntry 'ModuleEntry
    With this
        If .bFirst Then
            If .hResult = 0 Then
                Set .NNext = MNew.ModuleEntry(.Snapshot)
                .hResult = Module32First(.hSnapshot, ByVal .NNext.Ptr)
                .bFirst = False
                .Count = .Count + 1
                HasNextModuleEntry = .hResult <> 0 'True
            End If
        Else
            If .hResult <> 0 Then
                Set .NNext = MNew.ModuleEntry(.Snapshot)
                .hResult = Module32Next(.hSnapshot, ByVal .NNext.Ptr)
                .Count = .Count + 1
                HasNextModuleEntry = .hResult <> 0 'True
            End If
        End If
    End With
End Function

Public Function HasNextThreadEntry(this As SnapIter) As Boolean 'ISnapShotEntry 'ThreadEntry
    With this
        If .bFirst Then
            If .hResult = 0 Then
                Set .NNext = MNew.ThreadEntry(.Snapshot)
                .hResult = Thread32First(.hSnapshot, ByVal .NNext.Ptr)
                .bFirst = False
                .Count = .Count + 1
                HasNextThreadEntry = .hResult <> 0 'True
            End If
        Else
            If .hResult <> 0 Then
                Set .NNext = MNew.ThreadEntry(.Snapshot)
                .hResult = Thread32Next(.hSnapshot, ByVal .NNext.Ptr)
                .Count = .Count + 1
                HasNextThreadEntry = .hResult <> 0 'True
            End If
        End If
    End With
End Function

Public Function HasNextHeapList(this As SnapIter) As Boolean 'ISnapShotEntry 'HeapList
    With this
        If .bFirst Then
            If .hResult = 0 Then
                Set .NNext = MNew.HeapList(.Snapshot)
                .hResult = Heap32ListFirst(.hSnapshot, ByVal .NNext.Ptr)
                .bFirst = False
                .Count = .Count + 1
                HasNextHeapList = .hResult <> 0 'True
            End If
        Else
            If .hResult <> 0 Then
                Set .NNext = MNew.HeapList(.Snapshot)
                .hResult = Heap32ListNext(.hSnapshot, ByVal .NNext.Ptr)
                .Count = .Count + 1
                HasNextHeapList = .hResult <> 0 'True
            End If
        End If
    End With
End Function

Public Function HasNextHeapEntry(this As SnapIter) As Boolean 'ISnapShotEntry 'HeapEntry
    With this
        If .bFirst Then
            If .hResult = 0 Then
                Set .NNext = MNew.HeapEntry(.Snapshot, .Count)
                'Debug.Print .ProcessID & " " & .HeapID
                .hResult = Heap32First(ByVal .NNext.Ptr, .ProcessID, .HeapID)
                .bFirst = False
                .Count = .Count + 1
                HasNextHeapEntry = .hResult <> 0 'True
            End If
        Else
            If .hResult <> 0 Then
                Set .NNext = MNew.HeapEntry(.Snapshot, .Count)
                .hResult = Heap32Next(ByVal .NNext.Ptr)
                .Count = .Count + 1
                HasNextHeapEntry = .hResult <> 0 'True
            End If
        End If
    End With
End Function

