Attribute VB_Name = "MNew"
Option Explicit

Public Function Snapshot(ByVal Flags As SnapShotFlags, Optional ByVal PID As Long) As Snapshot
    Set Snapshot = New Snapshot: Snapshot.New_ Flags, PID
End Function

Public Function ProcessEntry(aSnapShot As Snapshot) As ProcessEntry
    Set ProcessEntry = New ProcessEntry: ProcessEntry.New_ aSnapShot
End Function

Public Function ModuleEntry(aSnapShot As Snapshot) As ModuleEntry
    Set ModuleEntry = New ModuleEntry: ModuleEntry.New_ aSnapShot
End Function

Public Function ThreadEntry(aSnapShot As Snapshot) As ThreadEntry
    Set ThreadEntry = New ThreadEntry: ThreadEntry.New_ aSnapShot
End Function

Public Function HeapList(aSnapShot As Snapshot) As HeapList
    Set HeapList = New HeapList: HeapList.New_ aSnapShot
End Function

Public Function HeapEntry(aSnapShot As Snapshot, ByVal Index As Long) As HeapEntry
    Set HeapEntry = New HeapEntry: HeapEntry.New_ aSnapShot, Index
End Function
