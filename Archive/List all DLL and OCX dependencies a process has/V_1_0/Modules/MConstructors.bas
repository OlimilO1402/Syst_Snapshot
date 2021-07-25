Attribute VB_Name = "MConstructors"
Option Explicit

Public Function New_Snapshot(ByVal Flags As SnapShotFlags, Optional ByVal PID As Long) As SnapShot
    Set New_Snapshot = New SnapShot
    Call New_Snapshot.NewC(Flags, PID)
End Function


Public Function New_ProcessEntry(aSnapShot As SnapShot) As ProcessEntry
    Set New_ProcessEntry = New ProcessEntry
    Call New_ProcessEntry.NewC(aSnapShot)
End Function
Public Function New_ModuleEntry(aSnapShot As SnapShot) As ModuleEntry
    Set New_ModuleEntry = New ModuleEntry
    Call New_ModuleEntry.NewC(aSnapShot)
End Function
Public Function New_ThreadEntry(aSnapShot As SnapShot) As ThreadEntry
    Set New_ThreadEntry = New ThreadEntry
    Call New_ThreadEntry.NewC(aSnapShot)
End Function
Public Function New_HeapList(aSnapShot As SnapShot) As HeapList
    Set New_HeapList = New HeapList
    Call New_HeapList.NewC(aSnapShot)
End Function
Public Function New_HeapEntry(aSnapShot As SnapShot, ByVal Index As Long) As HeapEntry
    Set New_HeapEntry = New HeapEntry
    Call New_HeapEntry.NewC(aSnapShot, Index)
End Function
