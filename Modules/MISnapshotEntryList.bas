Attribute VB_Name = "MISnapshotEntryList"
Option Explicit

'Public Sub ProcessesToListBox(col As Collection, aLB As ListBox)
'    Dim p As ProcessEntry
'    aLB.Clear
'    For Each p In col
'        aLB.AddItem p.Key
'    Next
'End Sub

Public Sub ISnapshotEntryListToListBox(col As Collection, aLB As ListBox)
    Dim ent As ISnapShotEntry 'ProcessEntry
    aLB.Clear
    If Not col Is Nothing Then
        For Each ent In col
            aLB.AddItem ent.Name 'Key
            aLB.ItemData(aLB.NewIndex) = ent.ID 'CLng(ent.Key)
        Next
    End If
End Sub

Public Function ContainsKey(col As Collection, Key As String) As Boolean
    On Error Resume Next
        If IsEmpty(col.Item(Key)) Then: 'Donothing
        ContainsKey = (Err.Number = 0)
    On Error GoTo 0
End Function


