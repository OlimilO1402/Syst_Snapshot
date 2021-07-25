VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SnapShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum SnapShotFlags
    SnapHeapEntry = &H0
    SnapHeapLists = &H1
    SnapProcesses = &H2
    SnapThreads = &H4
    SnapModules = &H8
    SnapAll = &HF
End Enum

Private m_Flags          As SnapShotFlags
Private m_ShapShotHandle As Long
Private m_ProcessEntries As Collection

Public Sub NewC(ByVal Flags As SnapShotFlags, Optional ByVal PID As Long)
    m_Flags = Flags
    'ist PID = 0 wird automatisch von der API der aktuelle Prozess verwendet.
    m_ShapShotHandle = CreateToolhelp32Snapshot(m_Flags, PID)
End Sub

Private Sub Class_Terminate()
    If m_ShapShotHandle <> 0 Then
        Call CloseHandle(m_ShapShotHandle)
    End If
End Sub
Public Property Get Handle() As Long
    'Debug.Print m_ShapShotHandle
    Handle = m_ShapShotHandle
End Property
Public Sub CloseSnap()
    Call Class_Terminate
End Sub

Public Property Get Flags() As SnapShotFlags
    Flags = m_Flags
End Property
Friend Function GetIterator(Optional aHeapList As HeapList) As SnapIter
    GetIterator = MSnapIter.New_SnapIter(Me, aHeapList)
End Function

Public Property Get ProcessEntries() As Collection
    If m_ProcessEntries Is Nothing Then
        If (m_Flags And SnapProcesses) = SnapProcesses Then
            Dim iter As SnapIter: iter = Me.GetIterator
            
            Dim col As New Collection
            While MSnapIter.HasNextProcessEntry(iter) '= True
                Call col.Add(iter.NNext, iter.NNext.Key)
            Wend
            
            'so und jetzt noch an jeden Prozess den Kindprozess anhängen
            Dim child As ProcessEntry
            Dim proc  As ProcessEntry
            Dim k As String
            For Each child In col
                If child.ParentProcessID <> 0 Then
                    k = CStr(child.ParentProcessID)
                    If ContainsKey(col, k) Then
                        Set proc = col.Item(k)
                        Call proc.AddChildProcess(child)
                    End If
                End If
            Next
            Set m_ProcessEntries = col
        End If
    End If
    Set ProcessEntries = m_ProcessEntries
End Property
