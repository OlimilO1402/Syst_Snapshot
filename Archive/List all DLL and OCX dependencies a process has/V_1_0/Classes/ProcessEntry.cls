VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISnapShotEntry
'Public Type PROCESSENTRY32
'    dwSize              As Long 'The size of the structure, in bytes. Before calling the Process32First function, set this member to sizeof(PROCESSENTRY32). If you do not initialize dwSize, Process32First fails.
'    cntUsage            As Long 'This member is no longer used and is always set to zero
'    th32ProcessID       As Long 'The process identifier
'    th32DefaultHeapID   As Long 'This member is no longer used and is always set to zero
'    th32ModuleID        As Long 'This member is no longer used and is always set to zero
'    cntThreads          As Long 'The number of execution threads started by the process.
'    th32ParentProcessID As Long 'The identifier of the process that created this process (its parent process).
'    pcPriClassBase      As Long 'The base priority of any threads created by this process.
'    dwFlags             As Long 'This member is no longer used, and is always set to zero
'    szExeFile           As String * MAX_PATH
'    'The name of the executable file for the process. To retrieve the full path to the executable file, call the Module32First function and check the szExePath member of the MODULEENTRY32 structure that is returned.
'    'However, if the calling process is a 32-bit process, you must call the QueryFullProcessImageName function to retrieve the full path of the executable file for a 64-bit process.
'End Type

Private m_ProcessEntry   As PROCESSENTRY32
Private m_Snapshot       As SnapShot
Private m_ChildProcesses As Collection
Private m_ModuleEntries  As Collection
Private m_ThreadEntries  As Collection
Private m_HeapLists      As Collection

Public Sub NewC(aSnapShot As SnapShot)
    Set m_Snapshot = aSnapShot
End Sub
Public Function KillProcess() As Boolean
    KillProcess = MProcess.TerminateProcessID(m_ProcessEntry.th32ProcessID)
End Function
Public Property Get Ptr() As Long
    Ptr = ISnapShotEntry_Ptr
End Property
Public Property Get Key() As String
    Key = ISnapShotEntry_Key
End Property
Public Property Get ProcessID() As Long
    ProcessID = m_ProcessEntry.th32ProcessID
End Property
Public Property Get ParentProcessID() As Long
    ParentProcessID = m_ProcessEntry.th32ParentProcessID
End Property
Public Property Get ThreadCount() As Long
    ThreadCount = m_ProcessEntry.cntThreads
End Property
Public Function ToString() As String
    Dim s As String
    With m_ProcessEntry
        s = s & "ProcessID:       " & .th32ProcessID & vbCrLf
        s = s & "Threads:         " & .cntThreads & vbCrLf
        s = s & "ParentProcessID: " & .th32ParentProcessID & vbCrLf
        s = s & "PriClassBase:    " & .pcPriClassBase & vbCrLf
        If .th32ProcessID = 17 Then
            s = s & "ExeFile:         " & GetString(GetString(.szExeFile)) & vbCrLf
        Else
            s = s & "ExeFile:         " & GetString(.szExeFile) & vbCrLf
        End If
    End With
    ToString = s
End Function
Public Function AddChildProcess(p As ProcessEntry) As Boolean
    If p.ParentProcessID = Me.ProcessID Then
        If m_ChildProcesses Is Nothing Then
            Set m_ChildProcesses = New Collection
        End If
        Call m_ChildProcesses.Add(p, CStr(p.ProcessID))
        AddChildProcess = True
    End If
End Function
Public Property Get ChildProcesses() As Collection
    Set ChildProcesses = m_ChildProcesses
End Property
'der Prozess kann Module und Threads auflisten
'Public Function GetModuleCollection() As Collection
'    Dim col As New Collection
'    Dim hResult As Long
'    Dim mde As ModuleEntry
'    Dim snap As SnapShot
'    'Set snap = New_Snapshot(SnapModules, m_ProcessEntry.th32ProcessID)
'    Set snap = m_Snapshot
'    Dim hsh As Long: hsh = snap.Handle
'
'    Set mde = New ModuleEntry
'    hResult = MSnapShot.Module32First(hsh, ByVal mde.Ptr)
'    Do While hResult <> 0
'        If mde.ProcessID = Me.ProcessID Then
'           col.Add mde
'           Set mde = New ModuleEntry
'        End If
'        hResult = MSnapShot.Module32Next(hsh, ByVal mde.Ptr)
'    Loop
'    Set GetModuleCollection = col
'End Function
'
'Public Function GetHeapListCollection() As Collection
'    Dim hResult As Long
'    Dim col As New Collection
'
'    Dim obj As HeapList
'
'    Set obj = New HeapList
'    hResult = Heap32ListFirst(m_Snapshot.Handle, ByVal obj.Ptr)
'
'    Do While Not hResult <> 0
'
'        col.Add obj
'        Set obj = New HeapList
'        hResult = Heap32ListNext(m_Snapshot.Handle, ByVal obj.Ptr)
'
'    Loop
'End Function
'
'
'Public Function GetThreadEntryCollection() As Collection
'
'End Function
'
Public Property Get ModuleEntries() As Collection
    If m_ModuleEntries Is Nothing Then
        If (m_Snapshot.Flags And SnapModules) = SnapModules Then
            Dim snp As SnapShot: Set snp = New_Snapshot(SnapModules, Me.ProcessID)
            
            Dim iter As SnapIter: iter = snp.GetIterator 'MSnapIter.New_SnapIter(snp) 'm_Snapshot)
            
            Dim col As New Collection
            Dim ent As ModuleEntry ': Set ent = MSnapIter.NextModuleEntry(iter)
            'Do While Not ent Is Nothing
            While MSnapIter.HasNextModuleEntry(iter)
               'If ent.ProcessID = m_ProcessEntry.th32ProcessID Then
               Set ent = iter.NNext
               If ent.ProcessModuleHandle Then
                   col.Add ent, ent.Key
               End If
               'Set ent = MSnapIter.NextModuleEntry(iter)
            'Loop
            Wend
            Set m_ModuleEntries = col
        End If
    End If
    Set ModuleEntries = m_ModuleEntries
End Property
Public Property Get ThreadEntries() As Collection
    If m_ProcessEntry.th32ProcessID <> 0 Then
        If m_ThreadEntries Is Nothing Then
            If (m_Snapshot.Flags And SnapThreads) = SnapThreads Then
                Dim iter As SnapIter: iter = m_Snapshot.GetIterator
                
                Dim col As New Collection
                Dim ent As ThreadEntry ': Set ent = MSnapIter.NextThreadEntry(iter)
                'Do While Not ent Is Nothing
                While MSnapIter.HasNextThreadEntry(iter)
                   Set ent = iter.NNext
                   If ent.OwnerProcessID = m_ProcessEntry.th32ProcessID Then
                       col.Add ent, ent.Key
                   End If
                   'Set ent = MSnapIter.NextThreadEntry(iter)
                'Loop
                Wend
                Set m_ThreadEntries = col
            End If
        End If
    End If
    Set ThreadEntries = m_ThreadEntries
End Property

Public Property Get HeapLists() As Collection
    If (m_Snapshot.Flags And SnapHeapLists) = SnapHeapLists Then
        If m_HeapLists Is Nothing Then
            Dim snp As SnapShot: Set snp = New_Snapshot(SnapHeapLists, m_ProcessEntry.th32ProcessID)
            
            Dim iter As SnapIter: iter = snp.GetIterator
            'Dim iter As SnapIter: iter = m_Snapshot.GetIterator
            
            Dim col As New Collection
            Dim ent As HeapList ': Set ent = MSnapIter.NextHeapList(iter)
            'Do While Not ent Is Nothing
            While MSnapIter.HasNextHeapList(iter)
                Set ent = iter.NNext
                'If ent.ProcessID = m_ProcessEntry.th32ProcessID Then
                    ent.FetchHeapEntries
                    col.Add ent, ent.Key
                'End If
                'Set ent = MSnapIter.NextHeapList(iter)
            'Loop
            Wend
            Set m_HeapLists = col
        End If
    End If
    Set HeapLists = m_HeapLists
End Property


Private Function ISnapShotEntry_ID() As Long
    ISnapShotEntry_ID = m_ProcessEntry.th32ProcessID
End Function
Private Function ISnapShotEntry_Key() As String
    ISnapShotEntry_Key = CStr(m_ProcessEntry.th32ProcessID)
End Function
Private Function ISnapShotEntry_Name() As String
    ISnapShotEntry_Name = GetString(m_ProcessEntry.szExeFile)
End Function
Private Function ISnapShotEntry_Ptr() As Long
    m_ProcessEntry.dwSize = LenB(m_ProcessEntry)
    ISnapShotEntry_Ptr = VarPtr(m_ProcessEntry)
End Function

