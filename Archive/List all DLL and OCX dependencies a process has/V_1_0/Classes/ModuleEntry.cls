VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModuleEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISnapShotEntry

'Public Type MODULEENTRY32
'    dwSize             As Long 'The size of the structure, in bytes. Before calling the Module32First function, set this member to sizeof(MODULEENTRY32). If you do not initialize dwSize, Module32First fails.
'    th32ModuleID       As Long 'This member is no longer used, and is always set to one.
'    th32ProcessID      As Long 'The identifier of the process whose modules are to be examined
'    GlblcntUsage       As Long 'The load count of the module, which is not generally meaningful, and usually equal to 0xFFFF
'    ProccntUsage       As Long 'The load count of the module (same as GlblcntUsage), which is not generally meaningful, and usually equal to 0xFFFF.
'    modBaseAddr        As Long 'The base address of the module in the context of the owning process.
'    modBaseSize        As Long 'The size of the module, in bytes.
'    hModule            As Long 'A handle to the module in the context of the owning process.
'    szModule           As String * MAX_MODULE_NAME32 'The module name.
'    szExePath          As String * MAX_PATH 'The module path.
'End Type
Private m_ModuleEntry As MODULEENTRY32
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

Public Property Get ProcessID() As Long
    ProcessID = m_ModuleEntry.th32ProcessID
End Property
'Public Property Get GlobalUsageLoadCount() As Long
'    GlobalUsageLoadCount = m_ModuleEntry.GlblcntUsage
'End Property
Public Property Get BaseAddress() As Long
    BaseAddress = m_ModuleEntry.modBaseAddr
End Property
Public Property Get BaseSize() As Long
    BaseSize = m_ModuleEntry.modBaseSize
End Property
Public Property Get ProcessModuleHandle() As Long
    ProcessModuleHandle = m_ModuleEntry.hModule
End Property
Public Property Get Name() As String
    Name = GetStringFromByteArr(m_ModuleEntry.szModule) ' ISnapShotEntry_Name '= GetString(m_ModuleEntry.szModule)
End Property
Public Property Get Path() As String
    Path = GetStringFromByteArr(m_ModuleEntry.szExePath) 'GetString(m_ModuleEntry.szExePath)
End Property

Public Function ToString() As String
    Dim s As String
    With m_ModuleEntry
        s = s & "ModuleID:       " & CStr(.th32ModuleID) & vbCrLf
        s = s & "ProcessID:      " & CStr(.th32ProcessID) & vbCrLf
        s = s & "glob.UsageLoad: " & CStr(.GlblcntUsage) & vbCrLf
        s = s & "proc.UsageLoad: " & CStr(.ProccntUsage) & vbCrLf
        s = s & "Mod. Base Addr: " & CStr(.modBaseAddr) & vbCrLf
        s = s & "Mod. Base Size: " & CStr(.modBaseSize) & vbCrLf
        s = s & "Module Handle:  " & CStr(.hModule) & vbCrLf
        s = s & "Module name:    " & Me.Name & vbCrLf 'GetString(.szModule) & vbCrLf
        s = s & "Exe path:       " & Me.Path & vbCrLf 'GetString(.szExePath) & vbCrLf
    End With
    ToString = s
End Function

Private Function ISnapShotEntry_ID() As Long
    ISnapShotEntry_ID = m_ModuleEntry.hModule
End Function
Public Function ISnapShotEntry_Key() As String
    ISnapShotEntry_Key = CStr(m_ModuleEntry.hModule)
End Function
Private Function ISnapShotEntry_Name() As String
    ISnapShotEntry_Name = GetStringFromByteArr(m_ModuleEntry.szModule)
End Function
Private Function ISnapShotEntry_Ptr() As Long
    m_ModuleEntry.dwSize = LenB(m_ModuleEntry)
    ISnapShotEntry_Ptr = VarPtr(m_ModuleEntry)
End Function

