Attribute VB_Name = "MSnapShot"
Option Explicit
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const MAX_MODULE_NAME32   As Long = 256
Public Const MAX_PATH            As Long = 260

'Toolhelp create snapshot flags
'siehe auch Enum SnapShotFlags def in class SnapShot
Public Const TH32CS_SNAPHEAPLIST As Long = &H1 ' Erzeugt zusätzlich einen Snapshot von Heaplisten
Public Const TH32CS_SNAPPROCESS  As Long = &H2 ' Erzeugt zusätzlich einen Snapshot von Prozessen
Public Const TH32CS_SNAPTHREAD   As Long = &H4 ' Erzeugt zusätzlich einen Snapshot von Threads
Public Const TH32CS_SNAPMODULE   As Long = &H8 ' Erzeugt zusätzlich einen Snapshot von Modulen
Public Const TH32CS_SNAPALL      As Long = &HF ' Erzeugt einen Snapshot aller Ressourcen
Public Const TH32CS_INHERIT      As Long = &H80000000 ' Erzeugt einen vererbbaren Snapshot

Public Type HEAPLIST32
    dwSize        As Long 'The size of the structure, in bytes. Before calling the Heap32ListFirst function, set this member. If you do not initialize dwSize, Heap32ListFirst will fail.
    th32ProcessID As Long 'The identifier of the process to be examined
    th32HeapID    As Long 'The heap identifier. This is not a handle, and has meaning only to the tool help functions
    dwFlags       As Long 'This member can be one of the following values.
End Type
Public Const HF32_DEFAULT As Long = 1
Public Type HEAPENTRY32
    dwSize        As Long 'The size of the structure, in bytes. Before calling the Heap32First function, set this member. If you do not initialize dwSize, Heap32First fails.'SIZE_T
    hHandle       As Long 'A handle to the heap block.
    dwAddress     As Long 'The linear address of the start of the block. ' As ULONG_PTR
    dwBlockSize   As Long 'The size of the heap block, in bytes ' As SIZE_T
    dwFlags       As Long 'This member can be one of the following values (see below LF32_...).
    dwLockCount   As Long 'This member is no longer used and is always set to zero
    dwReserved    As Long 'Reserved; do not use or alter.
    th32ProcessID As Long 'The identifier of the process that uses the heap
    th32HeapID    As Long 'The heap identifier. This is not a handle, and has meaning only to the tool help functions'ULONG_PTR
End Type
'dwFlags
Public Const LF32_FIXED    As Long = &H1 'The memory block has a fixed (unmovable) location.
Public Const LF32_FREE     As Long = &H2 'The memory block is not used.
Public Const LF32_MOVEABLE As Long = &H4 'The memory block location can be moved.

Public Type PROCESSENTRY32
    dwSize              As Long 'The size of the structure, in bytes. Before calling the Process32First function, set this member to sizeof(PROCESSENTRY32). If you do not initialize dwSize, Process32First fails.
    cntUsage            As Long 'This member is no longer used and is always set to zero
    th32ProcessID       As Long 'The process identifier
    th32DefaultHeapID   As Long 'This member is no longer used and is always set to zero
    th32ModuleID        As Long 'This member is no longer used and is always set to zero
    cntThreads          As Long 'The number of execution threads started by the process.
    th32ParentProcessID As Long 'The identifier of the process that created this process (its parent process).
    pcPriClassBase      As Long 'The base priority of any threads created by this process.
    dwFlags             As Long 'This member is no longer used, and is always set to zero
    szExeFile           As String * MAX_PATH
    'The name of the executable file for the process. To retrieve the full path to the executable file, call the Module32First function and check the szExePath member of the MODULEENTRY32 structure that is returned.
    'However, if the calling process is a 32-bit process, you must call the QueryFullProcessImageName function to retrieve the full path of the executable file for a 64-bit process.
End Type
Public Type THREADENTRY32
    dwSize             As Long 'The size of the structure, in bytes. Before calling the Thread32First function, set this member to sizeof(THREADENTRY32). If you do not initialize dwSize, Thread32First fails.
    cntUsage           As Long 'This member is no longer used and is always set to zero
    th32ThreadID       As Long 'The thread identifier, compatible with the thread identifier returned by the CreateProcess function
    th32OwnerProcessID As Long 'The identifier of the process that created the thread.
    tpBasePri          As Long 'The kernel base priority level assigned to the thread. The priority is a number from 0 to 31, with 0 representing the lowest possible thread priority. For more information, see KeQueryPriorityThread.
    tpDeltaPri         As Long 'This member is no longer used and is always set to zero.
    dwFlags            As Long 'This member is no longer used and is always set to zero.
End Type
Public Type MODULEENTRY32
    dwSize             As Long 'The size of the structure, in bytes. Before calling the Module32First function, set this member to sizeof(MODULEENTRY32). If you do not initialize dwSize, Module32First fails.
    th32ModuleID       As Long 'This member is no longer used, and is always set to one.
    th32ProcessID      As Long 'The identifier of the process whose modules are to be examined
    GlblcntUsage       As Long 'The load count of the module, which is not generally meaningful, and usually equal to 0xFFFF
    ProccntUsage       As Long 'The load count of the module (same as GlblcntUsage), which is not generally meaningful, and usually equal to 0xFFFF.
    modBaseAddr        As Long 'The base address of the module in the context of the owning process.
    modBaseSize        As Long 'The size of the module, in bytes.
    hModule            As Long 'A handle to the module in the context of the owning process.
    szModule(0 To MAX_MODULE_NAME32 - 1) As Byte ' As String * MAX_MODULE_NAME32  'The module name.
    szExePath(0 To MAX_PATH - 1)         As Byte ' As String * MAX_PATH 'The module path.
End Type

Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Declare Function Heap32ListFirst Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef pHeapList As Any) As Long
Public Declare Function Heap32ListNext Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef pHeapList As Any) As Long

Public Declare Function Heap32First Lib "kernel32.dll" (ByRef lphe As Any, ByVal th32ProcessID As Long, ByVal th32HeapID As Long) As Long 'ULONG_PTR
Public Declare Function Heap32Next Lib "kernel32.dll" (ByRef lphe As Any) As Long

Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, ByRef pProcess As Any) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef pProcess As Any) As Long

Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef pThread As Any) As Long
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef pThread As Any) As Long

Public Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, ByRef pModule As Any) As Long
Public Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef pModule As Any) As Long

Public Declare Function Toolhelp32ReadProcessMemory Lib "kernel32.dll" (ByVal th32ProcessID As Long, ByRef lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal cbRead As Long, ByRef lpNumberOfBytesRead As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private m_LB As ListBox
Private m_CB As CommandButton
Private m_Counter As Long

Public Function GetString(strval As String) As String
    GetString = StrConv(strval, vbUnicode)
    Dim pos As Long: pos = InStr(GetString, vbNullChar)
    If pos > 0 Then
        GetString = Left$(GetString, pos - 1)
    End If
End Function

Public Function GetStringFromByteArr(strval() As Byte) As String
    GetStringFromByteArr = StrConv(strval, vbUnicode)
    Dim pos As Long: pos = InStr(GetStringFromByteArr, vbNullChar)
    If pos > 0 Then
        GetStringFromByteArr = Left$(GetStringFromByteArr, pos - 1)
    End If
End Function

Private Sub Trace(s As String)
    m_Counter = m_Counter + 1
    If Not m_LB Is Nothing Then
        m_LB.AddItem s
    End If
    If (m_Counter Mod 20) = 0 Then
        m_CB.Caption = "Trav. " & CStr(m_Counter)
        DoEvents
    End If
End Sub
'Traversing the Heap List
Public Sub TraverseHeapList(aLB As ListBox, aCB As CommandButton, Optional ByVal ProcessID As Long)
    Set m_LB = aLB
    Set m_CB = aCB
    m_Counter = 0
    aLB.Clear
    aLB.Visible = False
    
    Dim hr As Long
'   HEAPLIST32 hl;
    Dim hl As HEAPLIST32
    'If ProcessID = 0 Then ProcessID = GetCurrentProcessId
'
'   HANDLE hHeapSnap = CreateToolhelp32Snapshot(TH32CS_SNAPHEAPLIST, GetCurrentProcessId());
    Dim hHeapSnap As Long
    hHeapSnap = CreateToolhelp32Snapshot(TH32CS_SNAPHEAPLIST, ProcessID)
    Trace "hHeapSnap: " & hHeapSnap
'
'   hl.dwSize = sizeof(HEAPLIST32);
    hl.dwSize = LenB(hl)
'
'   if ( hHeapSnap == INVALID_HANDLE_VALUE )
    If hHeapSnap = 0 Then
'   {
'      printf ("CreateToolhelp32Snapshot failed (%d)\n", GetLastError());
       'Debug.Print "CreateToolhelp32Snapshot failed: " & Err.LastDllError & Err.Description ' & vbCrLf
        Trace "CreateToolhelp32Snapshot failed (hHeapSnap = 0): " & Err.LastDllError & " " & Err.Description ' & vbCrLf
'      return;
        Exit Sub
'   }
    End If
'
'   if( Heap32ListFirst( hHeapSnap, &hl ) )
    hr = Heap32ListFirst(hHeapSnap, hl)
    If hr <> 0 Then
'   {
       Trace "Heap32ListFirst:  HeapID: " & hl.th32HeapID & " ProcessID: " & hl.th32ProcessID
'      Do
       Do
'      {
'         HEAPENTRY32 he;
          Dim he As HEAPENTRY32
'         ZeroMemory(&he, sizeof(HEAPENTRY32));
          Call ZeroMemory(he, LenB(he))
'         he.dwSize = sizeof(HEAPENTRY32);
          he.dwSize = LenB(he)

'         if( Heap32First( &he, GetCurrentProcessId(), hl.th32HeapID ) )
          hr = Heap32First(he, hl.th32ProcessID, hl.th32HeapID)
          'hr = Heap32First(he, pProcID, pheapID)
          If hr <> 0 Then
'         {
            'aLB.AddItem "Heap32First: " & hr & "   pProcID: " & pProcID & " pheapID: " & pheapID
             Trace "   Heap32First: " & hr & " hl.th32ProcessID: " & hl.th32ProcessID & " hl.th32HeapID: " & hl.th32HeapID & " Flags: " & hl.dwFlags
'            printf( "\nHeap ID: %d\n", hl.th32HeapID );
             'Debug.Print "Heap ID: " & hl.th32HeapID
             'aLB.AddItem "Heap ID: " & hl.th32HeapID
'            Do
             Do
'            {
'               printf( "Block size: %d\n", he.dwBlockSize );
                If he.th32HeapID = hl.th32HeapID Then
                    Trace "   Heap32Next:  Handle: " & he.hHandle & " ProcessID: " & he.th32ProcessID & " HeapID: " & he.th32HeapID & " BlockSize: " & he.dwBlockSize
                End If
'               he.dwSize = sizeof(HEAPENTRY32);
                he.dwSize = LenB(he)
'            } while( Heap32Next(&he) );
                hr = Heap32Next(he)
                'aLB.AddItem "Heap32Next: " & hr
             Loop While hr <> 0
'         }
          Else
             Trace "Cannot list first heapentry: " & CStr(Err.LastDllError) & " " & Err.Description
          End If
'         hl.dwSize = sizeof(HEAPLIST32);
          hl.dwSize = LenB(hl)
'      } while (Heap32ListNext( hHeapSnap, &hl ));
          hr = Heap32ListNext(hHeapSnap, hl)
          'aLB.AddItem "Heap32ListNext: " & hr
          If hr <> 0 Then
             Trace "Heap32ListNext:  HeapID: " & hl.th32HeapID & " ProcessID: " & hl.th32ProcessID
          End If
       Loop While hr <> 0
'   }
'   else printf ("Cannot list first heap (%d)\n", GetLastError());
    Else
       Trace "Cannot list first heaplist: " & CStr(Err.LastDllError) & " " & Err.Description
    End If
'
    aLB.Visible = True
    
    Call CloseHandle(hHeapSnap)
    
End Sub

Public Function GetProcessIDByName(aProcessName As String) As Long
    Dim hSnapshot As Long, RetVal As Long, slen As Long
    Dim Process As PROCESSENTRY32
    Dim ProcessName As String
  
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Function
    Process.dwSize = LenB(Process)
    RetVal = Process32First(hSnapshot, Process)
    
    Do While RetVal <> 0
        slen = InStr(Process.szExeFile, vbNullChar) - 1
        ProcessName = Left$(Process.szExeFile, slen)
        If StrComp(ProcessName, aProcessName, vbTextCompare) = 0 Then
            GetProcessIDByName = Process.th32ProcessID
            Exit Do
        End If
        RetVal = Process32Next(hSnapshot, Process)
    Loop
    Call CloseHandle(hSnapshot)
End Function


Public Function GetProcessModules(DependencyList() As String, Optional ByVal PID As Long) As Boolean
   Dim Me32 As MODULEENTRY32
   Dim lRet As Long
   Dim lhSnapShot As Long
   Dim iLen As Integer
   Dim sModule As String
Try: On Error GoTo Catch
   If PID = 0 Then PID = GetCurrentProcessId
   ReDim DependencyList(0) As String
   lhSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, CLng(PID))
   If lhSnapShot = 0 Then
      GetProcessModules = False
      Exit Function
   End If
   Me32.dwSize = Len(Me32)
   lRet = Module32First(lhSnapShot, Me32)
   Do While lRet
      If Me32.th32ProcessID = CLng(PID) Then
         With Me32
            iLen = InStr(.szExePath, Chr(0))
            sModule = IIf(iLen = 0, CStr(.szExePath), Left$(.szExePath, iLen - 1))
            If DependencyList(0) = "" Then
                DependencyList(0) = sModule
            Else
                ReDim Preserve DependencyList(UBound(DependencyList) + 1)
                DependencyList(UBound(DependencyList)) = sModule
            End If
         End With
      End If
      lRet = Module32Next(lhSnapShot, Me32)
   Loop
   Call CloseHandle(lhSnapShot)
   GetProcessModules = True
   Exit Function
Catch:
   GetProcessModules = False
End Function

'Traversing the Thread List
'BOOL ListProcessThreads( DWORD dwOwnerPID )
Public Function ListProcessThreads(ByVal dwOwnerPID As Long, ByRef strout() As String) As Boolean
'{
'  HANDLE hThreadSnap = INVALID_HANDLE_VALUE;
   Dim hThreadSnap As Long
'  THREADENTRY32 te32;
   Dim te32 As THREADENTRY32
'
'  // Take a snapshot of all running threads
'  hThreadSnap = CreateToolhelp32Snapshot( TH32CS_SNAPTHREAD, 0 );
   hThreadSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0&)
   
'  if( hThreadSnap == INVALID_HANDLE_VALUE )
   If (hThreadSnap = INVALID_HANDLE_VALUE) Then
'    return( FALSE );
       Exit Function
   End If
'
'  // Fill in the size of the structure before using it.
'  te32.dwSize = sizeof(THREADENTRY32 );
   te32.dwSize = LenB(te32)
'
'  // Retrieve information about the first thread,
'  // and exit if unsuccessful
'  if( !Thread32First( hThreadSnap, &te32 ) )
   If Thread32First(hThreadSnap, te32) <> 0 Then
'  {
'    printError( "Thread32First" );  // Show cause of failure
     Call printError("Thread32First")
'    CloseHandle( hThreadSnap );     // Must clean up the snapshot object!
     Call CloseHandle(hThreadSnap)
'    return( FALSE );
     Exit Function
'  }
   End If
'
'  // Now walk the thread list of the system,
'  // and display information about each thread
'  // associated with the specified process
'  Do
   Do
'  {
'    if( te32.th32OwnerProcessID == dwOwnerPID )
     If te32.th32OwnerProcessID = dwOwnerPID Then
'    {
'      printf( "\n\n     THREAD ID      = 0x%08X", te32.th32ThreadID );
'      printf( "\n     base priority  = %d", te32.tpBasePri );
'      printf( "\n     delta priority = %d", te32.tpDeltaPri );
'    }
     End If
'  } while( Thread32Next(hThreadSnap, &te32 ) );
   Loop While Thread32Next(hThreadSnap, te32) <> 0
'
'//  Don't forget to clean up the snapshot object.
'  CloseHandle( hThreadSnap );
   Call CloseHandle(hThreadSnap)
'  return( TRUE );
   ListProcessThreads = True
'}
End Function
'
'void printError(TCHAR * msg)
Private Function printError(msg As String) As String
'{
'  DWORD eNum;
   Dim eeNum As Long
'  TCHAR sysMsg[256];
   Dim sysMsg As String * 256
'  TCHAR* p;
   Dim p As Long
'
'  eNum = GetLastError( );
   eeNum = Err.LastDllError
'  FormatMessage( FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
'         NULL, eNum,
'         MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
'         sysMsg, 256, NULL );
'
'  // Trim the end of the line and terminate it with a null
'  p = sysMsg;
'  while( ( *p > 31 ) || ( *p == 9 ) )
'    ++p;
'  do { *p-- = 0; } while( ( p >= sysMsg ) &&
'                          ( ( *p == '.' ) || ( *p < 33 ) ) );
'
'  // Display the message
'  printf( "\n  WARNING: %s failed with error %d (%s)", msg, eNum, sysMsg );
   printError = "WARNING: " & msg & " failed with error " & eeNum & " " & sysMsg
'}
'
End Function


