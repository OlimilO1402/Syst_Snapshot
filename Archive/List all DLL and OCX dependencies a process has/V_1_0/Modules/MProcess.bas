Attribute VB_Name = "MProcess"
Option Explicit
Private Const SYNCHRONIZE       As Long = &H100000

Private Const PROCESS_TERMINATE As Long = &H1

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
        
Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
        
Private Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32.dll" ( _
    ByVal hProcess As Long, _
    ByRef lpExitCode As Long) As Long


Public Function TerminateProcessID(ByVal ProcessID As Long) As Boolean
    Dim hr As Long
    Dim hProcess As Long
    'hProcess = OpenProcess(0, 0, ProcessID)
    If hProcess = 0 Then
        hProcess = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, 0, ProcessID)
    End If
    If hProcess = 0 Then
        hProcess = OpenProcess(SYNCHRONIZE, 0, ProcessID)
    End If
    If hProcess = 0 Then
        hProcess = OpenProcess(0, 0, ProcessID)
    End If
    If hProcess = 0 Then
        hProcess = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, True, ProcessID)
    End If
    If hProcess = 0 Then
        hProcess = OpenProcess(SYNCHRONIZE, True, ProcessID)
    End If
    If hProcess = 0 Then
        hProcess = OpenProcess(0, True, ProcessID)
    End If
    
    If hProcess <> 0 Then
        Dim exitcode As Long
        hr = GetExitCodeProcess(hProcess, exitcode)
        If hr <> 0 Then
            hr = TerminateProcess(hProcess, exitcode)
            TerminateProcessID = (hr <> 0)
        Else
            hr = TerminateProcess(hProcess, PROCESS_TERMINATE)
            If hr = 0 Then hr = TerminateProcess(hProcess, 0)
            If hr = 0 Then hr = TerminateProcess(hProcess, 99)
            TerminateProcessID = (hr <> 0)
        End If
        Call CloseHandle(hProcess)
    Else
        MsgBox "Process-Handle leider 0"
    End If
End Function





'program ProcessTimesNoHandle;
'{$APPTYPE CONSOLE}
'uses
'  Windows,
'  SysUtils,
'  JwaNtStatus,
'  JwaWinType,
'  JwaNative;
'
'type
'  TCallBackProcess = function(ps: PSYSTEM_PROCESSES; dwUserData: DWORD): BOOL; stdcall;
'
'  PProcessTimeRecord = ^TProcessTimeRecord;
'  TProcessTimeRecord = record
'    PID: DWORD;
'    CreationTime,
'      KernelTime,
'      UserTime: LARGE_INTEGER;
'  end;
'
'function ListProcesses(Callback: TCallBackProcess; dwUserData: DWORD): Boolean;
'Var
'  Status: NTSTATUS;
'  Buffer: PVOID;
'  TempBuf: PSYSTEM_PROCESSES;
'  BufLen: ULONG;
'const
'  MinQuerySize = $10000;
'begin
'  Result := False;
'  BufLen := MinQuerySize;
'  Buffer := RtlAllocateHeap(NtpGetProcessHeap(), HEAP_ZERO_MEMORY, BufLen);
'  If (Assigned(Buffer)) Then
'  Try
'    Status := NtQuerySystemInformation(
'      SystemProcessesAndThreadsInformation,
'      Buffer,
'      BufLen,
'      nil);
'    while (Status = STATUS_INFO_LENGTH_MISMATCH) do
'    begin
'      // Double the size to allocate
'      BufLen := BufLen * 2;
'      TempBuf := RtlReAllocateHeap(NtpGetProcessHeap(), HEAP_ZERO_MEMORY, Buffer, BufLen);
'      If (Not Assigned(TempBuf)) Then
'        Exit; // And free "Buffer" inside finally clause
'      // Else assign the TempBuf to Buffer
'      Buffer := TempBuf;
'      // Try to query info again
'      Status := NtQuerySystemInformation(
'        SystemProcessesAndThreadsInformation,
'        Buffer,
'        BufLen,
'        nil);
'    end;
'    // TempBuf used for pointer arithmetics
'    TempBuf := Buffer;
'    If (NT_SUCCESS(Status)) Then
'    begin
'      while (True) do
'      begin
'        If (Assigned(Callback)) Then
'          If (Not Callback(TempBuf, dwUserData)) Then
'          // Exit loop if the callback signalled to do so.
'            Break;
'        // Break if there is no next entry
'        If (TempBuf ^ .NextEntryDelta = 0) Then
'          Break;
'        // Else go to next entry in list
'        TempBuf := PSYSTEM_PROCESSES(DWORD(TempBuf) + TempBuf^.NextEntryDelta);
'      end;
'      Result := True;
'    end;
'  finally
'    If (Assigned(Buffer)) Then
'      RtlFreeHeap(NtpGetProcessHeap(), 0, Buffer);
'  end;
'end;
'
'// This MUST NOT be a local function
'
'function CallBackProcess(ps: PSYSTEM_PROCESSES; ProcessTimeRecord: PProcessTimeRecord): BOOL; stdcall;
'begin
'  Result := True;
'  If (Assigned(ps)) Then
'    If (ps ^ .ProcessID = ProcessTimeRecord ^ .PID) Then
'    begin
'      ProcessTimeRecord^.CreationTime := ps^.CreateTime;
'      ProcessTimeRecord^.KernelTime := ps^.KernelTime;
'      ProcessTimeRecord^.UserTime := ps^.UserTime;
'      // FIXME: This is for debugging only. Of course not needed in production code
'      Writeln('PID = ', ps^.ProcessId, ' - parent = ', ps^.InheritedFromProcessId);
'      // Stop going through the list
'      Result := False;
'    end;
'end;
'
'// Instead of only taking the times, it would be easier and more effective to
'// take all information directly from the SYSTEM_PROCESS structures in the
'// callback!
'
'function GetProcessTimesByPid(
'  PID: DWORD;
'  var lpCreationTime: Windows.FILETIME;
'  var lpKernelTime: Windows.FILETIME;
'  Var lpUserTime: Windows.FILETIME
'  ): BOOL; stdcall;
'Var
'  times: TProcessTimeRecord;
'begin
'  times.PID := PID; // PID to search for
'  // We need to pass a pointer here!
'  Result := ListProcesses(@CallBackProcess, DWORD(@times));
'  lpCreationTime := Windows.FILETIME(times.CreationTime);
'  lpKernelTime := Windows.FILETIME(times.KernelTime);
'  lpUserTime := Windows.FILETIME(times.UserTime);
'end;
'
'Var
'  lpCreationTime,
'    lpKernelTime,
'    lpUserTime: Windows.FILETIME;
'  cst: SYSTEMTIME;
'begin
'  // Hardcoded PID for testing. This should be called for each PID found
'  If (GetProcessTimesByPid(2588, lpCreationTime, lpKernelTime, lpUserTime)) Then
'  begin
'    FileTimeToSystemTime(lpCreationTime, cst);
'    Writeln(Format('Process created: %.4d-%.2d-%.2dT%.2d:%.2d:%.2d.%.3d', [cst.wYear, cst.wMonth, cst.wDay, cst.wHour, cst.wMinute, cst.wSecond, cst.wMilliseconds]));
'  end;
'  Readln;
'end.
