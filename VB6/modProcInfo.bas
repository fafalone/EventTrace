Attribute VB_Name = "modProcInfo"
Option Explicit

'*******************************************************************************************
'modProcInfo  - Process information from pid.
'
'(revision 2, 2022/04/16) (c)2022 fafalone
'
'This is designed for high performance random bulk lookups of process names/paths
'given their pid or a thread id from them. It maintains a cache of previous lookups
'that is checked first.
'
'There are 5 public methods:
'
'     GetProcessInfoFromPID - Given a process id, sets arguments pointing to the
'                             executable name and the full path. Returns True if successful.
'
'
'  PrebuildFullProcessCache - To be called at loading or before heavy processing starts, this
'                             caches all currently running processes, so only new ones need
'                             to be looked up.
'
'           EnsurePidCached - If you know you'll need info about a pid, but not yet, cache it.
'
'    InvalidateProcessCache - Clears the pid info cache.
'
'          GetPidFromTidCTS - An extra utility function to get a process id from a thread id,
'                             in case one needed to support XP or wanted an alternative if
'                             GetProcessIdOfThread failed.
'
'WARNING: Thread IDs and Process IDs are reused. You cannot rely on them forever. This
'         is currently designed for short term use; in the future, there will be a mechanism
'         to add/remove processes and threads based on those events.
Public ProcInfoNoCache As Boolean 'Set this to True to disable caching if you're having
                                  'issues with pid reuse as a temporary mitigation.
'*******************************************************************************************

Private Type CachedProcess
    pid As Long
    ProgName As String
    FullPath As String
    iIcon As Long
End Type
Private ProcessCache() As CachedProcess
Private nCached As Long

 

Private Enum TH32CS_Flags
    TH32CS_SNAPHEAPLIST = &H1
    TH32CS_SNAPPROCESS = &H2
    TH32CS_SNAPTHREAD = &H4
    TH32CS_SNAPMODULE = &H8
    TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
    TH32CS_INHERIT = &H80000000
End Enum

Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type


Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE As Long = &H100000

Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF&)
Private Const PROCESS_TERMINATE = &H1 ' Enables using the process handle in the TerminateProcess function to terminate the process.
Private Const PROCESS_CREATE_THREAD = &H2   ' Enables using the process handle in the CreateRemoteThread function to create a thread in the process.
Private Const PROCESS_VM_OPERATION = &H8 ' Enables using the process handle in the VirtualProtectEx and WriteProcessMemory functions to modify the virtual memory of the process.
Private Const PROCESS_VM_READ = &H10     ' Enables using the process handle in the ReadProcessMemory function to read from the virtual memory of the process.
Private Const PROCESS_VM_WRITE = &H20 ' Enables using the process handle in the WriteProcessMemory function to write to the virtual memory of the process.
Private Const PROCESS_DUP_HANDLE = &H40   ' Enables using the process handle as either the source or target process in the DuplicateHandle function to duplicate a handle
Private Const PROCESS_SET_INFORMATION = &H200 ' Enables using the process handle in the SetPriorityClass function to set the priority class of the process.
Private Const PROCESS_QUERY_INFORMATION = &H400 ' Enables using the process handle in the GetExitCodeProcess and GetPriorityClass functions to read information from the process object.
Private Const PROCESS_QUERY_LIMITED_INFORMATION = &H1000

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As TH32CS_Flags, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, lpTE As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, lpTE As THREADENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function QueryFullProcessImageName Lib "kernel32" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As Long, lpdwSize As Long) As Long

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_flags) As Long
Private Enum SHGFI_flags
  SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
  SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
  ' Indicates that the function should not attempt to access the file specified by pszPath.
  ' Rather, it should act as if the file specified by pszPath exists with the file attributes
  ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
  ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
  SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
  SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
  SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                        ' containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
  SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
End Enum
Private Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type
'**************************
'Public functions

Public Sub PrebuildFullProcessCache()
'To optionally be called when loading, caches all currently running processes.
Dim tProcess As PROCESSENTRY32
Dim hr As Long
Dim Retry As Boolean
Dim hSnapshot As Long

hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

If hSnapshot Then
    tProcess.dwSize = Len(tProcess)
    hr = Process32First(hSnapshot, tProcess)
    If hr > 0& Then
        Do While hr > 0&
            ReDim Preserve ProcessCache(nCached)
            ProcessCache(nCached).ProgName = LCase$(Left$(tProcess.szExeFile, IIf(InStr(1, tProcess.szExeFile, Chr$(0)) > 0, InStr(1, tProcess.szExeFile, Chr$(0)) - 1, 0)))
            ProcessCache(nCached).ProgName = ProcessCache(nCached).ProgName & ":" & CStr(tProcess.th32ProcessID)
            hr = GetProcessFullPath(tProcess.th32ProcessID, ProcessCache(nCached).FullPath)
            If hr Then
                'Error
                ProcessCache(nCached).FullPath = "0x" & Hex$(hr)
            End If
            ProcessCache(nCached).pid = tProcess.th32ProcessID
            ProcessCache(nCached).iIcon = GetFileIconIndex(ProcessCache(nCached).FullPath, SHGFI_SMALLICON)
            If ProcessCache(nCached).iIcon < 2& Then ProcessCache(nCached).iIcon = 2 'Default program icon
            nCached = nCached + 1
            CloseHandle hSnapshot
            Exit Sub
            hr = Process32Next(hSnapshot, tProcess)
        Loop
    Else
        PostLog "Error calling Process32First"
    End If
Else
    PostLog "Error creating process snapshot."
End If
End Sub

Public Sub InvalidateProcessCache()
ReDim ProcessCache(0)
nCached = 0&
End Sub

Public Sub EnsurePidCached(ByVal pid As Long)
If pid = -1& Then Exit Sub
Dim i As Long
If nCached Then
    For i = 0 To UBound(ProcessCache)
        If ProcessCache(i).pid = pid Then Exit Sub
    Next i
    LoadPID pid
End If
End Sub

Public Function GetProcessInfoFromPID(ByVal pid As Long, lpName As String, lpFullPath As String, lpIcon As Long) As Boolean
On Error GoTo e0
If ProcInfoNoCache Then
    LoadPIDDirect pid, lpName, lpFullPath, lpIcon
    Exit Function
Else
    If nCached = 0 Then
        If LoadPID(pid) Then
            lpName = ProcessCache(0).ProgName
            lpFullPath = ProcessCache(0).FullPath
            lpIcon = ProcessCache(0).iIcon
            GetProcessInfoFromPID = True
        End If
        Exit Function
    End If
    Dim i As Long
    For i = 0 To nCached - 1
        If ProcessCache(i).pid = pid Then
            lpName = ProcessCache(i).ProgName
            lpFullPath = ProcessCache(i).FullPath
            lpIcon = ProcessCache(i).iIcon
            GetProcessInfoFromPID = True
            Exit Function
        End If
    Next i
    'Not cached, try to load
    If LoadPID(pid) Then
        lpName = ProcessCache(UBound(ProcessCache)).ProgName
        lpFullPath = ProcessCache(UBound(ProcessCache)).FullPath
        lpIcon = ProcessCache(UBound(ProcessCache)).iIcon
        GetProcessInfoFromPID = True
    End If
    Exit Function
End If
e0:
PostLog "GetProcessInfoFromPID.Error->" & GetErrorName(Err.Number)
End Function

Public Function GetPidFromTidCTS(tid As Long) As Long
'Alternative for XP and possible access issues
GetPidFromTidCTS = -1&
Dim tThread As THREADENTRY32
Dim hr As Long
Dim hSnapshot As Long

hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0&)

If hSnapshot Then
    tThread.dwSize = LenB(tThread)
    hr = Thread32First(hSnapshot, tThread)
    If hr > 0& Then
        Do While hr > 0&
            If tThread.th32ThreadID = tid Then
                GetPidFromTidCTS = tThread.th32OwnerProcessID
                CloseHandle hSnapshot
                Exit Function
            End If
        Loop
    End If
    CloseHandle hSnapshot
End If
End Function

Private Function LoadPID(pid As Long, Optional lTID As Long = 0&) As Boolean
On Error GoTo LoadPID_Err
Dim tProcess As PROCESSENTRY32
Dim hr As Long
Dim Retry As Boolean
Dim hSnapshot As Long

hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

If hSnapshot Then
    tProcess.dwSize = Len(tProcess)
    hr = Process32First(hSnapshot, tProcess)
    If hr > 0& Then
        Do While hr > 0&
            If tProcess.th32ProcessID = pid Then
                ReDim Preserve ProcessCache(nCached)
                ProcessCache(nCached).ProgName = LCase$(Left$(tProcess.szExeFile, IIf(InStr(1, tProcess.szExeFile, Chr$(0)) > 0, InStr(1, tProcess.szExeFile, Chr$(0)) - 1, 0)))
                ProcessCache(nCached).ProgName = ProcessCache(nCached).ProgName & ":" & CStr(pid)
                hr = GetProcessFullPath(pid, ProcessCache(nCached).FullPath)
                If hr Then
                    'Error
                    ProcessCache(nCached).FullPath = "0x" & Hex$(hr)
                End If
                ProcessCache(nCached).pid = pid
                ProcessCache(nCached).iIcon = GetFileIconIndex(ProcessCache(nCached).FullPath, SHGFI_SMALLICON)
                If ProcessCache(nCached).iIcon < 2& Then ProcessCache(nCached).iIcon = 2& 'Default program icon
                nCached = nCached + 1&
                LoadPID = True
                CloseHandle hSnapshot
                Exit Function
            End If
            hr = Process32Next(hSnapshot, tProcess)
        Loop
        'Process not found.
        ReDim Preserve ProcessCache(nCached)
        If pid = -1& Then
            ProcessCache(nCached).ProgName = "<unknown>:" & CStr(pid)
            ProcessCache(nCached).pid = pid
            ProcessCache(nCached).iIcon = 2& 'Default program icon
            nCached = nCached + 1&
        Else
            'We have a seemingly valid pid, try an alternative.
            Dim hProc As Long
            hr = GetProcessFullPath(pid, ProcessCache(nCached).FullPath)
            If hr Then
                'Error
                If pid <> -1& Then
                    PostLog "Non-error Pid not found in snapshot; error querying full path " & GetErrorName(hr)
                End If
                ProcessCache(nCached).ProgName = "<unknown>:" & CStr(pid)
                ProcessCache(nCached).pid = pid
                ProcessCache(nCached).iIcon = 2& 'Default program icon
                nCached = nCached + 1
               
            Else
                PostLog "Identified unknown process through GetProcessFullPath, pid=" & pid & ",path=" & ProcessCache(nCached).FullPath
                ProcessCache(nCached).ProgName = ProcessCache(nCached).FullPath & ":" & CStr(pid)
                ProcessCache(nCached).pid = pid
                ProcessCache(nCached).iIcon = 2& 'Default program icon
                nCached = nCached + 1
                
            End If
        
        End If
        LoadPID = True
        CloseHandle hSnapshot
        Exit Function
        
    Else
        PostLog "LoadPID->Process32First failed."
    End If
Else
    PostLog "LoadPID->Failed to create snapshot."
End If

CloseHandle hSnapshot
    
Exit Function

LoadPID_Err:
    PostLog "LoadPID.Error->" & Err.Description & ", 0x" & Hex$(Err.Number)
End Function

Private Function LoadPIDDirect(pid As Long, pName As String, pPath As String, pIcon As Long) As Boolean
'Gets pid info without caching
On Error GoTo LoadPID_Err
Dim tProcess As PROCESSENTRY32
Dim hr As Long
Dim Retry As Boolean
Dim hSnapshot As Long

hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

If hSnapshot Then
    tProcess.dwSize = Len(tProcess)
    hr = Process32First(hSnapshot, tProcess)
    If hr > 0& Then
        Do While hr > 0&
            If tProcess.th32ProcessID = pid Then
                pName = LCase$(Left$(tProcess.szExeFile, IIf(InStr(1, tProcess.szExeFile, Chr$(0)) > 0, InStr(1, tProcess.szExeFile, Chr$(0)) - 1, 0)))
                pName = pName & ":" & CStr(pid)
                hr = GetProcessFullPath(pid, pPath)
                If hr Then
                    'Error
                    pPath = "0x" & Hex$(hr)
                End If
                pIcon = GetFileIconIndex(pPath, SHGFI_SMALLICON)
                If pIcon < 2& Then pIcon = 2 'Default program icon
                LoadPIDDirect = True
                CloseHandle hSnapshot
                Exit Function
            End If
            hr = Process32Next(hSnapshot, tProcess)
        Loop
        'Process not found.
        If pid = -1& Then
            pName = "<unknown>:" & CStr(pid)
            pIcon = 2 'Default program icon
        Else
            'We have a seemingly valid pid, try an alternative.
            Dim hProc As Long
            hr = GetProcessFullPath(pid, pPath)
            If hr Then
                'Error
                If pid <> -1& Then
                    PostLog "Non-error Pid not found in snapshot; error querying full path " & GetErrorName(hr)
                End If
                pName = "<unknown>:" & CStr(pid)
                pIcon = 2 'Default program icon
            Else
                PostLog "Identified unknown process through GetProcessFullPath, pid=" & pid & ",path=" & pPath
                pName = pPath & ":" & CStr(pid)
                pIcon = 2 'Default program icon
            End If
        
        End If
        LoadPIDDirect = True
        CloseHandle hSnapshot
        Exit Function
        
    Else
        PostLog "LoadPID->Process32First failed."
    End If
Else
    PostLog "LoadPID->Failed to create snapshot."
End If

CloseHandle hSnapshot
    
Exit Function

LoadPID_Err:
    PostLog "LoadPIDDirect.Error->" & Err.Description & ", 0x" & Hex$(Err.Number)
End Function

Private Function GetProcessFullPath(pid As Long, pPath As String) As Long
On Error GoTo GetProcessFullPath_Err
Dim hProc As Long
Dim sBuf As String
Dim cb As Long
Dim lErr As Long
hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0&, pid)
lErr = Err.LastDllError
If hProc Then
    sBuf = String$(MAX_PATH_DOS, 0)
    Dim hr As Long
    cb = MAX_PATH_DOS
    hr = QueryFullProcessImageName(hProc, 0&, StrPtr(sBuf), cb)
    lErr = Err.LastDllError
    If hr Then
        pPath = Left$(sBuf, cb)
    Else
        GetProcessFullPath = lErr
    End If
Else
    GetProcessFullPath = lErr
End If

Exit Function

GetProcessFullPath_Err:
    PostLog "GetProcessFullPath.Error->" & Err.Description & ", 0x" & Hex$(Err.Number)
End Function

Public Function GetFileIconIndex(Path As String, uType As Long) As Long
  Dim sfi As SHFILEINFO
  Dim pidl As Long
  'pidl = PathToPidl(Path)
If SHGetFileInfo(Path, ByVal 0&, sfi, Len(sfi), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON Or SHGFI_TYPENAME) Then
    GetFileIconIndex = sfi.iIcon
  End If
End Function
