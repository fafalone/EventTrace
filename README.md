# VBEventTrace v2.1/TBEventTrace v2.2.4

![Screenshot](https://i.imgur.com/8F2HYde.jpg)

Event Tracing fo Windows (ETW) File Activity Monitor, VB6/twinBASIC x64 port

**Update (2023 Feb 17)**: Removed temporary fix as erase bug has been fixed. Corrected sign issues in some hex literals. 

**Update (2022 Dec 07)**: Applied temporary fix for twinBASIC bug that results in the ListView being erased on resize. Also, x64 binary no longer has a twinBASIC banner on startup as I'm now proudly supporting the project with a subscription :)

[Event Tracing for Windows (ETW)](https://docs.microsoft.com/en-us/windows/win32/etw/about-event-tracing) is a notoriously complex and unfriendly API, but it's extremely powerful. It allows access to messages from the NT Kernel Logger, which provides a profound level of detail about activity on the system. It provides details about many types of activity, but this first project will focus on File Activity. I also plan to follow this up with a monitor for TcpIp and Udp connections.

Given the complexity and unfriendliness that's given it the reputation of the world's worst API, why use it? You can find many projects that monitor file activity, using methods like SHChangeNotify, FindFirstChangeNotification, and monitoring open handles. But the reality is these are all high level methods that don't cover quite a bit of activity. The kernel logger shows activity coming from low level disk and file system drivers. This project started with me wanting to know what was causing idle hard drives to spin up, and none of the higher levels methods offered a clue. Programs like ProcessHacker and FileActivityView use the NT Kernel Logger as well, but I wanted two things: Better control over the process, and doing it in VB6. Why? Well, if you've seen my other projects, you know I'm excessively fond of going way beyond what VB6 was meant for both in terms of low level stuff and modern stuff.

### Intro

This project tracks most of the [DiskIO events](https://docs.microsoft.com/en-us/windows/win32/etw/diskio) and ([FileIo events](https://docs.microsoft.com/en-us/windows/win32/etw/fileio), providing a great deal of control over what events you watch and filtering them to find what you're looking for. It also looks up name and icon of the process that generated the activity (not always available). With no filtering or only light filtering, a tremendous amount of data is generated. The VB TextBox and ListView simply could not keep up with the rapid input, and all sorts of memory and display issues ensued where text and List Items disappeared. So while the project was already complicated to begin with, the only way to cope with this was to use an API-created Virtual ListView (created via API and using the LVS_OWNERDATA style so it only includes the data currently being displayed). The default mode looks only at DiskIO events, which represent physical disk activity, many FileIO events are only on in-memory caches. It's advisable to select 'Supplement', which uses info from FileIO events to correlate process attribution information, which is often missing from events.

This repository includes 2 versions: 

The original VB6 version: VB6 is x86 only, so this version was manually adjusted to operate with an x64 kernel. 

The [twinBASIC](https://github.com/twinbasic/twinbasic) version: can target either, but a 32bit build won't work on x64. twinBASIC is a successor to VB6 that's 100% backwards compatible as a goal, and brings in x64 compilation and new language features. It's 99% backwards compatible in language now, which lets this project run with only changing VB6-specific assembly thunks to regular subclassing. GUI and objects have a bit to go, but this project doesn't use much that's not available... see [this thread on VBForums](https://www.vbforums.com/showthread.php?897148-twinBASIC-x64-compatible-port-of-Event-Tracing-for-Windows-File-Activity-Monitor) for more info. Version 2.2.2 requires [twinBASIC Beta 122 or newer](https://github.com/twinbasic/twinbasic/releases) to open and build the source code.

### How It Works

[Have a read here](https://caseymuratori.com/blog_0025) for an introduction to setting up a Kernel Logger with ETW, and then realize it's even *more* complicated than that article suggests, because of some VB6 specific issues, and the hell on earth involved in interpreting the data.

Just starting the tracing session has 3 steps. You start with the EVENT_TRACE_PROPERTIES structure. Now, it's daunting enough on it's own. But when you read the article linked, you realize you have to have open bytes appended *after* the structure for Windows to copy the name into. Then the article doesn't touch on a recurring theme that was the source of a massive headache implementing it... in other languages, some ETW structures get automatically aligned along 8 byte intervals (a Byte is 1 byte, an Integer 2 bytes, a Long 4 bytes... alignment is making each of the largest type appear at, and the total size be, a multiple of its size). Not so in VB-- because of an arcane detail of alignment affecting LARGE_INTEGER: declaring it as two Longs is the standard definition, and indeed seems to *mostly* match the C/C++ def, except that mentions it's a union with a single 8 byte type. This means it triggers 8-byte alignment even if you don't use the QuadPart. Currency wouldn't help here, because VB masks it being a 2x4 byte UDT under the hood. The only true 8 byte type is Double, but that can't be easily substituted for ULONGLONG. It took quite a bit of crashing and failures to realize this, then properly pad the structures. The twinBASIC version of this project uses its LongLong type to remove the need for manual padding. The code uses it's own structure for the StartTrace function that looks like this:

```
Public Type EtpKernelTrace
    tProp As EVENT_TRACE_PROPERTIES
    padding(0 To 3) As Byte
    LoggerName(0 To 31) As Byte 'LenB(KERNEL_LOGGER_NAMEW)
    padding2(0 To 3) As Byte
End Type
```

Needed to include 4 bytes of padding after the structure, then add room for the name, then make sure it's all aligned to 8 byte intervals. In the VB6 version, this is done manually because it's targeting x64 Windows... these structures are passed directly to kernel mode WMI modules without being translated through the WOW64 layer. In the twinBASIC version, it's all handled by the compiler depending on whether you build for x86 or x64 Now we're ready to go, with tStruct being a module-level EtpKernelTrace var:

```
With tStruct.tProp
    .Wnode.Flags = WNODE_FLAG_TRACED_GUID
    .Wnode.ClientContext = 1&
    .Wnode.tGUID = SelectedGuid
    .Wnode.BufferSize = LenB(tStruct)
    .LogFileMode = EVENT_TRACE_REAL_TIME_MODE 'We're interested in doing real time monitoring, as opposed to processing a .etl file.
    If bUseNewLogMode Then
        .LogFileMode = .LogFileMode Or EVENT_TRACE_SYSTEM_LOGGER_MODE
    End If
    'The enable flags tell the system which classes of events we want to receive data for.
    .EnableFlags = EVENT_TRACE_FLAG_DISK_IO Or EVENT_TRACE_FLAG_DISK_FILE_IO Or EVENT_TRACE_FLAG_FILE_IO_INIT Or _
                    EVENT_TRACE_FLAG_DISK_IO_INIT Or EVENT_TRACE_FLAG_FILE_IO Or EVENT_TRACE_FLAG_NO_SYSCONFIG
    .FlushTimer = 1&
    .LogFileNameOffset = 0&
    .LoggerNameOffset = LenB(tStruct.tProp) + 4 'The logger name gets appended after the structure; but the system looks in 8 byte alignments,
                                                'so because of our padding, we tell it to start after an additional 4 bytes.
End With

'We're now ready to *begin* to start the trace. StartTrace is only 1/3rd of the way there...
hr = StartTraceW(gTraceHandle, StrPtr(SelectedName & vbNullChar), tStruct)
```

This begins to start a trace session. There's SelectedGuid and SelectedName because there's two options here. In Windows 7 and earlier, the name has to be "NT Kernel Logger", and the Guid has to be SystemTraceControlGuid. If you use that method, there can only be 1 such logger running. You have to stop other apps to run yours, and other apps will stop yours when you start them. On Windows 8 and newer, there can be several such loggers, and you supply a custom name and GUID, and inform it you want a kernel logger with the flag added with bUseNewLogMode. This project supports both methods. The EnableFlags are the event providers you want enabled. This project wants the disk and file io ones, but there's many others. Onto step 2...


```
Dim tLogfile As EVENT_TRACE_LOGFILEW
ZeroMemory tLogfile, LenB(tLogfile)
tLogfile.LoggerName = StrPtr(SelectedName & vbNullChar)
tLogfile.Mode = PROCESS_TRACE_MODE_REAL_TIME Or PROCESS_TRACE_MODE_EVENT_RECORD 'Prior to Windows Vista, EventRecordCallback wasn't available.
tLogfile.EventCallback = FARPROC(AddressOf EventRecordCallback) 'Further down, you can see the prototype for EventCallback for the older version.
gSessionHandle = OpenTraceW(tLogfile)
```

We have to tell it *again* we want to use real time mode, not a .etl log file, and at this point we supply a pointer to a callback that receives events. This project uses a newer type of callback available in Vista+, but has prototypes for the older one. Like a WndProc for subclassing, this has to be in a standard module (.bas); to put it in a class module/form/usercontrol, you'd need the kind of self-subclassing code like you find on the main form (but be careful copying/pasting that, it's been slightly modified and only works with Forms).

The final step is a single call: To `ProcessTrace`. Only then will you begin receiving events. But of course, this simple call couldn't be simple. ProcessTrace doesn't return until all messages have been processed, which in a real-time trace means indefinitely until you shut it off. So if you call it, execution stops. In that thread. In other languages, spinning off a new thread to call ProcessTrace is easy. In VB, it's painful. The VB6 version project makes use of The trick's VbTrickThreading project to launch a new thread for the ProcessTrace call. The downside here is that means event tracing is only possible in a compiled exe, making debugging difficult. The twinBASIC version is able to simply call `CreateThread` directly, as it natively supports multithreading (only via API like this for now, but eventually by language features).

Once you've called ProcessTrace, your callback begins receiving messages. We need to match them up with their provider, and then check the OpCode...

```
Public Sub EventRecordCallback(EventRecord As EVENT_RECORD)
'...
If IsEqualIID(EventRecord.EventHeader.ProviderId, DiskIoGuid) Then
    iCode = CLng(EventRecord.EventHeader.EventDescriptor.OpCode)
    
    'Some events use the same MOF structure and are processed similarly, so we group them together and separate
    'the codes for filtering and logging later.
    If (iCode = EVENT_TRACE_TYPE_IO_READ) Or (iCode = EVENT_TRACE_TYPE_IO_WRITE) Then
 ```
 
The EVENT_RECORD structure is also a nightmare. Many different parts of it had to having alignment padding added, and it tripped me up for a good long while. Extra thanks to The trick for helping me figure out the right alignment on this part.

From here, we're ready to process the data. The raw data is returned in MOF structures, e.g. this one for one of the Open/Create messages. There's ways to automate the processing of them, but that makes everything so far seem simple, and is the domain for a future project. For now, we manually process the raw data, which we copy from the pointer in .UserData in the event record. The documentation doesn't mention *at all* that even if you're running a 32bit application, these structures have 64bit sizes on 64bit Windows. The official documentation doesn't note which "uint32" types are pointers, and thus are 8 bytes instead of 4, so I had to go digging in some deep system files. The original 32bit structures are all included, but currently this project only works on 64bit Windows. It's possible to tell automatically via flags in the event record... perhaps in the future. EventRecord.EventHeader.Flags has flags EVENT_HEADER_FLAG_[32,64]_BIT_HEADER.

Here what the File Open/Create structure looks like, and how we set it up:

```
Public Type FileIo_Create64 'Event IDs: 64
    IrpPtr As Currency
    FileObject As Currency
    ttid As Long
    CreateOptions As CreateOpts
    FileAttributes As FILE_ATTRIBUTES
    ShareAccess As Long
    OpenPath(MAX_PATH) As Integer
End Type
```

The tB version uses the native LongPtr datatype instead of Currency (or where it's always 8 bytes, LongLong).

For VB6, fortunately VB has the Currency data type, which we also used for our event trace handles, which is 8 bytes. We can use this because there's no point where we have to interact a numeric representation of the value... it's just all raw bytes behind the scenes. Unfortunately, FileAttributes is only what's passed to the NtOpenFile API and not an actual query of the file's attributes, so is almost always 0 or FILE_ATTRIBUTES_NORMAL. We pick MAX_PATH for the size of the array, because using a fixed-size array avoids VB's internal SAFEARRAY type, which would make copying a structure from a language without it much more complicated. Converting a string of integer's to a normal string is trivial, but the real problems comes when you see what it is: files names look like \Device\HarddiskVolume1\folder\file.exe. To convert those into normal Win32 paths the project creates a map by querying each possible drive letter in the QueryDosDevice API, which returns a path like that for each drive.

Not all events contain a file name, so the project stores a record with the FileObject, which allows us to match other operations on the same file, and get the name. The documentation says we're supposed to receive event code 0 for names... but I've never seen that message come in. Perhaps on earlier Windows versions.

Perhaps the biggest problem in processing the data is that while there's an ProcessID and ThreadID in the event record's header, the process id is very often -1. Sometimes that information is returned in other events. This project goes through incredible lengths to correlate every with every other event in order to track down the process whenever possible. So many events will display -1 at first, and get updated later.

There's still a lot of work to be done in process attribution, and getting info about files already open before the trace starts. I attempted to copy ProcessHacker's use of a KernelRundownLogger, but so far have not been successful. I'll be look at other methods, but if I didn't put out a Version 1, who knows how long it would be.

Once we've captured the events, we store it in a the ActivityLog structure, which is the master data store for what's displayed on the ListView. 

### Options

You can see in the screenshot a number of options. There's the main controls for the trace; you don't really need to worry about 'Flush', it's there for completeness and shouldn't be needed. Stop is always enabled because in the event of crashes, you can stop previous sessions. You can save the trace; it saves what you see in the ListView, tab separated. There's options for which events you want to capture, whether to use the new logger method described earlier (Win8+), and the refresh interval for the ListView. The items aren't added to the ListView; they're stored in the ActivityLog structure, and the ListView is in virtual mode, so it only asks for what it's currently displaying. The refresh interval is how often it checks for new events and sets the last one as visible, creating a view that is always scrolled to the bottom but without the invisible items stored in the ListView itself, dramatically improving speed. (The greyed out option is for future work, not currently implemented)

Very important is the filtering system, if you're looking for certain activity. Each field allows multiple entries separated with a | (bar, it also accepts broken bars found on some keyboards). There's a button that displays a message explaining the syntax and the flow... the first thing checked is whether it's from a process we're interested in based on the process options. You can use DOS wildcards in the Process name field and File name fields, but not the paths at this point... for now the paths are strictly checked on a 'Starts with...' basis. After checking the process, then it checks 'Path must match', then 'Exclude paths', then 'File name must match', finally 'Exclude file name'.

Finally on the right there's a message log, which displays information about starting/stopping the trace, when a different function has correlated a previously unidentified process id, and any errors that arise.

Not shown: If you right click the ListView, there's a popup menu with options to open the selected items, show the selected items in Explorer, copy selected file names, copy all file names, copy the selected lines (tab separated), copy all lines, show properties of the process, and show the process in Explorer. 

### Requirements

PLEASE TAKE NOTE. This program has atypical requirements.

-VB6 version only: Windows Vista or newer 64bit. Although like all VB6 apps the app itself is 32bit, it handles data structures generated by the system, and is currently only coded to handle 64bit structures. To run on 32bit Windows, use the regular MOF structures instead of the x64 ones (and change the size checks at the start of each processing routine).

-twinBASIC version only: Windows Vista or newer; a 32bit build can only run on 32bit Windows versions. **UPDATED:** Version 2.2.2 requires [twinBASIC Beta 122 or newer](https://github.com/twinbasic/twinbasic/releases) to open and build from source.

-VB6 version only: This program can only start event tracing when compiled, due to the need for multithreading that cannot be done in a single thread.

-The NT Kernel Logger requires additional permissions- you need to be a member of the Administrators group (but not necessarily run as admin unless you got the Windows update that changed this), or be a member of the Performance Log Users group, or otherwise have permission to enable the SeSystemProfilePrivilege.

-There are no external dependencies. However, the demo uses a manifest for Common Controls 6.0 styles, and it's advised you also use them in any other project.

-VB6 version only: Unicode is supported in the ListView for displaying files etc, but the filter TextBoxes are just regular VB ones, so you'd need to replace those to use Unicode in filtering. twinBASIC: Unicode is supported in all areas.

Windows 10 is strongly recommended. I have not had the opportunity to test this on other OSs.

This API is *extremely* complicated and finicky, so there's bound to be bugs. Especially on other Windows versions. Let me know, I'll see what I can do.

Changlog:

```
'Applies to twinBASIC version only:
'Version 2.2.3 -Temporary workaround for twinBASIC API-control-erasure bug.
'
'Version 2.2.2 -Subclassing moved back to form and PictureBox now used for
'               'More Options' popup, taking advantage of new tB features to
'               make this more similar to the VB version.
'              -Bug fix: Default process cache option not selected; this was a
'                        visual glitch only, the logic still applied the right
'                        default if you didn't set it.
'
'Version 2.2 - Added option to display full process command line. (tB only)
'              -Bug fix: SimpleOp incorrect buffer size error (MOF has no
'               packing in data; tB inserted 4 extra bytes).
'              -Bug fix: Removed some incorrect padding; some not needed on x64
'                        and some not needed on x86. In fact, no manual
'                        alignment is neccessary at all in this version, 
'                        (besides removing it from MOF); it was just an 
'                        artifact of these structs going straight to the kernel,
'                        so manual x64 alignment is needed for the VB WOW64
'                        version (or if you convert this back, add it back in).
'
'Version 2.11 -Bug fix: Process Name did not support Unicode.
'             -Bug fix: ProcessPrebuildFullCache prematurely exited loop.
'
'Both versions:
'Version 2.1: -If you scroll away from the bottom, it will stay where you scroll
'              to instead of immediately going back to the bottom.
'
'             -Bug fix: FileIo_DirEnum pattern and InfoClass were reported wrong.
'
'Version 2.0: -Running as administrator is no longer strictly required. However,
'              you must still be a member of the administators group, a member
'              of the Performance Log Users group, or another user/group that has
'              access to the SeSystemProfilePrivilege, which the code now calls
'              AdjustTokenPrivileges to set.
'
'             -Project has been optimized to track disk read/write alone. Because
'              of caching, FileIo events don't neccessarily trigger disk activity,
'              so if you only wanted to watch disk activity, disabling everything
'              except DiskIO will enable DiskIO Exclusive Mode; FileIo events not
'              activated by the DiskIO flags will be disabled (and can't be enabled
'              while the trace is in progress, since this is done in the inital
'              flags), and DiskIO will properly attribute IO to Create events,
'              where when disabled, create events use FileIO_ReadWrite, which
'              doesn't reflect disk io because of caching.
'              DiskIO includes open/delete events that trigger disk activity.
'
'             -Added option to merge certain activities by the same process on the
'              same file. Create/open/read/write/delete can be merged. Multiple
'              events may still exist where an initial pid was -1 and then updated.
'              An additional option restricts this to same opcode only (combining
'              only read/write opcodes).
'
'             -The default configuration is now DiskIO Exclusive Mode with merging
'              enabled. Remember that enabling any FileIO operation takes it out
'              of that mode, and that FileIO operations do not represent disk read
'              and write activity due to caching.
'
'             -There's now an option to disable the initial rundown; you'll miss
'              activity on open files during the trace, but idle disks won't spin
'              up. If this option is used, the rundown may be performed at the end
'              of the trace (always, if old logger mode, optionally with new), and
'              the disk io logs will be scanned for missed events, and added then.
'
'             -Process caching now has an option to be disabled (never, always, or
'              only when running in DiskIO Exclusive Mode), for situations where
'              process id use may be a problem (processes rapidly being created
'              and exiting).
'
'             -Context switch tracking for process attribution is now implemented,
'              however I've found it entirely useless. It only works in cases where
'              the pid was already returned.
'
'             -Log sync interval can now be adjusted.
'
'             -Changed how filenames are read to no longer use a fixed buffer;
'              the MOF structure now uses String (variable length) and is set
'              by a Fill_<type> routine in modEventTrace that first copies the
'              data into a variable length byte array, then sets the MOF type
'              String. This allows long file name support without using a
'              massive fixed buffer.
'
'             -There's now an event for unattributed FileIo_ReadWrite events like
'              there is for DiskIo_ReadWrite, to catch read/writes on files already
'              open before the trace started.
'
'             -So few events were in the type enum I got rid of the enum and made
'              them byte constants to improve performance. Also added constants
'              for the FileIo opcodes.
'
'             -Added some more flags in the Misc column.
'
'             -Cleaned up some lingering potentially unsafe thread data accesses.
'
'             -There's reports of poor critical section performance on Win8+ with
'              the default dynamic spin count adjustment; trying 4000 instead per
'              MSDN recommendation as optimal; will see how it goes.
'
'             -Bug fix: AddActivity didn't add read/write size, so it was only ever
'                       non-zero if an update came in, which was generally never for
'                       FileIo since MSDN lied about FileKey being what you use to
'                       correlate FileIo_ReadWrite events.
'
'             -Bug fix: While read/write totals were updated, sync between the
'                       ActivityLog struct and threadmain copy consisted only of
'                       adding new items, so only updates that occured between sync
'                       calls were ever updated on the ListView.
'                       This also impacted pid updates.
'
'Version 1.2: -Now able to log read/write to files open before the trace started
'              by triggering a rundown. This comes up as 'DiskIO', which can
'              also be filtered. If you disable all events except DiskIO, the
'              list should be similar to what ProcessHacker does.
'
'             -Numbers only enforced on refresh interval textbox; can now update
'              while running by pressing enter. (Note: All other checkboxes except
'              'Use new logger' take effect immediately while running too, and
'              the 'Update' button on filters is for when it's running.)
'
'             -Fixed a number of small bugs.
'
'             -Added additional declares and corrected enums for modEventTrace.
'
'Version 1.1: -Fixed bug where old logger option didn't worked, sped up code
'              significantly by using a different wchar->string method and
'              alternative to VB's Replace() function, and added thread safety
'              to prevent crashes from too high an incoming event rate.
'
'             -There's now a separate activity log that the ListView displays
'              that's synchronized to the primary one inside a critical section
'              so the same memory isn't accessed by both threads at the same time,
'              which caused crashes in certain circumstances, on slower systems,
'              or very high event rates (such as the rundown with the old logger).
```

