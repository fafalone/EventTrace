VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Windows Event Tracing: File Activity Monitor"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1076
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbOptions 
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   3075
      ScaleHeight     =   3885
      ScaleWidth      =   3555
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3165
      Visible         =   0   'False
      Width           =   3555
      Begin VB.Frame Frame4 
         Height          =   3495
         Left            =   60
         TabIndex        =   43
         Top             =   270
         Width           =   3375
         Begin VB.CheckBox Check20 
            Caption         =   "Show merge count columns"
            Height          =   225
            Left            =   315
            TabIndex        =   60
            Top             =   2655
            Value           =   1  'Checked
            Width           =   2805
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Done"
            Height          =   360
            Left            =   2310
            TabIndex        =   59
            Top             =   3090
            Width           =   990
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Use context switch event to find pid"
            Height          =   255
            Left            =   75
            TabIndex        =   54
            ToolTipText     =   $"Form1.frx":0000
            Top             =   450
            Width           =   2850
         End
         Begin VB.CheckBox Check16 
            Caption         =   "end."
            Height          =   225
            Left            =   2250
            TabIndex        =   53
            ToolTipText     =   "Only optional with new logger mode, and only done if no initial rundown was done."
            Top             =   180
            Width           =   645
         End
         Begin VB.CheckBox Check15 
            Caption         =   "start"
            Height          =   225
            Left            =   1590
            TabIndex        =   52
            ToolTipText     =   $"Form1.frx":00A0
            Top             =   180
            Value           =   1  'Checked
            Width           =   645
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   60
            TabIndex        =   51
            Text            =   "1000"
            ToolTipText     =   "Press enter to update while running"
            Top             =   720
            Width           =   525
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   60
            TabIndex        =   50
            Text            =   "2500"
            Top             =   1035
            Width           =   525
         End
         Begin VB.PictureBox pbPCOpt 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   90
            ScaleHeight     =   315
            ScaleWidth      =   3240
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1605
            Width           =   3240
            Begin VB.OptionButton Option2 
               Caption         =   "Never"
               Height          =   285
               Index           =   0
               Left            =   0
               TabIndex        =   49
               Top             =   30
               Value           =   -1  'True
               Width           =   765
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Always"
               Height          =   285
               Index           =   1
               Left            =   915
               TabIndex        =   48
               Top             =   30
               Width           =   795
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Disk Only Mode"
               Height          =   285
               Index           =   2
               Left            =   1860
               TabIndex        =   47
               Top             =   30
               Width           =   1350
            End
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Merge create/open/read/write/delete events on same file from same process"
            Height          =   405
            Left            =   90
            TabIndex        =   45
            Top             =   1950
            Value           =   1  'Checked
            Width           =   3135
         End
         Begin VB.CheckBox Check19 
            Caption         =   "Only for same opcode"
            Height          =   255
            Left            =   315
            TabIndex        =   44
            Top             =   2370
            Width           =   2115
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Perform rundown at "
            Height          =   195
            Left            =   60
            TabIndex        =   58
            Top             =   195
            Width           =   1485
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "List refresh interval (ms)"
            Height          =   255
            Left            =   660
            TabIndex        =   57
            Top             =   735
            Width           =   1755
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Log refresh interval (ms)"
            Height          =   195
            Left            =   660
            TabIndex        =   56
            Top             =   1095
            Width           =   1770
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disable process caching"
            Height          =   195
            Left            =   90
            TabIndex        =   55
            Top             =   1395
            Width           =   1695
         End
      End
      Begin VB.Label lblOpts 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "More Options"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   41
         Top             =   30
         Width           =   1830
      End
      Begin VB.Shape shpSearchBk 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   0
         Top             =   0
         Width           =   3510
      End
   End
   Begin VB.Timer tmrLog 
      Left            =   13245
      Top             =   3300
   End
   Begin VB.Frame Frame3 
      Caption         =   "Trace Control"
      Height          =   2700
      Left            =   45
      TabIndex        =   26
      Top             =   15
      Width           =   1725
      Begin VB.CommandButton Command1 
         Caption         =   "Start Trace"
         Default         =   -1  'True
         Height          =   450
         Left            =   90
         TabIndex        =   32
         Top             =   210
         Width           =   1560
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop Trace"
         Height          =   450
         Left            =   120
         TabIndex        =   31
         Top             =   975
         Width           =   1560
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Flush buffer"
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   1470
         Width           =   1560
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save trace as..."
         Height          =   300
         Left            =   150
         TabIndex        =   29
         Top             =   2250
         Width           =   1560
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Clear log"
         Height          =   225
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pause collection"
         Height          =   300
         Left            =   150
         TabIndex        =   27
         Top             =   1905
         Width           =   1560
      End
      Begin VB.Line Line1 
         X1              =   270
         X2              =   1500
         Y1              =   1830
         Y2              =   1830
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   2700
      Left            =   1830
      TabIndex        =   8
      Top             =   15
      Width           =   2955
      Begin VB.CheckBox Check21 
         Caption         =   "Supplement"
         Height          =   255
         Left            =   1530
         TabIndex        =   62
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton Command10 
         Caption         =   "?"
         Height          =   360
         Left            =   2430
         TabIndex        =   61
         Top             =   525
         Width           =   420
      End
      Begin VB.CommandButton Command8 
         Caption         =   "More..."
         Height          =   300
         Left            =   1620
         TabIndex        =   42
         Top             =   2355
         Width           =   1290
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Capture DiskIO"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   $"Form1.frx":01B8
         Top             =   210
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox Check13 
         Caption         =   "FSCTL"
         Height          =   225
         Left            =   1530
         TabIndex        =   38
         Top             =   1470
         Width           =   1305
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Use multi-use logger (Win8+)"
         Height          =   225
         Left            =   135
         TabIndex        =   37
         Top             =   1935
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Don't show rundown events"
         Height          =   225
         Left            =   135
         TabIndex        =   19
         ToolTipText     =   "A rundown gives a list of all currently open files."
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Directory enum"
         Height          =   285
         Left            =   1530
         TabIndex        =   18
         Top             =   1200
         Width           =   1485
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Set info"
         Height          =   285
         Left            =   1530
         TabIndex        =   17
         Top             =   945
         Width           =   885
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Query"
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1185
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Rename"
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Delete"
         Height          =   285
         Left            =   1530
         TabIndex        =   13
         Top             =   690
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Write"
         Height          =   285
         Left            =   780
         TabIndex        =   12
         Top             =   945
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Read"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   945
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Open/Create"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Misc"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   1725
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capture FileIo events"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filters"
      Height          =   2700
      Left            =   4845
      TabIndex        =   1
      Top             =   15
      Width           =   6225
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3210
         ScaleHeight     =   255
         ScaleWidth      =   2805
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1965
         Width           =   2805
         Begin VB.OptionButton Option1 
            Caption         =   "Include only"
            Height          =   195
            Index           =   1
            Left            =   1095
            TabIndex        =   36
            Top             =   30
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Exclude"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   35
            Top             =   30
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Filter Flow and Syntax"
         Height          =   300
         Left            =   2040
         TabIndex        =   33
         Top             =   150
         Width           =   2970
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Ignore activity from self"
         Height          =   225
         Left            =   3240
         TabIndex        =   25
         ToolTipText     =   "Ignore activity generated by this process."
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   90
         TabIndex        =   24
         Top             =   2205
         Width           =   2900
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   90
         TabIndex        =   22
         Top             =   1665
         Width           =   2900
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5055
         TabIndex        =   20
         Top             =   150
         Width           =   1110
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3150
         TabIndex        =   7
         Top             =   1665
         Width           =   2900
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   1125
         Width           =   5685
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   5685
      End
      Begin VB.Line Line2 
         X1              =   3075
         X2              =   3075
         Y1              =   1425
         Y2              =   2505
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exclude file names:"
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   1980
         Width           =   1620
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File name must match:"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   1425
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process names or pids:"
         Height          =   195
         Left            =   3165
         TabIndex        =   6
         Top             =   1440
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exclude paths:"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File path must match:"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   1905
      End
   End
   Begin VB.Timer tmrRefresh 
      Left            =   540
      Top             =   3090
   End
   Begin VB.TextBox Text1 
      Height          =   2625
      Left            =   11130
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   4950
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'************************************************************************************
'VBEventTrace File Activity Monitor
'
'Version: 2.0 Last modified: 2022 May 11
'
'Monitor disk and file activity with the Windows Event Tracing NT Kernel Logger
'This will catch many activities missed by higher level methods like shell
'change notifications, directory change notifications, and
'scanning running processes for file handles.
'
'(c) 2022 Jon Johnson (aka fafalone)
'
'Help from many people, including VB Forum's own The trick, were invaluable
'for this project, because Microsoft's documentation is terrible.
'
'Notes:
'
'-You must run as administrator to use the NT Kernel Logger (and thus this app)
'
'-We begin receiving notifications after ProcessTrace is called; but
' ProcessTrace will not return until the trace is shut off and block execution.
' So, this project makes use of The trick's vbTrickThreading module to allow
' it to be called from a new thread. I've had some stability issues in the IDE,
' it may only run compiled.
'
'-I'm not sure about compatibility with Windows XP or Vista.
'
'-We respond to *a lot* of events, and the VB ListView was seriously choking.
' A regular API ListView might have had similar issues, so even though it
' adds a lot of complexity to an already monstrously complex demo project,
' I went for an API virtual ListView (LVS_OWNERDATA). ListView items are
' supplied by a callback, and only the ones being displayed are queried.
'
'Changelog
'
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
'
'************************************************************************************


Private nCurCt As Long
Private nLogPos As Long

Private bMergeColVis As Boolean

Private Const BS_PUSHLIKE = &H1000&
Private Const BS_NOTIFY = &H4000&
Private Const BCM_FIRST = &H1600
Private Const BM_SETSTATE = &HF3
Private Const BN_KILLFOCUS = 7&
Private Const BCM_SETSHIELD = (BCM_FIRST + &HC)
Private Const SB_BOTTOM = 7
Private Const EM_SCROLL As Integer = &HB5
Private Const ES_NUMBER = &H2000
Private Const SE_SYSTEM_PROFILE_NAME            As String = "SeSystemProfilePrivilege"
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As Any, ByRef ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValueW Lib "advapi32.dll" (ByVal StrPtrSystemName As Long, ByVal StrPtrName As Long, lpLuid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsUserAnAdmin Lib "shell32" () As Long
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function FindResourceW Lib "kernel32" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Any) As Long
Private Declare Function SHOpenFolderAndSelectItems Lib "shell32" (ByVal pidlFolder As Long, ByVal cidl As Long, ByVal apidl As Long, ByVal dwFlags As Long) As Long
Private Declare Function ILCreateFromPathW Lib "shell32" (ByVal pwszPath As Long) As Long
Private Declare Function ILFindLastID Lib "shell32" (ByVal pidl As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function CompareMemory Lib "ntdll" Alias "RtlCompareMemory" (Source1 As Any, Source2 As Any, ByVal Length As Long) As Long

Private Const GWL_STYLE As Long = (-16)
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Enum WinStylesEx
  WS_EX_DLGMODALFRAME = &H1
  WS_EX_NOPARENTNOTIFY = &H4
  WS_EX_TOPMOST = &H8
  WS_EX_ACCEPTFILES = &H10
  WS_EX_TRANSPARENT = &H20
  
  WS_EX_MDICHILD = &H40
  WS_EX_TOOLWINDOW = &H80
  WS_EX_WINDOWEDGE = &H100
  WS_EX_CLIENTEDGE = &H200
  WS_EX_CONTEXTHELP = &H400
  
  WS_EX_RIGHT = &H1000
  WS_EX_LEFT = &H0
  WS_EX_RTLREADING = &H2000
  WS_EX_LTRREADING = &H0
  WS_EX_LEFTSCROLLBAR = &H4000
  WS_EX_RIGHTSCROLLBAR = &H0
  
  WS_EX_CONTROLPARENT = &H10000
  WS_EX_STATICEDGE = &H20000
  WS_EX_APPWINDOW = &H40000
  
  WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
  WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum   ' WinStylesEx

Private clrDefBk As Long

Private bScrBtm As Boolean
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO) As Long
Private Const SB_HORZ As Long = 0
Private Const SB_VERT As Long = 1
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const FILE_END As Long = 2&
Private Const GENERIC_WRITE   As Long = &H40000000
Private Const FILE_SHARE_READ = &H1&
Private Const OPEN_ALWAYS As Long = 4&
Private Const CREATE_ALWAYS = 2&

Private Const OFASI_EDIT = &H1 'Initiate a rename (if single file)
Private Const OFASI_OPENDESKTOP = &H2 'Not used by this Demo, but highlights files on the desktop.

Private Const sCol0 As String = "Process"
Private Const sCol1 As String = "Event"
Private Const sCol2 As String = "File"
Private Const sCol3 As String = "Read"
Private Const sCol4 As String = "Write"
Private Const sCol5 As String = "Date added"
Private Const sCol6 As String = "Last r/w update"
Private Const sCol7 As String = "Notes"
Private Const sCol8 As String = "Ct/Opn"
Private Const sCol9 As String = "Del"

Private Const sSizeFmt_byte = "0 bytes"
Private Const sSizeFmt_kb = "#,##0 KB"
Private Const sSizeFmt_mb = "#,##0.00 MB"
Private Const sSizeFmt_gb = "#,##0.00 GB"
Private Const sSizeFmt_tb = "#,##0.00 TB"
Private Const sSizeFmt_pb = "#,##0.00 PB"

Private Const dtFormat As String = "yyyy-mm-dd Hh:nn:Ss"

Public hLVS As Long
Public himlSys16 As Long
Public himlSys32 As Long

Public Enum TOKEN_INFORMATION_CLASS
  TokenUser = 1
  TokenGroups
  TokenPrivileges
  TokenOwner
  TokenPrimaryGroup
  TokenDefaultDacl
  TokenSource
  TokenType
  TokenImpersonationLevel
  TokenStatistics
  TokenRestrictedSids
  TokenSessionId
  TokenGroupsAndPrivileges
  TokenSessionReference
  TokenSandBoxInert
  TokenAuditPolicy
  TokenOrigin
  TokenElevationType
  TokenLinkedToken
  TokenElevation
  TokenHasRestrictions
  TokenAccessInformation
  TokenVirtualizationAllowed
  TokenVirtualizationEnabled
  TokenIntegrityLevel
  TokenUIAccess
  TokenMandatoryPolicy
  TokenLogonSid
  TokenIsAppContainer
  TokenCapabilities
  TokenAppContainerSid
  TokenAppContainerNumber
  TokenUserClaimAttributes
  TokenDeviceClaimAttributes
  TokenRestrictedUserClaimAttributes
  TokenRestrictedDeviceClaimAttributes
  TokenDeviceGroups
  TokenRestrictedDeviceGroups
  TokenSecurityAttributes
  TokenIsRestricted
  MaxTokenInfoClass
End Enum

Private Const TOKEN_ASSIGN_PRIMARY As Long = &H1
Private Const TOKEN_DUPLICATE As Long = &H2
Private Const TOKEN_IMPERSONATE As Long = &H4
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_QUERY_SOURCE As Long = &H10
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_ADJUST_GROUPS As Long = &H40
Private Const TOKEN_ADJUST_DEFAULT As Long = &H80
Private Const TOKEN_ADJUST_SESSIONID As Long = &H100
Private Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT Or TOKEN_ADJUST_SESSIONID)

Private Type TOKEN_ELEVATION
    TokenIsElevated As Long
End Type
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long

Private Type LUID
    lowPart As Long
    highPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid       As LUID
    Attributes  As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount      As Long
    Privileges(0 To 1)  As LUID_AND_ATTRIBUTES
End Type

Private Enum SE_PRIVILEGE_ATTRIBUTES
'The attributes of a privilege can be a combination of the following values.
    SE_PRIVILEGE_ENABLED = &H2&                 'The privilege is enabled.
    SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1&      'The privilege is enabled by default.
    SE_PRIVILEGE_REMOVED = &H4&                 'Used to remove a privilege. For details, see AdjustTokenPrivileges.
    SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000   'The privilege was used to gain access to an object or service.
                                                ' This flag is used to identify the relevant privileges in a set passed by a client application
                                                ' that may contain unnecessary privileges.
                                                'PrivilegeCheck sets the Attributes member of each LUID_AND_ATTRIBUTES structure to
                                                ' SE_PRIVILEGE_USED_FOR_ACCESS if the corresponding privilege is enabled.
End Enum




Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer   ' e.g. = &h0000 = 0
   dwStrucVersionh As Integer   ' e.g. = &h0042 = .42
   dwFileVersionMSl As Integer   ' e.g. = &h0003 = 3
   dwFileVersionMSh As Integer   ' e.g. = &h0075 = .75
   dwFileVersionLSl As Integer   ' e.g. = &h0000 = 0
   dwFileVersionLSh As Integer   ' e.g. = &h0031 = .31
   dwProductVersionMSl As Integer   ' e.g. = &h0003 = 3
   dwProductVersionMSh As Integer   ' e.g. = &h0010 = .1
   dwProductVersionLSl As Integer   ' e.g. = &h0000 = 0
   dwProductVersionLSh As Integer   ' e.g. = &h0031 = .31
   dwFileFlagsMask As Long   ' = &h3F for version "0.42"
   dwFileFlags As Long   ' e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long   ' e.g. VOS_DOS_WINDOWS16
   dwFileType As Long   ' e.g. VFT_DRIVER
   dwFileSubtype As Long   ' e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long   ' e.g. 0
   dwFileDateLS As Long   ' e.g. 0
End Type
Private Type VS_VERSIONINFO_FIXED_PORTION
    wLength As Integer
    wValueLength As Integer
    wType As Integer
    szKey(1 To 16) As Integer   'Unicode "VS_VERSION_INFO" & vbNullChar.
    Padding1(1 To 1) As Integer 'Pad next field to DWORD boundary.
    Value As VS_FIXEDFILEINFO
End Type
Private Const RT_VERSION = 16

Private Enum ShowWindowTypes
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum
Private Enum SHELLEXECUTEMASK
    SEE_MASK_CLASSNAME = &H1
    SEE_MASK_CLASSKEY = &H3
    SEE_MASK_IDLIST = &H4
    SEE_MASK_INVOKEIDLIST = &HC
    SEE_MASK_ICON = &H10
    SEE_MASK_HOTKEY = &H20
    SEE_MASK_NOCLOSEPROCESS = &H40
    SEE_MASK_CONNECTNETDRV = &H80
    SEE_MASK_FLAG_DDEWAIT = &H100
    SEE_MASK_DOENVSUBST = &H200
    SEE_MASK_FLAG_NO_UI = &H400
    SEE_MASK_UNICODE = &H4000
    SEE_MASK_NO_CONSOLE = &H8000
    SEE_MASK_ASYNCOK = &H100000
    SEE_MASK_HMONITOR = &H200000
    SEE_MASK_NOZONECHECKS = &H800000
    SEE_MASK_NOQUERYCLASSSTORE = &H1000000
    SEE_MASK_WAITFORINPUTIDLE = &H2000000
    SEE_MASK_FLAG_LOG_USAGE = &H4000000
    SEE_MASK_FLAG_HINST_IS_SITE = &H8000000
End Enum
Private Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As SHELLEXECUTEMASK
  hWnd As Long
  lpVerb As Long
  lpFile As Long
  lpParameters As Long
  lpDirectory As Long
  nShow As ShowWindowTypes
  hInstApp As Long
  lpIDList As Long
  lpClass As Long
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type
Private Declare Function ShellExecuteExW Lib "shell32.dll" (lpExecInfo As SHELLEXECUTEINFO) As Long

Private Type ResultFolder
    sPath As String
    sFiles() As String
End Type
'**************************************************
'Menu APIs

Private Const widOpen As Long = 2500&
Private Const widShow As Long = 2501&
Private Const widProcProp As Long = 2502&
Private Const widProcOpen As Long = 2503&
Private Const widCopySel As Long = 2504&
Private Const widCopyAll As Long = 2505&
Private Const widCopySelLn As Long = 2506&
Private Const widCopyAllLn As Long = 2507&
Private Const mnOpen As String = "&Open item"
Private Const mnShow As String = "&Show item in Explorer"
Private Const mnProcProp As String = "&Process properties"
Private Const mnProcOpen As String = "Sho&w process in Explorer"
Private Const mnCopySel As String = "&Copy selected file names"
Private Const mnCopyAll As String = "Copy &all file names"
Private Const mnCopySelLn As String = "&Copy selected lines"
Private Const mnCopyAllLn As String = "Copy &all lines"


Private Type MENUITEMINFOW
  cbSize As Long
  fMask As MII_Mask
  fType As MF_Type              ' MIIM_TYPE
  fState As MF_State             ' MIIM_STATE
  wID As Long                       ' MIIM_ID
  hSubMenu As Long            ' MIIM_SUBMENU
  hbmpChecked As Long      ' MIIM_CHECKMARKS
  hbmpUnchecked As Long  ' MIIM_CHECKMARKS
  dwItemData As Long          ' MIIM_DATA
  dwTypeData As Long        ' MIIM_TYPE
  cch As Long                       ' MIIM_TYPE
  hbmpItem As Long
End Type
Private Enum MenuFlags
  MF_INSERT = &H0
  MF_ENABLED = &H0
  MF_UNCHECKED = &H0
  MF_BYCOMMAND = &H0
  MF_STRING = &H0
  MF_UNHILITE = &H0
  MF_GRAYED = &H1
  MF_DISABLED = &H2
  MF_BITMAP = &H4
  MF_CHECKED = &H8
  MF_POPUP = &H10
  MF_MENUBARBREAK = &H20
  MF_MENUBREAK = &H40
  MF_HILITE = &H80
  MF_CHANGE = &H80
  MF_END = &H80                    ' Obsolete -- only used by old RES files
  MF_APPEND = &H100
  MF_OWNERDRAW = &H100
  MF_DELETE = &H200
  MF_USECHECKBITMAPS = &H200
  MF_BYPOSITION = &H400
  MF_SEPARATOR = &H800
  MF_REMOVE = &H1000
  MF_DEFAULT = &H1000
  MF_SYSMENU = &H2000
  MF_HELP = &H4000
  MF_RIGHTJUSTIFY = &H4000
  MF_MOUSESELECT = &H8000&
End Enum
Private Enum MII_Mask
  MIIM_STATE = &H1
  MIIM_ID = &H2
  MIIM_SUBMENU = &H4
  MIIM_CHECKMARKS = &H8
  MIIM_TYPE = &H10
  MIIM_DATA = &H20
  MIIM_BITMAP = &H80
  MIIM_STRING = &H40
End Enum
Private Enum MF_Type
  MFT_STRING = MF_STRING
  MFT_BITMAP = MF_BITMAP
  MFT_MENUBARBREAK = MF_MENUBARBREAK
  MFT_MENUBREAK = MF_MENUBREAK
  MFT_OWNERDRAW = MF_OWNERDRAW
  MFT_RADIOCHECK = &H200
  MFT_SEPARATOR = MF_SEPARATOR
  MFT_RIGHTORDER = &H2000
  MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
End Enum
Private Enum MF_State
  MFS_GRAYED = &H3
  MFS_DISABLED = MFS_GRAYED
  MFS_CHECKED = MF_CHECKED
  MFS_HILITE = MF_HILITE
  MFS_ENABLED = MF_ENABLED
  MFS_UNCHECKED = MF_UNCHECKED
  MFS_UNHILITE = MF_UNHILITE
  MFS_DEFAULT = MF_DEFAULT
End Enum
Private Enum TPM_wFlags
  TPM_LEFTBUTTON = &H0
  TPM_RIGHTBUTTON = &H2
  TPM_LEFTALIGN = &H0
  TPM_CENTERALIGN = &H4
  TPM_RIGHTALIGN = &H8
  TPM_TOPALIGN = &H0
  TPM_VCENTERALIGN = &H10
  TPM_BOTTOMALIGN = &H20

  TPM_HORIZONTAL = &H0         ' Horz alignment matters more
  TPM_VERTICAL = &H40            ' Vert alignment matters more
  TPM_NONOTIFY = &H80           ' Don't send any notification msgs
  TPM_RETURNCMD = &H100
  
  TPM_HORPOSANIMATION = &H400
  TPM_HORNEGANIMATION = &H800
  TPM_VERPOSANIMATION = &H1000
  TPM_VERNEGANIMATION = &H2000
  TPM_NOANIMATION = &H4000
End Enum

Private Declare Function InsertMenuItemW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpmii As MENUITEMINFOW) As Boolean
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As TPM_wFlags, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lpRC As Any) As Long



'**************************************************************************
'This form makes use of self-subclassing. The order of the functions at the end
'cannot be changed.

' === Subclassing ========================================================
' Subclasing by Paul Caton
Private z_scFunk            As Collection   'hWnd/thunk-address collection
Private z_hkFunk            As Collection   'hook/thunk-address collection
Private z_cbFunk            As Collection   'callback/thunk-address collection
Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

' Declarations:
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpFN As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualQuery Lib "kernel32.dll" (ByVal addr As Long, pMBI As Any, ByVal lenMBI As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Enum eThunkType
    SubclassThunk = 0
    CallbackThunk = 2
End Enum

Private Enum eMsgWhen                                                   'When to callback
  MSG_BEFORE = 1                                                        'Callback before the original WndProc
  MSG_AFTER = 2                                                         'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
End Enum

Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows

Private Enum eAllMessages
    ALL_MESSAGES = -1     'All messages will callback
End Enum

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Enum CodePageTypeEnumn
    cptUnknown = 0
    cptVbClass = 0
    cptVbDataReport = 1
    cptVbFormOrMDI = 2
    cptVbPropertyPage = 3
    cptVbUserControl = 4
End Enum
'***************************************************************************
'
'----------------------------END DECLARE SECTION---------------------------
'
'***************************************************************************

'-SelfSub code------------------------------------------------------------------------------------
'-The following routines are exclusively for the ssc_Subclass routines----------------------------
Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False, _
                    Optional ByVal bIsAPIwindow As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '** NOTE: Self-sub code has been modified by fafalone to handle UserControls with >512 procedures
    '** NOTE: For performance reasons, currently configured only for use on a Form. See vOffset in
    '**       GetAddressOfEx to change if you copy this code for your own project. -fafalone
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4  'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                 'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
    
  ' This is the complete listing of thunk offset values and what they point/relate to.
  ' Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim z_ScMem       As Long           'Thunk base address
    Dim nAddr         As Long
    Dim nid           As Long
    Dim nMyID         As Long
    Dim bIDE          As Boolean

    If IsWindow(lng_hWnd) = 0 Then      'Ensure the window handle is valid
        Call zError(SUB_NAME, "Invalid window handle")
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nid              'Get the process ID associated with the window handle
    If nid <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        Call zError(SUB_NAME, "Window handle belongs to another process")
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = GetAddressOfEx(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        Call zError(SUB_NAME, "Callback method not found")
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                'Ensure the allocation succeeded
    
      If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
      Call z_scFunk.Add(z_ScMem, "h" & lng_hWnd)                'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'z_Sc (0) thru z_Sc(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    z_Sc(18) = &HD231C031: z_Sc(19) = &HBBE58960: z_Sc(21) = &H21E8F631: z_Sc(22) = &HE9000001: z_Sc(23) = &H12C&: z_Sc(24) = &HD231C031: z_Sc(25) = &HBBE58960: z_Sc(27) = &H3FFF631: z_Sc(28) = &H75047339: z_Sc(29) = &H2873FF23: z_Sc(30) = &H751C53FF: z_Sc(31) = &HC433913: z_Sc(32) = &H53FF2274: z_Sc(33) = &H13D0C: z_Sc(34) = &H18740000: z_Sc(35) = &H875C085: z_Sc(36) = &H820443C7: z_Sc(37) = &H90000000: z_Sc(38) = &H87E8&: z_Sc(39) = &H22E900: z_Sc(40) = &H90900000: z_Sc(41) = &H2C7B8B4A: z_Sc(42) = &HE81C7589: z_Sc(43) = &H90&: z_Sc(44) = &H75147539: z_Sc(45) = &H6AE80F: z_Sc(46) = &HD2310000: z_Sc(47) = &HE8307B8B: z_Sc(48) = &H7C&: z_Sc(49) = &H7D810BFF: z_Sc(50) = &H8228&: z_Sc(51) = &HC7097500: z_Sc(52) = &H80000443: z_Sc(53) = &H90900000: z_Sc(54) = &H44753339: z_Sc(55) = &H74047339: z_Sc(56) = &H2473FF3F: z_Sc(57) = &HFFFFFC68
    z_Sc(58) = &H2475FFFF: z_Sc(59) = &H811453FF: z_Sc(60) = &H82047B: z_Sc(61) = &HC750000: z_Sc(62) = &H74387339: z_Sc(63) = &H2475FF07: z_Sc(64) = &H903853FF: z_Sc(65) = &H81445B89: z_Sc(66) = &H484443: z_Sc(67) = &H73FF0000: z_Sc(68) = &H646844: z_Sc(69) = &H56560000: z_Sc(70) = &H893C53FF: z_Sc(71) = &H90904443: z_Sc(72) = &H10C261: z_Sc(73) = &H53E8&: z_Sc(74) = &H3075FF00: z_Sc(75) = &HFF2C75FF: z_Sc(76) = &H75FF2875: z_Sc(77) = &H2473FF24: z_Sc(78) = &H891053FF: z_Sc(79) = &H90C31C45: z_Sc(80) = &H34E30F8B: z_Sc(81) = &H1078C985: z_Sc(82) = &H4C781: z_Sc(83) = &H458B0000: z_Sc(84) = &H75AFF228: z_Sc(85) = &H90909023: z_Sc(86) = &H8D144D8D: z_Sc(87) = &H8D503443: z_Sc(88) = &H75FF1C45: z_Sc(89) = &H2C75FF30: z_Sc(90) = &HFF2875FF: z_Sc(91) = &H51502475: z_Sc(92) = &H2073FF52: z_Sc(93) = &H902853FF: z_Sc(94) = &H909090C3: z_Sc(95) = &H74447339: z_Sc(96) = &H4473FFF7
    z_Sc(97) = &H4053FF56: z_Sc(98) = &HC3447389: z_Sc(99) = &H89285D89: z_Sc(100) = &H45C72C75: z_Sc(101) = &H800030: z_Sc(102) = &H20458B00: z_Sc(103) = &H89145D89: z_Sc(104) = &H81612445: z_Sc(105) = &H4C4&: z_Sc(106) = &H1862FF00

    ' cache callback related pointers & offsets
      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_EBX_TMR) = z_ScMem                                             'Patch the thunk data address
      z_Sc(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If bIdeSafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      z_Sc(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      z_Sc(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If z_Sc(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      Call RtlMoveMemory(z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&)
      ssc_Subclass = True                                                     'Indicate success
    Else
        Call zError(SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError)
    End If
 Exit Function                                                                'Exit ssc_Subclass
    
CatchDoubleSub:
 Call zError(SUB_NAME, "Window handle is already subclassed")
      
ReleaseMemory:
      Call VirtualFree(z_ScMem, 0, MEM_RELEASE)                               'ssc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    Call zTerminateThunks(SubclassThunk)
End Sub

'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public, can be removed & zUnthunk can be called instead
    Call zUnThunk(lng_hWnd, SubclassThunk)
End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    Dim z_ScMem       As Long                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim m As Long
      For m = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(m))                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(m), IDX_BTABLE, z_ScMem) = False Then 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(m), IDX_ATABLE, z_ScMem) = False Then 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)   'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim m As Long
      For m = LBound(Messages) To UBound(Messages) ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(m))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then            'If the message is to be removed from the before original WndProc table...
              Call zDelMsg(Messages(m), IDX_BTABLE, z_ScMem) 'Remove the message to the before table
            End If
            If When And MSG_AFTER Then                       'If message is to be removed from the after original WndProc table...
              zDelMsg Messages(m), IDX_ATABLE, z_ScMem       'Remove the message to the after table
            End If
        End Select
      Next
    End If
End Sub

'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your window procedure
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                                     'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE, z_ScMem) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function

'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType) As Long
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zGet_lParamUser = zData(IDX_PARM_USER, z_ScMem)               'Get the lParamUser callback parameter
        End If
    End If
End Function

'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType, ByVal newValue As Long)
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zData(IDX_PARM_USER, z_ScMem) = newValue                      'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long) As Boolean
      Dim nCount As Long                            'Table entry count
      Dim nBase  As Long
      Dim i      As Long                            'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, z_ScMem)                'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                   'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                       'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                    'Get the current table entry count
        For i = 1 To nCount                         'Loop through the table entries
          If zData(i, nBase) = 0 Then               'If the element is free...
            zData(i, nBase) = uMsg                  'Use this element
            GoTo Bail                               'Bail
          ElseIf zData(i, nBase) = uMsg Then        'If the message is already in the table...
            GoTo Bail                               'Bail
          End If
        Next i                                      'Next message table entry
    
        nCount = i                                  'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                'Check for message table overflow
          Call zError("zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values")
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                            'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                 'Store the new table entry count
Bail:
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                     'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = uMsg Then                                        'If the message is found...
            zData(i, nBase) = 0                                                 'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
       ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
Bail:
End Sub

'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Private Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
                     Optional ByVal nOrdinal As Long = 1, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True, _
                     Optional ByVal bIsTimerCallback As Boolean) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '* bIsTimerCallback - optional, set to true for extra protection when used as a SetTimer callback
    '       If True, timer will be destroyed when IDE/app terminates. See scb_ReleaseCallback.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters and their types are dependent on the individual callback procedures
    
    Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
    Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
    Const INDX_OWNER    As Long = 0                             'Thunk data index of the Owner object's vTable address
    Const INDX_CALLBACK As Long = 1                             'Thunk data index of the EbMode function address
    Const INDX_EBMODE   As Long = 2                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_BADPTR   As Long = 3                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_KT       As Long = 4                             'Thunk data index of the KillTimer function address
    Const INDX_EBX      As Long = 6                             'Thunk code patch index of the thunk data
    Const INDX_PARAMS   As Long = 18                            'Thunk code patch index of the number of parameters expected in callback
    Const INDX_PARAMLEN As Long = 24                            'Thunk code patch index of the bytes to be released after callback
    Const PROC_OFF      As Long = &H14                          'Thunk offset to the callback execution address

    Dim z_ScMem       As Long                                   'Thunk base address
    Dim z_Cb()    As Long                                       'Callback thunk array
    Dim nValue    As Long
    Dim nCallback As Long
    Dim bIDE      As Boolean
      
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
        z_ScMem = z_cbFunk.Item("h" & ObjPtr(oCallback) & "." & nOrdinal) 'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = z_ScMem + PROC_OFF  'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                     ' validate parameters
        Call zError(SUB_NAME, "Invalid Parameter count")
        Exit Function
    End If
    If oCallback Is Nothing Then
        Set oCallback = Me
    End If
    nCallback = GetAddressOfEx(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        Call zError(SUB_NAME, "Callback address not found.")
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        
    If z_ScMem = 0& Then
        Call zError(SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError)  ' oops
        Exit Function
    End If
    Call z_cbFunk.Add(z_ScMem, "h" & ObjPtr(oCallback) & "." & nOrdinal) 'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long          'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(5) = &HBB60E089: z_Cb(7) = &H73FFC589: z_Cb(8) = &HC53FF04: z_Cb(9) = &H59E80A74: z_Cb(10) = &HE9000000
    z_Cb(11) = &H30&: z_Cb(12) = &H87B81: z_Cb(13) = &H75000000: z_Cb(14) = &H9090902B: z_Cb(15) = &H42DE889: z_Cb(16) = &H50000000: z_Cb(17) = &HB9909090: z_Cb(19) = &H90900AE3
    z_Cb(20) = &H8D74FF: z_Cb(21) = &H9090FAE2: z_Cb(22) = &H53FF33FF: z_Cb(23) = &H90909004: z_Cb(24) = &H2BADC261: z_Cb(25) = &H3D0853FF: z_Cb(26) = &H1&: z_Cb(27) = &H23DCE74: z_Cb(28) = &H74000000: z_Cb(29) = &HAE807
    z_Cb(30) = &H90900000: z_Cb(31) = &H4589C031: z_Cb(32) = &H90DDEBFC: z_Cb(33) = &HFF0C75FF: z_Cb(34) = &H53FF0475: z_Cb(35) = &HC310&

    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                    'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                         'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                    'Cache ordinal used for zTerminateThunks
      
    If bIdeSafety = True Then                               'If the user wants IDE protection
        Debug.Assert zInIDE(bIDE)
        If bIDE = True Then z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False) 'Store the EbMode function address in the thunk data
    End If
    If bIsTimerCallback Then
        z_Cb(INDX_KT) = zFnAddr("user32", "KillTimer", False)
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                         'Set the parameter count
    Call RtlMoveMemory(VarPtr(z_Cb(INDX_PARAMLEN)) + 2, VarPtr(nParamCount * 4), 2&)

    z_Cb(INDX_EBX) = z_ScMem                                'Set the data address relative to virtual memory pointer

    Call RtlMoveMemory(z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN) 'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + PROC_OFF                       'Thunk code start address
End Function

Private Sub scb_ReleaseCallback(ByVal nOrdinal As Long, Optional ByVal oCallback As Object)
    ' can be made public, can be removed & zUnThunk can be called instead
    ' NEVER call this from within the callback routine itself
    
    ' oCallBack is the object containing nOrdinal to be released
    ' if oCallback was already closed (say it was a class or form), then you won't be
    '   able to release it here, but it will be released when zTerminateThunks is
    '   eventually called
    
    ' Special Warning. If the callback thunk is used for a recurring callback (i.e., Timer),
    ' then ensure you terminate what is using the callback before releasing the thunk,
    ' otherwise you are subject to a crash when that item tries to callback to zeroed memory
    Call zUnThunk(nOrdinal, CallbackThunk, oCallback)
End Sub

Private Sub scb_TerminateCallbacks()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    Call zTerminateThunks(CallbackThunk)
End Sub


'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim z_ScMem       As Long         'Thunk base address
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
        If oCallback Is Nothing Then Set oCallback = Me
        colID = "h" & ObjPtr(oCallback) & "." & vFuncTarget
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
        colID = "h" & vFuncTarget
    Else
        Call zError("zMap_Vfunction", "Invalid thunk type passed")
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        Call zError("zMap_VFunction", "Thunk hasn't been initialized")
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            z_ScMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(z_ScMem) Then z_ScMem = 0&
            zMap_VFunction = z_ScMem
        End If
    End If
    Exit Function                                   'Exit returning the thunk address
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then Call zError("zMap_VFunction", "Thunk type for " & vType & " does not exist")
End Function

' sets/retrieves data at the specified offset for the specified memory address
Private Property Get zData(ByVal nIndex As Long, ByVal z_ScMem As Long) As Long
  Call RtlMoveMemory(VarPtr(zData), z_ScMem + (nIndex * 4), 4)
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal z_ScMem As Long, ByVal nValue As Long)
  Call RtlMoveMemory(z_ScMem + (nIndex * 4), VarPtr(nValue), 4)
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  ' App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  Call MsgBox(sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine)
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

Private Function GetAddressOfEx(ByRef VbCodePage As Object, ByVal nOrdinal As Long, _
                            Optional ByVal CodePageType As CodePageTypeEnumn = cptUnknown, _
                            Optional ByRef nMethodCount As Long, Optional ByRef nLastMethodOffset As Long) As Long
    
    ' Routine is basically an AddressOf function for VB code pages (forms, classes, etc)
    
    If nOrdinal < 1 Then Exit Function
    If VbCodePage Is Nothing Then Exit Function
    
    ' redesigned but based on Paul Caton's GetAddressOfEx method that can find function
    '   pointers within VB forms, classes, etc. Redesign includes merging 2 routines,
    '   using known VB offsets, and use of VirtualQuery over IsBadCodePtr API. This
    '   revised logic is slower than Caton's latest versions, but will prevent
    '   unintended page-guard activation and will not fail to return results in
    '   cases where Caton's GetAddressOfEx would due to his built-in loop restrictions.
    
    ' Modify the routine for large-address-aware application, as needed, i.e., pointer-safe math.
    
    ' Parameters
    '   :: VbCodePage is the VB class module (form, class, etc) containing the method ordinal
    '   :: nOrdinal is the ordinal whose function pointer/address is to be returned
    '       ordinals are always one-bound and counted from the bottom of the code page
    '       the last method is ordinal #1, second to last is #2, etc
    '       keep public methods near top of code page & private/friend near bottom because
    '       VB will move public ones closer to top during runtime, offsetting your ordinals.
    '   :: CodePageType when passed can help the function scan the code page more efficiently
    '   :: nMethodCount is returned with the number of user-defined methods in the code page
    '   :: nLastMethodOffset is returned with the address after the last user-defined method
    ' Return value
    '   If success, the function pointer will be returned, else zero is returned
    '   If zero is returned, nMethodCount and nLastMethodOffset may not be updated
    
    ' How this method works...
    ' With known offsets and function signatures, finding what we want is pretty easy.
    ' The function signature is simply the 1st byte of the function's code
    ' 1) If a function pointer is zero, then this is expected
    '       Seen typically when a code page uses Implements keyword
    ' Otherwise, there are four byte values we are interested in (the signature)
    ' 2) Byte &H33  start of XOR instruction in native code (always when in IDE)
    ' 3) Byte &HE9  start of XOR instruction in P-Code (only when compiled in P-Code)
    ' 4) Byte &H81  start of ADD instruction, regardless of P-Code usage
    ' 5) Byte &H58  start of POP instruction, regardless of P-Code usage
    
    Dim bSig As Byte, bVal As Byte
    Dim nAddr&, vOffset&, nFirst&
    Dim nMethod&, nAttempts&, n&
    Dim minAddrV&, maxAddrV&, minAddrM&, maxAddrM&
    Dim MBI&(0 To 6)          ' faux MEMORY_BASIC_INFORMATION structure
    ' (0) BaseAddress member    minimum range of committed memory (same protection)
    ' (3) Range member          maximum range BaseAddress+Range
    ' (5) Protect member
    ' This structure is key to not crashing while probing memory addresses.
    ' The Protect member of the structure is examined after each call. If it
    ' contains &H101 (mask), then the address is a page-guard or has no-access.
    ' Otherwise, if it contains &HFE (mask) then the address is readable.
    
    ' Step 1. Probe the passed code page to find the first user-defined method.
    ' The probe is quite fast. The outer For:Next loop helps to quickly filter the
    ' passed code page via the known offsets. The inner DO loop will execute up to
    ' four times to find the right code page offset as needed. After found, it will
    ' execute as little as one time or several times, depending on Implements usage
    ' and number of Public variables declared within the code page. That inner loop
    ' has a fudge-factor built in should some signature exist that is not known yet.
    ' However, no others have been found, to date, after the known offsets of the
    ' correct code page.
    
    If CodePageType <= cptUnknown Or CodePageType > cptVbUserControl Then
        n = 0: nAttempts = 4
    Else
        n = CodePageType: nAttempts = n
    End If
    CopyMemory nAddr, ByVal ObjPtr(VbCodePage), 4 ' host VTable
    
'    For n = n To nAttempts                      ' search in ascending order of offsets
'        Select Case n
'            Case 0: vOffset = nAddr + &H1C      ' known offset for VB Class,DataEnvironment,Add-in,DHTMLPage
'            Case 1: vOffset = nAddr + &H9C      ' known offset for VB DataReport
'            Case 2: vOffset = nAddr + &H6F8     ' known offset for VB Form, MDI
'            Case 3: vOffset = nAddr + &H710     ' known offset for VB Property Page
'            Case 4: vOffset = nAddr + &H7A4     ' known offset for VB UserControl
'        End Select
        
        vOffset = nAddr + &H6F8
        
        nAttempts = 4                           ' fudge-factor
        Do
            ' First validate the VTable slot address. If invalid, unsupported code page type
            If vOffset < minAddrV Or vOffset > maxAddrV Then
                MBI(5) = 0: VirtualQuery vOffset, MBI(0), 28
                If (MBI(5) And &HFE) = 0 Or (MBI(5) And &H101) <> 0 Then Exit Do ' Exit For
                minAddrV = MBI(0): maxAddrV = minAddrV + MBI(3)  ' set min/max range
            End If
            CopyMemory nMethod, ByVal vOffset, 4 ' get function address at VTable slot
            If nMethod <> 0 Then                ' zero = implemented, skip
            
                ' Next validate the function pointer. If invalid, unsupported code page type
                If nMethod < minAddrM Or nMethod > maxAddrM Then
                    MBI(5) = 0: VirtualQuery nMethod, MBI(0), 28
                    If (MBI(5) And &HFE) = 0 Or (MBI(5) And &H101) <> 0 Then Exit Do ' For
                    minAddrM = MBI(0): maxAddrM = minAddrM + MBI(3)  ' set min/max range
                End If
            
                CopyMemory bVal, ByVal nMethod, 1 ' get the 1st byte of that method
                If bVal = &H33 Or bVal = &HE9 Then
                    nFirst = vOffset            ' cache the location of first user-defined method
                    bSig = bVal: Exit Do ' For       ' cache the function signature & done
                ElseIf bVal <> &H81 Then        ' if not one of these 4 signatures, decrement attempts
                    If bVal <> &H58 Then nAttempts = nAttempts - 1
                End If
            End If
            vOffset = vOffset + 4               ' look at next VTable slot
        Loop Until nAttempts = 0
'    Next
    
    If nFirst = 0 Then Exit Function           ' failure
    ' If failure, then likely one of two reasons:
    ' 1) Unsupported code page
    ' 2) Code page has no user-defined methods
    
    ' Step 2. Find the last user-defined method.
    ' VB stacks user-defined methods contiguously, back to back. So, to find the last method,
    ' we simply need to keep looking until a signature no longer matches or we hit end of page.
    Do
        ' Validate the VTable slot address. If invalid, end of code page & done
        vOffset = vOffset + 4
        If vOffset < minAddrV Or vOffset > maxAddrV Then
            MBI(5) = 0: VirtualQuery vOffset, MBI(0), 28
            If (MBI(5) And &HFE) = 0 Or (MBI(5) And &H101) <> 0 Then Exit Do
            minAddrV = MBI(0): maxAddrV = minAddrV + MBI(3)  ' set min/max range
        End If
        
        CopyMemory nMethod, ByVal vOffset, 4    ' get function pointer at VTable slot
        If nMethod = 0 Then Exit Do             ' if zero, done because doesn't match our signature
        
        ' Validate the function pointer. If invalid, end of code page & done
        If nMethod < minAddrM Or nMethod > maxAddrM Then
            MBI(5) = 0: VirtualQuery nMethod, MBI(0), 28
            If (MBI(5) And &HFE) = 0 Or (MBI(5) And &H101) <> 0 Then Exit Do
            minAddrM = MBI(0): maxAddrM = minAddrM + MBI(3)  ' set min/max range
        End If
        
        CopyMemory bVal, ByVal nMethod, 1       ' get function's signature
    Loop Until bVal <> bSig                     ' done when doesn't match our signature
    
    ' Now set the optional parameter values
    nMethodCount = (vOffset - nFirst) \ 4
    nLastMethodOffset = vOffset
    
    ' Return the function pointer for requested ordinal, if a valid ordinal
    If nOrdinal <= nMethodCount Then
        CopyMemory GetAddressOfEx, ByVal vOffset - (nOrdinal * 4), 4
    End If

End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                    'Start address
  nLimit = nAddr + 128                                              'Probe eight entries
  Do While nAddr < nLimit                                           'While we've not reached our probe depth
    Call RtlMoveMemory(VarPtr(nEntry), nAddr, 4)                    'Get the vTable entry
    
    If nEntry <> 0 Then                                             'If not an implemented interface
      Call RtlMoveMemory(VarPtr(bVal), nEntry, 1)                   'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                            'Check for a native or pcode method signature
        nMethod = nAddr                                             'Store the vTable entry
        bSub = bVal                                                 'Store the found method signature
        zProbe = True                                               'Indicate success
        Exit Do                                                     'Return
      End If
    End If
    nAddr = nAddr + 4                                               'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType, Optional ByVal oCallback As Object)
    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Callbacks:    the ordinal of the callback
    '       ensure KillTimer is already called, if any callback used for SetTimer
    ' oCallback only used when vType is CallbackThunk
    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
    
    Dim z_ScMem       As Long                       'Thunk base address
    
    z_ScMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, z_ScMem) = 1                  'Set the shutdown indicator
            Call zDelMsg(ALL_MESSAGES, IDX_BTABLE, z_ScMem)   'Delete all before messages
            Call zDelMsg(ALL_MESSAGES, IDX_ATABLE, z_ScMem)   'Delete all after messages
        End If
        Call z_scFunk.Remove("h" & thunkID)                   'Remove the specified thunk from the collection
    Case CallbackThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            Call VirtualFree(z_ScMem, 0, MEM_RELEASE)   'Release allocated memory
        End If
        Call z_cbFunk.Remove("h" & ObjPtr(oCallback) & "." & thunkID) 'Remove the specified thunk from the collection
    End Select
End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)
    ' Terminates all thunks of a specific type
    ' Any subclassing, recurring callbacks should have already been canceled
    Dim i As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim z_ScMem       As Long                           'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          z_ScMem = .Item(i)                          'Get the thunk address
          If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), SubclassThunk    'Unsubclass
                Case CallbackThunk
                    ' zUnThunk expects object not pointer, convert pointer to object
                    Call RtlMoveMemory(VarPtr(oCallback), VarPtr(zData(INDX_OWNER, z_ScMem)), 4&)
                    Call zUnThunk(zData(IDX_CALLBACKORDINAL, z_ScMem), CallbackThunk, oCallback) ' release callback
                    ' remove the object pointer reference
                    Call RtlMoveMemory(VarPtr(oCallback), VarPtr(INDX_OWNER), 4&)
            End Select
          End If
        Next i                                        'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If
End Sub
'==============================================================================
'End of Self-Subclass procedures
'==============================================================================


Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True: bValue = True
End Function

Private Function SetNumbersOnly(hEdit As Long) As Long
'makes an edit control numbers only
Dim dwStyle As Long

dwStyle = GetWindowLong(hEdit, GWL_STYLE)
dwStyle = dwStyle Or ES_NUMBER
Call SetWindowLong(hEdit, GWL_STYLE, dwStyle)

End Function

Private Function MakePushButton(hWnd As Long) As Long
Dim dwStyle As Long

dwStyle = GetWindowLong(hWnd, GWL_STYLE)

dwStyle = dwStyle Or BS_PUSHLIKE Or BS_NOTIFY

Call SetWindowLong(hWnd, GWL_STYLE, dwStyle)
Call UpdateWindow(hWnd)

End Function

Private Sub ToggleButtonState(hWnd As Long, lState As Long)

Call SendMessage(hWnd, BM_SETSTATE, lState, ByVal 0&)

End Sub

Private Sub Check1_Click()
bEventCreate = (Check1.Value = vbChecked)
End Sub

Private Sub Check10_Click()
bIgnoreSelf = (Check10.Value = vbChecked)
End Sub

Private Sub Check12_Click()
bUseNewLogMode = (Check12.Value = vbChecked)
End Sub

Private Sub Check13_Click()
bEventFsctl = (Check13.Value = vbChecked)
End Sub

Private Sub Check14_Click()
bEventDiskIO = (Check14.Value = vbChecked)
End Sub

Private Sub Check15_Click()
bUseInitRd = (Check15.Value = vbChecked)
End Sub

Private Sub Check16_Click()
bUseEndRd = (Check16.Value = vbChecked)
End Sub

Private Sub Check17_Click()
bEnableCSwitch = (Check17.Value = vbChecked)
End Sub

Private Sub Check18_Click()
bMergeSameFile = (Check18.Value = vbChecked)
End Sub

Private Sub Check19_Click()
bMergeSameCode = (Check19.Value = vbChecked)
End Sub

Private Sub Check2_Click()
bEventRead = (Check2.Value = vbChecked)
End Sub

Private Sub Check20_Click()
If bMergeColVis Then
    Call SendMessage(hLVS, LVM_DELETECOLUMN, 9&, ByVal 0&)
    Call SendMessage(hLVS, LVM_DELETECOLUMN, 8&, ByVal 0&)
    bMergeColVis = False
Else
    AddMergeCols
    bMergeColVis = True
End If
End Sub

Private Sub Check21_Click()
bSupDIOE = (Check21.Value = vbChecked)
End Sub

Private Sub Check3_Click()
bEventWrite = (Check3.Value = vbChecked)
End Sub

Private Sub Check4_Click()
bEventDelete = (Check4.Value = vbChecked)
End Sub

Private Sub Check5_Click()
bEventRename = (Check5.Value = vbChecked)
End Sub

Private Sub Check6_Click()
bEventQuery = (Check6.Value = vbChecked)
End Sub

Private Sub Check7_Click()
bEventSetInfo = (Check7.Value = vbChecked)
End Sub

Private Sub Check8_Click()
bEventDirEnum = (Check8.Value = vbChecked)
End Sub

Private Sub Check9_Click()
bEventNoRundown = (Check9.Value = vbChecked)
End Sub

Private Sub SetFilterOptions()
bEventCreate = (Check1.Value = vbChecked)
bEventRead = (Check2.Value = vbChecked)
bEventWrite = (Check3.Value = vbChecked)
bEventDelete = (Check4.Value = vbChecked)
bEventRename = (Check5.Value = vbChecked)
bEventQuery = (Check6.Value = vbChecked)
bEventSetInfo = (Check7.Value = vbChecked)
bEventDirEnum = (Check8.Value = vbChecked)
bEventFsctl = (Check13.Value = vbChecked)
bEventDiskIO = (Check14.Value = vbChecked)
bEventNoRundown = (Check9.Value = vbChecked)
bUseInitRd = (Check15.Value = vbChecked)
bUseEndRd = (Check16.Value = vbChecked)
bEnableCSwitch = (Check17.Value = vbChecked)
bMergeSameFile = (Check18.Value = vbChecked)
bMergeSameCode = (Check19.Value = vbChecked)
bSupDIOE = (Check21.Value = vbChecked)


If bEventCreate Or bEventRead Or bEventWrite Or bEventDelete Or bEventQuery Or bEventSetInfo Or bEventRename Or bEventDirEnum Or bEventFsctl Then
    DiskIOExclusive = False
Else
    DiskIOExclusive = True
    Check1.Enabled = False
    Check2.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    Check5.Enabled = False
    Check6.Enabled = False
    Check7.Enabled = False
    Check8.Enabled = False
    Check13.Enabled = False
    Check21.Enabled = False
    PostLog "Note: Running in DiskIO exclusive mode. FileIO events cannot be enabled/disabled while running."
End If

If Option2(0).Value = True Then
    ProcInfoNoCache = False
ElseIf Option2(1).Value = True Then
    ProcInfoNoCache = True
ElseIf Option2(2).Value = True Then
    If DiskIOExclusive Then
        ProcInfoNoCache = True
    Else
        ProcInfoNoCache = False
    End If
End If

sFilterPath = Text2.Text
sFilterPathExc = Text3.Text
sFilterFile = Text7.Text
sFilterFileExc = Text8.Text
sFilterProc = Text4.Text

bProcIsInc = (Option1(1).Value = True)
bIgnoreSelf = (Check10.Value = vbChecked)
End Sub

Private Sub Command1_Click()
bStopping = False
Command1.Enabled = False
SetFilterOptions
tmrRefresh.Interval = CLng(Text5.Text)
tmrRefresh.Enabled = True
tmrLog.Interval = CLng(Text6.Text)
bScrBtm = True
If Check11.Value = vbChecked Then
    ClearBuffers
End If
If InitTrace() Then
    Command5.Enabled = True
    PostLog "Successfully started trace."
Else
    Command1.Enabled = True
    PostLog "Failed to start trace."
End If
End Sub

Private Sub Command10_Click()
Dim sMsg As String
sMsg = "Physical DiskIO will capture activity reading/writing from the disk. This includes an open/create not cached, read/write, and delete." & vbCrLf & _
    "The FileIO events include all events, where many of these events will be done on cached copies, and don't involve physically accessing the disk." & vbCrLf & vbCrLf & _
    "This version includes an optimized DiskIO exclusive mode, where only that event set is checked, to monitor disk activity only. The 'Supplement' " & _
    "option enables the FileIO events but filters them out from display; this will aid in process attribution, but may result in open/create messages " & _
    "from cached files. The difference between this option and manually enabling FileIO events is the optimizations added for DiskIO exclusive mode."
    
MsgBox sMsg, vbInformation + vbOKOnly, App.Title

End Sub

Private Sub Command2_Click()

If bInitRdDone = False Then
    If bUseNewLogMode = True Then
        If bUseEndRd = True Then
            Dim enpm As ENABLE_TRACE_PARAMETERS
            enpm.Version = ENABLE_TRACE_PARAMETERS_VERSION_2
            Dim het As Long
            het = EnableTraceEx2(gTraceHandle, KernelRundownGuid, EVENT_CONTROL_CODE_ENABLE_PROVIDER, TRACE_LEVEL_NONE, &H10 / 10000, 0@, 0&, enpm)
            PostLog "EnableTraceEx2(Enable)=0x" & Hex$(het)
        End If
    End If
End If

bStopping = True
EndTrace gSessionHandle, gTraceHandle

Dim hr As Long
hr = WaitForSingleObject(hThreadWait, 5000)
If hr = WAIT_OBJECT_0 Then
    CloseHandle hThreadWait
    tmrRefresh.Enabled = False
    If bUseInitRd = False Then
        If bInitRdDone = False Then
            DoEndRundown
        End If
    End If
    SyncRecordsAndUpdate 'Do one last refresh to get the last data
ElseIf hr = WAIT_TIMEOUT Then
    PostLog "Error: Timed out waiting for ProcessTrace thread exit."
ElseIf hr = WAIT_FAILED Then
    PostLog "Error: WAIT_FAILED, " & GetErrorName(Err.LastDllError)
End If

Command5.Enabled = False
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
Check8.Enabled = True
Check13.Enabled = True
Check21.Enabled = False

Dim IsIDE As Boolean
Debug.Assert MakeTrue(IsIDE)
If Not IsIDE Then
    Command1.Enabled = True
End If
End Sub

Private Sub Command3_Click()
FlushTrace gTraceHandle
End Sub

Private Sub Command4_Click()
SaveFileActivity
End Sub

Private Sub Command5_Click()
SetFilterOptions
SetFilters
End Sub

Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bPauseCol Then
    ToggleButtonState Command6.hWnd, 0
    bPauseCol = False
Else
    ToggleButtonState Command6.hWnd, 1
    bPauseCol = True
End If
End Sub


Private Sub Command7_Click()
Dim sMsg As String
sMsg = "Filters are processed in the following manner:" & vbCrLf & _
    "(1) If Self is excluded, any activity from this process is ignored." & vbCrLf & _
    "(2) Event process is checked, if included based on process filter, then " & vbCrLf & _
    "(3) File path is checked against the inclusion filters, if not, the event is ignored, if it is, " & vbCrLf & _
    "(4) File path is checked against the exclusion filters, if it matches an exclusion, the event if ignored, otherwise," & vbCrLf & _
    "(5) File name is checked against inclusion filters, if it doesn't match one, the event is ignored, otherwise," & vbCrLf & _
    "(6) File name is checked against the exclusion filter, if it matches one, the event is ignored." & vbCrLf & vbCrLf & _
    "Multiple entries are separated with a | (bar), file and process names can use DOS wildcards, process IDs must be exact, " & _
    "comparisons are case insensitive, process ids are indicated a > (caret)." & vbCrLf & vbCrLf & _
    "Note: If an activity initially occurs with a pid of -1, if the process is identified later through correlation, that will not be filtered even if the pid should be."
MsgBox sMsg, vbOKOnly + vbInformation, "Filter flow and syntax"

End Sub
Private Function LongToDouble(ByVal Lng As Long) As Double
    If Lng And &H80000000 = 0 Then
        LongToDouble = CDbl(Lng)
    Else
        LongToDouble = (Lng Xor &H80000000) + (2 ^ 31)
    End If
End Function

Private Function IsProcessElevated() As Boolean
Dim hToken As Long

If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hToken) = 0 Then Exit Function
 
Dim tElv As TOKEN_ELEVATION
Dim dwSize As Long
If GetTokenInformation(hToken, TokenElevation, tElv, Len(tElv), dwSize) = 0 Then
    PostLog "Error getting token information."
End If

If tElv.TokenIsElevated Then IsProcessElevated = True

If hToken Then CloseHandle hToken
End Function

Private Sub Command8_Click()
SetParent pbOptions.hWnd, Me.hWnd
pbOptions.Visible = True
Dim dwFrEx As Long
dwFrEx = GetWindowLong(pbOptions.hWnd, GWL_EXSTYLE)
dwFrEx = dwFrEx Or WS_EX_DLGMODALFRAME
SetWindowLong pbOptions.hWnd, GWL_EXSTYLE, dwFrEx

SetWindowPos pbOptions.hWnd, 0&, Frame2.Left + 10&, 10&, 0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED

End Sub

Private Sub Command9_Click()
pbOptions.Visible = False
End Sub

Private Sub Form_Load()
InitializeCriticalSectionAndSpinCount oCS, 4000&
InitializeCriticalSectionAndSpinCount oCS2, 4000&
ReDim DispActLog(0&)
tmrLog.Interval = 2500
tmrLog.Enabled = True
If Not IsProcessElevated() Then
    If AdjustPrivileges() = False Then
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command6.Enabled = False
        MsgBox "This program uses the NT Kernel Logger, which requires you to either run this program as Administrator or " & _
            "be a member of the Administrators group, a member of the Performance Log Users group, or another user/group " & _
            "with access to the SeSystemProfilePrivilege.", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
End If

Dim IsWow64 As Long
Dim hProc As Long
hProc = OpenProcess(PROCESS_QUERY_INFORMATION, 0&, GetCurrentProcessId())
If hProc Then
    Call IsWow64Process(hProc, IsWow64)
    If IsWow64 Then
        ReadWindowsVersion
        If bIsWin8OrGreater Then
            bUseNewLogMode = True
        Else
            Check12.Value = vbUnchecked
            Check12.Enabled = False
        End If
        hBtnPause = Command6.hWnd
        SetNumbersOnly Text5.hWnd
        SetNumbersOnly Text6.hWnd
        MakePushButton hBtnPause
        If ssc_Subclass(Me.hWnd, , 1) Then
         Call ssc_AddMsg(Me.hWnd, MSG_BEFORE, ALL_MESSAGES)
        End If
        InitLV
'       PrebuildFullProcessCache
        Dim IsIDE As Boolean
        Debug.Assert MakeTrue(IsIDE)
        If IsIDE Then
            Command1.Enabled = False
        End If
    Else
        Form1.Enabled = False
        MsgBox "This program is only compatible with 64 bit Windows.", vbCritical + vbOKOnly, App.Title
    End If
    CloseHandle hProc
Else
    Form1.Enabled = False
    MsgBox "Error accessing process.", vbCritical + vbOKOnly, App.Title
End If
'DumpMap
End Sub
 
Private Sub Form_Resize()
On Error Resume Next
If Form1.Height < 5000 Then Form1.Height = 5000
Text1.Width = Form1.ScaleWidth - 730& - 15

Dim rc As RECT
GetClientRect Me.hWnd, rc

SetWindowPos hLVS, 0&, 0&, 0&, rc.Right - 10&, rc.Bottom - 204&, SWP_NOMOVE Or SWP_NOZORDER
End Sub

Private Sub Form_Terminate()
DestroyWindow hLVS
Call ssc_Terminate
DeleteCriticalSection oCS
DeleteCriticalSection oCS2
End Sub

Private Sub Form_Unload(Cancel As Integer)
If gTraceHandle Then
    EndTrace gSessionHandle, gTraceHandle
End If
End Sub

Private Sub lblOpts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call ReleaseCapture
    Call SendMessage(pbOptions.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Option2_Click(Index As Integer)
If Option2(0).Value = True Then
    ProcInfoNoCache = False
ElseIf Option2(1).Value = True Then
    ProcInfoNoCache = True
ElseIf Option2(2).Value = True Then
    If DiskIOExclusive Then
        ProcInfoNoCache = True
    Else
        ProcInfoNoCache = False
    End If
End If
End Sub

Private Sub pbOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call ReleaseCapture
    Call SendMessage(pbOptions.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo Text5_KeyPress_Err
If KeyAscii = vbKeyReturn Then
    tmrRefresh.Interval = CLng(Text5.Text)
End If
Exit Sub

Text5_KeyPress_Err:
    PostLog "Text5_KeyPress.Error->" & Err.Description & ", 0x" & Hex$(Err.Number)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
On Error GoTo Text6_KeyPress_Err
If KeyAscii = vbKeyReturn Then
    tmrLog.Interval = CLng(Text6.Text)
End If
Exit Sub

Text6_KeyPress_Err:
    PostLog "Text6_KeyPress.Error->" & Err.Description & ", 0x" & Hex$(Err.Number)

End Sub

Private Sub tmrLog_Timer()
EnterCriticalSection oCS2
sFullLogLocal = sFullLog
LeaveCriticalSection oCS2
cbLog = LenB(sFullLogLocal)
If nLogPos <> cbLog Then
    Text1.Text = sFullLogLocal
    nLogPos = cbLog
    SendMessage Text1.hWnd, EM_SCROLL, SB_BOTTOM, ByVal 0&
End If
End Sub

Private Sub tmrRefresh_Timer()
SyncRecordsAndUpdate
End Sub

Private Sub SyncRecordsAndUpdate()
'Synchronize data with the ActivityLog
'This involves accessing the log and counter the other thread
'is constantly accessing, so we need a critical section to enforce
'this thread having exclusive access while it synchronizes the
'main activity log to the copy we use for displaying the data.
On Error GoTo e0
EnterCriticalSection oCS
If nAcEv Then
    If nAcEv <> nDspAc Then
        Dim i As Long, start As Long
        start = UBound(DispActLog)
        ReDim Preserve DispActLog(UBound(ActivityLog))
        For i = 0& To (start - 1&) 'Sync updatables
            DispActLog(i).sProcess = ActivityLog(i).sProcess
            DispActLog(i).intProcId = ActivityLog(i).intProcId
            DispActLog(i).intProcPath = ActivityLog(i).intProcPath
            DispActLog(i).cRead = ActivityLog(i).cRead
            DispActLog(i).cWrite = ActivityLog(i).cWrite
            DispActLog(i).OpenCount = ActivityLog(i).OpenCount
            DispActLog(i).DeleteCount = ActivityLog(i).DeleteCount
            DispActLog(i).dtMod = ActivityLog(i).dtMod
            If CompareMemory(DispActLog(i), ActivityLog(i), cbALCompRgn) = cbALCompRgn Then
                DispActLog(i).bChanged = True
            Else
                DispActLog(i).bChanged = False
            End If
        Next i
        For i = start To (UBound(ActivityLog))
            DispActLog(i) = ActivityLog(i)
        Next i
        nDspAc = nAcEv
    End If
End If
LeaveCriticalSection oCS
'We're now done with data the other thread is using.
On Error GoTo e1
If nDspAc <> nCurCt Then
    ListView_SetItemCount hLVS, nDspAc
    If bScrBtm Then ListView_EnsureVisible hLVS, nDspAc - 1&, 0& 'If we're scrolled to the bottom, keep it there
                                                                 'Otherwise, the user has scrolled up to look at something, leave the view alone
    nCurCt = nDspAc
Else
    'No change in items, but RW totals may have changed, so we want redraws anyway
    ListView_RedrawItems hLVS, 0&, nDspAc
    UpdateWindow hLVS
End If
Exit Sub
e0:
LeaveCriticalSection oCS
PostLog "RefreshError nAcEv=" & nAcEv & ",nDspAc=" & nDspAc
Exit Sub
e1:
PostLog "RefreshNonCritError"
End Sub

Private Function CreateListView(hWndParent As Long, idd As Long, dwStyle As Long, dwExStyle As Long, X As Long, Y As Long, CX As Long, CY As Long) As Long
 
Dim hwndLV As Long
hwndLV = CreateWindowEx(dwExStyle, WC_LISTVIEW, "", dwStyle, X, Y, CX, CY, hWndParent, idd, App.hInstance, 0)
If ssc_Subclass(hwndLV, , 2, , , True, True) Then
    Call ssc_AddMsg(hwndLV, MSG_BEFORE, ALL_MESSAGES)
End If
  
CreateListView = hwndLV
End Function
Private Sub InitLV()
  Dim dwStyle As Long, dwStyle2 As Long
  Dim i As Long
  Dim rc As RECT
  GetClientRect Me.hWnd, rc
  
hLVS = CreateListView(Me.hWnd, IDD_LISTVIEW, _
                    LVS_REPORT Or LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_ALIGNTOP Or LVS_OWNERDATA Or _
                    WS_VISIBLE Or WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, WS_EX_CLIENTEDGE, 5&, 200&, rc.Right - 10&, rc.Bottom - 204&)


dwStyle2 = GetWindowLong(hLVS, GWL_STYLE)
If ((dwStyle2 And LVS_SHAREIMAGELISTS) = False) Then
    Call SetWindowLong(hLVS, GWL_STYLE, dwStyle2 Or LVS_SHAREIMAGELISTS)
End If

Dim dwStyleEx As LVStylesEx
dwStyleEx = LVS_EX_JUSTIFYCOLUMNS Or LVS_EX_DOUBLEBUFFER Or LVS_EX_FULLROWSELECT Or LVS_EX_LABELTIP Or LVS_EX_HEADERDRAGDROP
Call SendMessage(hLVS, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal dwStyleEx)

Shell_GetImageLists himlSys32, himlSys16
Call ListView_SetImageList(hLVS, himlSys16, LVSIL_SMALL)

Dim swt1 As String
Dim swt2 As String
swt1 = "explorer"
swt2 = ""
Call SetWindowTheme(hLVS, StrPtr(swt1), 0&)

Dim lvcol As LVCOLUMNW

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 140
lvcol.cchTextMax = Len(sCol0) + 1
lvcol.pszText = StrPtr(sCol0)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 0&, lvcol)

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 110
lvcol.cchTextMax = Len(sCol1) + 1
lvcol.pszText = StrPtr(sCol1)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 1&, lvcol)

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 500
lvcol.cchTextMax = Len(sCol2) + 1
lvcol.pszText = StrPtr(sCol2)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 2&, lvcol)

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 60
lvcol.cchTextMax = Len(sCol3) + 1
lvcol.pszText = StrPtr(sCol3)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 3&, lvcol)

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 60
lvcol.cchTextMax = Len(sCol4) + 1
lvcol.pszText = StrPtr(sCol4)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 4&, lvcol)

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 140
lvcol.cchTextMax = Len(sCol5) + 1
lvcol.pszText = StrPtr(sCol5)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 5&, lvcol)

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 140
lvcol.cchTextMax = Len(sCol6) + 1
lvcol.pszText = StrPtr(sCol6)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 6&, lvcol)

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 180
lvcol.cchTextMax = Len(sCol7) + 1
lvcol.pszText = StrPtr(sCol7)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 7&, lvcol)

If Check20.Value = vbChecked Then
    AddMergeCols
    bMergeColVis = True
End If

clrDefBk = SendMessage(hLVS, LVM_GETBKCOLOR, 0&, ByVal 0&)

End Sub

Private Sub AddMergeCols()
Dim lvcol As LVCOLUMNW

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 60
lvcol.cchTextMax = Len(sCol8) + 1
lvcol.pszText = StrPtr(sCol8)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 8&, lvcol)

lvcol.Mask = LVCF_TEXT Or LVCF_WIDTH
lvcol.CX = 60
lvcol.cchTextMax = Len(sCol9) + 1
lvcol.pszText = StrPtr(sCol9)
Call SendMessage(hLVS, LVM_INSERTCOLUMNW, 9&, lvcol)

End Sub

Private Function AdjustPrivileges() As Boolean
Dim hToken As Long
Dim lRet As Long
lRet = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
If lRet Then
    PostLog "AdjustPrivileges::Got process token."
    
    If SetPrivilege(hToken, SE_SYSTEM_PROFILE_NAME, True) Then
        AdjustPrivileges = True
        PostLog "AdjustPrivileges::Enabled system profile privilege."
    Else
        PostLog "AdjustPrivileges::Failed to enable system profile privilege."
    End If
    
    CloseHandle hToken
Else
    PostLog "AdjustPrivileges::Failed to open process token."
End If
End Function

Private Function SetPrivilege(hToken As Long, ByVal sPriv As String, ByVal bEnable As Boolean) As Boolean
Dim tLUID As LUID
Dim tTP As TOKEN_PRIVILEGES
Dim tTPprv As TOKEN_PRIVILEGES
Dim cb As Long
Dim lRet As Long
Dim lastErr As Long
SetPrivilege = False

If LookupPrivilegeValueW(0&, StrPtr(sPriv), tLUID) = 0 Then
  lastErr = Err.LastDllError
  PostLog "SetPrivilege::LookupPrivilegeValue failed. LastDllError=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
  Exit Function
End If

With tTP
  .PrivilegeCount = 1
  .Privileges(0).pLuid = tLUID
  If bEnable Then
    .Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
  Else
    .Privileges(0).Attributes = 0
  End If
End With

lRet = AdjustTokenPrivileges(hToken, False, tTP, Len(tTPprv), tTPprv, cb)
lastErr = Err.LastDllError
If lastErr = 0 Then
  SetPrivilege = True
Else
  PostLog "SetPrivilege::Error code=" & GetErrorName(lastErr) & " (0x" & Hex$(Err.LastDllError) & "), return=0x" & Hex$(lRet)
End If
End Function

Private Sub GenerateTestData()
ReDim ActivityLog(5)
With ActivityLog(0)
    .sFile = "C:\dir1\file1.txt"
    .iType = atFileCreate
    .sProcess = "chrome.exe"
    .intProcPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    .iIcon = 2 'GetFileIconIndex("C:\procexp.exe", SHGFI_SMALLICON)
End With
With ActivityLog(1)
    .sFile = "C:\dir1\file2.txt"
    .iType = atFileCreate
    .sProcess = "chrome.exe"
    .intProcPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    .iIcon = 2
End With
With ActivityLog(2)
    .sFile = "C:\dir2\file1.txt"
    .iType = atFileCreate
    .sProcess = "chrome.exe"
    .intProcPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    .iIcon = 2
End With
With ActivityLog(3)
    .sFile = "C:\dir2\file2.txt"
    .iType = atFileCreate
    .sProcess = "chrome.exe"
    .intProcPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    .iIcon = 2
End With
With ActivityLog(4)
    .sFile = "C:\dir3\file1.txt"
    .iType = atFileCreate
    .sProcess = "chrome.exe"
    .intProcPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    .iIcon = 2
End With
With ActivityLog(5)
    .sFile = "C:\dir3\file2.txt"
    .iType = atFileDelete
    .sProcess = "chrome.exe"
    .intProcPath = "C:\addd.exe"
    .iIcon = 2
End With

 
End Sub

Private Sub OpenFolders(sFiles() As String, Optional bRename As Boolean = False)

If sFiles(0) = "" Then Exit Sub 'caller is responsible for ensuring array has been dim'd and contains valid info

Dim tRes() As ResultFolder
Dim apidl() As Long
Dim ppidl As Long
Dim pidlFQ() As Long
Dim i As Long, j As Long

GetResultsByFolder sFiles, tRes

'Now each entry in tRes is a folder, and its .sFiles member contains every file
'in the original list that is in that folder. So for every folder, we now need to
'create a pidl for the folder itself, and an array of all the relative pidls for the
'files. Two helper APIs replace what used to be tons of pidl-related support
'code before XP. After we've got the pidls, they're handed off to the API
For i = 0 To UBound(tRes)
    ReDim apidl(UBound(tRes(i).sFiles))
    ReDim pidlFQ(UBound(tRes(i).sFiles))
    For j = 0 To UBound(tRes(i).sFiles)
        pidlFQ(j) = ILCreateFromPathW(StrPtr(tRes(i).sFiles(j))) 'ILCreateFromPathW gives us Unicode support
        apidl(j) = ILFindLastID(pidlFQ(j))
    Next
    ppidl = ILCreateFromPathW(StrPtr(tRes(i).sPath))

    Dim dwFlag As Long
    If bRename Then
        dwFlag = OFASI_EDIT
    End If
    Call SHOpenFolderAndSelectItems(ppidl, UBound(apidl) + 1, VarPtr(apidl(0)), dwFlag)
    'Vista+ has the dwFlags to start renaming (single file) or select on desktop; there's no valid flags on XP

    'now we need to free all the pidls we created, otherwise it's a memory leak
    CoTaskMemFree ppidl
    For j = 0 To UBound(pidlFQ)
        CoTaskMemFree pidlFQ(j) 'per MSDN, child ids obtained w/ ILFindLastID don't need ILFree, so just free FQ
    Next
Next
        
End Sub

Private Sub GetResultsByFolder(sSelFullPath() As String, tResFolders() As ResultFolder)
Dim i As Long
Dim sPar As String
Dim k As Long, cn As Long, fc As Long
ReDim tResFolders(0)

For i = 0 To UBound(sSelFullPath)
    sPar = Left$(sSelFullPath(i), InStrRev(sSelFullPath(i), "\") - 1)
    k = RFExists(sPar, tResFolders)
    If k >= 0 Then 'there's already a file in this folder, so just add a new file to the folders list
        cn = UBound(tResFolders(k).sFiles)
        cn = cn + 1
        ReDim Preserve tResFolders(k).sFiles(cn)
        tResFolders(k).sFiles(cn) = sSelFullPath(i)
    Else 'create a new folder entry
        ReDim Preserve tResFolders(fc)
        ReDim tResFolders(fc).sFiles(0)
        tResFolders(fc).sPath = sPar
        tResFolders(fc).sFiles(0) = sSelFullPath(i)
        fc = fc + 1
    End If
Next
End Sub

Private Function RFExists(sPath As String, tResFolders() As ResultFolder) As Long
Dim i As Long
For i = 0 To UBound(tResFolders)
    If tResFolders(i).sPath = sPath Then
        RFExists = i
        Exit Function
    End If
Next
RFExists = -1
End Function

Private Sub SaveFileActivity()
Dim sWrite As String
Dim sPat As String

sPat = "Text File (*.txt)|*.txt" & vbNullChar & "All (*.*)| *.*" & vbNullChar & vbNullChar

If GetSaveName(sWrite, , , sPat, , App.Path, "Save activity log", ".txt", Me.hWnd) Then
    Dim sCpy() As String
    Dim cn As Long
    Dim sOut As String
    Dim i As Long
    cn = 0
    For i = 0& To UBound(DispActLog)
        ReDim Preserve sCpy(cn)
        sCpy(cn) = ItemString(i)
        Debug.Print sCpy(cn)
        cn = cn + 1
    Next i
    sOut = Join(sCpy, vbCrLf)
    
    WriteStrToFile sOut, sWrite
End If


End Sub
Private Sub WriteStrToFile(sIn As String, szFile As String, Optional bAppend As Boolean = False)
'Simple function to write a single string to file as-is
Dim hFile As Long
Dim RetVal As Long
Dim lngBytesWritten As Long

If Not bAppend Then
    hFile = CreateFileW(StrPtr(szFile), GENERIC_WRITE, FILE_SHARE_READ, _
                ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0&)
 
    If hFile = -1 Then PostLog "ERROR: Logfile could not be opened for writing."
Else
    'open the file for appending
    hFile = CreateFileW(StrPtr(szFile), GENERIC_WRITE, FILE_SHARE_READ, _
                ByVal 0&, OPEN_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0&)
    
    If hFile = -1 Then PostLog "ERROR: Logfile could not be opened for writing."
End If
If hFile Then
    'we need to move to EOF
    RetVal = SetFilePointer(hFile, 0&, 0&, FILE_END)
    RetVal = WriteFile(hFile, ByVal sIn, Len(sIn), lngBytesWritten, ByVal 0&)
    RetVal = CloseHandle(hFile)
End If
End Sub

Private Sub ShowLVMenu()

On Error GoTo ShowLVMenu_Err

Dim mii As MENUITEMINFOW
Dim hMenu As Long
Dim ptCur As POINTAPI
Dim idCmd As Long

hMenu = CreatePopupMenu()

With mii
    .cbSize = LenB(mii)
    .fMask = MIIM_ID Or MIIM_STRING Or MIIM_STATE
    .wID = widOpen
    .dwTypeData = StrPtr(mnOpen)
    .cch = Len(mnOpen)
    If nAcEv = 0& Then .fState = MFS_DISABLED
    Call InsertMenuItemW(hMenu, 0&, True, mii)
    
    .fMask = MIIM_ID Or MIIM_STRING Or MIIM_STATE
    .wID = widShow
    .dwTypeData = StrPtr(mnShow)
    .cch = Len(mnShow)
    If nAcEv = 0& Then .fState = MFS_DISABLED
    Call InsertMenuItemW(hMenu, 1&, True, mii)
    
    .fMask = MIIM_ID Or MIIM_STRING Or MIIM_STATE
    .wID = widProcProp
    .dwTypeData = StrPtr(mnProcProp)
    .cch = Len(mnProcProp)
    If nAcEv = 0& Then .fState = MFS_DISABLED
    Call InsertMenuItemW(hMenu, 2&, True, mii)
    
    .fMask = MIIM_ID Or MIIM_STRING Or MIIM_STATE
    .wID = widProcOpen
    .dwTypeData = StrPtr(mnProcOpen)
    .cch = Len(mnProcOpen)
    If nAcEv = 0& Then .fState = MFS_DISABLED
    Call InsertMenuItemW(hMenu, 3&, True, mii)
    
    .fMask = MIIM_ID Or MIIM_STRING Or MIIM_STATE
    .wID = widCopySel
    .dwTypeData = StrPtr(mnCopySel)
    .cch = Len(mnCopySel)
    If nAcEv = 0& Then .fState = MFS_DISABLED
    Call InsertMenuItemW(hMenu, 4&, True, mii)
    
    .fMask = MIIM_ID Or MIIM_STRING Or MIIM_STATE
    .wID = widCopyAll
    .dwTypeData = StrPtr(mnCopyAll)
    .cch = Len(mnCopyAll)
    If nAcEv = 0& Then .fState = MFS_DISABLED
    Call InsertMenuItemW(hMenu, 5&, True, mii)
    
    .fMask = MIIM_ID Or MIIM_STRING Or MIIM_STATE
    .wID = widCopySelLn
    .dwTypeData = StrPtr(mnCopySelLn)
    .cch = Len(mnCopySelLn)
    If nAcEv = 0& Then .fState = MFS_DISABLED
    Call InsertMenuItemW(hMenu, 6&, True, mii)
    
    .fMask = MIIM_ID Or MIIM_STRING Or MIIM_STATE
    .wID = widCopyAllLn
    .dwTypeData = StrPtr(mnCopyAllLn)
    .cch = Len(mnCopyAllLn)
    If nAcEv = 0& Then .fState = MFS_DISABLED
    Call InsertMenuItemW(hMenu, 7&, True, mii)
    
    
End With

Call GetCursorPos(ptCur)

idCmd = TrackPopupMenu(hMenu, TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_HORIZONTAL Or TPM_RETURNCMD, ptCur.X, ptCur.Y, 0, hLVS, 0)

Dim i As Long, cn As Long
Dim sCpy() As String
If idCmd Then
    Select Case idCmd
        Case widOpen
            i = LVI_NOITEM
            Do
                i = SendMessage(hLVS, LVM_GETNEXTITEM, i, ByVal LVNI_SELECTED)
                If (i <> LVI_NOITEM) Then
                    Dim sei As SHELLEXECUTEINFO
                    With sei
                        .cbSize = LenB(sei)
                        .fMask = SEE_MASK_UNICODE
                        .lpFile = DispActLog(i).sFile
                        .lpVerb = StrPtr("open")
                        .hWnd = hLVS
                        .nShow = SW_SHOWNORMAL
                    End With
                    Call ShellExecuteExW(sei)
                End If
            Loop Until (i = LVI_NOITEM)
        
        Case widShow
            Dim sShow() As String
            i = LVI_NOITEM
            Do
                i = SendMessage(hLVS, LVM_GETNEXTITEM, i, ByVal LVNI_SELECTED)
                If (i <> LVI_NOITEM) Then
                    ReDim Preserve sShow(cn)
                    sShow(cn) = DispActLog(i).sFile
                    cn = cn + 1
                End If
            Loop Until (i = LVI_NOITEM)
            
            OpenFolders sShow
            
        Case widProcProp
            i = ListView_GetSelectedItem(hLVS)
            Dim seip As SHELLEXECUTEINFO
            With seip
                .cbSize = LenB(sei)
                .fMask = SEE_MASK_UNICODE
                .lpFile = DispActLog(i).sFile
                .lpVerb = StrPtr("properties")
                .hWnd = hLVS
                .nShow = SW_SHOWNORMAL
            End With
            Call ShellExecuteExW(seip)
           
        Case widProcOpen
            Dim pidl As Long
            i = ListView_GetSelectedItem(hLVS)
            If DispActLog(i).intProcId <> -1& Then
                pidl = ILCreateFromPathW(DispActLog(i).intProcPath)
                If pidl Then
                    Call SHOpenFolderAndSelectItems(VarPtr(0&), 1&, VarPtr(pidl), 0&)
                    CoTaskMemFree pidl
                Else
                    PostLog "Couldn't find process."
                End If
            Else
                PostLog "Invalid process specified."
            End If
            
        Case widCopySel
            i = LVI_NOITEM
            Do
                i = SendMessage(hLVS, LVM_GETNEXTITEM, i, ByVal LVNI_SELECTED)
                If (i <> LVI_NOITEM) Then
                    ReDim Preserve sCpy(cn)
                    sCpy(cn) = DispActLog(i).sFile
                    cn = cn + 1
                End If
            Loop Until (i = LVI_NOITEM)
            Clipboard.Clear
            Clipboard.SetText Join(sCpy, vbCrLf)
            
         Case widCopyAll
            For i = 0& To UBound(DispActLog)
                ReDim Preserve sCpy(cn)
                sCpy(cn) = DispActLog(i).sFile
                cn = cn + 1
            Next i
            Clipboard.Clear
            Clipboard.SetText Join(sCpy, vbCrLf)
            
        Case widCopySelLn
            i = LVI_NOITEM
            Do
                i = SendMessage(hLVS, LVM_GETNEXTITEM, i, ByVal LVNI_SELECTED)
                If (i <> LVI_NOITEM) Then
                    ReDim Preserve sCpy(cn)
                    sCpy(cn) = ItemString(i)
                    cn = cn + 1
                End If
            Loop Until (i = LVI_NOITEM)
            Clipboard.Clear
            Clipboard.SetText Join(sCpy, vbCrLf)
         
         Case widCopyAllLn
            For i = 0& To UBound(DispActLog)
                ReDim Preserve sCpy(cn)
                sCpy(cn) = ItemString(i)
                cn = cn + 1
            Next i
            Clipboard.Clear
            Clipboard.SetText Join(sCpy, vbCrLf)
    End Select
End If


Exit Sub

ShowLVMenu_Err:
    PostLog "ShowLVMenu.Error->" & Err.Description & ", 0x" & Hex$(Err.Number)

End Sub


Private Function SystemTimeToDate(st As SYSTEMTIME) As Date
SystemTimeToDate = DateSerial(st.wYear, st.wMonth, st.wDay) + TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function
Private Function FormatTime(syst As SYSTEMTIME) As String
If syst.wYear = 0 Then
    FormatTime = vbNullString
    Exit Function
End If
Dim dt As Date
dt = SystemTimeToDate(syst)
FormatTime = Format$(dt, dtFormat) '& "." & Format$("###", syst.wMilliseconds)
End Function

Private Function ItemString(i As Long) As String
Dim sOut As String

sOut = DispActLog(i).sProcess & vbTab
Select Case DispActLog(i).iType
    Case atFileCreate: sOut = sOut & "Open/Create:" & CStr(DispActLog(i).iCode) & vbTab
    Case atFileAccess: sOut = sOut & "DiskIO:" & CStr(DispActLog(i).iCode) & vbTab
    Case atFileQuery: sOut = sOut & "DeleteFile:" & CStr(DispActLog(i).iCode) & vbTab
    Case atFileDelete: sOut = sOut & "RenameFile:" & CStr(DispActLog(i).iCode) & vbTab
    Case atFileRename: sOut = sOut & "QueryFile:" & CStr(DispActLog(i).iCode) & vbTab
    Case atFileSetInfo: sOut = sOut & "SetFileInfo:" & CStr(DispActLog(i).iCode) & vbTab
    Case atDirEnum: sOut = sOut & "DirEnum:" & CStr(DispActLog(i).iCode) & vbTab
    Case atDirChange: sOut = sOut & "DirChange:" & CStr(DispActLog(i).iCode) & vbTab
    Case atDirDel: sOut = sOut & "DeleteDir:" & CStr(DispActLog(i).iCode) & vbTab
    Case atDirRename: sOut = sOut & "RenameDir:" & CStr(DispActLog(i).iCode) & vbTab
    Case atDirSetLink: sOut = sOut & "SetLinkDir:" & CStr(DispActLog(i).iCode) & vbTab
    Case atDirNotify: sOut = sOut & "DirNotify:" & CStr(DispActLog(i).iCode) & vbTab
    Case atFileFsctl: sOut = sOut & "FSCTL:" & CStr(DispActLog(i).iCode) & vbTab
    Case atRundown: sOut = sOut & "Rundown:" & CStr(DispActLog(i).iCode) & vbTab
    Case atOpenFileRW: sOut = sOut & "OpenFileRW:" & CStr(DispActLog(i).iCode) & vbTab
End Select
sOut = sOut & DispActLog(i).sFile & vbTab
sOut = sOut & FormatSize(DispActLog(i).cRead) & vbTab
sOut = sOut & FormatSize(DispActLog(i).cWrite) & vbTab
sOut = sOut & FormatTime(DispActLog(i).dtStart) & vbTab
sOut = sOut & FormatTime(DispActLog(i).dtMod) & vbTab
sOut = sOut & DispActLog(i).sMisc
If bMergeColVis Then
    sOut = sOut & vbTab
    sOut = sOut & DispActLog(i).OpenCount & vbTab
    sOut = sOut & DispActLog(i).DeleteCount
End If
ItemString = sOut
End Function

Private Function DoLVNotify(hWnd As Long, lParam As Long) As Long
Dim sText As String, sSubText As String
Dim tNMH As NMHDR
CopyMemory tNMH, ByVal lParam, Len(tNMH)

Select Case tNMH.code
    Case WM_NOTIFYFORMAT
        DoLVNotify = NFR_UNICODE
        Exit Function
        
    Case LVN_GETEMPTYMARKUP
        Dim nmlvem As NMLVEMPTYMARKUP
        Debug.Print "LVN_GETEMPTYMARKUP"
        CopyMemory ByVal VarPtr(nmlvem), ByVal lParam, LenB(nmlvem)
        CopyMemory nmlvem.szMarkup(0), ByVal StrPtr(sEmpty), LenB(sEmpty)
        nmlvem.dwFlags = EMF_CENTERED
        CopyMemory ByVal lParam, ByVal VarPtr(nmlvem), LenB(nmlvem)

   Case LVN_GETDISPINFOW
        Dim LVDI As NMLVDISPINFO
        CopyMemory ByVal VarPtr(LVDI), ByVal lParam, LenB(LVDI)
        With LVDI.Item
            If (.Mask And LVIF_IMAGE) = LVIF_IMAGE Then
                .iImage = DispActLog(.iItem).iIcon
           End If
            If (.Mask And LVIF_TEXT) = LVIF_TEXT Then
            Select Case .iSubItem
                 Case 0:  .pszText = StrPtr(DispActLog(.iItem).sProcess)
                 Case 1
                    If DispActLog(.iItem).iType = atFileCreate Then .pszText = StrPtr("Open/Create:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atFileAccess Then .pszText = StrPtr("DiskIO:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atFileDelete Then .pszText = StrPtr("DeleteFile:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atFileRename Then .pszText = StrPtr("RenameFile:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atFileQuery Then .pszText = StrPtr("QueryFile:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atFileSetInfo Then .pszText = StrPtr("SetFileInfo:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atDirChange Then .pszText = StrPtr("DirChange:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atDirEnum Then .pszText = StrPtr("DirEnum:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atDirDel Then .pszText = StrPtr("DeleteDir:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atDirRename Then .pszText = StrPtr("RenameDir:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atDirSetLink Then .pszText = StrPtr("SetLinkDir:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atFileFsctl Then .pszText = StrPtr("FSCTL:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atDirNotify Then .pszText = StrPtr("DirNotify:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atRundown Then .pszText = StrPtr("Rundown:" & CStr(DispActLog(.iItem).iCode))
                    If DispActLog(.iItem).iType = atOpenFileRW Then .pszText = StrPtr("OpenFileRW:" & CStr(DispActLog(.iItem).iCode))
                Case 2: .pszText = StrPtr(DispActLog(.iItem).sFile)
                Case 3: .pszText = StrPtr(FormatSize(DispActLog(.iItem).cRead))
                Case 4: .pszText = StrPtr(FormatSize(DispActLog(.iItem).cWrite))
                Case 5: .pszText = StrPtr(FormatTime(DispActLog(.iItem).dtStart))
                Case 6: .pszText = StrPtr(FormatTime(DispActLog(.iItem).dtMod))
                Case 7: .pszText = StrPtr(DispActLog(.iItem).sMisc)
                Case 8: .pszText = StrPtr(CStr(DispActLog(.iItem).OpenCount))
                Case 9: .pszText = StrPtr(CStr(DispActLog(.iItem).DeleteCount))
            End Select
            End If
        End With

        CopyMemory ByVal lParam, ByVal VarPtr(LVDI), LenB(LVDI)
                   
        Case LVN_ENDSCROLL
            Dim tsc As SCROLLINFO
            tsc.cbSize = LenB(tsc)
            tsc.fMask = SIF_ALL
            GetScrollInfo hLVS, SB_VERT, tsc
            If (tsc.nPage + tsc.nPos) >= (tsc.nMax - 1&) Then
                bScrBtm = True
            Else
                bScrBtm = False
            End If
                
                   
    Case NM_RCLICK
        If tNMH.hWndFrom = hLVS Then
            ShowLVMenu
        End If
       
'    Case NM_DBLCLK
'            Dim isel As Long
'            isel = ListView_GetSelectedItem(hLVS)
'            dbgitemstr isel

'        Case NM_CUSTOMDRAW
'            Dim NMLVCD As NMLVCUSTOMDRAW
'            CopyMemory NMLVCD, ByVal lParam, Len(NMLVCD)
'            With NMLVCD.NMCD
'                Select Case .dwDrawStage
'
'                    Case CDDS_PREPAINT
'                        DoLVNotify = CDRF_NOTIFYITEMDRAW
'                        Exit Function
'
'                    Case CDDS_ITEMPREPAINT
'                        DoLVNotify = CDRF_NOTIFYSUBITEMDRAW
'                        Exit Function
'
'                    Case CDDS_ITEMPREPAINT Or CDDS_SUBITEM
'                        If NMLVCD.iSubItem = 6& Then
'                            If DispActLog(.dwItemSpec).bChanged Then
'                                NMLVCD.ClrTextBk = vbYellow
'                            Else
'                                NMLVCD.ClrTextBk = clrDefBk
'                            End If
'
'                            CopyMemory ByVal lParam, NMLVCD, LenB(NMLVCD)
'                            DoLVNotify = CDRF_NOTIFYSUBITEMDRAW 'Or CDRF_NEWFONT
'                            Exit Function
'                        End If
'                End Select
'            End With


End Select
        
End Function

Private Sub dbgitemstr(i As Long)
Dim sl As String, slb As String, slbt As String
Dim c As Long
Dim bt() As Byte
If DispActLog(i).iCode = 72 Then
    sl = Len(DispActLog(i).sMisc)
    slb = LenB(DispActLog(i).sMisc)
    bt = DispActLog(i).sMisc
    For c = 0 To UBound(bt)
        slbt = slbt & Hex$(bt(c)) & ","
    Next

Else
    sl = Len(DispActLog(i).sFile)
    slb = LenB(DispActLog(i).sFile)
    bt = DispActLog(i).sFile
    For c = 0 To UBound(bt)
        slbt = slbt & Hex$(bt(c)) & ","
    Next
End If
MsgBox "Len=" & sl & ",LenB=" & slb & ",ub=" & UBound(bt) & vbCrLf & "dump=" & slbt
End Sub

Private Function FormatSize(crSz As Currency) As String
If crSz = 0@ Then
    FormatSize = vbNullString
Else
    FormatSize = FormatFileSizeCurExB(crSz, False, sSizeFmt_byte, False, sSizeFmt_kb, _
                                                       False, sSizeFmt_mb, False, sSizeFmt_gb, False, sSizeFmt_tb, False, sSizeFmt_pb)
End If
End Function
Private Function FormatFileSizeCurExB(curB As Currency, _
                                     bNoBytes As Boolean, sByteFmt As String, _
                                     bNoKilo As Boolean, sKiloFmt As String, _
                                     bNoMega As Boolean, sMegaFmt As String, _
                                     bNoGiga As Boolean, sGigaFmt As String, _
                                     bNoTera As Boolean, sTeraFmt As String, _
                                     bNoPeta As Boolean, sPetaFmt As String) As String ' bNoExa As Boolean, bNoZetta As Boolean, bNoYotta As Boolean) As String
Dim sName As String
Dim dblKBs As Currency

'cannot currently work with exabytes and above due to limit of currency data type
On Error GoTo e0

If (bNoKilo = True) And (bNoMega = True) And (bNoGiga = True) And (bNoTera = True) And (bNoPeta = True) Then
    sName = Format$(curB, sByteFmt)
    GoTo chkunit
End If

If (curB < 1024) And (bNoBytes = False) Then
    sName = Format$(curB, sByteFmt)
Else
    dblKBs = curB / 1024@
    If ((dblKBs > 999999999999@) And (bNoPeta = False)) Or ((bNoTera = True) And (bNoGiga = True) And (bNoMega = True) And (bNoKilo = True) And (bNoBytes = True)) Then
        'file size >1TB and user-pref allows using GB unit
        dblKBs = dblKBs / 1024@ / 1024@ / 1024@ / 1024@
        sName = Format$(dblKBs, sPetaFmt)
    ElseIf ((dblKBs > 999999999@) And (bNoTera = False)) Or ((bNoGiga = True) And (bNoMega = True) And (bNoKilo = True) And (bNoBytes = True)) Then
        'file size >1TB and user-pref allows using GB unit
        dblKBs = dblKBs / 1024@ / 1024@ / 1024@
        sName = Format$(dblKBs, sTeraFmt)
    ElseIf ((dblKBs > 999999@) And (bNoGiga = False)) Or ((bNoMega = True) And (bNoKilo = True) And (bNoBytes = True)) Then
        'file size >1GB and user-pref allows using GB unit
        dblKBs = dblKBs / 1024@ / 1024@
        sName = Format$(dblKBs, sGigaFmt)
    ElseIf ((dblKBs > 999@) And (bNoMega = False)) Or ((bNoKilo = True) And (bNoBytes = True)) Then
        dblKBs = dblKBs / 1024@
        sName = Format$(dblKBs, sMegaFmt)
    ElseIf (bNoKilo = False) Then
        sName = Format$(dblKBs, sKiloFmt)
    Else
        sName = Format$(curB, sByteFmt)
    End If
End If

chkunit:

    FormatFileSizeCurExB = sName

On Error GoTo 0
Exit Function

e0:
PostLog "FormatFileSizeCurExB.Error->" & Err.Description & " (" & Err.Number & ")"

End Function

Private Sub ReadWindowsVersion()
'GetVersion[Ex] does not work with Win8 and above, so we'll go by kernel32 version
'GetFileVersionInfo does not work with some versions of Win10 and above.

Dim hMod As Long
Dim hRes As Long

hMod = LoadLibraryW(StrPtr("kernel32.dll"))
If hMod Then
    hRes = FindResourceW(hMod, StrPtr("#1"), RT_VERSION)
    If hRes Then
        Dim hGbl As Long
        hGbl = LoadResource(hMod, hRes)
        If (hGbl) Then
            Dim lpRes As Long
            lpRes = LockResource(hGbl)
            If lpRes Then
                Dim tVerInfo As VS_VERSIONINFO_FIXED_PORTION
                CopyMemory tVerInfo, ByVal lpRes, Len(tVerInfo)
                If tVerInfo.Value.dwFileVersionMSh >= 6& Then
                    bIsWinVistaOrGreater = True
                    If tVerInfo.Value.dwFileVersionMSl >= 1& Then bIsWin7OrGreater = True
                    If tVerInfo.Value.dwFileVersionMSl >= 2& Then bIsWin8OrGreater = True: bIsWin7OrGreater = True
                    If (tVerInfo.Value.dwFileVersionMSl = 4&) Or (tVerInfo.Value.dwFileVersionMSh >= 10&) Then
                        bIsWin7OrGreater = True
                        bIsWin8OrGreater = True
                        bIsWin10OrGreater = True
                    End If
                End If
            End If
        End If
    End If
    FreeLibrary hMod
End If
End Sub

'***************************************************************
'SUBCLASSING ROUTINES: DO NOT REORDER THESE OR ADD NEW METHODS BEYOND HERE

'@2
Private Sub LVSWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                      ByRef lParamUser As Long)
Select Case uMsg
    Case WM_NOTIFYFORMAT
        'Enable Unicode support
        lReturn = NFR_UNICODE
        bHandled = True
        Exit Sub
    Case WM_NOTIFY
        Dim tNMH As NMHDR
        CopyMemory tNMH, ByVal lParam, Len(tNMH)
        
        Select Case tNMH.code
            Case WM_NOTIFYFORMAT
                lReturn = NFR_UNICODE
                bHandled = True
        End Select
End Select

End Sub

'@1
Private Sub F1WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                      ByRef lParamUser As Long)
                      
Select Case uMsg
        Case WM_NOTIFYFORMAT
            'Enable Unicode support
            lReturn = NFR_UNICODE
            bHandled = True
            Exit Sub

        Case WM_NOTIFY
            Dim dwRtn As Long
            If (wParam = IDD_LISTVIEW) Then
                dwRtn = DoLVNotify(hWnd, lParam)
            End If
            If dwRtn Then
              lReturn = dwRtn
              bHandled = True
              Exit Sub
            End If
            
        Case WM_COMMAND
            Dim lCode As Long
            lCode = HiWord(wParam)
            Select Case lCode
                Case BN_KILLFOCUS
                    
                    If lParam = hBtnPause Then
                        Debug.Print "BN_KILLFOCUS " & bPauseCol
                        If bPauseCol Then
                            Call SendMessage(hBtnPause, BM_SETSTATE, 1&, ByVal 0&)
                        End If
                    End If
    
                    lReturn = 1
                    bHandled = True
                    Exit Sub
            End Select
End Select

End Sub
'*************************************************************
'WARNING: DO NOT ADD NEW CODE BELOW THIS LINE
