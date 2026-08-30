VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPerformanceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' CLASS: cPERFORMANCEMANAGER
'------------------------------------------------------------------------------
' PURPOSE
'   High-precision timing and benchmark-support utility for VBA on Windows.
'
'   The class provides:
'     - multiple timing backends under a session-bound interface
'     - low-overhead numeric elapsed-time measurement
'     - human-readable elapsed-time diagnostics
'     - optional formatting of an already measured elapsed-seconds value
'     - optional shared Excel "time-waster" suppression for benchmark runs and
'       for general Excel/VBA performance improvement in heavy procedures
'     - structured named checkpoints with numeric delta/cumulative reporting
'
' WHY THIS EXISTS
'   - VBA Timer() has limited resolution and rolls over at midnight.
'   - Windows provides monotonic and high-resolution timing sources that are
'     better suited to instrumentation and benchmarking.
'   - A single, consistent API reduces timing mistakes, duplicated code, and
'     backend-specific handling across modules.
'   - Excel benchmark runs and heavy workbook automation often benefit from
'     temporarily suppressing expensive Application behaviors such as
'     ScreenUpdating, EnableEvents, DisplayAlerts, Calculation, and Cursor
'     changes.
'
' COMPATIBILITY
'   - Windows only for API-backed timing methods (2..5).
'   - On non-Windows hosts, only Timer()/Now()-based methods are conceptually
'     portable, if retained.
'   - Requires appropriate VBA7 / Win64 conditional compilation for 32-bit and
'     64-bit Office.
'
' TIMER METHODS (iMethod)
'   1 = Timer
'       Seconds since midnight; rolls over every 24 hours.
'
'   2 = GetTickCount / GetTickCount64
'       Milliseconds since boot; the 32-bit variant wraps at approximately
'       49.7 days.
'
'   3 = timeGetTime
'       Millisecond counter with 32-bit rollover semantics; can benefit from
'       a 1ms multimedia timer resolution request.
'
'   4 = timeGetSystemTime (MMTIME / TIME_MS)
'       Millisecond counter source treated with 32-bit millisecond-counter
'       rollover semantics.
'
'   5 = QueryPerformanceCounter (QPC)   [DEFAULT / RECOMMENDED]
'       High-resolution monotonic counter for benchmark-grade timing.
'
'   6 = Now() * 86400
'       Wall-clock seconds; useful for diagnostics, not recommended as the
'       primary monotonic benchmark source.
'
' PERFORMANCE GUIDANCE
'   - Prefer ElapsedSeconds() for low-overhead numeric measurement.
'   - Prefer method 5 (QPC) for the primary benchmark path.
'   - ElapsedTime() is presentation-oriented:
'       * when called without ElapsedSecondsIn, it measures and formats
'       * when called with ElapsedSecondsIn, it formats an existing value
'         without taking a second timing sample
'   - AlignToNextTick is benchmark-specific and should be used only when its
'     extra polling overhead is justified.
'
' IMPLEMENTATION NOTES
'   - Timing is session-bound:
'       StartTimer establishes the active backend and elapsed-time reads are
'       validated against that same backend.
'   - QPC ticks and frequency are stored as Currency to preserve stable tick
'     precision while keeping the implementation efficient in VBA.
'   - 32-bit millisecond-counter sources require rollover-safe elapsed-time
'     arithmetic.
'   - timeBeginPeriod(1) for method 3 affects system-wide timer resolution and
'     can increase power usage; it is enabled only when needed and released
'     through ResetEnvironment, with Class_Terminate as a fallback safety net.
'   - Excel "time-waster" suppression is managed through a shared scope model so
'     overlapping instances can coexist without corrupting Application state.
'
' ELAPSED-TIME VALIDATION POLICY
'   A computed elapsed time can only be negative when the underlying timing
'   source moved backwards, which makes the measurement invalid rather than
'   merely imprecise.
'
'   Methods 1, 2, 3, 4 and 6 therefore route their result through
'   Elapsed_Validate after rollover correction:
'     - strict mode     => raises ERR_CPM_NEGATIVE_ELAPSED
'     - non-strict mode => clamps the result to zero
'
'   Method 5 (QPC) deliberately skips this check. QPC is monotonic, and method 5
'   is the default hot path where the extra call is not justified.
'
' INSTANCE KEY POLICY
'   The shared time-waster key for this instance is obtained once from
'   PM_TW_NewInstanceKey and cached in m_InstanceKey.
'
'   Object addresses (ObjPtr) must NOT be used for this purpose. VBA reuses heap
'   addresses, so a session that outlived its instance could be silently
'   inherited by an unrelated instance later allocated at the same address.
'
' ERROR NUMBERS
'   All raised error numbers are declared as named constants in the
'   HARDENING CONSTANTS section. No literal vbObjectError offsets appear in the
'   body of this class, so the full set of numbers in use can be reviewed in one
'   place and collisions remain visible.
'
' UPDATED
'   2026-08-15
'
' AUTHOR
'   Daniele Penza
'
' VERSION: 1.3.0
'==============================================================================

'------------------------------------------------------------------------------
' CLASS SETTINGS
'------------------------------------------------------------------------------
    Option Explicit     'Force explicit declaration of all variables

'------------------------------------------------------------------------------
' ENUMS
'------------------------------------------------------------------------------
    'Bitmask flags used to exempt specific Excel Application settings from
    'shared TW suppression.
    '
    'Example:
    '   Except:=ScreenUpdating Or EnableEvents
    Public Enum TW_Enum
        None = 0
        ScreenUpdating = 1
        EnableEvents = 2
        DisplayAlerts = 4
        Calculation = 8
        Cursor = 16
    End Enum

    'Timing backend selector.
    '
    'The values are unchanged from earlier releases, so StartTimer 5 and
    'StartTimer cPM_MethodQPC are identical. The enum exists so the choice is
    'legible at the call site and selectable from IntelliSense.
    Public Enum cPM_TimerMethod
        cPM_MethodSession = 0           'Use the backend bound to the active session
        cPM_MethodTimer = 1             'VBA Timer, seconds since midnight
        cPM_MethodTickCount = 2         'GetTickCount / GetTickCount64
        cPM_MethodTimeGetTime = 3       'winmm timeGetTime
        cPM_MethodSystemTime = 4        'winmm timeGetSystemTime via MMTIME
        cPM_MethodQPC = 5               'QueryPerformanceCounter - default
        cPM_MethodNow = 6               'Now() * 86400, wall clock
    End Enum

    'Pause strategy selector.
    '
    'These share the numbers 1 to 4 with the timing backends above and mean
    'something entirely different. Separate enum types make that difference
    'legible at the call site, but VBA enum members and enum-typed parameters
    'have Long semantics: a cPM_TimerMethod constant passed where a
    'cPM_PauseMethod is expected compiles and runs. The separation documents
    'intent; it is not a compile-time type boundary.
    Public Enum cPM_PauseMethod
        cPM_PauseSleep = 1              'Sleep API, lowest overhead, no DoEvents
        cPM_PauseTimerLoop = 2          'Timer delta with DoEvents
        cPM_PauseAppWait = 3            'Application.Wait plus a Timer bound
        cPM_PauseNowLoop = 4            'Now loop with DoEvents plus a Timer bound
    End Enum

    'Outcome of the most recent native timing read.
    '
    'A non-OK value means the reported elapsed value is not a measurement. In
    'non-strict mode the class returns 0 rather than raising, so callers need a
    'way to distinguish "zero seconds elapsed" from "no valid value was obtained".
    '
    'The name says "read" because most failures are native read failures, but
    'cPM_ReadElapsedInvalid covers the arithmetic validation that follows a
    'successful read. Both make the returned zero meaningless in the same way.
    Public Enum cPM_ReadStatus
        cPM_ReadOK = 0
        cPM_ReadQpcFailed = 1
        cPM_ReadSystemTimeFailed = 2
        cPM_ReadSystemTimeFormatInvalid = 3
        cPM_ReadFallbackToMethod2 = 4
        cPM_ReadElapsedInvalid = 5
    End Enum

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    'SMPTE is retained only to preserve the expected MMTIME structure layout.
        Private Type SMPTE
            Hour        As Byte
            Min         As Byte
            Sec         As Byte
            Frame       As Byte
            Fps         As Byte
            Dummy       As Byte
            Pad(2)      As Byte
        End Type

    'MMTIME is used by timeGetSystemTime.
    '
    'The native Windows MMTIME structure contains:
    '   wType As UINT
    '   union { ms/sample/cb/ticks/smpte/midi }
    '
    'The union must occupy 8 bytes because its largest member is the SMPTE
    'layout. VBA does not support C-style unions, so the safest compact mapping
    'for TIME_MS is:
    '   wType    => 4 bytes
    '   Units    => 4 bytes, used as the TIME_MS millisecond value
    '   Reserved => 4 bytes, pads the union to 8 bytes
    '
    'Total LenB(MMTIME) = 12 bytes.
        Private Type MMTIME
            wType       As Long     'Requested/returned time format
            Units       As Long     'Milliseconds when wType = CPM_TIME_MS
            Reserved    As Long     'Pads the native union to 8 bytes
        End Type

    'Structured checkpoint row captured within one active timing session
        Private Type tCheckpointRow
            Seq                     As Long      'Checkpoint sequence number
            NameText                As String    'Checkpoint label
            NoteText                As String    'Optional checkpoint note
            DeltaSecondsValue       As Double    'Elapsed seconds since previous checkpoint
            ElapsedSecondsValue     As Double    'Cumulative elapsed seconds since StartTimer
            MethodID                As cPM_TimerMethod   'Active method identifier
            MethodNameText          As String    'Active method descriptive name
            RunLabelText            As String    'Optional run label
        End Type

'
'------------------------------------------------------------------------------
'
'                               API DECLARATIONS
'
'------------------------------------------------------------------------------
'
    'Win32 / Win64 API declarations used for:
    '  - timing backends
    '  - timer-resolution management
    '  - coarse sleep / scheduler diagnostics
    '
    'Conditional compilation handles:
    '  - VBA7 vs pre-VBA7
    '  - Win64 vs Win32

    '---------------------------------------------------------------------------
    ' HIGH-RESOLUTION TIMING (KERNEL32)
    '   QueryPerformanceFrequency / QueryPerformanceCounter
    '---------------------------------------------------------------------------
        'QPC provides a high-resolution monotonic counter:
        ' - QueryPerformanceFrequency returns ticks per second
        ' - QueryPerformanceCounter returns the current tick value
        '
        'Currency is used as efficient fixed-point 64-bit storage in VBA and is
        'a common, stable representation for QPC values in VBA implementations.

        #If VBA7 Then
            Private Declare PtrSafe Function QueryPerformanceFrequency Lib _
                "kernel32" (ByRef Frequency As Currency) As Long

            Private Declare PtrSafe Function QueryPerformanceCounter Lib _
                "kernel32" (ByRef Counter As Currency) As Long
        #Else
            Private Declare Function QueryPerformanceFrequency Lib _
                "kernel32" (ByRef Frequency As Currency) As Long

            Private Declare Function QueryPerformanceCounter Lib "kernel32" _
                (ByRef Counter As Currency) As Long
        #End If

    '---------------------------------------------------------------------------
    ' MILLISECOND UPTIME COUNTERS (KERNEL32)
    '   GetTickCount / GetTickCount64
    '---------------------------------------------------------------------------
        'System uptime counter in milliseconds since boot:
        ' - Win64: GetTickCount64 is effectively non-rolling for practical use
        ' - Win32: GetTickCount has 32-bit rollover semantics

        #If VBA7 Then
            #If Win64 Then
                Private Declare PtrSafe Function GetTickCount64 Lib _
                    "kernel32" () As LongLong
            #Else
                Private Declare PtrSafe Function GetTickCount Lib "kernel32" _
                    () As Long
            #End If
        #Else
            Private Declare Function GetTickCount Lib "kernel32" () As Long
        #End If

    '---------------------------------------------------------------------------
    ' SLEEP / SYSTEM CLOCK INFO (KERNEL32)
    '   Sleep / GetSystemTimeAdjustment
    '---------------------------------------------------------------------------
        'Sleep suspends the current thread for the requested number of
        'milliseconds and is used for pause/coarse-wait behavior.

        #If VBA7 Then
            Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal _
                dwMilliseconds As Long)
        #Else
            Private Declare Sub Sleep Lib "kernel32" (ByVal _
                dwMilliseconds As Long)
        #End If

        'GetSystemTimeAdjustment exposes the nominal system clock increment.
        'lpTimeIncrement is used for coarse tick-interval diagnostics.

        #If VBA7 Then
            Private Declare PtrSafe Function GetSystemTimeAdjustment Lib _
                "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, _
                lpTimeAdjustmentDisabled As Long) As Long
        #Else
            Private Declare Function GetSystemTimeAdjustment Lib "kernel32" _
                (lpTimeAdjustment As Long, lpTimeIncrement As Long, _
                lpTimeAdjustmentDisabled As Long) As Long
        #End If

    '---------------------------------------------------------------------------
    ' MULTIMEDIA TIMERS (WINMM)
    '   timeGetTime / timeBeginPeriod / timeEndPeriod
    '---------------------------------------------------------------------------
        'Multimedia timer APIs:
        ' - timeGetTime       : millisecond counter with 32-bit rollover semantics
        ' - timeBeginPeriod(1): requests 1ms timer resolution (global effect)
        ' - timeEndPeriod(1)  : releases the requested timer resolution

        #If VBA7 Then
            Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () _
                As Long

            Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" _
                (ByVal uPeriod As Long) As Long

            Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" _
                (ByVal uPeriod As Long) As Long
        #Else
            Private Declare Function timeGetTime Lib "winmm.dll" () As Long

            Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal _
                uPeriod As Long) As Long

            Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal _
                uPeriod As Long) As Long
        #End If

    '---------------------------------------------------------------------------
    ' MULTIMEDIA SYSTEM TIME (WINMM)
    '   timeGetSystemTime using MMTIME
    '---------------------------------------------------------------------------
        'timeGetSystemTime returns its value through an MMTIME structure.
        'The class configures wType = 1 (TIME_MS), so Units is interpreted as
        'milliseconds.

        #If VBA7 Then
            Private Declare PtrSafe Function timeGetSystemTime Lib _
                "winmm.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long
        #Else
            Private Declare Function timeGetSystemTime Lib "winmm.dll" _
                (lpTime As MMTIME, ByVal uSize As Long) As Long
        #End If

'------------------------------------------------------------------------------
' HARDENING CONSTANTS
'------------------------------------------------------------------------------
    'TIME_MS format identifier for MMTIME
        Private Const CPM_TIME_MS                       As Long = 1

    'Finite-loop guard for tick-alignment polling
        Private Const CPM_ALIGN_MAX_SPINS               As Long = 1000000

    'Error numbers raised by this class.
    '
    'Public so callers can trap a specific condition by name rather than by a
    'magic offset. The values are unchanged from earlier releases, so code
    'already comparing against vbObjectError + N keeps working.
    '
    'The gaps between groups are deliberate: a new error can join any group
    'without disturbing the numbering of the others.
    Public Enum cPM_Error
        'Method validation and session binding
        ERR_CPM_INVALID_START_METHOD = vbObjectError + 1000
        ERR_CPM_QPC_UNAVAILABLE = vbObjectError + 1001
        ERR_CPM_NO_ACTIVE_SESSION = vbObjectError + 1002
        ERR_CPM_INVALID_ELAPSED_METHOD = vbObjectError + 1003
        ERR_CPM_METHOD_MISMATCH = vbObjectError + 1004
        ERR_CPM_RESOLVED_METHOD_INVALID = vbObjectError + 1005

        'Timing source and environment
        ERR_CPM_QPC_READ_FAILED = vbObjectError + 1007
        ERR_CPM_TIMEBEGINPERIOD_FAILED = vbObjectError + 1008
        ERR_CPM_TIMEENDPERIOD_FAILED = vbObjectError + 1009
        ERR_CPM_ALIGN_TICK_TIMEOUT = vbObjectError + 1010
        ERR_CPM_TIMEGETSYSTEMTIME_FAILED = vbObjectError + 1011
        ERR_CPM_TIMEGETSYSTEMTIME_FORMAT = vbObjectError + 1012
        ERR_CPM_NEGATIVE_ELAPSED = vbObjectError + 1013
        ERR_CPM_START_CAPTURE_FAILED = vbObjectError + 1014

        'Checkpoints and structured reporting
        ERR_CPM_CHECKPOINT_NO_SESSION = vbObjectError + 1020
        ERR_CPM_RUNLABEL_AFTER_CAPTURE = vbObjectError + 1021

        'Statistics and the measurement harness
        ERR_CPM_STATS_NO_SAMPLES = vbObjectError + 1030
        ERR_CPM_STATS_BAD_PERCENTILE = vbObjectError + 1031
        ERR_CPM_MEASURE_BLANK_PROC = vbObjectError + 1032
        ERR_CPM_STATS_UNDEFINED_CV = vbObjectError + 1033
        ERR_CPM_STATS_BAD_CV_THRESHOLD = vbObjectError + 1034
        ERR_CPM_MEASURE_NO_BASELINE = vbObjectError + 1035
        ERR_CPM_MEASURE_NO_VALID_SAMPLES = vbObjectError + 1036
        ERR_CPM_STATS_OUT_OF_DOMAIN = vbObjectError + 1037
    End Enum

    'Coefficient-of-variation above which a sample set is reported as noisy
        Private Const CPM_CV_CONTAMINATION_THRESHOLD    As Double = 0.25
'------------------------------------------------------------------------------
' STATE
'------------------------------------------------------------------------------
    'Internal class state for:
    '  - active timing sessions and elapsed-time computation
    '  - QPC and non-QPC timing infrastructure
    '  - shared TW participation for this instance

    '--------------------------------------------------------------------------
    ' METHOD NAMES (1..6)
    '--------------------------------------------------------------------------
        Private m_MethodName(1 To 6) As String   'Human-readable timing backend labels

    '--------------------------------------------------------------------------
    ' GENERAL TIMING STATE (METHODS 1 / 2 / 3 / 4 / 6)
    '--------------------------------------------------------------------------
        Private m_T1    As Double    'Start timestamp for the active session
        Private m_T2    As Double    'End timestamp captured during elapsed-time evaluation
        Private m_ET    As Double    'Last computed elapsed time in seconds

    '--------------------------------------------------------------------------
    ' ACTIVE TIMER SESSION STATE
    '--------------------------------------------------------------------------
        Private m_ActiveMethodID  As cPM_TimerMethod   'Backend currently bound to the active session
        Private m_HasActiveSession  As Boolean   'TRUE once StartTimer has established a session
        Private m_QpcAvailable  As Boolean   'TRUE when QPC is available on this machine

        'Strict mode:
        ' - TRUE  => raise on invalid / mismatched timer usage
        ' - FALSE => coerce / fallback where allowed
        Private m_StrictMode    As Boolean

    '--------------------------------------------------------------------------
    ' QPC TIMING STATE (METHOD 5)
    '--------------------------------------------------------------------------
        'QPC values are stored as Currency to preserve stable tick precision in
        'VBA while keeping the implementation efficient.
        Private m_qpcStart  As Currency      'QPC start tick
        Private m_qpcEnd    As Currency      'QPC end tick
        Private m_qpcFreq   As Currency      'QPC frequency (ticks per second)

    '--------------------------------------------------------------------------
    ' MMTIME INSTANCE FOR timeGetSystemTime
    '--------------------------------------------------------------------------
        'Reusable structure configured once and reused for method 4.
        Private m_Mmt       As MMTIME

    '--------------------------------------------------------------------------
    ' timeBeginPeriod(1) LIFECYCLE GUARD (METHOD 3)
    '--------------------------------------------------------------------------
        'Tracks whether this instance successfully requested 1ms multimedia timer
        'resolution and therefore must later balance it with timeEndPeriod(1).
        Private m_TimePeriodOn As Boolean

    '--------------------------------------------------------------------------
    ' SHARED "TIME WASTER" SESSION STATE
    '--------------------------------------------------------------------------
        Private m_TW_Active   As Boolean   'TRUE when this instance is registered in shared TW suppression
        Private m_Except      As TW_Enum   'Current exemption mask requested by this instance

        'Collision-proof shared-TW key for this instance.
        'Assigned lazily on first use so instances that never request TW
        'suppression consume no key.
        Private m_InstanceKey As String

    '--------------------------------------------------------------------------
    ' CHECKPOINT / STRUCTURED REPORT STATE
    '--------------------------------------------------------------------------
        Private m_RunLabel              As String                'Optional run label for the current timing session
        Private m_CheckpointCount       As Long                  'Number of captured checkpoints
        Private m_LastCheckpointS       As Double                'Last cumulative checkpoint elapsed seconds
        Private m_Checkpoints()         As tCheckpointRow        'Checkpoint storage array

    '--------------------------------------------------------------------------
    ' NATIVE READ STATUS
    '--------------------------------------------------------------------------
        'Outcome of the most recent native read performed by this instance.
        'Reset at the start of every public timing operation.
        Private m_LastReadStatus As cPM_ReadStatus

    '--------------------------------------------------------------------------
    ' FAULT INJECTION (REGRESSION SUITE ONLY)
    '--------------------------------------------------------------------------
        'One-shot flags used by the regression suite to force a native read to
        'fail. Each flag is cleared by the read that consumes it, so a forced
        'failure can never leak into a later test case.
        Private m_TestFailNextQpcRead        As Boolean
        Private m_TestFailNextSystemTimeRead As Boolean
        Private m_TestForceNextSystemTimeFormatInvalid As Boolean
        Private m_TestForceNextNegativeElapsed As Boolean
        'Measured-cycle worker seams. Start and endpoint are counted separately
        'because they arm at different moments and prove different things: a
        'start failure produces a fallback the harness must reject, an endpoint
        'failure produces a read the harness must exclude.
        Private m_TestWorkerStartFailures As Long
        Private m_TestWorkerEndpointFailures As Long

'
'------------------------------------------------------------------------------
'
'                                   LIFECYCLE
'
'------------------------------------------------------------------------------
'

Private Sub Class_Initialize()
'
'==============================================================================
'                             CLASS INITIALIZE
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes the performance manager class state:
'     - Captures and validates the QPC frequency once.
'     - Initializes the MMTIME structure for timeGetSystemTime.
'     - Populates the method name table.
'     - Resets lifecycle/session flags deterministically.
'
' WHY THIS EXISTS
'   This routine establishes a known, reproducible baseline for each new
'   cPerformanceManager instance so that timing sessions and TW lifecycle
'   handling always begin from a clean state.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Queries QueryPerformanceFrequency once and records availability.
'   - Configures MMTIME to return milliseconds.
'   - Loads the timer method labels used by MethodName().
'   - Resets timing-session state.
'   - Resets TW lifecycle flags for this instance.
'   - Delegates checkpoint / report state setup to Checkpoint_ClearState.
'
' ERROR POLICY
'   Does not raise an error if QPC is unavailable.
'   Instead, records availability in m_QpcAvailable for later validation.
'
' DEPENDENCIES
'   - QueryPerformanceFrequency
'   - MMTIME
'   - TW_Enum
'   - Checkpoint_ClearState
'
' NOTES
'   Checkpoint state is initialized through Checkpoint_ClearState rather than
'   duplicated here, so there is exactly one definition of a clean checkpoint
'   baseline.
'
'   m_InstanceKey is deliberately left blank. TW_InstanceKey assigns it lazily
'   on first use.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC      As Long         'Return code from QueryPerformanceFrequency

'------------------------------------------------------------------------------
' INITIALIZE (TIMING INFRASTRUCTURE)
'------------------------------------------------------------------------------
    'Capture the QPC frequency once and record backend availability
        RC = QueryPerformanceFrequency(m_qpcFreq)
        m_QpcAvailable = (RC <> 0 And m_qpcFreq <> 0)
    'Configure the reusable MMTIME structure for millisecond output
        m_Mmt.wType = CPM_TIME_MS

    'Load the human-readable timing backend labels
        m_MethodName(cPM_MethodTimer) = "Timer"
        m_MethodName(cPM_MethodTickCount) = "GetTickCount / GetTickCount64"
        m_MethodName(cPM_MethodTimeGetTime) = "timeGetTime"
        m_MethodName(cPM_MethodSystemTime) = "timeGetSystemTime"
        m_MethodName(cPM_MethodQPC) = "QPC"
        m_MethodName(cPM_MethodNow) = "Now()"

'------------------------------------------------------------------------------
' INITIALIZE (SESSION / LIFECYCLE STATE)
'------------------------------------------------------------------------------
    'Guard for timeBeginPeriod/timeEndPeriod lifecycle
        m_TimePeriodOn = False
    'Default exemption mask for TW_Turn_OFF
        m_Except = TW_Enum.None
    'Initialize active timer session state
        m_ActiveMethodID = 0
    'Mark that no timer session is active yet
        m_HasActiveSession = False
    'Default strict-mode policy
        m_StrictMode = True
    'Clear cached generic timing values deterministically
        m_T1 = 0#
        m_T2 = 0#
        m_ET = 0#
    'Clear cached QPC timing values deterministically
        m_qpcStart = 0@
        m_qpcEnd = 0@
    'Indicates whether this instance currently owns an active shared TW session
        m_TW_Active = False
    'Leave the shared-TW key unassigned until first use
        m_InstanceKey = vbNullString

'------------------------------------------------------------------------------
' INITIALIZE (CHECKPOINT / REPORT STATE)
'------------------------------------------------------------------------------
    'Establish a clean checkpoint baseline through the single shared helper
        Checkpoint_ClearState

End Sub

Private Sub Class_Terminate()
'
'==============================================================================
'                              CLASS TERMINATE
'------------------------------------------------------------------------------
' PURPOSE
'   Final safety-net cleanup for environment changes made by this instance.
'
' WHY THIS EXISTS
'   Class_Terminate is the last opportunity for this instance to release
'   environment changes if the caller forgot to invoke ResetEnvironment
'   explicitly.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Calls ResetEnvironment as a best-effort cleanup step during object teardown.
'
' ERROR POLICY
'   Suppresses teardown errors to avoid destabilizing object termination.
'
' DEPENDENCIES
'   - ResetEnvironment
'
' NOTES
'   ResetEnvironment remains the primary cleanup contract.
'   Class_Terminate is only the final fallback.
'
'   Errors here are necessarily invisible, because a terminator cannot raise.
'   A hard End statement bypasses this routine entirely, which is why
'   PM_TW_EndAllSessions exists as the documented manual recovery path.
'
' UPDATED
'   2026-08-15
'==============================================================================

    'Destructor cleanup must be best-effort and non-fatal
        On Error Resume Next
    'Use the explicit cleanup routine as the termination safety net
        ResetEnvironment
    'Return to normal error handling semantics
        On Error GoTo 0
End Sub

'
'------------------------------------------------------------------------------
'
'                           PUBLIC: CORE TIMING API
'
'------------------------------------------------------------------------------
'

Public Sub StartTimer(Optional ByVal iMethod As cPM_TimerMethod = cPM_MethodQPC, _
                      Optional ByVal AlignToNextTick As Boolean = False, _
                      Optional ByVal RunLabel As String = vbNullString)
'
'==============================================================================
'                                  START TIMER
'------------------------------------------------------------------------------
' PURPOSE
'   Starts a timing session and binds the selected backend to it.
'
' WHY THIS EXISTS
'   Timing must begin from a well-defined backend, and that backend must remain
'   bound to the session so later elapsed reads can be validated against it.
'
' INPUTS
'   iMethod (optional)
'     Requested timing backend.
'
'   AlignToNextTick (optional)
'     TRUE  => wait for the next observable tick before capturing the start
'     FALSE => capture the start timestamp immediately
'
'   RunLabel (optional)
'     Optional run label. Applied after the checkpoint reset, which is the
'     reliable way to label a run. Calling SetRunLabel BEFORE StartTimer has no
'     effect, because the reset discards it.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Normalizes and validates the requested method.
'   - Attempts to capture a start timestamp for that method.
'   - On failure in strict mode, raises and leaves the previous session intact.
'   - On failure in non-strict mode, falls back to method 2 and captures again
'     BEFORE anything is committed.
'   - Commits class state only once a valid start timestamp exists.
'   - Starts a fresh checkpoint session and applies the optional run label.
'
' ERROR POLICY
'   - Method validation is delegated to Method_NormalizeStart.
'   - A failed start capture raises in strict mode, before any state changes.
'
' CRITICAL INVARIANT
'   No session is ever committed whose start value was not successfully read.
'
'   The previous implementation committed m_qpcStart = 0 after a failed QPC read
'   in non-strict mode. A subsequent successful end read then measured from zero
'   and reported something resembling machine uptime.
'
' NON-STRICT FALLBACK
'   When the requested backend cannot be read, non-strict mode binds method 2 and
'   records cPM_ReadFallbackToMethod2 in LastReadStatus. The resulting session is
'   fully valid; the status exists so the caller can see the coercion rather than
'   having it happen silently.
'
' DEPENDENCIES
'   - Method_NormalizeStart
'   - Start_TryCapture
'   - Checkpoint_ClearState
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NewMethod       As cPM_TimerMethod   'Method to bind to the new session
    Dim NewT1           As Double           'Generic start timestamp
    Dim NewQpcStart     As Currency         'QPC start tick for method 5
    Dim CaptureStatus   As cPM_ReadStatus   'Outcome of the capture attempt
    Dim FinalStatus     As cPM_ReadStatus   'Status published to the caller

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'This operation describes its own read outcome
        FinalStatus = cPM_ReadOK

    'Normalize and validate the requested method
        NewMethod = Method_NormalizeStart(iMethod)

'------------------------------------------------------------------------------
' CAPTURE START TIMESTAMP BEFORE COMMITTING ANY STATE
'------------------------------------------------------------------------------
    'Attempt the requested backend
        If Not Start_TryCapture(NewMethod, AlignToNextTick, _
                                NewT1, NewQpcStart, CaptureStatus) Then

            'Strict mode: fail without disturbing the existing session
                If m_StrictMode Then
                    m_LastReadStatus = CaptureStatus
                    Err.Raise Read_StatusToErrorNumber(CaptureStatus), _
                              "cPerformanceManager.StartTimer", _
                              Read_StatusToText(CaptureStatus) & _
                              " The timing session was not started."
                End If

            'Non-strict mode: fall back to method 2 BEFORE committing anything
                NewMethod = cPM_MethodTickCount
                NewQpcStart = 0@

                If Not Start_TryCapture(NewMethod, AlignToNextTick, _
                                        NewT1, NewQpcStart, CaptureStatus) Then
                    'Method 2 has no failure path; this is defensive only
                        m_LastReadStatus = CaptureStatus
                        Err.Raise ERR_CPM_START_CAPTURE_FAILED, _
                                  "cPerformanceManager.StartTimer", _
                                  "No timing backend could supply a valid start timestamp."
                End If

            'Record that the requested backend was coerced
                FinalStatus = cPM_ReadFallbackToMethod2
        End If

'------------------------------------------------------------------------------
' COMMIT NEW SESSION STATE
'------------------------------------------------------------------------------
    'Bind the backend that actually produced the start timestamp
        m_ActiveMethodID = NewMethod
    'Mark the timer session as started
        m_HasActiveSession = True
    'Store the validated start timestamp
        m_T1 = NewT1
    'Clear cached end/elapsed values for this new session
        m_T2 = 0#
        m_ET = 0#
        m_qpcEnd = 0@
    'Store the QPC start tick only for a genuine QPC session
        If NewMethod = cPM_MethodQPC Then
            m_qpcStart = NewQpcStart
        Else
            m_qpcStart = 0@
        End If
    'Publish the outcome of this operation
        m_LastReadStatus = FinalStatus
    'Start a fresh checkpoint session for the new timing run
        Checkpoint_ClearState

'------------------------------------------------------------------------------
' APPLY OPTIONAL RUN LABEL
'------------------------------------------------------------------------------
    'Apply the supplied run label after the checkpoint reset has cleared it
        If Len(RunLabel) > 0 Then
            m_RunLabel = RunLabel
        End If

End Sub


Private Function Start_TryCapture( _
    ByVal MethodID As cPM_TimerMethod, _
    ByVal AlignToNextTick As Boolean, _
    ByRef T1Out As Double, _
    ByRef QpcOut As Currency, _
    ByRef StatusOut As cPM_ReadStatus) _
    As Boolean
'
'==============================================================================
'                              START TRY CAPTURE
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to capture a start timestamp for one backend, without touching any
'   class state.
'
' WHY THIS EXISTS
'   StartTimer needs to try a backend, and possibly a second one, before
'   committing anything. Isolating the capture keeps that logic readable and
'   guarantees a failed attempt has no side effects.
'
' INPUTS
'   MethodID          - backend to attempt
'   AlignToNextTick   - whether to align the capture to the next tick
'   T1Out             - receives the generic start timestamp, on success only
'   QpcOut            - receives the QPC start tick for method 5, on success only
'   StatusOut         - receives the failure reason, or cPM_ReadOK
'
' RETURNS
'   Boolean
'     TRUE  => a valid start timestamp was captured
'     FALSE => capture failed; the outputs are untouched
'
' ERROR POLICY
'   Read failures are reported through the return value. A strict-mode alignment
'   timeout may still raise from the alignment helpers.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Assume success until a read reports otherwise
        StatusOut = cPM_ReadOK

'------------------------------------------------------------------------------
' QPC PATH
'------------------------------------------------------------------------------
    'Method 5 works in Currency ticks and has its own capture path
        If MethodID = cPM_MethodQPC Then
            If AlignToNextTick Then
                If Not QPC_TryGetNextTick(QpcOut) Then
                    StatusOut = cPM_ReadQpcFailed
                    Start_TryCapture = False
                    Exit Function
                End If
            Else
                If Not QPC_TryReadTick(QpcOut) Then
                    StatusOut = cPM_ReadQpcFailed
                    Start_TryCapture = False
                    Exit Function
                End If
            End If

            T1Out = CDbl(QpcOut)
            Start_TryCapture = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' NON-QPC PATH
'------------------------------------------------------------------------------
    'Every other backend works in seconds
        If AlignToNextTick Then
            Start_TryCapture = Timestamp_TryGetNextSeconds(MethodID, T1Out, StatusOut)
        Else
            Start_TryCapture = Timestamp_TryGetCurrentSeconds(MethodID, T1Out, StatusOut)
        End If

End Function


Public Property Get ElapsedSeconds(Optional ByVal iMethod As cPM_TimerMethod = cPM_MethodSession) As Double
'
'==============================================================================
'                               ELAPSED SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the elapsed time in seconds for the active timing session.
'
' INPUTS
'   iMethod (optional)
'     0   => use the active timer session method
'     1-6 => explicit method to validate against the active session
'
' RETURNS
'   Double
'     Elapsed seconds since StartTimer, or 0 when a read failed in non-strict
'     mode.
'
' BEHAVIOR
'   - Resets LastReadStatus, so it always describes this read.
'   - Resolves and validates the method against the active session.
'   - Delegates the arithmetic to Elapsed_ComputeSeconds.
'   - Caches m_T2, m_qpcEnd and m_ET ONLY when the read succeeded.
'
' ERROR POLICY
'   - Session and method validation is delegated to Method_ResolveElapsed.
'   - A failed read publishes LastReadStatus first, then raises in strict mode
'     or returns 0 in non-strict mode.
'
' NOTES
'   A returned 0 with LastReadStatus <> cPM_ReadOK is NOT a measurement of zero
'   seconds. It means no valid reading was obtained. Callers that need to tell
'   the two apart must check LastReadStatus.
'
'   LastReadStatus is assigned before any raise, so it describes this read on
'   both paths. A caller that traps a strict-mode error can inspect it for the
'   specific condition rather than decoding the error number.
'
'   On a failed read the cached T2 and ET are deliberately left untouched, so a
'   failure cannot destroy the last known good measurement.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResolvedMethod  As cPM_TimerMethod          'Validated method for this read
    Dim EndStamp        As Double           'Raw end timestamp
    Dim QpcEnd          As Currency         'Raw QPC end tick
    Dim Seconds         As Double           'Computed elapsed seconds
    Dim Status          As cPM_ReadStatus   'Outcome of the read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'This read describes its own outcome
        m_LastReadStatus = cPM_ReadOK
        EndStamp = m_T2
        QpcEnd = m_qpcEnd

'------------------------------------------------------------------------------
' RESOLVE METHOD
'------------------------------------------------------------------------------
    'Resolve and validate the method against the active session
        ResolvedMethod = Method_ResolveElapsed(iMethod)

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Compute elapsed seconds through the single shared backend reader
        Seconds = Elapsed_ComputeSeconds(ResolvedMethod, EndStamp, QpcEnd, Status)

'------------------------------------------------------------------------------
' HANDLE A FAILED READ
'------------------------------------------------------------------------------
    'Publish the outcome BEFORE deciding whether to fail, so a caller that traps
    'a strict-mode error still finds the status describing this read
        If Status <> cPM_ReadOK Then
            m_LastReadStatus = Status

            If m_StrictMode Then
                Err.Raise Read_StatusToErrorNumber(Status), _
                          "cPerformanceManager.ElapsedSeconds", _
                          Read_StatusToText(Status) & _
                          " No elapsed time could be computed."
            End If

            'Non-strict: preserve the last known good cache and return zero
                ElapsedSeconds = 0#
                Exit Property
        End If

'------------------------------------------------------------------------------
' CACHE AND ASSIGN RESULT
'------------------------------------------------------------------------------
    'Cache the raw endpoints and the computed elapsed value
        m_T2 = EndStamp
        m_qpcEnd = QpcEnd
        m_ET = Seconds

    'Return the computed elapsed seconds
        ElapsedSeconds = Seconds

End Property


Public Function ElapsedTime( _
    Optional ByVal iMethod As cPM_TimerMethod = cPM_MethodSession, _
    Optional ByVal ElapsedSecondsIn As Variant) _
    As String
'
'==============================================================================
'                                ELAPSED TIME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a human-readable elapsed time string:
'     H:nn:ss - mmm milliseconds - uuu microseconds - nnn nanoseconds
'
' WHY THIS EXISTS
'   This routine supports two related use cases:
'
'     - when no numeric elapsed value is supplied, it measures elapsed time
'       through ElapsedSeconds()
'
'     - when a numeric elapsed value is already available, it formats that
'       value directly without taking a second timing sample
'
'   This avoids the double-measurement problem in callers that already captured
'   ElapsedSeconds() separately for logging or further numeric processing
'
' INPUTS
'   iMethod (optional)
'     0   => use the active timer session method
'     1-6 => explicit method to validate against the active session
'
'   ElapsedSecondsIn (optional)
'     Optional numeric elapsed-seconds value to format directly
'
' RETURNS
'   String
'     The formatted elapsed time string
'
' BEHAVIOR
'   - If ElapsedSecondsIn is omitted, measures elapsed time through
'     ElapsedSeconds(iMethod)
'   - If ElapsedSecondsIn is supplied, formats that value directly
'   - Splits elapsed time into whole seconds and fractional seconds
'   - Decomposes the fractional part into millisecond / microsecond /
'     nanosecond display groups
'   - Formats the whole-second duration manually so hours do not wrap at 24
'
' ERROR POLICY
'   Delegates timer/session validation to ElapsedSeconds() when measurement is
'   requested
'
' DEPENDENCIES
'   - ElapsedSeconds (only when ElapsedSecondsIn is omitted)
'
' NOTES
'   - Nanosecond output is presentation precision only and does not imply actual
'     measurement resolution
'   - When ElapsedSecondsIn is supplied, no new timing sample is taken
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sec         As Double   'Elapsed seconds to format
    Dim WholeS      As Long     'Whole seconds component (floor)
    Dim Frac        As Double   'Fractional seconds component in [0,1)
    Dim NS_Total    As Long     'Fractional part scaled to nanoseconds
    Dim Ms          As Long     'Milliseconds component
    Dim Us          As Long     'Microseconds component
    Dim NN          As Long     'Nanoseconds component
    Dim Hours       As Long     'Whole hours in the elapsed duration
    Dim Minutes     As Long     'Minute component
    Dim Seconds     As Long     'Second component
    Dim sMS         As String   'Zero-padded milliseconds text
    Dim sUS         As String   'Zero-padded microseconds text
    Dim sNN         As String   'Zero-padded nanoseconds text

'------------------------------------------------------------------------------
' RESOLVE ELAPSED SECONDS TO FORMAT
'------------------------------------------------------------------------------
    'Use the supplied elapsed-seconds value when provided
        If Not IsMissing(ElapsedSecondsIn) Then
            Sec = CDbl(ElapsedSecondsIn)
    'Otherwise measure elapsed seconds now
        Else
            Sec = ElapsedSeconds(iMethod)
        End If

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Defensively clamp negative elapsed values
        If Sec < 0# Then Sec = 0#

'------------------------------------------------------------------------------
' SPLIT WHOLE / FRACTIONAL SECONDS
'------------------------------------------------------------------------------
    'Extract whole seconds and the fractional remainder
        WholeS = CLng(Fix(Sec))
        Frac = Sec - CDbl(WholeS)
    'Defensive clamp for negative fractional jitter
        If Frac < 0# Then Frac = 0#

'------------------------------------------------------------------------------
' CONVERT FRACTIONAL SECONDS TO ns/ms/us/nnn
'------------------------------------------------------------------------------
    'Convert fractional seconds to integer nanoseconds using floor semantics
        NS_Total = CLng(Fix(Frac * 1000000000#))
    'Defensively clamp unexpected negative values
        If NS_Total < 0 Then
            NS_Total = 0
        End If
    'Defensively clamp unexpected overflow to one-second boundary
        If NS_Total > 999999999 Then
            NS_Total = 999999999
        End If
    'Decompose nanoseconds into mmm / uuu / nnn display groups
        Ms = NS_Total \ 1000000
        Us = (NS_Total \ 1000) Mod 1000
        NN = NS_Total Mod 1000

'------------------------------------------------------------------------------
' DECOMPOSE WHOLE-SECOND DURATION
'------------------------------------------------------------------------------
    'Compute duration hours/minutes/seconds without 24-hour wraparound
        Hours = WholeS \ 3600
        Minutes = (WholeS \ 60) Mod 60
        Seconds = WholeS Mod 60

'------------------------------------------------------------------------------
' FORMAT COMPONENTS
'------------------------------------------------------------------------------
    'Format each sub-second group as exactly three digits
        sMS = Right$("000" & CStr(Ms), 3)
        sUS = Right$("000" & CStr(Us), 3)
        sNN = Right$("000" & CStr(NN), 3)

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the final formatted duration string
        ElapsedTime = CStr(Hours) & ":" & _
                      Right$("00" & CStr(Minutes), 2) & ":" & _
                      Right$("00" & CStr(Seconds), 2) & " - " & _
                      sMS & " milliseconds - " & _
                      sUS & " microseconds - " & _
                      sNN & " nanoseconds"

End Function

'
'------------------------------------------------------------------------------
'
'                       PUBLIC: SESSION / STATE INSPECTION
'
'------------------------------------------------------------------------------
'

Public Property Get T1() As Double
'
'==============================================================================
'                                      T1
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last captured raw start timestamp for the most recent timing
'   session established by StartTimer().
'
' WHY THIS EXISTS
'   Useful for diagnostics, verification, and unit tests when the caller wants
'   to inspect the internal raw start capture without triggering formatting or
'   recomputation logic.
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     The internal m_T1 value exactly as currently stored.
'
' BEHAVIOR
'   - Returns the raw start timestamp without modification.
'   - The meaning of the value depends on the active timing backend used when
'     StartTimer() was called:
'       * Method 1 (Timer)             : seconds since midnight
'       * Method 2 (GetTickCount / GetTickCount64) : seconds since boot (wrapper-normalized)
'       * Method 3 (timeGetTime)       : seconds since boot (wrapper-normalized)
'       * Method 4 (timeGetSystemTime) : seconds from a millisecond source
'       * Method 5 (QPC)               : CDbl(m_qpcStart), i.e. raw QPC ticks
'       * Method 6 (Now)               : Now() * 86400 as stored by the class
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   - m_T1
'
' NOTES
'   - Prefer ElapsedSeconds() for the canonical numeric elapsed-time result.
'   - Prefer ElapsedTime() for display-oriented diagnostics.
'   - T1 is an inspection/debugging surface, not the primary timing API.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the raw start timestamp.
        T1 = m_T1
End Property

Public Property Get T2() As Double
'
'==============================================================================
'                                      T2
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last captured raw end timestamp for the most recent timing
'   session as evaluated by ElapsedSeconds() or by ElapsedTime() when that
'   function internally performs a timing read
'
' WHY THIS EXISTS
'   Useful for diagnostics, verification, and unit tests when the caller wants
'   to inspect the internal raw end capture directly
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     The internal m_T2 value exactly as currently stored.
'
' BEHAVIOR
'   - Returns the raw end timestamp without modification.
'   - m_T2 is updated when ElapsedSeconds() is evaluated.
'   - m_T2 is also updated when ElapsedTime() is called without
'     ElapsedSecondsIn, because that path internally measures elapsed time.
'   - m_T2 is also updated by Checkpoint(), which measures internally.
'   - When ElapsedTime() is called with ElapsedSecondsIn, no new timing sample
'     is taken and m_T2 is not updated by that formatting step.
'   - The meaning of the value depends on the timing backend:
'       * Methods 1 / 2 / 3 / 4 / 6 : stored in seconds
'       * Method 5 (QPC)             : stored as CDbl(m_qpcEnd), i.e. raw QPC ticks
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   - m_T2
'
' NOTES
'   - If ElapsedSeconds() has not yet been called after StartTimer(), T2 may
'     still reflect a prior session or its reset value.
'   - T2 is an inspection/debugging surface, not the primary timing API.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the raw end timestamp.
        T2 = m_T2
End Property

Public Property Get ET() As Double
'
'==============================================================================
'                                      ET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last computed elapsed time in seconds for the most recent
'   timing session
'
' WHY THIS EXISTS
'   Allows callers, diagnostics, and tests to inspect the cached elapsed value
'   without re-running timing logic or invoking formatting code
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     The internal m_ET value exactly as currently stored.
'
' BEHAVIOR
'   - Returns the cached elapsed-seconds value without modification.
'   - m_ET is set by ElapsedSeconds().
'   - m_ET is also updated when ElapsedTime() is called without
'     ElapsedSecondsIn, because that path internally measures elapsed time.
'   - m_ET is also updated by Checkpoint(), which measures internally.
'   - When ElapsedTime() is called with ElapsedSecondsIn, the supplied value is
'     formatted directly and m_ET is not updated by that formatting step.
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   - m_ET
'
' NOTES
'   - For the canonical numeric timing result, prefer calling
'     ElapsedSeconds() directly.
'   - ET is an inspection/debugging surface and cache exposure, not the
'     primary timing API.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the cached elapsed seconds.
        ET = m_ET
End Property

Public Property Get StrictMode() As Boolean
'------------------------------------------------------------------------------
' STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the class raises errors for invalid/mismatched timer usage.
'------------------------------------------------------------------------------
    'Return the current strict-mode flag.
        StrictMode = m_StrictMode
End Property

Public Property Let StrictMode(ByVal Value As Boolean)
'------------------------------------------------------------------------------
' STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Sets whether the class raises errors for invalid/mismatched timer usage.
'------------------------------------------------------------------------------
    'Persist the strict-mode flag.
        m_StrictMode = Value
End Property

Public Property Get ActiveMethodID() As cPM_TimerMethod
'------------------------------------------------------------------------------
'                               ACTIVE METHOD ID
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the active timer method selected by the most recent StartTimer().
'------------------------------------------------------------------------------
    'Return the current active method.
        ActiveMethodID = m_ActiveMethodID
End Property

Public Property Get HasActiveSession() As Boolean
'------------------------------------------------------------------------------
'                              HAS ACTIVE SESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when a timer session is currently active.
'------------------------------------------------------------------------------
    'Return the current session-start flag.
        HasActiveSession = m_HasActiveSession
End Property

Public Property Get MethodName( _
    ByVal Idx As cPM_TimerMethod) _
    As String
'
'==============================================================================
'                                 METHOD NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the human-readable name of a timing backend by method index.
'
' WHY THIS EXISTS
'   Provides a stable mapping from the numeric timing-method identifier used by
'   StartTimer() / ElapsedSeconds() to a readable backend label for diagnostics,
'   reporting, and UI/debug output.
'
' INPUTS
'   Idx
'     Timing method index:
'       1 = "Timer"
'       2 = "GetTickCount / GetTickCount64"
'       3 = "timeGetTime"
'       4 = "timeGetSystemTime"
'       5 = "QPC"
'       6 = "Now()"
'
' RETURNS
'   String
'     The corresponding method label when Idx is valid (1..6);
'     otherwise vbNullString.
'
' BEHAVIOR
'   - Validates the supplied method index.
'   - Returns the stored label from m_MethodName() when the index is valid.
'   - Returns vbNullString for out-of-range indices.
'
' ERROR POLICY
'   Does not raise errors.
'   Invalid indices return vbNullString.
'
' DEPENDENCIES
'   - m_MethodName
'
' NOTES
'   This property is intended for display/diagnostic use and does not perform
'   timing-method normalization.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Guard against invalid indices to avoid subscript errors
        If (Idx < cPM_MethodTimer) Or (Idx > cPM_MethodNow) Then
            MethodName = vbNullString
            Exit Property
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the corresponding method label.
        MethodName = m_MethodName(Idx)
End Property

Public Property Get RunLabel() As String
'
'==============================================================================
'                                  RUN LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current run label stored on this instance
'
' WHY THIS EXISTS
'   Structured reporting helpers and diagnostics may need to inspect the
'   current run label without exporting the full checkpoint report
'
' INPUTS
'   None
'
' RETURNS
'   String
'     Current run label
'
' ERROR POLICY
'   Does not raise errors
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the current run label
        RunLabel = m_RunLabel

End Property

'
'------------------------------------------------------------------------------
'
'                   PUBLIC: DIAGNOSTICS / BENCHMARK HELPERS
'
'------------------------------------------------------------------------------
'

Public Function OverheadMeasurement_Seconds( _
    Optional ByVal iMethod As cPM_TimerMethod = cPM_MethodQPC, _
    Optional ByVal Iterations As Long = 1000) _
    As Double
'
'==============================================================================
'                        OVERHEAD MEASUREMENT IN SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Estimates the average near-empty measurement overhead for the selected timer
'   backend across multiple iterations.
'
' WHY THIS EXISTS
'   This function provides the numeric benchmark primitive behind
'   OverheadMeasurement_Text().
'
' INPUTS
'   iMethod (optional)
'     Timing backend to benchmark.
'
'   Iterations (optional)
'     Number of near-empty timing cycles to average.
'
' RETURNS
'   Double
'     Average near-empty measurement overhead in seconds.
'
' BEHAVIOR
'   - Normalizes and validates the requested timer method.
'   - Warms up the timing path once before the measured loop.
'   - Repeats a minimal StartTimer / ElapsedSeconds cycle.
'   - Always attempts to clean up the isolated worker instance.
'
' ERROR POLICY
'   Raises validation/runtime errors normally after cleanup has been attempted.
'
' DEPENDENCIES
'   - Method_NormalizeStart
'   - StartTimer
'   - ElapsedSeconds
'   - ResetEnvironment
'
' NOTES
'   The measurement is performed on an isolated temporary cPerformanceManager
'   instance so that the caller's checkpoint/report state is not disturbed.
'
'   This routine reports the arithmetic MEAN only. Timing distributions are
'   right-skewed, so a single scheduler preemption during the loop shifts the
'   result and the caller cannot distinguish a clean run from a contaminated
'   one. Treat the value as indicative rather than as a reliable floor.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K                   As Long                  'Loop counter
    Dim TotalS              As Double                'Accumulated elapsed seconds
    Dim ResolvedMethod      As cPM_TimerMethod               'Validated timer method
    Dim DiscardedElapsedS   As Double                'Throwaway elapsed value used only for warm-up
    Dim cPM_Work            As cPerformanceManager   'Isolated worker instance
    Dim ErrNumber           As Long                  'Captured error number
    Dim ErrSource           As String                'Captured error source
    Dim ErrDescription      As String                'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route runtime failures through cleanup
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Enforce a minimum of one iteration
        If Iterations < 1 Then Iterations = 1
    'Normalize and validate the timer method using the current instance policy
        ResolvedMethod = Method_NormalizeStart(iMethod)

'------------------------------------------------------------------------------
' INITIALIZE ISOLATED WORKER
'------------------------------------------------------------------------------
    'Create an isolated worker instance
        Set cPM_Work = New cPerformanceManager
    'Mirror the current strict-mode policy into the isolated worker
        cPM_Work.StrictMode = m_StrictMode

'------------------------------------------------------------------------------
' WARM UP
'------------------------------------------------------------------------------
    'Warm up the timing path once before measurement
        cPM_Work.StartTimer ResolvedMethod, False
        DiscardedElapsedS = cPM_Work.ElapsedSeconds()

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Run a repeated near-empty timing loop and accumulate the measured values
        For K = 1 To Iterations
            cPM_Work.StartTimer ResolvedMethod, False
            TotalS = TotalS + cPM_Work.ElapsedSeconds()
        Next K

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the average measured overhead
        OverheadMeasurement_Seconds = TotalS / CDbl(Iterations)

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Release the isolated worker instance
        On Error Resume Next
        If Not cPM_Work Is Nothing Then cPM_Work.ResetEnvironment
        Set cPM_Work = Nothing
        On Error GoTo 0
        Exit Function

Err_Handler:
'------------------------------------------------------------------------------
' CAPTURE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Capture the original error before cleanup
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description

'------------------------------------------------------------------------------
' CLEAN UP AFTER FAILURE
'------------------------------------------------------------------------------
    'Cleanup must be best-effort so that we can preserve the original error
        On Error Resume Next
        If Not cPM_Work Is Nothing Then cPM_Work.ResetEnvironment
        Set cPM_Work = Nothing
        On Error GoTo 0

'------------------------------------------------------------------------------
' RE-RAISE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Re-raise the original error
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

Public Property Get OverheadMeasurement_Text( _
    Optional ByVal iMethod As cPM_TimerMethod = cPM_MethodQPC, _
    Optional ByVal Iterations As Long = 1000, _
    Optional ByVal OverheadSecondsIn As Variant) _
    As String
'
'==============================================================================
'                         OVERHEAD MEASUREMENT TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a human-readable report of the average near-empty measurement
'   overhead for the selected timer backend
'
' WHY THIS EXISTS
'   OverheadMeasurement_Seconds() provides the numeric benchmark primitive
'
'   This property exists as the presentation-oriented companion for diagnostics,
'   logs, and quick inspection
'
'   When an already measured overhead value is supplied, this routine formats
'   that value directly without taking a second measurement
'
' INPUTS
'   iMethod (optional)
'     Timing backend whose near-empty average measurement overhead should be
'     reported
'
'   Iterations (optional)
'     Number of near-empty timing cycles to average
'     Values less than 1 are coerced to 1 by OverheadMeasurement_Seconds()
'
'   OverheadSecondsIn (optional)
'     Optional already measured overhead value to format directly
'
' RETURNS
'   String
'     A human-readable measurement-overhead report
'
' BEHAVIOR
'   - Normalizes and validates the requested timer method
'   - If OverheadSecondsIn is omitted, measures average near-empty overhead
'   - If OverheadSecondsIn is supplied, formats that value directly
'   - Returns a readable string containing the backend name and average
'     measurement overhead in seconds
'
' ERROR POLICY
'   Validation behavior is delegated to Method_NormalizeStart()
'
' DEPENDENCIES
'   - Method_NormalizeStart
'   - OverheadMeasurement_Seconds
'   - MethodName
'
' NOTES
'   This property mirrors the design of ElapsedTime:
'     - measure when no numeric value is supplied
'     - format directly when a numeric value is supplied
'
'   The reported figure is a mean. See the note on OverheadMeasurement_Seconds
'   regarding its sensitivity to scheduler outliers.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AvgS            As Double    'Average measurement overhead in seconds
    Dim ResolvedMethod  As cPM_TimerMethod   'Validated timer method

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Normalize and validate the timer method
        ResolvedMethod = Method_NormalizeStart(iMethod)

'------------------------------------------------------------------------------
' RESOLVE NUMERIC OVERHEAD VALUE
'------------------------------------------------------------------------------
    'Use the supplied numeric overhead value when provided
        If Not IsMissing(OverheadSecondsIn) Then
            AvgS = CDbl(OverheadSecondsIn)
    'Otherwise measure average near-empty overhead now
        Else
            AvgS = OverheadMeasurement_Seconds(ResolvedMethod, Iterations)
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a readable diagnostic string
        OverheadMeasurement_Text = "Average measurement overhead - " & _
                                   MethodName(ResolvedMethod) & " - " & _
                                   Format$(AvgS, "0.000000000") & " seconds"

End Property

Public Property Get Get_SystemTickInterval() As String
'
'==============================================================================
'                           GET SYSTEM TICK INTERVAL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the nominal system clock tick interval in milliseconds, derived from
'   GetSystemTimeAdjustment.
'
' WHY THIS EXISTS
'   This property provides a coarse diagnostic of the system clock increment
'   that can affect low-resolution timing sources and general scheduler
'   granularity.
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     Human-readable tick-interval text.
'
' BEHAVIOR
'   - Calls GetSystemTimeAdjustment to retrieve the nominal system clock
'     increment.
'   - Converts lpTimeIncrement from 100-nanosecond units to milliseconds.
'   - Returns a deterministic "(unknown)" string if the API call fails.
'
' ERROR POLICY
'   Does not raise errors.
'   Returns a deterministic string when the underlying API call fails.
'
' DEPENDENCIES
'   - GetSystemTimeAdjustment
'
' NOTES
'   - lpTimeIncrement from GetSystemTimeAdjustment is expressed in
'     100-nanosecond units, hence division by 10,000 to obtain milliseconds.
'   - This is not the same as QPC resolution and does not guarantee actual
'     scheduling precision.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC As Long     'Return code from GetSystemTimeAdjustment
    Dim lA As Long     'lpTimeAdjustment (unused here)
    Dim LB As Long     'lpTimeIncrement (100ns units)
    Dim lC As Long     'lpTimeAdjustmentDisabled (unused here)

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Query system time adjustment parameters.
        RC = GetSystemTimeAdjustment(lA, LB, lC)

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a deterministic string when the API call fails.
        If RC = 0 Then
            Get_SystemTickInterval = "Tick Interval = (unknown)"
            Exit Property
        End If

    'Convert tick increment (100ns units) to milliseconds.
        Get_SystemTickInterval = "Tick Interval = " & (LB / 10000#) & " milliseconds"
End Property

Public Property Get QPC_Get_SystemTickInterval() As String
'
'==============================================================================
'                       QPC GET SYSTEM TICK INTERVAL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the QPC tick interval in seconds per tick.
'
' WHY THIS EXISTS
'   This property provides a human-readable diagnostic for the effective
'   granularity of the QueryPerformanceCounter timebase.
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     Human-readable QPC tick interval text.
'
' BEHAVIOR
'   - Fetches the true QPC frequency via QPC_GetFrequencyHz().
'   - Returns "(unknown)" when frequency is unavailable.
'   - Otherwise returns the reciprocal of the QPC frequency in seconds/tick.
'
' ERROR POLICY
'   Does not raise errors.
'   Returns a deterministic string when QPC frequency is unavailable.
'
' DEPENDENCIES
'   - QPC_GetFrequencyHz
'
' NOTES
'   This is a presentation-oriented diagnostic property.
'   Use numeric helpers when machine-readable values are required.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Hz As Double      'True QPC frequency in ticks/second

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Fetch the true QPC frequency
        Hz = QPC_GetFrequencyHz()

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a deterministic string when frequency is unavailable
        If Hz = 0# Then
            QPC_Get_SystemTickInterval = "QPC Tick interval is: (unknown)"
            Exit Property
        End If

    'Return seconds-per-tick
        QPC_Get_SystemTickInterval = "QPC Tick interval is: " & (1# / Hz)
End Property

Public Property Get QPC_FrequencyPerSecond() As String
'
'==============================================================================
'                          QPC FREQUENCY PER SECOND
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the QPC frequency as a human-readable string.
'
' WHY THIS EXISTS
'   This property provides a presentation-oriented diagnostic for the effective
'   QueryPerformanceCounter frequency used by the class.
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     Human-readable QPC frequency text.
'
' BEHAVIOR
'   - Fetches the true QPC frequency via QPC_GetFrequencyHz().
'   - Returns "(unknown)" when frequency is unavailable.
'   - Otherwise returns ticks-per-second as readable text.
'
' ERROR POLICY
'   Does not raise errors.
'   Returns a deterministic string when QPC frequency is unavailable.
'
' DEPENDENCIES
'   - QPC_GetFrequencyHz
'
' NOTES
'   This is a presentation-oriented diagnostic property.
'   Use a numeric helper/property when machine-readable frequency values are
'   required.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Hz As Double      'True QPC frequency in ticks/second

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Fetch the true QPC frequency
        Hz = QPC_GetFrequencyHz()

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a deterministic string when frequency is unavailable
        If Hz = 0# Then
            QPC_FrequencyPerSecond = "QPC Tick frequency is: (unknown)"
            Exit Property
        End If

    'Return ticks-per-second as text
        QPC_FrequencyPerSecond = "QPC Tick frequency is: " & Hz & " ticks/second"
End Property

Public Property Get QPC_FrequencyPerSecond_Value() As Double
'
'==============================================================================
'                       QPC FREQUENCY PER SECOND VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the QueryPerformanceCounter frequency as a numeric Double.
'
' WHY THIS EXISTS
'   This property is the machine-readable companion to
'   QPC_FrequencyPerSecond(). It provides the true QPC frequency in
'   ticks/second for calculations, assertions, diagnostics, or external
'   reporting that should not depend on formatted text.
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     True QPC frequency in ticks/second.
'     Returns 0 when QPC frequency is unavailable.
'
' BEHAVIOR
'   Delegates directly to QPC_GetFrequencyHz().
'
' ERROR POLICY
'   Does not raise errors.
'   Returns 0 when QPC frequency information is unavailable.
'
' DEPENDENCIES
'   - QPC_GetFrequencyHz
'
' NOTES
'   Use QPC_FrequencyPerSecond() when human-readable text is desired.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the true QPC frequency in ticks/second.
        QPC_FrequencyPerSecond_Value = QPC_GetFrequencyHz()
End Property

'
'------------------------------------------------------------------------------
'
'                   PUBLIC: EXECUTION CONTROL / ENVIRONMENT
'
'------------------------------------------------------------------------------
'

Public Sub Pause( _
    ByVal dSeconds As Double, _
    Optional ByVal iMethod As cPM_PauseMethod = cPM_PauseSleep)
'
'==============================================================================
'                                     PAUSE
'------------------------------------------------------------------------------
' PURPOSE
'   Pauses execution for dSeconds (limited to <= 3600 seconds) using one of the
'   supported waiting mechanisms.
'
' WHY THIS EXISTS
'   This routine provides a centralized pause utility with multiple waiting
'   strategies so callers can choose between lower overhead, better UI
'   responsiveness, and simpler blocking behavior.
'
' INPUTS
'   dSeconds
'     Desired pause duration in seconds.
'     - <= 0   : no-op (returns immediately)
'     - > 3600 : rejected (returns immediately) to avoid excessive blocking
'
'   iMethod (optional)
'     Selects the pause mechanism:
'       1 => Sleep API (kernel32) (millisecond granularity; no DoEvents)
'       2 => Timer + DoEvents loop (coarser; yields to Excel/UI)
'       3 => Application.Wait for coarse waiting + Timer bound and guard
'       4 => Now + DoEvents loop + Timer bound and guard
'       Else => treated like method 4
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Method 1 is the lowest-overhead pause mechanism and does not call
'     DoEvents.
'   - Method 2 uses Timer() delta logic and yields via DoEvents.
'   - Method 3 issues coarse Application.Wait calls only while enough budget
'     remains for one, then finishes with a Timer()-based loop.
'   - Method 4 runs a Date/Now-flavored yielding loop that is itself bounded by
'     elapsed time, then finishes with a Timer()-based loop.
'
' OVERSHOOT POLICY
'   Now and Application.Wait both operate at one-second granularity, and
'   WaitTill is computed in floating point where 1/86400 is not exactly
'   representable. A coarse loop left unbounded can therefore wait a full extra
'   second before its condition clears.
'
'   Every coarse wait is now bounded by the same Timer() measurement used by the
'   trailing guard loop, so no strategy can exceed the requested duration by
'   more than DoEvents granularity.
'
'   The trailing guard loops remain, because they protect against the opposite
'   error: a coarse mechanism that returns early. A guard loop can only add
'   time, so it cannot correct an overshoot - the bound has to be inside the
'   coarse loop itself.
'
' ERROR POLICY
'   Invalid, non-positive, or excessively large pause durations are treated as
'   no-op exits.
'
' DEPENDENCIES
'   - Sleep
'   - Timer
'   - Application.Wait
'   - Now
'   - DoEvents
'
' NOTES
'   - The iMethod values here are pause mechanisms and are NOT the same
'     numbering as the timing backends used by StartTimer.
'   - Every elapsed comparison is midnight-safe because it measures a delta
'     rather than comparing Timer() against an absolute target.
'   - A single Application.Wait can consume up to two seconds, because it rounds
'     up to the next whole-second boundary before adding its interval. Method 3
'     therefore issues a coarse wait only while at least that much budget
'     remains; for requests under roughly two seconds it behaves like method 2.
'   - This routine uses local elapsed-time variables and does not overwrite the
'     class-level m_ET timing cache.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PauseTime   As Double   'Requested pause duration in seconds
    Dim StartS      As Double   'Timer() snapshot at pause start
    Dim ElapsedS    As Double   'Elapsed seconds since StartS
    Dim WaitTill    As Date     'Target timestamp for coarse Date/Wait paths
    Dim Ms          As Long     'Milliseconds for Sleep (method 1)
    Dim WholeS      As Long     'Whole-second coarse-wait portion (method 3)
    Dim K           As Long     'Loop counter for repeated Wait calls (method 3)

'------------------------------------------------------------------------------
' VALIDATE / CONTROL
'------------------------------------------------------------------------------
    'Reject non-positive pauses (no-op)
        If dSeconds <= 0# Then Exit Sub
    'Hard cap to avoid accidental long blocking pauses
        If dSeconds > 3600# Then Exit Sub

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Persist the requested pause duration locally
        PauseTime = dSeconds

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    Select Case iMethod

        Case cPM_PauseSleep
            'Sleep takes milliseconds; coerce seconds to Long milliseconds
                Ms = CLng(PauseTime * 1000#)

            'Sleep(0) yields the remainder of the thread's timeslice; allow it explicitly
                Sleep Ms

        Case cPM_PauseTimerLoop
            'Capture the Timer() start snapshot
                StartS = Timer

            'Loop until the requested pause duration has elapsed
                Do
                    'Compute elapsed seconds since the start snapshot
                        ElapsedS = Timer - StartS

                    'Correct midnight rollover when Timer() wraps at 24 hours
                        If ElapsedS < 0# Then ElapsedS = ElapsedS + 86400#

                    'Exit once the requested pause duration has been reached
                        If ElapsedS >= PauseTime Then Exit Do

                    'Yield to Excel/UI while waiting
                        DoEvents
                Loop

        Case cPM_PauseAppWait
            'Application.Wait is coarse and second-based, so use it only for the
            'whole-second portion and then finish with a Timer()-based loop
                WholeS = CLng(Fix(PauseTime))

            'Capture the Timer() start snapshot before any waiting begins, so the
            'coarse portion and the guard measure from the same origin
                StartS = Timer

            'Issue coarse waits only while enough budget remains for one
                For K = 1 To WholeS
                    'Compute elapsed seconds since the pause started
                        ElapsedS = Timer - StartS

                    'Correct midnight rollover when Timer() wraps at 24 hours
                        If ElapsedS < 0# Then ElapsedS = ElapsedS + 86400#

                    'A single Application.Wait can consume up to two seconds
                    'because it rounds up to the next whole-second boundary.
                    'Stop issuing coarse waits once less than that remains.
                        If (PauseTime - ElapsedS) < 2# Then Exit For

                    'Apply one coarse whole-second wait
                        Application.Wait Now + TimeSerial(0, 0, 1)
                Next K

            'Finish the remainder precisely using Timer() delta logic
                Do
                    'Compute elapsed seconds since the pause started
                        ElapsedS = Timer - StartS

                    'Correct midnight rollover when Timer() wraps at 24 hours
                        If ElapsedS < 0# Then ElapsedS = ElapsedS + 86400#

                    'Exit once the total requested pause duration has been reached
                        If ElapsedS >= PauseTime Then Exit Do

                    'Yield to Excel/UI while waiting
                        DoEvents
                Loop

        Case Else
            'Preserve the Date/Now-flavored yielding path, but bound it by elapsed
            'time so a coarse one-second tick cannot overshoot the request
                WaitTill = Now + (PauseTime / 86400#)

            'Capture the Timer() start snapshot before any waiting begins
                StartS = Timer

            'Run the coarse Date-based yielding loop, bounded by elapsed time
                Do While Now < WaitTill
                    'Compute elapsed seconds since the pause started
                        ElapsedS = Timer - StartS

                    'Correct midnight rollover when Timer() wraps at 24 hours
                        If ElapsedS < 0# Then ElapsedS = ElapsedS + 86400#

                    'Leave the coarse loop rather than waiting for the next tick
                        If ElapsedS >= PauseTime Then Exit Do

                    'Yield to Excel/UI while waiting
                        DoEvents
                Loop

            'Finish the remainder precisely using Timer() delta logic
                Do
                    'Compute elapsed seconds since the pause started
                        ElapsedS = Timer - StartS

                    'Correct midnight rollover when Timer() wraps at 24 hours
                        If ElapsedS < 0# Then ElapsedS = ElapsedS + 86400#

                    'Exit once the total requested pause duration has been reached
                        If ElapsedS >= PauseTime Then Exit Do

                    'Yield to Excel/UI while waiting
                        DoEvents
                Loop

    End Select
End Sub
Public Sub ResetEnvironment()
'
'==============================================================================
'                              RESET ENVIRONMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Explicitly restores environment changes made by this instance:
'     - releases timeBeginPeriod(1) if enabled
'     - ends this instance's shared TW session if active
'
' WHY THIS EXISTS
'   This is the primary explicit cleanup routine for cPerformanceManager.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - If this instance enabled 1ms multimedia timer resolution, balances it with
'     timeEndPeriod 1.
'   - Clears m_TimePeriodOn only when timeEndPeriod succeeds.
'   - Ends this instance's shared TW session when active.
'   - Raises after TW cleanup if timer-resolution release failed in strict mode.
'
' ERROR POLICY
'   Raises in strict mode if timeEndPeriod(1) fails.
'
' DEPENDENCIES
'   - timeEndPeriod
'   - PM_TW_EndSession
'   - TW_InstanceKey
'   - TW_Enum
'
' NOTES
'   This routine remains idempotent.
'
'   TW cleanup is deliberately performed BEFORE any strict-mode raise, so a
'   timer-resolution failure cannot leave Excel suppressed.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC              As Long     'Return code from timeEndPeriod
    Dim ReleaseFailed   As Boolean  'TRUE when timeEndPeriod fails

'------------------------------------------------------------------------------
' BALANCE timeBeginPeriod(1) IF IT WAS ENABLED
'------------------------------------------------------------------------------
    'Release the previously requested 1ms multimedia timer resolution
        If m_TimePeriodOn Then
            RC = timeEndPeriod(1)
            If RC = 0 Then
                m_TimePeriodOn = False
            Else
                ReleaseFailed = True
            End If
        End If

'------------------------------------------------------------------------------
' END THIS INSTANCE'S SHARED TW SESSION IF ACTIVE
'------------------------------------------------------------------------------
    'End only this instance's TW participation
        If m_TW_Active Then
            PM_TW_EndSession TW_InstanceKey()
            m_TW_Active = False
            m_Except = TW_Enum.None
        End If

'------------------------------------------------------------------------------
' REPORT TIMER-RELEASE FAILURE AFTER OTHER CLEANUP
'------------------------------------------------------------------------------
    'Raise only after the TW session has also been cleaned up
        If ReleaseFailed And m_StrictMode Then
            Err.Raise ERR_CPM_TIMEENDPERIOD_FAILED, _
                      "cPerformanceManager.ResetEnvironment", _
                      "timeEndPeriod(1) failed."
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                      PUBLIC: SHARED TIME-WASTER CONTROL
'
'------------------------------------------------------------------------------
'

Public Sub TW_Turn_OFF( _
    Optional ByVal Except As TW_Enum = TW_Enum.None)
'
'==============================================================================
'                           TURN OFF TIMEWASTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Starts or updates this instance's shared Excel "time waster" suppression
'   session.
'
' WHY THIS EXISTS
'   Application-level TW settings are global to Excel, not instance-local.
'   Therefore this routine registers the current instance with the shared
'   module-level TW scope manager rather than trying to manage Application state
'   independently inside the class.
'
' INPUTS
'   Except (optional)
'     Bitmask of TW_Enum flags that this instance does NOT want disabled.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Registers this instance in the shared TW manager if not already active.
'   - Updates this instance's effective mask if already active.
'   - Commits local instance state only after the shared manager succeeds.
'
' ERROR POLICY
'   Raises errors normally.
'   Registration/update failures in the shared TW manager should not be silently
'   ignored during the normal execution path.
'
' DEPENDENCIES
'   - PM_TW_BeginSession
'   - TW_InstanceKey
'   - TW_Enum
'
' NOTES
'   - This routine is effectively idempotent at the instance level:
'       * first call => begins this instance's shared session
'       * subsequent calls => update this instance's exemption mask
'   - The effective Excel state is determined by the shared manager from the
'     aggregate disable-mask of all active instances.
'   - Local state is committed AFTER PM_TW_BeginSession returns. A failed
'     registration therefore cannot leave this instance reporting an exemption
'     mask that was never applied.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' BEGIN / UPDATE SHARED SESSION
'------------------------------------------------------------------------------
    'Register or update this instance in the shared TW manager, passing this
    'instance's strictness so the manager can decide whether an unhonourable
    'Calculation request should raise or be exempted
        PM_TW_BeginSession TW_InstanceKey(), CLng(Except), m_StrictMode

'------------------------------------------------------------------------------
' COMMIT INSTANCE STATE
'------------------------------------------------------------------------------
    'Persist this instance's exemption mask only after successful registration
        m_Except = Except
    'Mark this instance as actively participating in the shared TW scope
        m_TW_Active = True

End Sub

Public Sub TW_Turn_ON()
'
'==============================================================================
'                              TURN ON TIMEWASTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Ends this instance's shared TW suppression session.
'
' WHY THIS EXISTS
'   TW suppression is managed through a shared module-level scope manager
'   because the underlying Excel Application settings are global. Therefore
'   "turning TW back on" means removing only this instance's participation and
'   letting the shared manager determine the correct resulting effective state.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - No-op if this instance is not currently active.
'   - If this instance is active, removes it from the shared TW manager.
'   - If other instances remain active, the shared manager recomputes the
'     effective state from the remaining aggregate disable-mask.
'   - If this is the final active instance, the original baseline is restored.
'   - Clears this instance's local TW activity state and exemption mask after
'     successful removal from the shared manager.
'
' ERROR POLICY
'   Raises errors normally.
'   Explicit shared-session teardown should fail loudly during the normal
'   execution path if it cannot be completed.
'
' DEPENDENCIES
'   - PM_TW_EndSession
'   - TW_InstanceKey
'   - TW_Enum
'
' NOTES
'   - This routine is idempotent:
'       * if m_TW_Active = False, nothing happens
'   - This routine does not restore Excel state directly.
'     That responsibility belongs to the shared TW manager.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' END SHARED SESSION (IF ACTIVE)
'------------------------------------------------------------------------------
    'Only end a session if this instance is currently active in the shared manager
        If m_TW_Active Then
            PM_TW_EndSession TW_InstanceKey()

            m_TW_Active = False
            m_Except = TW_Enum.None
        End If
End Sub

Public Property Get TW_IsActive() As Boolean
'------------------------------------------------------------------------------
' TW_IsActive
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if this instance currently has an active shared TW session.
'------------------------------------------------------------------------------
    'Return this instance's TW activity state.
        TW_IsActive = m_TW_Active
End Property

Public Property Get TW_ActiveSessionCount() As Long
'------------------------------------------------------------------------------
' TW_ACTIVESESSIONCOUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of active shared TW sessions across all instances.
'------------------------------------------------------------------------------
    'Return the global active shared TW session count.
        TW_ActiveSessionCount = PM_TW_ActiveCount()
End Property

'
'------------------------------------------------------------------------------
'
'                 PUBLIC: CHECKPOINTS / STRUCTURED REPORTING
'
'------------------------------------------------------------------------------
'

Public Sub SetRunLabel( _
    ByVal RunLabel As String)
'
'==============================================================================
'                                SET RUN LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Stores an optional run label used by later checkpoint captures
'
' WHY THIS EXISTS
'   Structured checkpoint output is easier to interpret when rows can be tagged
'   with a user-supplied run label such as:
'
'     - ImportWorkflow
'     - ScenarioA
'     - Baseline
'     - TW_Off
'
'   To keep report output unambiguous, the run label must be set before the
'   first checkpoint is captured in the current timing session
'
' INPUTS
'   RunLabel
'     User-defined label for the current timing run
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Stores the supplied text as the current run label
'   - Raises if the caller attempts to change the run label after one or more
'     checkpoints have already been captured in the current session
'
' ERROR POLICY
'   Raises when the caller attempts to change the run label after checkpoint
'   capture has already begun in the current timing session
'
' NOTES
'   This routine must be called AFTER StartTimer. Starting a session clears the
'   run label, so a label set beforehand is discarded.
'
'   To set the label as part of starting the session, use the RunLabel argument
'   of StartTimer instead:
'
'       cPM.StartTimer 5, False, "Baseline run"
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE SESSION CHECKPOINT STATE
'------------------------------------------------------------------------------
    'Reject run-label changes after checkpoint capture has already begun
        If m_CheckpointCount > 0 Then
            Err.Raise ERR_CPM_RUNLABEL_AFTER_CAPTURE, _
                      "cPerformanceManager.SetRunLabel", _
                      "RunLabel must be set before the first checkpoint of the current timing session."
        End If

'------------------------------------------------------------------------------
' ASSIGN STATE
'------------------------------------------------------------------------------
    'Store the supplied run label
        m_RunLabel = RunLabel

End Sub

Public Sub ClearCheckpoints()
'
'==============================================================================
'                              CLEAR CHECKPOINTS
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the structured checkpoint state for the current instance
'
' WHY THIS EXISTS
'   Although a new StartTimer session already resets checkpoint state, callers
'   may also want to clear the current structured checkpoint / report state
'   explicitly without recreating the object
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Clears the current run label
'   - Clears the checkpoint counter
'   - Resets the last checkpoint elapsed-seconds marker
'   - Reinitializes checkpoint storage
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Checkpoint_ClearState
'
' NOTES
'   This does NOT end the active timing session. Cumulative checkpoint times
'   captured afterwards are still measured from the original StartTimer.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' CLEAR STATE
'------------------------------------------------------------------------------
    'Clear all checkpoint / report state explicitly
        Checkpoint_ClearState

End Sub

Public Sub Checkpoint( _
    ByVal CheckpointName As String, _
    Optional ByVal NoteText As String = vbNullString)
'
'==============================================================================
'                                  CHECKPOINT
'------------------------------------------------------------------------------
' PURPOSE
'   Captures one named checkpoint within the current timing session.
'
' INPUTS
'   CheckpointName   - user-defined checkpoint label
'   NoteText         - optional descriptive note stored with the row
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Requires an active timing session.
'   - Captures cumulative elapsed seconds without disturbing cached state.
'   - Abandons the capture entirely if the underlying read fails.
'   - Appends one structured checkpoint row and advances the marker.
'
' ERROR POLICY
'   - Raises when no timing session is active.
'   - A failed read publishes LastReadStatus first, then raises in strict mode
'     or abandons the capture in non-strict mode.
'
' NOTES
'   On a failed read, the checkpoint count and the last-checkpoint marker are
'   left unchanged. A failed measurement must not appear in a report, and must
'   not distort the delta of the next successful checkpoint.
'
'   Taking a checkpoint does not overwrite the cached T2 / ET values.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentS            As Double           'Current cumulative elapsed seconds
    Dim DeltaS              As Double           'Elapsed seconds since previous checkpoint
    Dim NextSeq             As Long             'Next checkpoint sequence number
    Dim EffectiveName       As String           'Normalized checkpoint label
    Dim ResolvedMethod      As cPM_TimerMethod          'Validated method for this read
    Dim DiscardEndStamp     As Double           'Raw endpoint, discarded by design
    Dim DiscardQpcEnd       As Currency         'Raw QPC endpoint, discarded by design
    Dim Status              As cPM_ReadStatus   'Outcome of the read

'------------------------------------------------------------------------------
' VALIDATE SESSION
'------------------------------------------------------------------------------
    'Reject checkpoint capture without an active timing session
        If Not m_HasActiveSession Then
            Err.Raise ERR_CPM_CHECKPOINT_NO_SESSION, _
                      "cPerformanceManager.Checkpoint", _
                      "Checkpoint cannot be captured before StartTimer."
        End If

'------------------------------------------------------------------------------
' MEASURE WITHOUT MUTATING CACHED STATE
'------------------------------------------------------------------------------
    'This capture describes its own read outcome
        m_LastReadStatus = cPM_ReadOK

    'Resolve the session method and read elapsed seconds directly
        ResolvedMethod = Method_ResolveElapsed(cPM_MethodSession)
        CurrentS = Elapsed_ComputeSeconds(ResolvedMethod, _
                                          DiscardEndStamp, DiscardQpcEnd, Status)

    'Publish the outcome BEFORE deciding whether to fail, so a caller that traps
    'a strict-mode error still finds the status describing this read
        If Status <> cPM_ReadOK Then
            m_LastReadStatus = Status

            If m_StrictMode Then
                Err.Raise Read_StatusToErrorNumber(Status), _
                          "cPerformanceManager.Checkpoint", _
                          Read_StatusToText(Status) & _
                          " The checkpoint was not captured."
            End If

            'Non-strict: abandon the capture, leaving the count and the
            'last-checkpoint marker untouched
                Exit Sub
        End If

'------------------------------------------------------------------------------
' NORMALIZE LABEL
'------------------------------------------------------------------------------
    'Substitute a generated label when the caller supplied a blank name
        EffectiveName = Trim$(CheckpointName)
        If Len(EffectiveName) = 0 Then
            EffectiveName = Checkpoint_DefaultName(m_CheckpointCount + 1)
        End If

'------------------------------------------------------------------------------
' COMPUTE SEQUENCE AND DELTA
'------------------------------------------------------------------------------
    'Resolve the next sequence number
        NextSeq = m_CheckpointCount + 1

    'Compute the delta since the previous checkpoint
        DeltaS = CurrentS - m_LastCheckpointS

    'Defensively clamp a negative delta
        If DeltaS < 0# Then
            DeltaS = 0#
        End If

'------------------------------------------------------------------------------
' APPEND CHECKPOINT ROW
'------------------------------------------------------------------------------
    'Ensure storage can hold the new row
        Checkpoint_EnsureCapacity NextSeq

    'Populate the structured checkpoint row
        With m_Checkpoints(NextSeq)
            .Seq = NextSeq
            .NameText = EffectiveName
            .NoteText = NoteText
            .DeltaSecondsValue = DeltaS
            .ElapsedSecondsValue = CurrentS
            .MethodID = m_ActiveMethodID
            .MethodNameText = MethodName(m_ActiveMethodID)
            .RunLabelText = m_RunLabel
        End With

'------------------------------------------------------------------------------
' UPDATE STATE
'------------------------------------------------------------------------------
    'Advance the checkpoint counter and last-checkpoint marker
        m_CheckpointCount = NextSeq
        m_LastCheckpointS = CurrentS

End Sub


Public Property Get CheckpointCount() As Long
'
'==============================================================================
'                              CHECKPOINT COUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of captured structured checkpoints
'
' WHY THIS EXISTS
'   Callers often need a lightweight way to know whether checkpoint rows exist
'   and how many were captured
'
' INPUTS
'   None
'
' RETURNS
'   Long
'     Number of captured checkpoints
'
' ERROR POLICY
'   Does not raise errors
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the number of captured checkpoints
        CheckpointCount = m_CheckpointCount

End Property

Public Function ReportAsArray() As Variant
'
'==============================================================================
'                               REPORT AS ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Exports the structured checkpoint report as a 2D Variant array
'
' WHY THIS EXISTS
'   A 2D array is the most reusable machine-readable report surface because it
'   can be:
'
'     - written directly to a worksheet range
'     - consumed by helper modules
'     - post-processed by calling code
'
' INPUTS
'   None
'
' RETURNS
'   Variant
'     Two-dimensional 1-based array including a header row
'
' BEHAVIOR
'   - Returns a 2D array with one header row
'   - Returns one additional row per checkpoint
'   - Uses this column order:
'
'       1 = RunLabel
'       2 = Seq
'       3 = Checkpoint
'       4 = Note
'       5 = MethodID
'       6 = MethodName
'       7 = DeltaSeconds
'       8 = CumulativeSeconds
'
'   - When no checkpoints exist, still returns a header-only 2D array
'
' ERROR POLICY
'   Raises errors normally
'
' NOTES
'   M_cPM_TimeWasters.cPM_Report_WriteToRange formats columns 7 and 8 through
'   its own PM_RPT_COL_* constants. Changing the column order here requires
'   updating those constants as well.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Arr()       As Variant      'Export array
    Dim RowCount    As Long         'Total export rows including the header
    Dim i           As Long         'Checkpoint loop index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Size the export array to include the header row
        RowCount = m_CheckpointCount + 1
        ReDim Arr(1 To RowCount, 1 To 8)

'------------------------------------------------------------------------------
' BUILD HEADER ROW
'------------------------------------------------------------------------------
    'Populate the header labels
        Arr(1, 1) = "RunLabel"
        Arr(1, 2) = "Seq"
        Arr(1, 3) = "Checkpoint"
        Arr(1, 4) = "Note"
        Arr(1, 5) = "MethodID"
        Arr(1, 6) = "MethodName"
        Arr(1, 7) = "DeltaSeconds"
        Arr(1, 8) = "CumulativeSeconds"

'------------------------------------------------------------------------------
' BUILD DATA ROWS
'------------------------------------------------------------------------------
    'Populate one export row per captured checkpoint
        For i = 1 To m_CheckpointCount
            With m_Checkpoints(i)
                Arr(i + 1, 1) = .RunLabelText
                Arr(i + 1, 2) = .Seq
                Arr(i + 1, 3) = .NameText
                Arr(i + 1, 4) = .NoteText
                Arr(i + 1, 5) = .MethodID
                Arr(i + 1, 6) = .MethodNameText
                Arr(i + 1, 7) = .DeltaSecondsValue
                Arr(i + 1, 8) = .ElapsedSecondsValue
            End With
        Next i

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the structured report array
        ReportAsArray = Arr

End Function

Public Function ReportAsText() As String
'
'==============================================================================
'                                REPORT AS TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the structured checkpoint report as an aligned multiline text table
'
' WHY THIS EXISTS
'   This provides a quick developer-facing report surface for:
'
'     - Immediate Window output
'     - ad hoc diagnostics
'     - lightweight reporting without worksheet output
'
' INPUTS
'   None
'
' RETURNS
'   String
'     Multiline text table
'
' BEHAVIOR
'   - Returns a short message when no checkpoints exist
'   - Sizes every column to its widest value, so nothing is truncated and the
'     columns line up whatever the checkpoint names happen to be
'   - Right-aligns the numeric columns and left-aligns the text ones
'
' ERROR POLICY
'   Raises errors normally
'
' NOTES
'   Column widths are computed from the data rather than fixed, because a fixed
'   width would either truncate a long checkpoint name or waste space on short
'   ones. The cost is a second pass over the rows, which is irrelevant for a
'   diagnostic surface
'
'   The layout is developer-facing and not a stable contract. Use ReportAsArray
'   for anything machine-readable
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S           As String       'Accumulated report text
    Dim i           As Long         'Checkpoint loop index
    Dim wSeq        As Long         'Width of the sequence column
    Dim wName       As Long         'Width of the checkpoint-name column
    Dim wDelta      As Long         'Width of the delta-seconds column
    Dim wCum        As Long         'Width of the cumulative-seconds column
    Dim wMethod     As Long         'Width of the method-name column
    Dim wNote       As Long         'Width of the note column
    Dim DeltaText   As String       'Formatted delta for the current row
    Dim CumText     As String       'Formatted cumulative for the current row

'------------------------------------------------------------------------------
' HANDLE EMPTY REPORT
'------------------------------------------------------------------------------
    'Return a short message when no checkpoints were captured
        If m_CheckpointCount = 0 Then
            ReportAsText = "No checkpoints captured."
            Exit Function
        End If

'------------------------------------------------------------------------------
' SIZE THE COLUMNS
'------------------------------------------------------------------------------
    'Start from the header widths so a header is never clipped
        wSeq = Len("Seq")
        wName = Len("Checkpoint")
        wDelta = Len("DeltaSeconds")
        wCum = Len("CumulativeSeconds")
        wMethod = Len("MethodName")
        wNote = Len("Note")

    'Widen each column to its widest value
        For i = 1 To m_CheckpointCount
            With m_Checkpoints(i)
                If Len(CStr(.Seq)) > wSeq Then wSeq = Len(CStr(.Seq))
                If Len(.NameText) > wName Then wName = Len(.NameText)
                If Len(Format$(.DeltaSecondsValue, "0.000000000")) > wDelta Then _
                    wDelta = Len(Format$(.DeltaSecondsValue, "0.000000000"))
                If Len(Format$(.ElapsedSecondsValue, "0.000000000")) > wCum Then _
                    wCum = Len(Format$(.ElapsedSecondsValue, "0.000000000"))
                If Len(.MethodNameText) > wMethod Then wMethod = Len(.MethodNameText)
                If Len(.NoteText) > wNote Then wNote = Len(.NoteText)
            End With
        Next i

'------------------------------------------------------------------------------
' BUILD HEADER
'------------------------------------------------------------------------------
    'Start the report title
        S = "CHECKPOINT REPORT"
    'Append the optional run label when present
        If Len(m_RunLabel) > 0 Then
            S = S & " | RunLabel=" & m_RunLabel
        End If

    'Append the column legend, padded to the resolved widths
        S = S & vbCrLf & _
            Report_PadLeft("Seq", wSeq) & " | " & _
            Report_PadRight("Checkpoint", wName) & " | " & _
            Report_PadLeft("DeltaSeconds", wDelta) & " | " & _
            Report_PadLeft("CumulativeSeconds", wCum) & " | " & _
            Report_PadRight("MethodName", wMethod) & " | " & _
            "Note"

    'Append a rule so the header reads as a table heading
        S = S & vbCrLf & _
            String$(wSeq, "-") & "-+-" & _
            String$(wName, "-") & "-+-" & _
            String$(wDelta, "-") & "-+-" & _
            String$(wCum, "-") & "-+-" & _
            String$(wMethod, "-") & "-+-" & _
            String$(wNote, "-")

'------------------------------------------------------------------------------
' BUILD DATA ROWS
'------------------------------------------------------------------------------
    'Append one aligned line per checkpoint
        For i = 1 To m_CheckpointCount
            With m_Checkpoints(i)
                DeltaText = Format$(.DeltaSecondsValue, "0.000000000")
                CumText = Format$(.ElapsedSecondsValue, "0.000000000")

                S = S & vbCrLf & _
                    Report_PadLeft(CStr(.Seq), wSeq) & " | " & _
                    Report_PadRight(.NameText, wName) & " | " & _
                    Report_PadLeft(DeltaText, wDelta) & " | " & _
                    Report_PadLeft(CumText, wCum) & " | " & _
                    Report_PadRight(.MethodNameText, wMethod) & " | " & _
                    .NoteText
            End With
        Next i

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the aligned structured report text
        ReportAsText = S

End Function

Private Function Report_PadRight( _
    ByVal Text As String, _
    ByVal Width As Long) _
    As String
'
'==============================================================================
'                              REPORT PAD RIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Left-aligns text within a fixed width by padding on the right
'
' INPUTS
'   Text   - value to place
'   Width  - column width
'
' RETURNS
'   String
'
' BEHAVIOR
'   Text longer than the width is returned unchanged, so a column can widen
'   rather than silently losing content
'
' ERROR POLICY
'   Does not raise
'
' UPDATED
'   2026-08-16
'==============================================================================

    'Pad on the right, never truncate
        If Len(Text) >= Width Then
            Report_PadRight = Text
        Else
            Report_PadRight = Text & String$(Width - Len(Text), " ")
        End If

End Function

Private Function Report_PadLeft( _
    ByVal Text As String, _
    ByVal Width As Long) _
    As String
'
'==============================================================================
'                               REPORT PAD LEFT
'------------------------------------------------------------------------------
' PURPOSE
'   Right-aligns text within a fixed width by padding on the left
'
' WHY THIS EXISTS
'   Numeric columns read far better right-aligned, because the decimal points
'   line up and magnitudes become comparable at a glance
'
' INPUTS
'   Text   - value to place
'   Width  - column width
'
' RETURNS
'   String
'
' BEHAVIOR
'   Text longer than the width is returned unchanged
'
' ERROR POLICY
'   Does not raise
'
' UPDATED
'   2026-08-16
'==============================================================================

    'Pad on the left, never truncate
        If Len(Text) >= Width Then
            Report_PadLeft = Text
        Else
            Report_PadLeft = String$(Width - Len(Text), " ") & Text
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                   PRIVATE: SESSION / VALIDATION HELPERS
'
'------------------------------------------------------------------------------
'

Friend Function UInt32ToDouble( _
    ByVal Value32 As Long) _
    As Double
'
'==============================================================================
'                              UINT32 TO DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Converts a signed VBA Long containing a native unsigned 32-bit value into
'   the corresponding logical unsigned Double value.
'
' WHY THIS EXISTS
'   Several Windows millisecond counters are returned through 32-bit DWORD
'   fields. VBA stores those fields as signed Long values, so values above
'   2,147,483,647 appear negative unless they are mapped back into the unsigned
'   range.
'
' INPUTS
'   Value32
'     Signed VBA Long containing a native unsigned 32-bit value.
'
' CALLER CONTRACT
'   Value32 MUST have originated as a DWORD returned by a Windows API. This
'   routine reinterprets the bit pattern; it does not convert a signed quantity.
'
'   Passing a genuinely signed value - an elapsed delta, a difference between
'   two counters, any arithmetic result - returns a number close to 4.29e9
'   rather than the negative value intended, and that number will look
'   plausible. There is no way to detect the misuse from the value alone, which
'   is why the contract is stated rather than validated.
'
' RETURNS
'   Double
'     Logical unsigned value in the range 0 to 4,294,967,295.
'
' BEHAVIOR
'   - Returns Value32 directly when nonnegative.
'   - Adds 2^32 when Value32 is negative.
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   None.
'
' NOTES
'   4294967296# is exactly 2^32 represented as a Double literal.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Map signed Long representation back to logical unsigned DWORD space
        If Value32 < 0 Then
            UInt32ToDouble = 4294967296# + CDbl(Value32)
        Else
            UInt32ToDouble = CDbl(Value32)
        End If

End Function

Friend Function Elapsed_Validate( _
    ByVal Seconds As Double, _
    ByRef StatusOut As cPM_ReadStatus) _
    As Double
'
'==============================================================================
'                              ELAPSED VALIDATE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a computed elapsed-seconds value after rollover correction, and
'   reports whether it is a real measurement
'
' WHY THIS EXISTS
'   Once rollover correction has been applied, a negative elapsed time can only
'   mean the underlying timing source moved backwards. That makes the whole
'   measurement invalid rather than merely imprecise.
'
'   This matters specifically on Win64 method 2, where RolloverSeconds returns 0
'   because GetTickCount64 is used. Without this check a negative value would be
'   returned to the caller unchanged and silently treated as a real duration.
'
' INPUTS
'   Seconds
'     Computed elapsed seconds after any applicable rollover correction
'
'   StatusOut
'     Receives cPM_ReadElapsedInvalid when the value had to be clamped, and is
'     otherwise left untouched so an earlier failure is not overwritten
'
' RETURNS
'   Double
'     The validated elapsed-seconds value, clamped to zero when invalid
'
' BEHAVIOR
'   - Returns the value unchanged when it is zero or positive
'   - Clamps a negative value to zero and reports cPM_ReadElapsedInvalid
'
' ERROR POLICY
'   NEVER raises. Strict-mode policy is applied by the public operation, which
'   publishes the status first so the two public surfaces agree.
'
' WHY THE CLAMP IS STATUS-BEARING
'   Clamping alone produced "zero seconds, status OK", which is indistinguishable
'   from a genuine measurement of zero. A backward-clock event could therefore be
'   recorded as a valid checkpoint, or kept as the fastest sample in a benchmark.
'
' NOTES
'   Method 5 (QPC) deliberately does not call this helper. QPC is monotonic and
'   method 5 is the default hot path, so the extra call is not justified there.
'
'   Declared Friend rather than Private so the regression suite can test this
'   arithmetic directly. Friend keeps it invisible outside the project.
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' FAULT INJECTION
'------------------------------------------------------------------------------
    'Consume the one-shot forced-negative flag, if the suite armed it. A clock
    'that moves backwards cannot be arranged on demand, so the condition is
    'injected here rather than left permanently untested end to end
        If m_TestForceNextNegativeElapsed Then
            m_TestForceNextNegativeElapsed = False
            Seconds = -1#
        End If

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Clamp a backwards-moving timing source and say that is what happened
        If Seconds < 0# Then
            StatusOut = cPM_ReadElapsedInvalid
            Seconds = 0#
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the validated elapsed-seconds value
        Elapsed_Validate = Seconds

End Function

Private Function Method_NormalizeStart( _
    ByVal iMethod As cPM_TimerMethod) _
    As cPM_TimerMethod
'
'==============================================================================
'                          NORMALIZE START METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Normalizes the method requested by StartTimer and validates QPC availability.
'
' WHY THIS EXISTS
'   StartTimer should not contain repeated validation logic for method range and
'   backend availability. This helper centralizes that validation and returns a
'   method that is safe for use by the timing session.
'
' INPUTS
'   iMethod
'     Requested timer method identifier.
'
' RETURNS
'   Integer
'     A validated timer method identifier.
'
' BEHAVIOR
'   - Accepts methods 1..6 as valid.
'   - In non-strict mode:
'       * invalid method values default to 5 (QPC)
'       * if QPC is unavailable, method 5 falls back to 2
'   - In strict mode:
'       * invalid method values raise an error
'       * requesting QPC when unavailable raises an error
'
' ERROR POLICY
'   - Raises on invalid method values when m_StrictMode = True
'   - Raises when QPC is requested but unavailable and m_StrictMode = True
'
' DEPENDENCIES
'   - m_QpcAvailable
'   - m_StrictMode
'
' NOTES
'   Method 2 is the non-strict fallback when QPC is unavailable because it is a
'   practical general-purpose monotonic timing source.
'
'   The QPC availability check here is what makes the division by m_qpcFreq in
'   ElapsedSeconds unreachable with a zero denominator.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE METHOD RANGE
'------------------------------------------------------------------------------
    'Reject or coerce out-of-range method values
        If (iMethod < cPM_MethodTimer) Or (iMethod > cPM_MethodNow) Then
            If m_StrictMode Then
                Err.Raise ERR_CPM_INVALID_START_METHOD, _
                          "cPerformanceManager.Method_NormalizeStart", _
                          "Invalid timer method passed to StartTimer."
            Else
                iMethod = cPM_MethodQPC
            End If
        End If

'------------------------------------------------------------------------------
' VALIDATE BACKEND AVAILABILITY
'------------------------------------------------------------------------------
    'Reject or downgrade a QPC request when QPC is unavailable
        If iMethod = cPM_MethodQPC And Not m_QpcAvailable Then
            If m_StrictMode Then
                Err.Raise ERR_CPM_QPC_UNAVAILABLE, _
                          "cPerformanceManager.Method_NormalizeStart", _
                          "QPC is unavailable on this system."
            Else
                iMethod = cPM_MethodTickCount
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the validated method
        Method_NormalizeStart = iMethod
End Function

Private Function Method_ResolveElapsed( _
    ByVal iMethod As cPM_TimerMethod) _
    As cPM_TimerMethod
'
'==============================================================================
'                          RESOLVE ELAPSED METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves which timer method ElapsedSeconds should use:
'     - 0 => use the active session method
'     - 1..6 => validate against the active session method
'
' WHY THIS EXISTS
'   Once StartTimer establishes a timing session, elapsed-time reads must remain
'   bound to that same backend. This helper centralizes the validation rules
'   that prevent cross-method misuse.
'
' INPUTS
'   iMethod
'     Requested elapsed-time method:
'       0   => use active session method
'       1-6 => explicit method to validate against the active session
'
' RETURNS
'   Integer
'     The validated method that ElapsedSeconds should use.
'
' BEHAVIOR
'   - Raises or fails fast if no session has been started.
'   - Resolves method 0 to the active session method.
'   - Rejects invalid explicit method values.
'   - Rejects explicit methods that do not match the active session method.
'
' ERROR POLICY
'   - Raises when ElapsedSeconds is called before StartTimer if m_StrictMode = True
'   - Raises on invalid explicit method values if m_StrictMode = True
'   - Raises on explicit method/session mismatch if m_StrictMode = True
'   - In non-strict mode, returns the active method when possible; otherwise 0
'
' DEPENDENCIES
'   - m_HasActiveSession
'   - m_ActiveMethodID
'   - m_StrictMode
'
' NOTES
'   This routine is one of the key safeguards that turns the class from a loose
'   timer wrapper into a real timing-session manager.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE ACTIVE SESSION
'------------------------------------------------------------------------------
    'Reject elapsed-time reads before a session has been started
        If Not m_HasActiveSession Then
            If m_StrictMode Then
                Err.Raise ERR_CPM_NO_ACTIVE_SESSION, _
                          "cPerformanceManager.Method_ResolveElapsed", _
                          "ElapsedSeconds called before StartTimer."
            Else
                Method_ResolveElapsed = 0
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' RESOLVE IMPLICIT METHOD
'------------------------------------------------------------------------------
    'Method 0 means "use the active session method"
        If iMethod = cPM_MethodSession Then
            Method_ResolveElapsed = m_ActiveMethodID
            Exit Function
        End If

'------------------------------------------------------------------------------
' VALIDATE EXPLICIT METHOD RANGE
'------------------------------------------------------------------------------
    'Reject out-of-range explicit method values
        If (iMethod < cPM_MethodTimer) Or (iMethod > cPM_MethodNow) Then
            If m_StrictMode Then
                Err.Raise ERR_CPM_INVALID_ELAPSED_METHOD, _
                          "cPerformanceManager.Method_ResolveElapsed", _
                          "Invalid timer method passed to ElapsedSeconds."
            Else
                Method_ResolveElapsed = m_ActiveMethodID
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' VALIDATE SESSION BINDING
'------------------------------------------------------------------------------
    'Reject explicit methods that do not match the active session method
        If iMethod <> m_ActiveMethodID Then
            If m_StrictMode Then
                Err.Raise ERR_CPM_METHOD_MISMATCH, _
                          "cPerformanceManager.Method_ResolveElapsed", _
                          "ElapsedSeconds method does not match the active StartTimer method."
            Else
                Method_ResolveElapsed = m_ActiveMethodID
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the validated method
        Method_ResolveElapsed = iMethod
End Function

Private Function TW_InstanceKey() As String
'
'==============================================================================
'                               TW INSTANCE KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a stable per-instance key used by the shared TW scope manager.
'
' WHY THIS EXISTS
'   The shared TW manager must distinguish overlapping cPerformanceManager
'   instances so it can maintain the correct aggregate disable-mask and remove
'   the correct instance when a shared TW session ends.
'
'   Earlier revisions derived this key from ObjPtr(Me). That is unsafe: VBA
'   reuses heap addresses, so a session left behind by a destroyed instance
'   could be silently inherited by an unrelated instance later allocated at the
'   same address. Keys are therefore issued by PM_TW_NewInstanceKey instead.
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     The cached shared-TW key for this instance.
'
' BEHAVIOR
'   - Requests a new key from PM_TW_NewInstanceKey on first use.
'   - Caches the key for the lifetime of this instance.
'   - Returns the cached key on every subsequent call.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - PM_TW_NewInstanceKey (M_cPM_TimeWasters)
'   - m_InstanceKey
'
' NOTES
'   - Assignment is lazy, so instances that never request TW suppression consume
'     no key.
'   - The returned key is intended only for live in-memory instance tracking.
'   - It is stable for the lifetime of the object instance.
'   - It is not intended as a persisted or cross-session identifier.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ISSUE KEY ON FIRST USE
'------------------------------------------------------------------------------
    'Request and cache a collision-proof key the first time it is needed
        If Len(m_InstanceKey) = 0 Then
            m_InstanceKey = PM_TW_NewInstanceKey()
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the cached instance key
        TW_InstanceKey = m_InstanceKey

End Function

'
'------------------------------------------------------------------------------
'
'                       PRIVATE: TIMING SOURCE HELPERS
'
'------------------------------------------------------------------------------
'

Private Function QPC_TryReadTick( _
    ByRef TickOut As Currency) _
    As Boolean
'
'==============================================================================
'                              QPC TRY READ TICK
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the current QueryPerformanceCounter tick and reports whether the read
'   succeeded.
'
' WHY THIS EXISTS
'   This is the ONLY place in the class that calls QueryPerformanceCounter.
'   Centralizing the call means there is exactly one point where a failed read
'   can be detected, and no caller can accidentally consume an unread value.
'
' INPUTS
'   TickOut
'     Receives the raw QPC tick. Written ONLY when the read succeeds, so a
'     failed read cannot overwrite a caller's existing value with a sentinel.
'
' RETURNS
'   Boolean
'     TRUE  => the read succeeded and TickOut holds a valid tick
'     FALSE => the read failed and TickOut is untouched
'
' BEHAVIOR
'   - Consumes and clears the one-shot fault-injection flag when set.
'   - Calls QueryPerformanceCounter into a local.
'   - Assigns TickOut only on a non-zero API return code.
'
' ERROR POLICY
'   Never raises. The caller decides whether a failure is fatal.
'
' NOTES
'   The fault-injection flag is deliberately one-shot: it is cleared by the read
'   that consumes it. A sticky flag would make test ordering significant and
'   would introduce exactly the kind of hidden state this class exists to avoid.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Tick As Currency    'Local target, so TickOut is written only on success

'------------------------------------------------------------------------------
' FAULT INJECTION
'------------------------------------------------------------------------------
    'Consume the one-shot forced-failure flag, if the suite armed it
        If m_TestFailNextQpcRead Then
            m_TestFailNextQpcRead = False
            QPC_TryReadTick = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Report failure without touching the caller's variable
        If QueryPerformanceCounter(Tick) = 0 Then
            QPC_TryReadTick = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Publish the tick only now that the read is known to have succeeded
        TickOut = Tick
        QPC_TryReadTick = True

End Function


Private Function QPC_GetFrequencyHz() As Double
'
'==============================================================================
'                             QPC FREQUENCY HZ
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the QPC frequency in true ticks/second.
'
' WHY THIS EXISTS
'   The class stores m_qpcFreq as Currency for compatibility with the API usage
'   pattern in the QPC routines. This helper centralizes the conversion back to
'   true frequency units so that QPC diagnostics and related calculations remain
'   correct and consistent.
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     True QPC frequency in ticks/second.
'     Returns 0 when QPC frequency is unavailable.
'
' BEHAVIOR
'   - Returns 0 when the stored QPC frequency is unavailable.
'   - Otherwise undoes Currency's 10,000 scaling and returns the true frequency.
'
' ERROR POLICY
'   Does not raise errors.
'   Returns 0 when frequency information is unavailable.
'
' DEPENDENCIES
'   - m_qpcFreq
'
' NOTES
'   m_qpcFreq is stored as Currency, so we must undo Currency's 10,000 scaling.
'
'   Note that ElapsedSeconds does NOT use this helper. There the Currency
'   scaling cancels between the tick delta and the frequency, so the raw
'   Currency division is already correct.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Return 0 when frequency is unavailable
        If m_qpcFreq = 0 Then
            QPC_GetFrequencyHz = 0#
            Exit Function
        End If

    'Undo Currency scaling to obtain the true QPC frequency in Hz
        QPC_GetFrequencyHz = CDbl(m_qpcFreq) * 10000#
End Function

Private Function Get_TickCountSeconds() As Double
'
'==============================================================================
'                              GET TICKCOUNT SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the system uptime tick counter in seconds as a Double.
'
'     - Win64: uses GetTickCount64 (no practical rollover)
'     - Win32: uses GetTickCount (32-bit unsigned ms counter; wraps at 2^32 ms)
'
' WHY THIS EXISTS
'   VBA Long is signed. On 32-bit, GetTickCount returns a signed Long even
'   though the underlying counter is logically unsigned, so once the high bit is
'   set the returned value reads as negative.
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     Seconds since boot (subject to rollover on 32-bit).
'
' BEHAVIOR
'   - On Win64, converts the 64-bit millisecond counter directly.
'   - On 32-bit, delegates the signed-to-unsigned reinterpretation to
'     UInt32ToDouble rather than repeating it inline.
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   - GetTickCount
'   - GetTickCount64 (Win64/VBA7)
'   - UInt32ToDouble
'   - Conditional compilation symbols: VBA7, Win64
'
' BITNESS COVERAGE
'   The 32-bit branches cannot execute during a 64-bit test run, so they were
'   previously untestable in full. They no longer contain any arithmetic of
'   their own: the reinterpretation is UInt32ToDouble, which is Friend and is
'   verified at all four boundary values on whichever bitness the suite runs.
'
'   What remains bitness-specific here is the declaration binding and the API
'   return itself, not the semantics. That is a materially smaller surface, but
'   it is not zero, which is why 32-bit execution evidence is still required
'   before a major release.
'
' NOTES
'   The pre-VBA7 branch targets Office 2007 and earlier. VBA7 arrived with
'   Office 2010, so that branch is retained for completeness rather than because
'   any supported host reaches it. Whether to keep carrying it is a separate
'   decision from this one.
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
#If VBA7 Then
    #If Win64 Then
        '64-bit: GetTickCount64 returns a 64-bit millisecond counter; convert to seconds
            Get_TickCountSeconds = CDbl(GetTickCount64()) / 1000#
    #Else
        '32-bit: GetTickCount returns a DWORD in a signed Long, so reinterpret
        'it through the shared helper rather than repeating the correction here
            Get_TickCountSeconds = UInt32ToDouble(GetTickCount()) / 1000#
    #End If
#Else
        'Pre-VBA7: 32-bit only; identical reinterpretation
            Get_TickCountSeconds = UInt32ToDouble(GetTickCount()) / 1000#
#End If
End Function

Private Function SystemTime_TryReadMs( _
    ByRef ValueOut As Double, _
    ByRef StatusOut As cPM_ReadStatus) _
    As Boolean
'
'==============================================================================
'                           SYSTEM TIME TRY READ MS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the timeGetSystemTime millisecond counter and reports whether the read
'   succeeded.
'
' WHY THIS EXISTS
'   This is the ONLY place in the class that calls timeGetSystemTime.
'
'   The previous Get_SystemTimeMs returned 0 on failure in non-strict mode. That
'   sentinel then entered elapsed arithmetic, produced a negative delta, and was
'   "corrected" by adding the 32-bit rollover period - converting a failed read
'   into a large positive duration that passed validation.
'
'   Returning a status instead makes that impossible.
'
' INPUTS
'   ValueOut
'     Receives the logical unsigned millisecond value. Written ONLY on success.
'
'   StatusOut
'     Receives the specific failure reason, or cPM_ReadOK.
'
' RETURNS
'   Boolean
'     TRUE  => the read succeeded and ValueOut holds a valid value
'     FALSE => the read failed and ValueOut is untouched
'
' BEHAVIOR
'   - Consumes and clears the one-shot fault-injection flag when set.
'   - Resets the reusable MMTIME structure and requests TIME_MS.
'   - Distinguishes an API failure from an unexpected returned format.
'   - Assigns ValueOut only when both checks pass.
'
' ERROR POLICY
'   Never raises. The caller decides whether a failure is fatal.
'
' DEPENDENCIES
'   - timeGetSystemTime
'   - UInt32ToDouble
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC As Long    'Return code from timeGetSystemTime

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Assume success until a check fails
        StatusOut = cPM_ReadOK

'------------------------------------------------------------------------------
' FAULT INJECTION
'------------------------------------------------------------------------------
    'Consume the one-shot forced-failure flag, if the suite armed it
        If m_TestFailNextSystemTimeRead Then
            m_TestFailNextSystemTimeRead = False
            StatusOut = cPM_ReadSystemTimeFailed
            SystemTime_TryReadMs = False
            Exit Function
        End If

    'Consume the one-shot forced wrong-format flag, if the suite armed it.
    'This models a successful API call that returns a format other than
    'TIME_MS, which is a distinct condition from an outright failure.
        If m_TestForceNextSystemTimeFormatInvalid Then
            m_TestForceNextSystemTimeFormatInvalid = False
            StatusOut = cPM_ReadSystemTimeFormatInvalid
            SystemTime_TryReadMs = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' PREPARE STRUCTURE
'------------------------------------------------------------------------------
    'Reset the reusable structure before each API call
        m_Mmt.wType = CPM_TIME_MS
        m_Mmt.Units = 0&
        m_Mmt.Reserved = 0&

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Populate m_Mmt using TIME_MS format
        RC = timeGetSystemTime(m_Mmt, LenB(m_Mmt))

    'Report an API failure without touching the caller's variable
        If RC <> 0 Then
            StatusOut = cPM_ReadSystemTimeFailed
            SystemTime_TryReadMs = False
            Exit Function
        End If

    'Report an unexpected format separately from an outright failure
        If m_Mmt.wType <> CPM_TIME_MS Then
            StatusOut = cPM_ReadSystemTimeFormatInvalid
            SystemTime_TryReadMs = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Publish the value only now that the read is known to be valid
        ValueOut = UInt32ToDouble(m_Mmt.Units)
        SystemTime_TryReadMs = True

End Function


Friend Function RolloverSeconds( _
    ByVal iMethod As cPM_TimerMethod) _
    As Double
'
'==============================================================================
'                              ROLLOVER SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the rollover period (in seconds) for the selected timing backend.
'
' WHY THIS EXISTS
'   Methods 2, 3, and 4 do not share the same rollover behavior on Win64:
'     - Method 2 uses GetTickCount64 on Win64 => effectively no rollover
'     - Methods 3 and 4 still behave like 32-bit millisecond counters
'
' INPUTS
'   iMethod
'     Timing backend identifier.
'
' RETURNS
'   Double
'     Rollover period in seconds for the specified backend.
'     Returns 0 when rollover correction is not applicable here.
'
' BEHAVIOR
'   - For method 2:
'       * returns 0 on Win64/VBA7 because GetTickCount64 is used
'       * returns 2^32 milliseconds expressed in seconds otherwise
'   - For methods 3 and 4:
'       * returns 2^32 milliseconds expressed in seconds
'   - For all other methods:
'       * returns 0
'
' ERROR POLICY
'   Does not raise errors.
'   Unsupported/non-rollover methods simply return 0.
'
' DEPENDENCIES
'   - Conditional compilation symbols: VBA7, Win64
'
' NOTES
'   - The rollover value 2^32 milliseconds equals 4294967296# / 1000# seconds.
'   - This helper is intentionally limited to the backends whose rollover
'     correction is handled centrally in ElapsedSeconds().
'   - Returning 0 for the Win64 method 2 path means a negative elapsed value
'     survives the correction step unchanged. Elapsed_Validate is what catches
'     that case.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    Select Case iMethod
        Case cPM_MethodTickCount
            #If VBA7 Then
                #If Win64 Then
                        RolloverSeconds = 0#
                #Else
                        RolloverSeconds = 4294967296# / 1000#
                #End If
            #Else
                    RolloverSeconds = 4294967296# / 1000#
            #End If
        Case cPM_MethodTimeGetTime, cPM_MethodSystemTime
                RolloverSeconds = 4294967296# / 1000#
        Case Else
                RolloverSeconds = 0#

    End Select
End Function

Private Sub EnsureTimerResolution1ms()
'
'==============================================================================
'                        ENSURE TIMER RESOLUTION 1MS
'------------------------------------------------------------------------------
' PURPOSE
'   Requests 1ms multimedia timer resolution exactly once for this instance
'   when using timeGetTime-based timing paths.
'
' WHY THIS EXISTS
'   timeGetTime() can benefit from timer-resolution behavior. This helper
'   centralizes the one-time request logic so callers do not repeatedly invoke
'   timeBeginPeriod(1) in hot paths.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - If 1ms timer resolution has not yet been requested by this instance,
'     calls timeBeginPeriod 1.
'   - Marks m_TimePeriodOn = True only when the request succeeds.
'   - Raises in strict mode when the request fails.
'
' ERROR POLICY
'   - Raises in strict mode when timeBeginPeriod(1) fails.
'   - Continues without the resolution request in non-strict mode.
'
' DEPENDENCIES
'   - timeBeginPeriod
'   - m_TimePeriodOn
'   - m_StrictMode
'
' NOTES
'   ResetEnvironment is responsible for balancing a successful request with
'   timeEndPeriod 1.
'
'   This request affects system-wide timer resolution and can increase power
'   consumption while held.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC As Long    'Return code from timeBeginPeriod

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Exit immediately when this instance already owns the resolution request
        If m_TimePeriodOn Then Exit Sub
    'Request 1ms timer resolution
        RC = timeBeginPeriod(1)
    'Handle request failure explicitly
        If RC <> 0 Then
            If m_StrictMode Then
                Err.Raise ERR_CPM_TIMEBEGINPERIOD_FAILED, _
                          "cPerformanceManager.EnsureTimerResolution1ms", _
                          "timeBeginPeriod(1) failed."
            End If
            Exit Sub
        End If
    'Record successful ownership of the resolution request
        m_TimePeriodOn = True

End Sub

Private Function Timestamp_TryGetCurrentSeconds( _
    ByVal iMethod As cPM_TimerMethod, _
    ByRef ValueOut As Double, _
    ByRef StatusOut As cPM_ReadStatus) _
    As Boolean
'
'==============================================================================
'                      TIMESTAMP TRY GET CURRENT SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current timestamp in seconds for a non-QPC backend, reporting
'   whether the underlying read succeeded.
'
' WHY THIS EXISTS
'   Replaces Get_CurrentTimestampSeconds, which could not report a method-4
'   failure and therefore returned a sentinel that later entered arithmetic.
'
' INPUTS
'   iMethod    - non-QPC backend selector
'   ValueOut   - receives the timestamp in seconds; written ONLY on success
'   StatusOut  - receives the failure reason, or cPM_ReadOK
'
' RETURNS
'   Boolean
'     TRUE  => ValueOut holds a valid timestamp
'     FALSE => the read failed and ValueOut is untouched
'
' BEHAVIOR
'   - Backends 1, 2, 3 and 6 cannot fail and always report success.
'   - Backend 4 delegates to SystemTime_TryReadMs and propagates its status.
'
' ERROR POLICY
'   Never raises.
'
' NOTES
'   Method 5 is not handled here. QPC has its own dedicated capture path because
'   it works in Currency ticks rather than seconds.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ms As Double    'Raw millisecond value for backend 4

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Assume success until a read reports otherwise
        StatusOut = cPM_ReadOK

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    Select Case iMethod

        Case cPM_MethodTimer
                ValueOut = Timer
                Timestamp_TryGetCurrentSeconds = True

        Case cPM_MethodTickCount
                ValueOut = Get_TickCountSeconds()
                Timestamp_TryGetCurrentSeconds = True

        Case cPM_MethodTimeGetTime
                EnsureTimerResolution1ms
                ValueOut = UInt32ToDouble(timeGetTime()) / 1000#
                Timestamp_TryGetCurrentSeconds = True

        Case cPM_MethodSystemTime
                'The only non-QPC backend with a real failure path
                If Not SystemTime_TryReadMs(Ms, StatusOut) Then
                    Timestamp_TryGetCurrentSeconds = False
                    Exit Function
                End If
                ValueOut = Ms / 1000#
                Timestamp_TryGetCurrentSeconds = True

        Case cPM_MethodNow
                ValueOut = Now() * 86400#
                Timestamp_TryGetCurrentSeconds = True

        Case Else
                ValueOut = Timer
                Timestamp_TryGetCurrentSeconds = True

    End Select

End Function


Private Function Timestamp_TryGetNextSeconds( _
    ByVal iMethod As cPM_TimerMethod, _
    ByRef ValueOut As Double, _
    ByRef StatusOut As cPM_ReadStatus) _
    As Boolean
'
'==============================================================================
'                       TIMESTAMP TRY GET NEXT SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Aligns to the next observable tick for a non-QPC backend, reporting whether
'   the underlying reads succeeded.
'
' INPUTS
'   iMethod    - non-QPC backend selector
'   ValueOut   - receives the aligned timestamp; written ONLY on success
'   StatusOut  - receives the failure reason, or cPM_ReadOK
'
' RETURNS
'   Boolean
'     TRUE  => ValueOut holds a valid aligned timestamp
'     FALSE => a read failed and ValueOut is untouched
'
' BEHAVIOR
'   - Spins until the selected counter advances.
'   - Stops after CPM_ALIGN_MAX_SPINS iterations.
'   - Abandons alignment immediately if a backend-4 read fails.
'
' ERROR POLICY
'   - Raises ERR_CPM_ALIGN_TICK_TIMEOUT in strict mode if the spin guard is hit.
'   - Falls back to a current timestamp in non-strict mode.
'   - A failed READ is reported through the return value, not by raising.
'
' NOTES
'   Alignment failure and read failure are deliberately different outcomes. A
'   spin-guard timeout means the clock did not advance in time; a read failure
'   means no value was obtained at all.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim A           As Double   'Initial observed value
    Dim B           As Double   'Next observed value
    Dim SpinCount   As Long     'Finite-loop guard counter

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Assume success until a read reports otherwise
        StatusOut = cPM_ReadOK

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    Select Case iMethod

        Case cPM_MethodTimer
                A = Timer
                Do
                    SpinCount = SpinCount + 1
                    B = Timer
                    If B <> A Then
                        ValueOut = B
                        Timestamp_TryGetNextSeconds = True
                        Exit Function
                    End If
                Loop Until SpinCount >= CPM_ALIGN_MAX_SPINS

        Case cPM_MethodTickCount
                A = Get_TickCountSeconds()
                Do
                    SpinCount = SpinCount + 1
                    B = Get_TickCountSeconds()
                    If B <> A Then
                        ValueOut = B
                        Timestamp_TryGetNextSeconds = True
                        Exit Function
                    End If
                Loop Until SpinCount >= CPM_ALIGN_MAX_SPINS

        Case cPM_MethodTimeGetTime
                EnsureTimerResolution1ms
                A = UInt32ToDouble(timeGetTime())
                Do
                    SpinCount = SpinCount + 1
                    B = UInt32ToDouble(timeGetTime())
                    If B <> A Then
                        ValueOut = B / 1000#
                        Timestamp_TryGetNextSeconds = True
                        Exit Function
                    End If
                Loop Until SpinCount >= CPM_ALIGN_MAX_SPINS

        Case cPM_MethodSystemTime
                'Abandon alignment entirely if the initial read fails
                If Not SystemTime_TryReadMs(A, StatusOut) Then
                    Timestamp_TryGetNextSeconds = False
                    Exit Function
                End If

                Do
                    SpinCount = SpinCount + 1
                    'A failure mid-alignment is still a failure
                    If Not SystemTime_TryReadMs(B, StatusOut) Then
                        Timestamp_TryGetNextSeconds = False
                        Exit Function
                    End If
                    If B <> A Then
                        ValueOut = B / 1000#
                        Timestamp_TryGetNextSeconds = True
                        Exit Function
                    End If
                Loop Until SpinCount >= CPM_ALIGN_MAX_SPINS

        Case cPM_MethodNow
                'Now() cannot be meaningfully aligned at this granularity
                ValueOut = Now() * 86400#
                Timestamp_TryGetNextSeconds = True
                Exit Function

        Case Else
                Timestamp_TryGetNextSeconds = _
                    Timestamp_TryGetCurrentSeconds(iMethod, ValueOut, StatusOut)
                Exit Function

    End Select

'------------------------------------------------------------------------------
' ALIGNMENT FAILURE
'------------------------------------------------------------------------------
    'The clock did not advance within the spin guard
        If m_StrictMode Then
            Err.Raise ERR_CPM_ALIGN_TICK_TIMEOUT, _
                      "cPerformanceManager.Timestamp_TryGetNextSeconds", _
                      "Next-tick alignment did not complete within the spin guard."
        End If

    'Non-strict: settle for a current timestamp
        Timestamp_TryGetNextSeconds = _
            Timestamp_TryGetCurrentSeconds(iMethod, ValueOut, StatusOut)

End Function


Private Function QPC_TryGetNextTick( _
    ByRef TickOut As Currency) _
    As Boolean
'
'==============================================================================
'                           QPC TRY GET NEXT TICK
'------------------------------------------------------------------------------
' PURPOSE
'   Aligns to the next QueryPerformanceCounter tick, reporting whether the reads
'   succeeded.
'
' INPUTS
'   TickOut
'     Receives the first tick strictly greater than the initial read. Written
'     ONLY on success.
'
' RETURNS
'   Boolean
'     TRUE  => TickOut holds a valid aligned tick
'     FALSE => a QPC read failed and TickOut is untouched
'
' BEHAVIOR
'   - Captures an initial tick.
'   - Spins until a strictly greater tick is observed.
'   - Stops after CPM_ALIGN_MAX_SPINS attempts.
'
' ERROR POLICY
'   - Raises ERR_CPM_ALIGN_TICK_TIMEOUT in strict mode if the spin guard is hit.
'   - Returns the initial tick in non-strict mode on spin-guard timeout.
'   - A failed READ is reported through the return value, not by raising, so the
'     caller can fall back to another backend before committing any state.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TickA       As Currency     'Initial tick
    Dim TickB       As Currency     'Next tick
    Dim SpinCount   As Long         'Finite-loop guard counter

'------------------------------------------------------------------------------
' INITIAL READ
'------------------------------------------------------------------------------
    'A failed initial read means no alignment is possible
        If Not QPC_TryReadTick(TickA) Then
            QPC_TryGetNextTick = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' ALIGN TO NEXT TICK
'------------------------------------------------------------------------------
    'Spin until the counter advances strictly beyond the initial read
        Do
            SpinCount = SpinCount + 1

            If Not QPC_TryReadTick(TickB) Then
                QPC_TryGetNextTick = False
                Exit Function
            End If

            If TickB > TickA Then
                TickOut = TickB
                QPC_TryGetNextTick = True
                Exit Function
            End If
        Loop Until SpinCount >= CPM_ALIGN_MAX_SPINS

'------------------------------------------------------------------------------
' ALIGNMENT FAILURE
'------------------------------------------------------------------------------
    'The counter did not advance within the spin guard
        If m_StrictMode Then
            Err.Raise ERR_CPM_ALIGN_TICK_TIMEOUT, _
                      "cPerformanceManager.QPC_TryGetNextTick", _
                      "QPC next-tick alignment did not complete within the spin guard."
        End If

    'Non-strict: the initial tick is still a valid reading
        TickOut = TickA
        QPC_TryGetNextTick = True

End Function


'
'------------------------------------------------------------------------------
'
'                PRIVATE: CHECKPOINT / REPORT HELPER ROUTINES
'
'------------------------------------------------------------------------------
'

Private Sub Checkpoint_ClearState()
'
'==============================================================================
'                           CHECKPOINT CLEAR STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Resets the structured checkpoint state for a fresh timing session
'
' WHY THIS EXISTS
'   Checkpoints and run label belong to a timing session. Starting a new session
'   should therefore reset the structured reporting state deterministically.
'
'   This same helper is also used by ClearCheckpoints() and by Class_Initialize,
'   so there is exactly one definition of a clean checkpoint baseline.
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Clears the current run label
'   - Clears the checkpoint count
'   - Clears the last checkpoint marker
'   - Reinitializes checkpoint storage to a small first capacity
'
' ERROR POLICY
'   Raises errors normally
'
' NOTES
'   StartTimer calls this routine and then reapplies its optional RunLabel
'   argument, which is why a label supplied to StartTimer survives the reset but
'   one set beforehand through SetRunLabel does not.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' RESET STATE
'------------------------------------------------------------------------------
    'Clear the current run label
        m_RunLabel = vbNullString
    'Clear the checkpoint counter
        m_CheckpointCount = 0
    'Clear the last checkpoint elapsed-seconds marker
        m_LastCheckpointS = 0#
    'Reinitialize checkpoint storage with a small first capacity
        ReDim m_Checkpoints(1 To 1)

End Sub

Private Sub Checkpoint_EnsureCapacity( _
    ByVal RequiredCount As Long)
'
'==============================================================================
'                        CHECKPOINT ENSURE CAPACITY
'------------------------------------------------------------------------------
' PURPOSE
'   Ensures that checkpoint storage can hold the requested number of rows
'
' WHY THIS EXISTS
'   Checkpoints are appended dynamically. This helper centralizes array growth
'   logic and avoids repeated ReDim Preserve on every single row
'
' INPUTS
'   RequiredCount
'     Minimum required checkpoint capacity
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Leaves storage unchanged when current capacity is already sufficient
'   - Otherwise grows storage using a simple doubling strategy
'
' ERROR POLICY
'   Raises errors normally
'
' NOTES
'   Storage starts at a capacity of 1, so the first growth occurs on the second
'   checkpoint. Doubling keeps the amortized cost of ReDim Preserve low across
'   long checkpoint runs.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentCapacity     As Long                 'Current checkpoint storage capacity
    Dim NewCapacity         As Long                 'New checkpoint storage capacity

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the current checkpoint storage capacity
        CurrentCapacity = UBound(m_Checkpoints)

'------------------------------------------------------------------------------
' EXIT EARLY WHEN CAPACITY IS ALREADY SUFFICIENT
'------------------------------------------------------------------------------
    'Leave storage unchanged when the requested capacity is already available
        If RequiredCount <= CurrentCapacity Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPUTE NEW CAPACITY
'------------------------------------------------------------------------------
    'Start from a simple doubling growth rule
        NewCapacity = CurrentCapacity * 2
    'Ensure the new capacity is never smaller than the required count
        If NewCapacity < RequiredCount Then
            NewCapacity = RequiredCount
        End If

'------------------------------------------------------------------------------
' GROW STORAGE
'------------------------------------------------------------------------------
    'Resize checkpoint storage while preserving existing rows
        ReDim Preserve m_Checkpoints(1 To NewCapacity)

End Sub

Private Function Checkpoint_DefaultName( _
    ByVal Seq As Long) _
    As String
'
'==============================================================================
'                         CHECKPOINT DEFAULT NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a generated fallback checkpoint name for blank checkpoint labels
'
' WHY THIS EXISTS
'   Structured reports are easier to read when every checkpoint row has a
'   non-blank label, even when the caller omitted one
'
' INPUTS
'   Seq
'     Sequence number that will be embedded in the generated label
'
' RETURNS
'   String
'     Generated checkpoint label
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a generated checkpoint label
        Checkpoint_DefaultName = "Checkpoint " & CStr(Seq)

End Function

'
'------------------------------------------------------------------------------
'
'                 PUBLIC: MEASUREMENT HARNESS / STATISTICS
'
'------------------------------------------------------------------------------
'
'   Statistics are computed over a plain Double() sample vector so callers can
'   keep, export, or post-process the raw per-run values.
'
'   Typical use:
'
'       Dim S() As Double
'       S = cPM.MeasureProcedure("RebuildPivotCache", 30, 3)
'       Debug.Print cPM.Stats_Text(S)
'
'   Report MEDIAN and MIN in preference to MEAN. Timing distributions are
'   right-skewed: one scheduler preemption inflates the mean while leaving the
'   median almost untouched.
'

Public Function MeasureProcedure( _
    ByVal ProcedureName As String, _
    Optional ByVal Iterations As Long = 30, _
    Optional ByVal WarmupIterations As Long = 3, _
    Optional ByVal iMethod As cPM_TimerMethod = cPM_MethodQPC, _
    Optional ByRef FailedReadsOut As Long, _
    Optional ByRef LastFailureStatusOut As cPM_ReadStatus, _
    Optional ByRef RejectedSamplesOut As Long) _
    As Double()
'
'==============================================================================
'                              MEASURE PROCEDURE
'------------------------------------------------------------------------------
' PURPOSE
'   Runs a named VBA procedure repeatedly and returns the per-run elapsed times
'
' WHY THIS EXISTS
'   A single timing sample cannot show whether a result is stable. Benchmarking
'   requires repeated runs and the full vector of per-run values, so the caller
'   can inspect the distribution rather than trusting one number
'
'   Building that loop by hand in every caller is exactly the duplicated,
'   error-prone code this class exists to remove
'
' INPUTS
'   ProcedureName
'     Name of a Public Sub in a standard module, as accepted by Application.Run
'
'   Iterations (optional)
'     Number of measured runs. Coerced to a minimum of 1
'
'   WarmupIterations (optional)
'     Number of discarded runs executed before measurement begins.
'     Coerced to a minimum of 0
'
'   iMethod (optional)
'     Timing backend to use
'
'   FailedReadsOut (optional, ByRef)
'     Native read failures, warm-up included
'
'   LastFailureStatusOut (optional, ByRef)
'     Status of the most recent failure or rejection
'
'   RejectedSamplesOut (optional, ByRef)
'     Measured cycles rejected because the requested backend fell back
'
' RETURNS
'   Double()
'     1-based vector of per-run elapsed seconds, every element measured on the
'     requested backend. In non-strict mode a cycle that fails to read, or that
'     falls back to another backend, contributes NO element, so the vector may
'     be shorter than Iterations
'
' BEHAVIOR
'   - Validates the procedure name and coerces the loop counts
'   - Normalizes and validates the requested timer method
'   - Runs the warm-up iterations and discards their timings
'   - Runs the measured iterations, classifying each before recording it
'   - Always attempts to clean up the isolated worker instance
'
' ERROR POLICY
'   - Raises on a blank procedure name
'   - Raises ERR_CPM_MEASURE_NO_VALID_SAMPLES when no measured read succeeded
'   - Propagates any error raised by the measured procedure, after cleanup
'   - Timer validation is delegated to Method_NormalizeStart
'
' MEASURED-CYCLE CLASSIFICATION
'   Every measured cycle ends in exactly one of three states.
'
'   RETAINED
'     The requested backend started and the endpoint read succeeded. The value
'     is appended to the vector.
'
'   REJECTED
'     The requested backend could not start and the session fell back to
'     method 2. The read would have succeeded, on the wrong clock. Retaining it
'     would mix backends inside a vector the caller believes is homogeneous, so
'     the observation is discarded and counted in RejectedSamplesOut.
'
'     A rejected cycle does NOT run the workload. See BEHAVIORAL CORRECTION.
'
'   FAILED
'     The requested backend started, but the endpoint read failed. In non-strict
'     mode that read returns zero, and storing the zero would make the failure
'     look like the fastest run in the set, materially distorting the minimum
'     and, on small samples, the median. The sample is excluded and counted in
'     FailedReadsOut.
'
'   The three are counted separately because their remedies differ. A failed
'   read says the host's clock is unreliable; a rejection says this host cannot
'   start the backend that was asked for. One counter could not say which.
'
' COUNT EQUATION
'
'     Iterations = Stats_Count(returned vector)
'                + measured native failures
'                + RejectedSamplesOut
'
'   FailedReadsOut counts native read failures INCLUDING warm-up, so it can
'   exceed the measured native-failure term above. Measured native failures are
'   therefore derived, not published:
'
'     measured native failures = Iterations - Stats_Count - RejectedSamplesOut
'
'   RejectedSamplesOut excludes warm-up. That asymmetry is deliberate: a warm-up
'   read that fails is evidence the host cannot be measured reliably, whereas a
'   warm-up fallback determines no returned observation, because every measured
'   iteration starts and validates a new session of its own.
'
' EVIDENCE BOUNDARY
'   The optional outputs are publication targets, not running totals. They are
'   authoritative in exactly two situations:
'
'     - a normal return; and
'     - the aggregate raise below, where they are published before raising
'
'   On any other propagated error they have not reached their publication point.
'   That covers the blank-name raise, the worker raising in strict mode, the
'   defensive internal-invariant raise, and anything a future change adds.
'
'   The boundary is enforced rather than merely documented. Both counts are set
'   to -1 on entry, before any validation that can raise:
'
'     FailedReadsOut < 0     this run published nothing; read Err instead
'     FailedReadsOut >= 0    the outputs are authoritative
'
'   LastFailureStatusOut is meaningful only while the counts are non-negative.
'
'   Writing zeros instead would repeat the mistake this class already corrected
'   for failed reads: it makes an aborted run indistinguishable from a flawless
'   one, in the one situation where the caller most needs to tell them apart.
'   Err_Handler deliberately does not touch the outputs, so evidence already
'   published by the aggregate raise survives being re-raised.
'
' BEHAVIORAL CORRECTION IN 1.4.0
'   v1.3.0 read the worker's status only after ElapsedSeconds, which resets it
'   on entry. A start fallback was therefore erased before it could be seen, and
'   the method-2 observation was retained as though it came from the requested
'   backend. Vectors could silently mix backends.
'
'   v1.4.0 captures the start transaction immediately after StartTimer and
'   rejects the cycle. Two consequences for existing callers:
'
'     - a non-strict run on a host that cannot start the requested backend now
'       returns a shorter vector, or raises ERR_CPM_MEASURE_NO_VALID_SAMPLES,
'       where v1.3.0 returned mixed-backend numbers; and
'     - a rejected cycle does not dispatch the workload, so a procedure with
'       side effects executes fewer times than Iterations requested. v1.3.0 ran
'       it every time. Executing a workload whose observation cannot be kept
'       buys nothing and costs both time and those side effects.
'
' DEPENDENCIES
'   - Method_NormalizeStart
'   - Application.Run
'   - StartTimer
'   - ElapsedSeconds
'   - ResetEnvironment
'
' TRUSTED INPUT
'   ProcedureName is passed to Application.Run, which executes it. Supplying a
'   name that came from a worksheet cell, a configuration sheet, a file or any
'   other untrusted source is equivalent to letting that source run arbitrary
'   code inside this VBA project.
'
'   Qualification narrows a different risk - a same-named procedure in another
'   open workbook being selected instead of the intended one - but it does not
'   address this one. Validating the string would not help either: any
'   syntactically valid procedure name is executable by design, so a filter
'   would look like protection without being any.
'
'   Pass literals, or names your own code controls.
'
' NOTES
'   - Application.Run only reaches Public procedures in standard modules. It
'     cannot call Private procedures, class methods, or procedures in a module
'     declared Option Private Module
'   - Application.Run adds a dispatch cost to every sample. Measure it on the
'     target machine rather than assuming a figure; it varies by Excel version,
'     bitness and load. This harness suits procedures measured in milliseconds
'     or longer
'   - MeasureOverhead_Samples measures the backend timing cycle only and does
'     NOT dispatch through Application.Run, so it is not a matched baseline for
'     this routine. For a matched baseline, run MeasureProcedure against an
'     empty Public Sub
'   - Measurement runs on an isolated worker instance, so the caller's session
'     and checkpoint state are untouched
'   - The trap this routine fell into is open to any caller. LastReadStatus
'     describes the most recent read operation and is reset by each one, so a
'     start fallback is gone as soon as ElapsedSeconds runs. ActiveMethodID is
'     the durable indicator: compare it against the backend you asked for
'
' UPDATED
'   2026-08-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Samples()           As Double                'Per-run elapsed seconds
    Dim K                   As Long                  'Loop counter
    Dim ResolvedMethod      As cPM_TimerMethod       'Validated timer method
    Dim cPM_Work            As cPerformanceManager   'Isolated worker instance
    Dim DiscardedElapsedS   As Double                'Throwaway warm-up value
    Dim RunTarget           As String                'Workbook-qualified procedure name
    Dim FailedReads         As Long                  'Native read failures, warm-up included
    Dim RejectedSamples     As Long                  'Measured cycles rejected for backend fallback
    Dim ValidCount          As Long                  'Measured reads that produced a value
    Dim LastFailure         As cPM_ReadStatus        'Status of the most recent failure or rejection
    Dim ThisSample          As Double                'Elapsed value for the current iteration
    Dim StartStatus         As cPM_ReadStatus        'Status of the start transaction
    Dim StartedMethod       As cPM_TimerMethod       'Backend the start transaction actually bound
    Dim EndpointStatus      As cPM_ReadStatus        'Status of the endpoint read
    Dim ErrNumber           As Long                  'Captured error number
    Dim ErrSource           As String                'Captured error source
    Dim ErrDescription      As String                'Captured error description

'------------------------------------------------------------------------------
' SENTINEL THE OUTPUTS
'------------------------------------------------------------------------------
    'Mark the optional outputs as NOT PUBLISHED before anything below can raise.
    'A run that ends before its cycles are classified must not leave the caller
    'reading zeros, which are indistinguishable from a flawless run. A negative
    'count says this run never reached its publication point, and Err carries
    'the diagnosis instead.
    '
    'This has to precede the validation below, or a blank-name raise would leave
    'zeros, or whatever the caller's variables happened to hold already
        FailedReadsOut = -1
        RejectedSamplesOut = -1
        LastFailureStatusOut = cPM_ReadOK

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank procedure name
        If Len(Trim$(ProcedureName)) = 0 Then
            Err.Raise ERR_CPM_MEASURE_BLANK_PROC, _
                      "cPerformanceManager.MeasureProcedure", _
                      "ProcedureName cannot be blank."
        End If

    'Coerce the loop counts into a usable range
        If Iterations < 1 Then Iterations = 1
        If WarmupIterations < 0 Then WarmupIterations = 0

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route runtime failures through cleanup
        On Error GoTo Err_Handler

    'The outputs stay sentinelled until classification completes. Only the two
    'publication points below overwrite them: the normal return, and the
    'aggregate raise which publishes before raising
        LastFailure = cPM_ReadOK

    'Normalize and validate the timer method using the current instance policy
        ResolvedMethod = Method_NormalizeStart(iMethod)

    'Resolve the target deterministically rather than leaving it to whichever
    'workbook happens to be active
        RunTarget = Measure_QualifyProcedureName(ProcedureName)

    'Size the result vector
        ReDim Samples(1 To Iterations)

'------------------------------------------------------------------------------
' INITIALIZE ISOLATED WORKER
'------------------------------------------------------------------------------
    'Create an isolated worker instance
        Set cPM_Work = New cPerformanceManager
    'Mirror the current strict-mode policy into the isolated worker
        cPM_Work.StrictMode = m_StrictMode

'------------------------------------------------------------------------------
' WARM UP
'------------------------------------------------------------------------------
    'Run and discard the warm-up iterations
        For K = 1 To WarmupIterations
            cPM_Work.StartTimer ResolvedMethod, False
            Application.Run RunTarget
            DiscardedElapsedS = cPM_Work.ElapsedSeconds()

            'A warm-up read that fails is still evidence the host cannot be
            'measured reliably, so it is counted even though its value is
            'discarded either way
                If cPM_Work.LastReadStatus <> cPM_ReadOK Then
                    FailedReads = FailedReads + 1
                    LastFailure = cPM_Work.LastReadStatus
                End If
        Next K

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Record one elapsed value per measured iteration
        For K = 1 To Iterations

            'Arm a forced start failure before the session opens, so the
            'fallback it produces is the one this cycle is classified on
                Measure_ArmWorkerStartFault cPM_Work, ResolvedMethod

            cPM_Work.StartTimer ResolvedMethod, False

            'Capture the whole start transaction before anything can overwrite
            'it. ElapsedSeconds resets LastReadStatus on entry, so reading the
            'start status afterwards reports the endpoint's outcome and loses
            'the fallback entirely. That erasure is the v1.3.0 defect
                StartStatus = cPM_Work.LastReadStatus
                StartedMethod = cPM_Work.ActiveMethodID

            If StartStatus = cPM_ReadFallbackToMethod2 Then

                'The read succeeded, on the wrong backend. Retaining it would
                'mix method-2 observations into a vector the caller believes is
                'homogeneous, so the cycle is rejected rather than failed. No
                'workload runs: an observation that cannot be kept is not worth
                'the time or the side effects of producing it
                    RejectedSamples = RejectedSamples + 1
                    LastFailure = StartStatus

            ElseIf StartedMethod <> ResolvedMethod Then

                'Unreachable under the current start transaction, which either
                'binds the requested backend, falls back and says so, or raises.
                'If it is ever reached the session is not what it claims to be,
                'and no cPM_ReadStatus describes that truthfully. Raise rather
                'than invent a rejection status
                    Err.Raise ERR_CPM_METHOD_MISMATCH, _
                              "cPerformanceManager.MeasureProcedure", _
                              "Internal invariant violated: StartTimer reported success but " & _
                              "bound method " & CLng(StartedMethod) & _
                              " instead of expected method " & CLng(ResolvedMethod) & "."

            Else

                Application.Run RunTarget

                'Arm a forced endpoint failure only after the workload returns,
                'so the injected fault lands on the measured read
                    Measure_ArmWorkerEndpointFault cPM_Work, ResolvedMethod

                ThisSample = cPM_Work.ElapsedSeconds()
                EndpointStatus = cPM_Work.LastReadStatus

                'A failed read returns zero in non-strict mode. Storing that
                'zero would make the failure look like the fastest run in the
                'set, so the sample is excluded rather than recorded.
                '
                'Strict mode never reaches here: the worker mirrors the caller's
                'policy and raises, which propagates through Err_Handler
                    If EndpointStatus <> cPM_ReadOK Then
                        FailedReads = FailedReads + 1
                        LastFailure = EndpointStatus
                    Else
                        ValidCount = ValidCount + 1
                        Samples(ValidCount) = ThisSample
                    End If

            End If
        Next K

    'An empty vector is not a measurement, and every statistic over it would
    'raise anyway. Fail here, where the reason is still known.
    '
    'The outputs are published first. Without that the all-rejected path would
    'destroy the only evidence explaining it, leaving a caller who traps the
    'error unable to tell a broken clock from a backend this host cannot start
        If ValidCount = 0 Then
            FailedReadsOut = FailedReads
            LastFailureStatusOut = LastFailure
            RejectedSamplesOut = RejectedSamples
            m_TestWorkerStartFailures = 0
            m_TestWorkerEndpointFailures = 0
            Err.Raise ERR_CPM_MEASURE_NO_VALID_SAMPLES, _
                      "cPerformanceManager.MeasureProcedure", _
                      Measure_NoValidSamplesText(Iterations, RejectedSamples, _
                                                 ResolvedMethod, LastFailure)
        End If

    'Trim to the readings actually obtained
        If ValidCount < Iterations Then
            ReDim Preserve Samples(1 To ValidCount)
        End If

    'The worker is released below, so its outcome has to be surfaced here or it
    'is lost. The caller's own LastReadStatus is deliberately not touched: it
    'describes the caller's own reads, not a run delegated to a worker
        FailedReadsOut = FailedReads
        LastFailureStatusOut = LastFailure
        RejectedSamplesOut = RejectedSamples
        m_TestWorkerStartFailures = 0
        m_TestWorkerEndpointFailures = 0

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the full per-run sample vector
        MeasureProcedure = Samples

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Release the isolated worker instance
        On Error Resume Next
        If Not cPM_Work Is Nothing Then cPM_Work.ResetEnvironment
        Set cPM_Work = Nothing
        On Error GoTo 0
        Exit Function

Err_Handler:
'------------------------------------------------------------------------------
' CAPTURE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Capture the original error before cleanup
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description

'------------------------------------------------------------------------------
' CLEAN UP AFTER FAILURE
'------------------------------------------------------------------------------
    'Cleanup must be best-effort so that we can preserve the original error
        On Error Resume Next
        If Not cPM_Work Is Nothing Then cPM_Work.ResetEnvironment
        Set cPM_Work = Nothing
        On Error GoTo 0

    'Disarm on the error path as well, or a forced failure armed for this run
    'leaks into the next one and produces a result nobody asked for. The
    'optional outputs are deliberately NOT written here: at this point they hold
    'nothing the caller could act on, and publishing zeros would say that
    'nothing failed
        m_TestWorkerStartFailures = 0
        m_TestWorkerEndpointFailures = 0

'------------------------------------------------------------------------------
' RE-RAISE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Re-raise the original error
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

Public Function MeasureBaseline( _
    ByVal EmptyProcedureName As String, _
    Optional ByVal Iterations As Long = 30, _
    Optional ByVal WarmupIterations As Long = 3, _
    Optional ByVal iMethod As cPM_TimerMethod = cPM_MethodQPC, _
    Optional ByRef FailedReadsOut As Long, _
    Optional ByRef LastFailureStatusOut As cPM_ReadStatus, _
    Optional ByRef RejectedSamplesOut As Long) _
    As Double()
'
'==============================================================================
'                              MEASURE BASELINE
'------------------------------------------------------------------------------
' PURPOSE
'   Measures the cost of the harness itself, through the same dispatch path that
'   MeasureProcedure uses
'
' WHY THIS EXISTS
'   MeasureOverhead_Samples measures the timing cycle only. It does not call
'   Application.Run, so subtracting it from a MeasureProcedure result removes
'   the wrong quantity and leaves the dispatch cost in place
'
'   This routine runs an empty procedure through the identical path, so its
'   median is a baseline that can legitimately be subtracted
'
' INPUTS
'   EmptyProcedureName
'     Name of an empty Public Sub in a standard module
'
'   Iterations, WarmupIterations, iMethod
'     As for MeasureProcedure
'
'   FailedReadsOut, LastFailureStatusOut, RejectedSamplesOut (optional, ByRef)
'     As for MeasureProcedure, forwarded unchanged from the delegated run
'
' RETURNS
'   Double()
'     1-based vector of per-run dispatch cost in seconds, every element measured
'     on the requested backend
'
' BEHAVIOR
'   Delegates to MeasureProcedure, so qualification, isolation, classification
'   and failure handling are identical by construction rather than by
'   duplication
'
' ERROR POLICY
'   Raises ERR_CPM_MEASURE_NO_BASELINE on a blank procedure name, with a message
'   naming what to create. Everything else propagates from MeasureProcedure,
'   including ERR_CPM_MEASURE_NO_VALID_SAMPLES when no baseline cycle survived
'
' EVIDENCE
'   Before v1.4.0 this routine exposed no failure evidence at all. A caller
'   could see that a workload vector had been shortened, but had no way to ask
'   whether the baseline it was subtracting had been shortened for the same
'   reason, or by how much. Comparing a clean 30-sample workload against a
'   baseline quietly reduced to 4 samples is not a comparison, and nothing said
'   so.
'
'   The three optional outputs are forwarded from the delegated run and carry
'   the same meaning and the same evidence boundary as MeasureProcedure's,
'   including the entry sentinel of -1 on both counts. This routine applies that
'   sentinel itself, before its own blank-name validation, so an error raised
'   here rather than in the delegated run still leaves the caller able to tell
'   that nothing was published.
'
' TRUSTED INPUT
'   EmptyProcedureName reaches Application.Run through MeasureProcedure and is
'   executed. The same caution applies: pass literals, or names your own code
'   controls, never a value read from a worksheet or a file.
'
' WHY THE CALLER SUPPLIES THE PROCEDURE
'   Application.Run cannot reach a procedure in a module declared
'   Option Private Module, which is what the companion TW module uses to keep
'   its internals out of the Macro dialog. An empty procedure shipped with the
'   component would therefore need a third file whose entire content is two
'   lines, and the import contract is deliberately two files.
'
'   Add this once, anywhere in a standard module:
'
'       Public Sub cPM_BaselineEmpty()
'       End Sub
'
' USAGE
'       Dim Work() As Double, Base() As Double
'       Work = cPM.MeasureProcedure("MyWorkload", 30, 3)
'       Base = cPM.MeasureBaseline("cPM_BaselineEmpty", 30, 3)
'
'       Debug.Print "Net: " & (cPM.Stats_Median(Work) - cPM.Stats_Median(Base))
'
' NOTES
'   Subtract medians, not means. Both vectors are right-skewed, and subtracting
'   two means compounds the contamination in each
'
'   The baseline is machine-specific and worth re-measuring on each host
'
'   Measure the baseline on the same backend as the workload. A baseline
'   rejected for fallback while the workload was not, or the reverse, is not a
'   matched pair no matter how close the medians look
'
' UPDATED
'   2026-08-30
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Sentinel the outputs before the validation below can raise, so a blank-name
    'error cannot leave the caller reading a clean-run shape. MeasureProcedure
    'sentinels them again on entry; doing it here as well is what makes the
    'guarantee hold for errors raised before delegation
        FailedReadsOut = -1
        RejectedSamplesOut = -1
        LastFailureStatusOut = cPM_ReadOK

    'Reject a blank name with instructions rather than a bare failure
        If Len(Trim$(EmptyProcedureName)) = 0 Then
            Err.Raise ERR_CPM_MEASURE_NO_BASELINE, _
                      "cPerformanceManager.MeasureBaseline", _
                      "EmptyProcedureName cannot be blank. Add an empty Public Sub to " & _
                      "a standard module and pass its name, for example: " & _
                      "Public Sub cPM_BaselineEmpty() / End Sub"
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Measure through the identical path so the baseline is genuinely matched,
    'and forward its evidence. Without this the caller can see that a workload
    'vector was shortened but not whether the baseline it is being compared
    'against was shortened for the same reason
        MeasureBaseline = MeasureProcedure(EmptyProcedureName, Iterations, _
                                           WarmupIterations, iMethod, _
                                           FailedReadsOut, LastFailureStatusOut, _
                                           RejectedSamplesOut)

End Function

Public Function MeasureOverhead_Samples( _
    Optional ByVal Iterations As Long = 1000, _
    Optional ByVal WarmupIterations As Long = 10, _
    Optional ByVal iMethod As cPM_TimerMethod = cPM_MethodQPC) _
    As Double()
'
'==============================================================================
'                          MEASURE OVERHEAD SAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the per-run near-empty measurement overhead as a sample vector
'
' WHY THIS EXISTS
'   OverheadMeasurement_Seconds returns only the arithmetic mean, which hides
'   the shape of the distribution. This routine returns every sample so the
'   caller can report median, minimum, and spread
'
'   The minimum of this vector is the practical resolution floor of the selected
'   backend: intervals shorter than it cannot be measured meaningfully
'
' INPUTS
'   Iterations (optional)
'     Number of near-empty timing cycles to record. Coerced to a minimum of 1
'
'   WarmupIterations (optional)
'     Number of discarded cycles before measurement. Coerced to a minimum of 0
'
'   iMethod (optional)
'     Timing backend to benchmark
'
' RETURNS
'   Double()
'     1-based vector of per-cycle overhead in seconds. A cycle whose read failed
'     contributes no element, so the vector may be shorter than Iterations
'
' BEHAVIOR
'   - Coerces the loop counts
'   - Normalizes and validates the requested timer method
'   - Runs and discards the warm-up cycles
'   - Records one elapsed value per measured cycle
'   - Always attempts to clean up the isolated worker instance
'
' ERROR POLICY
'   Raises validation/runtime errors normally after cleanup has been attempted
'
' DEPENDENCIES
'   - Method_NormalizeStart
'   - StartTimer
'   - ElapsedSeconds
'   - ResetEnvironment
'
' NOTES
'   This is the distribution-aware companion to OverheadMeasurement_Seconds.
'   The warm-up default is 10 rather than 1, because a single discarded cycle is
'   not enough to settle code paths and caches
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Samples()           As Double                'Per-cycle overhead in seconds
    Dim K                   As Long                  'Loop counter
    Dim ResolvedMethod      As cPM_TimerMethod               'Validated timer method
    Dim cPM_Work            As cPerformanceManager   'Isolated worker instance
    Dim DiscardedElapsedS   As Double                'Throwaway warm-up value
    Dim ValidCount          As Long                  'Measured cycles that produced a value
    Dim LastFailure         As cPM_ReadStatus        'Status of the most recent failed read
    Dim ThisSample          As Double                'Elapsed value for the current cycle
    Dim ErrNumber           As Long                  'Captured error number
    Dim ErrSource           As String                'Captured error source
    Dim ErrDescription      As String                'Captured error description

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Coerce the loop counts into a usable range
        If Iterations < 1 Then Iterations = 1
        If WarmupIterations < 0 Then WarmupIterations = 0

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route runtime failures through cleanup
        On Error GoTo Err_Handler

    'Normalize and validate the timer method using the current instance policy
        ResolvedMethod = Method_NormalizeStart(iMethod)

    'Size the result vector
        ReDim Samples(1 To Iterations)

'------------------------------------------------------------------------------
' INITIALIZE ISOLATED WORKER
'------------------------------------------------------------------------------
    'Create an isolated worker instance
        Set cPM_Work = New cPerformanceManager
    'Mirror the current strict-mode policy into the isolated worker
        cPM_Work.StrictMode = m_StrictMode

'------------------------------------------------------------------------------
' WARM UP
'------------------------------------------------------------------------------
    'Run and discard the warm-up cycles
        For K = 1 To WarmupIterations
            cPM_Work.StartTimer ResolvedMethod, False
            DiscardedElapsedS = cPM_Work.ElapsedSeconds()
        Next K

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Record one elapsed value per measured cycle
        For K = 1 To Iterations
            cPM_Work.StartTimer ResolvedMethod, False
            ThisSample = cPM_Work.ElapsedSeconds()

            'Exclude a failed read rather than storing its zero, which would
            'otherwise become the minimum and be read as the resolution floor
                If cPM_Work.LastReadStatus <> cPM_ReadOK Then
                    LastFailure = cPM_Work.LastReadStatus
                Else
                    ValidCount = ValidCount + 1
                    Samples(ValidCount) = ThisSample
                End If
        Next K

    'An empty vector is not a measurement
        If ValidCount = 0 Then
            Err.Raise ERR_CPM_MEASURE_NO_VALID_SAMPLES, _
                      "cPerformanceManager.MeasureOverhead_Samples", _
                      "No measured cycle produced a value. " & _
                      Read_StatusToText(LastFailure) & _
                      " All " & CStr(Iterations) & " cycle(s) failed."
        End If

    'Trim to the cycles actually obtained
        If ValidCount < Iterations Then
            ReDim Preserve Samples(1 To ValidCount)
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the full per-cycle sample vector
        MeasureOverhead_Samples = Samples

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Release the isolated worker instance
        On Error Resume Next
        If Not cPM_Work Is Nothing Then cPM_Work.ResetEnvironment
        Set cPM_Work = Nothing
        On Error GoTo 0
        Exit Function

Err_Handler:
'------------------------------------------------------------------------------
' CAPTURE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Capture the original error before cleanup
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description

'------------------------------------------------------------------------------
' CLEAN UP AFTER FAILURE
'------------------------------------------------------------------------------
    'Cleanup must be best-effort so that we can preserve the original error
        On Error Resume Next
        If Not cPM_Work Is Nothing Then cPM_Work.ResetEnvironment
        Set cPM_Work = Nothing
        On Error GoTo 0

'------------------------------------------------------------------------------
' RE-RAISE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Re-raise the original error
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

Public Function Stats_Count( _
    ByRef Samples() As Double) _
    As Long
'
'==============================================================================
'                                 STATS COUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of elements in a sample vector, or 0 when uninitialized
'
' WHY THIS EXISTS
'   Calling UBound on an uninitialized dynamic array raises. Every statistics
'   routine needs a safe emptiness test, so that check lives here once
'
' INPUTS
'   Samples
'     Sample vector to inspect
'
' RETURNS
'   Long
'     Element count, or 0 when the array has never been dimensioned
'
' BEHAVIOR
'   - Attempts to read the array bounds
'   - Returns 0 when the bounds read fails
'
' ERROR POLICY
'   Does not raise errors
'
' NOTES
'   This is the only routine in the statistics set that tolerates an
'   uninitialized array. All others raise through Stats_RequireSamples
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Reading bounds of an uninitialized array raises, so probe defensively
        On Error Resume Next
        Stats_Count = UBound(Samples) - LBound(Samples) + 1
        If Err.Number <> 0 Then
            Err.Clear
            Stats_Count = 0
        End If
        On Error GoTo 0

End Function

Public Function Stats_Min( _
    ByRef Samples() As Double) _
    As Double
'
'==============================================================================
'                                  STATS MIN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the smallest value in a sample vector
'
' WHY THIS EXISTS
'   The minimum is the single most informative benchmark statistic. It is the
'   run least disturbed by scheduling, and therefore the closest estimate of the
'   true cost of the measured work
'
' INPUTS
'   Samples
'     Sample vector
'
' RETURNS
'   Double
'     Smallest observed value
'
' ERROR POLICY
'   Raises on an empty sample vector
'
' DEPENDENCIES
'   - Stats_RequireSamples
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long     'Loop index
    Dim Lo      As Double   'Running minimum

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_Min"

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Seed the running minimum from the first element
        Lo = Samples(LBound(Samples))

    'Retain the smallest observed value
        For i = LBound(Samples) To UBound(Samples)
            If Samples(i) < Lo Then Lo = Samples(i)
        Next i

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the smallest observed value
        Stats_Min = Lo

End Function

Public Function Stats_Max( _
    ByRef Samples() As Double) _
    As Double
'
'==============================================================================
'                                  STATS MAX
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the largest value in a sample vector
'
' WHY THIS EXISTS
'   The maximum shows the worst observed case and, compared against the median,
'   indicates how badly the run was disturbed
'
' INPUTS
'   Samples
'     Sample vector
'
' RETURNS
'   Double
'     Largest observed value
'
' ERROR POLICY
'   Raises on an empty sample vector
'
' DEPENDENCIES
'   - Stats_RequireSamples
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long     'Loop index
    Dim Hi      As Double   'Running maximum

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_Max"

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Seed the running maximum from the first element
        Hi = Samples(LBound(Samples))

    'Retain the largest observed value
        For i = LBound(Samples) To UBound(Samples)
            If Samples(i) > Hi Then Hi = Samples(i)
        Next i

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the largest observed value
        Stats_Max = Hi

End Function

Public Function Stats_Mean( _
    ByRef Samples() As Double) _
    As Double
'
'==============================================================================
'                                  STATS MEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the arithmetic mean of a sample vector
'
' WHY THIS EXISTS
'   The mean is familiar and is retained for continuity with
'   OverheadMeasurement_Seconds
'
' INPUTS
'   Samples
'     Sample vector
'
' RETURNS
'   Double
'     Arithmetic mean
'
' ERROR POLICY
'   Raises on an empty sample vector
'
' DEPENDENCIES
'   - Stats_RequireSamples
'
' NOTES
'   Prefer Stats_Median for reporting. A single scheduler preemption shifts the
'   mean materially while barely moving the median
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long     'Loop index
    Dim Total   As Double   'Accumulated sum

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_Mean"

'------------------------------------------------------------------------------
' ACCUMULATE
'------------------------------------------------------------------------------
    'Sum every sample
        For i = LBound(Samples) To UBound(Samples)
            Total = Total + Samples(i)
        Next i

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the arithmetic mean
        Stats_Mean = Total / CDbl(Stats_Count(Samples))

End Function

Public Function Stats_Median( _
    ByRef Samples() As Double) _
    As Double
'
'==============================================================================
'                                STATS MEDIAN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the median of a sample vector
'
' WHY THIS EXISTS
'   The median is the headline statistic for timing work. It is insensitive to
'   the occasional very slow run that scheduling produces, so it represents the
'   typical case far better than the mean
'
' INPUTS
'   Samples
'     Sample vector
'
' RETURNS
'   Double
'     Median value.
'     For an even element count, the mean of the two central values
'
' BEHAVIOR
'   - Sorts a private copy of the samples
'   - Returns the central value, or the mean of the two central values
'
' ERROR POLICY
'   Raises on an empty sample vector
'
' DEPENDENCIES
'   - Stats_RequireSamples
'   - Stats_SortedCopy
'
' NOTES
'   The caller's array is never reordered
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sorted()    As Double   'Sorted working copy
    Dim N           As Long     'Element count
    Dim Mid         As Long     'Central index

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_Median"

'------------------------------------------------------------------------------
' SORT A PRIVATE COPY
'------------------------------------------------------------------------------
    'Work on a 1-based sorted copy so the caller's array is untouched
        Sorted = Stats_SortedCopy(Samples)
        N = UBound(Sorted)

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Odd count returns the central value
        If (N Mod 2) = 1 Then
            Stats_Median = Sorted((N + 1) \ 2)
    'Even count returns the mean of the two central values
        Else
            Mid = N \ 2
            Stats_Median = (Sorted(Mid) + Sorted(Mid + 1)) / 2#
        End If

End Function

Public Function Stats_Percentile( _
    ByRef Samples() As Double, _
    ByVal Percent As Double) _
    As Double
'
'==============================================================================
'                              STATS PERCENTILE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the nearest-rank percentile of a sample vector
'
' WHY THIS EXISTS
'   A high percentile such as P95 shows the tail of the distribution. Comparing
'   P95 against the median is the quickest way to see whether a benchmark run
'   was disturbed
'
' INPUTS
'   Samples
'     Sample vector
'
'   Percent
'     Requested percentile in the range 0 to 100
'
' RETURNS
'   Double
'     Sample value at the requested nearest rank
'
' BEHAVIOR
'   - Sorts a private copy of the samples
'   - Computes the nearest rank as Ceiling(Percent / 100 * N)
'   - Clamps the rank into the valid index range
'
' ERROR POLICY
'   - Raises on an empty sample vector
'   - Raises when Percent falls outside 0 to 100
'
' DEPENDENCIES
'   - Stats_RequireSamples
'   - Stats_SortedCopy
'
' NOTES
'   Nearest-rank is used rather than interpolation, so every returned value is
'   an actually observed measurement rather than a synthetic one
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sorted()    As Double   'Sorted working copy
    Dim N           As Long     'Element count
    Dim Rank        As Long     'Nearest rank index

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_Percentile"

    'Reject an out-of-range percentile
        If (Percent < 0#) Or (Percent > 100#) Then
            Err.Raise ERR_CPM_STATS_BAD_PERCENTILE, _
                      "cPerformanceManager.Stats_Percentile", _
                      "Percent must be between 0 and 100."
        End If

'------------------------------------------------------------------------------
' SORT A PRIVATE COPY
'------------------------------------------------------------------------------
    'Work on a 1-based sorted copy so the caller's array is untouched
        Sorted = Stats_SortedCopy(Samples)
        N = UBound(Sorted)

'------------------------------------------------------------------------------
' RESOLVE NEAREST RANK
'------------------------------------------------------------------------------
    'Ceiling of Percent/100 * N, using -Int(-x) because Int floors toward -inf
        Rank = CLng(-Int(-((Percent / 100#) * CDbl(N))))

    'Clamp the rank into the valid index range
        If Rank < 1 Then Rank = 1
        If Rank > N Then Rank = N

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the sample at the resolved rank
        Stats_Percentile = Sorted(Rank)

End Function

Public Function Stats_StdDev( _
    ByRef Samples() As Double) _
    As Double
'
'==============================================================================
'                                STATS STDDEV
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the sample standard deviation of a sample vector
'
' WHY THIS EXISTS
'   Spread is what distinguishes a stable measurement from a noisy one. Two runs
'   can share a mean while one is trustworthy and the other is not
'
' INPUTS
'   Samples
'     Sample vector
'
' RETURNS
'   Double
'     Sample standard deviation.
'     Returns 0 when fewer than two samples are present
'
' BEHAVIOR
'   - Computes the mean
'   - Accumulates squared deviations
'   - Divides by N-1 and returns the square root
'
' ERROR POLICY
'   Raises on an empty sample vector
'
' DEPENDENCIES
'   - Stats_RequireSamples
'   - Stats_Mean
'
' NOTES
'   The N-1 divisor is used because these are samples of a process, not a
'   complete population
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i           As Long     'Loop index
    Dim N           As Long     'Element count
    Dim Avg         As Double   'Arithmetic mean
    Dim SumSq       As Double   'Accumulated squared deviations
    Dim Deviation   As Double   'Deviation of the current sample

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_StdDev"

    'A single sample has no spread
        N = Stats_Count(Samples)
        If N < 2 Then
            Stats_StdDev = 0#
            Exit Function
        End If

'------------------------------------------------------------------------------
' ACCUMULATE SQUARED DEVIATIONS
'------------------------------------------------------------------------------
    'Resolve the mean once
        Avg = Stats_Mean(Samples)

    'Sum the squared deviations from the mean
        For i = LBound(Samples) To UBound(Samples)
            Deviation = Samples(i) - Avg
            SumSq = SumSq + (Deviation * Deviation)
        Next i

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the sample standard deviation
        Stats_StdDev = Sqr(SumSq / CDbl(N - 1))

End Function

Public Function Stats_CoefficientOfVariation( _
    ByRef Samples() As Double) _
    As Double
'
'==============================================================================
'                     STATS COEFFICIENT OF VARIATION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the coefficient of variation: standard deviation divided by mean
'
' WHY THIS EXISTS
'   Standard deviation alone cannot be compared across benchmarks of different
'   magnitudes. Expressing spread relative to the mean gives a dimensionless
'   noise measure that is comparable between runs
'
' INPUTS
'   Samples
'     Sample vector
'
' RETURNS
'   Double
'     Coefficient of variation
'
' BEHAVIOR
'   Delegates the validity decision to Stats_TryCoefficientOfVariation and
'   raises when the quantity is undefined
'
' ERROR POLICY
'   - Raises on an empty sample vector
'   - Raises ERR_CPM_STATS_UNDEFINED_CV when the mean is not strictly positive
'
' BEHAVIOURAL CORRECTION IN 1.3.0
'   Earlier releases returned 0 when the mean was zero. For non-negative timing
'   data a zero mean means every sample was zero, which is what a run of failed
'   non-strict reads produces. Returning 0 then made Stats_IsContaminated report
'   that run as perfectly stable - certifying a failed measurement as clean.
'
'   The quantity is undefined here, so the routine now says so.
'
' DEPENDENCIES
'   - Stats_RequireSamples
'   - Stats_TryCoefficientOfVariation
'
' NOTES
'   As a rough guide: below 0.05 is a clean run, above
'   CPM_CV_CONTAMINATION_THRESHOLD the result should not be trusted without
'   re-running on a quieter machine
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cv      As Double   'Computed coefficient of variation

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_CoefficientOfVariation"

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
    'Reject a sample set for which the quantity is not defined
        If Not Stats_TryCoefficientOfVariation(Samples, Cv) Then
            Err.Raise ERR_CPM_STATS_UNDEFINED_CV, _
                      "cPerformanceManager.Stats_CoefficientOfVariation", _
                      "The coefficient of variation is undefined because the sample " & _
                      "mean is not strictly positive. For timing data this means every " & _
                      "observation was zero, which usually indicates failed reads " & _
                      "rather than a fast operation."
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return spread expressed relative to the mean
        Stats_CoefficientOfVariation = Cv

End Function

Public Function Stats_IsContaminated( _
    ByRef Samples() As Double, _
    Optional ByVal CvThreshold As Double = CPM_CV_CONTAMINATION_THRESHOLD) _
    As Boolean
'
'==============================================================================
'                            STATS IS CONTAMINATED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when a sample vector is too noisy to be trusted
'
' WHY THIS EXISTS
'   A benchmark result that looks precise can still be worthless if the machine
'   was busy during the run. Callers should be told this explicitly rather than
'   having to infer it from the numbers
'
' INPUTS
'   Samples
'     Sample vector
'
'   CvThreshold (optional)
'     Coefficient-of-variation threshold above which the run is reported as
'     contaminated. Must not be negative
'
' RETURNS
'   Boolean
'     TRUE  => spread exceeds the threshold, or validity cannot be established
'     FALSE => spread is within the threshold
'
' BEHAVIOR
'   - Rejects a negative threshold
'   - Returns TRUE when the coefficient of variation is undefined
'   - Otherwise compares relative spread against the threshold
'
' ERROR POLICY
'   - Raises on an empty sample vector
'   - Raises ERR_CPM_STATS_BAD_CV_THRESHOLD on a negative threshold
'   - Does NOT raise when the coefficient of variation is undefined
'
' FAIL-SAFE POLICY
'   This routine deliberately differs from Stats_CoefficientOfVariation. That
'   function answers a mathematical question and raises when the answer does not
'   exist. This one answers a judgement question - "should I trust this run?" -
'   and a sample set whose validity cannot be established is exactly the kind
'   that should not be trusted.
'
'   Reporting FALSE there would certify a failed measurement as clean, which is
'   the defect this correction removes.
'
' DEPENDENCIES
'   - Stats_RequireSamples
'   - Stats_TryCoefficientOfVariation
'
' NOTES
'   This is a heuristic, not a statistical test. It flags runs worth repeating;
'   it does not prove a result is wrong
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cv      As Double   'Computed coefficient of variation

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_IsContaminated"

    'Reject a negative threshold. A negative threshold would report every
    'sample set as contaminated, since the coefficient of variation is never
    'negative, which is a silently useless answer rather than an error
        If CvThreshold < 0# Then
            Err.Raise ERR_CPM_STATS_BAD_CV_THRESHOLD, _
                      "cPerformanceManager.Stats_IsContaminated", _
                      "CvThreshold cannot be negative."
        End If

'------------------------------------------------------------------------------
' FAIL SAFE WHEN VALIDITY CANNOT BE ESTABLISHED
'------------------------------------------------------------------------------
    'Treat an undefined coefficient of variation as suspect rather than clean
        If Not Stats_TryCoefficientOfVariation(Samples, Cv) Then
            Stats_IsContaminated = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Compare relative spread against the supplied threshold
        Stats_IsContaminated = (Cv > CvThreshold)

End Function

Public Function Stats_Text( _
    ByRef Samples() As Double, _
    Optional ByVal Caption As String = vbNullString) _
    As String
'
'==============================================================================
'                                 STATS TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a readable multiline summary of a sample vector
'
' WHY THIS EXISTS
'   This is the presentation companion to the numeric statistics routines, for
'   Immediate Window output and quick inspection
'
' INPUTS
'   Samples
'     Sample vector
'
'   Caption (optional)
'     Optional heading for the summary block
'
' RETURNS
'   String
'     Multiline statistics summary
'
' BEHAVIOR
'   - Reports N, median, min, P95, max, mean, standard deviation and CV
'   - Appends a warning line when the run looks contaminated
'
' ERROR POLICY
'   Raises on an empty sample vector
'
' DEPENDENCIES
'   - The Stats_* routines
'
' NOTES
'   Median and minimum are listed first deliberately. They are the statistics
'   that should drive conclusions; the mean is reported last for continuity
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S       As String   'Accumulated summary text
    Dim Cv      As Double   'Computed coefficient of variation

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty sample vector
        Stats_RequireSamples Samples, "Stats_Text"

'------------------------------------------------------------------------------
' BUILD HEADER
'------------------------------------------------------------------------------
    'Start with the optional caption
        S = "TIMING STATISTICS"
        If Len(Caption) > 0 Then
            S = S & " | " & Caption
        End If
        S = S & vbCrLf & "Samples           : " & CStr(Stats_Count(Samples))

'------------------------------------------------------------------------------
' BUILD PRIMARY STATISTICS
'------------------------------------------------------------------------------
    'Report the robust statistics first
        S = S & vbCrLf & "Median            : " & Format$(Stats_Median(Samples), "0.000000000") & " s"
        S = S & vbCrLf & "Min               : " & Format$(Stats_Min(Samples), "0.000000000") & " s"
        S = S & vbCrLf & "P95               : " & Format$(Stats_Percentile(Samples, 95#), "0.000000000") & " s"
        S = S & vbCrLf & "Max               : " & Format$(Stats_Max(Samples), "0.000000000") & " s"

'------------------------------------------------------------------------------
' BUILD SPREAD STATISTICS
'------------------------------------------------------------------------------
    'Report the mean and spread afterwards
        S = S & vbCrLf & "Mean              : " & Format$(Stats_Mean(Samples), "0.000000000") & " s"
        S = S & vbCrLf & "StdDev            : " & Format$(Stats_StdDev(Samples), "0.000000000") & " s"
    'Report the coefficient of variation, or say plainly that it is undefined.
    'Stats_Text is a diagnostic surface and must not raise on the very sample
    'set a caller is trying to diagnose.
        If Stats_TryCoefficientOfVariation(Samples, Cv) Then
            S = S & vbCrLf & "CoeffOfVariation  : " & Format$(Cv, "0.0000")
        Else
            S = S & vbCrLf & "CoeffOfVariation  : undefined (mean is not positive)"
        End If

'------------------------------------------------------------------------------
' APPEND CONTAMINATION WARNING
'------------------------------------------------------------------------------
    'Tell the caller explicitly when the run looks too noisy to trust, and
    'distinguish "noisy" from "not demonstrably valid at all"
        If Not Stats_TryCoefficientOfVariation(Samples, Cv) Then
            S = S & vbCrLf & "WARNING           : sample validity cannot be established; " & _
                             "inspect the observations."
        ElseIf Cv > CPM_CV_CONTAMINATION_THRESHOLD Then
            S = S & vbCrLf & "WARNING           : high variance; re-run on a quieter machine."
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the assembled summary
        Stats_Text = S

End Function

'
'------------------------------------------------------------------------------
'
'              PRIVATE: BACKEND READER / STATISTICS HELPERS
'
'------------------------------------------------------------------------------
'

Private Function Elapsed_ComputeSeconds( _
    ByVal ResolvedMethod As cPM_TimerMethod, _
    ByRef EndStampOut As Double, _
    ByRef QpcEndOut As Currency, _
    ByRef StatusOut As cPM_ReadStatus) _
    As Double
'
'==============================================================================
'                          ELAPSED COMPUTE SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Computes elapsed seconds for an already resolved backend, without writing to
'   any cached class state.
'
' WHY THIS EXISTS
'   Backend dispatch lives here and nowhere else, so adding or changing a backend
'   is a single edit. Keeping it free of side effects also lets Checkpoint
'   measure without clobbering the cached T2 and ET values.
'
' INPUTS
'   ResolvedMethod  - backend already validated by Method_ResolveElapsed
'   EndStampOut     - receives the raw end timestamp, on success only
'   QpcEndOut       - receives the raw QPC end tick, on success only
'   StatusOut       - receives the failure reason, or cPM_ReadOK
'
' RETURNS
'   Double
'     Elapsed seconds, or 0 when the read failed in non-strict mode.
'
' CRITICAL INVARIANT
'   A failed read NEVER enters timestamp arithmetic.
'
'   The previous implementation let a failed method-4 read return 0, which then
'   produced a negative delta and had the 32-bit rollover period added to it -
'   turning a failure into a large positive duration that passed validation.
'
'   Every read is now checked before any subtraction or rollover correction.
'
' ERROR POLICY
'   - NEVER raises for a read failure. A failed read is reported through
'     StatusOut so the caller can publish it before deciding whether to fail.
'   - Still raises for an invalid resolved method, which is a defensive branch
'     with no corresponding read status, and through Elapsed_Validate for a
'     backwards-moving clock, which is arithmetic validation rather than a read.
'   - Never applies rollover correction to a failed read.
'
' WHY THE READ POLICY LIVES IN THE CALLER
'   Strict mode previously raised from here, before StatusOut could be returned.
'   The public operation had already reset LastReadStatus to OK, so a caller who
'   trapped the strict error and then inspected the status was told the last read
'   succeeded.
'
'   Reporting the status and letting the caller decide keeps the two public
'   surfaces - Err and LastReadStatus - in agreement. StartTimer already works
'   this way through Start_TryCapture.
'
' NOTES
'   Method 5 deliberately skips Elapsed_Validate. QPC is monotonic and method 5
'   is the default hot path.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim QpcNow      As Currency     'Raw QPC tick for the current read
    Dim EndStamp    As Double       'Raw end timestamp for the current read
    Dim Ms          As Double       'Raw millisecond value for backend 4
    Dim Seconds     As Double       'Computed elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Assume success until a read reports otherwise
        StatusOut = cPM_ReadOK

'------------------------------------------------------------------------------
' DISPATCH - READ FIRST, ARITHMETIC ONLY AFTER A VALID READ
'------------------------------------------------------------------------------
    Select Case ResolvedMethod

        Case cPM_MethodTimer
                EndStamp = Timer
                Seconds = EndStamp - m_T1
                'Correct midnight rollover for Timer()
                    If Seconds < 0# Then Seconds = Seconds + 86400#
                Seconds = Elapsed_Validate(Seconds, StatusOut)

        Case cPM_MethodTickCount
                EndStamp = Get_TickCountSeconds()
                Seconds = EndStamp - m_T1
                'Correct 32-bit millisecond-counter rollover where applicable
                    If Seconds < 0# Then Seconds = Seconds + RolloverSeconds(cPM_MethodTickCount)
                Seconds = Elapsed_Validate(Seconds, StatusOut)

        Case cPM_MethodTimeGetTime
                EnsureTimerResolution1ms
                EndStamp = UInt32ToDouble(timeGetTime()) / 1000#
                Seconds = EndStamp - m_T1
                'Correct 32-bit millisecond-counter rollover
                    If Seconds < 0# Then Seconds = Seconds + RolloverSeconds(cPM_MethodTimeGetTime)
                Seconds = Elapsed_Validate(Seconds, StatusOut)

        Case cPM_MethodSystemTime
                'Validate the read BEFORE any arithmetic touches it.
                'The status is returned rather than raised, so the caller can
                'publish it before deciding whether to fail.
                If Not SystemTime_TryReadMs(Ms, StatusOut) Then
                    Elapsed_ComputeSeconds = 0#
                    Exit Function
                End If

                EndStamp = Ms / 1000#
                Seconds = EndStamp - m_T1
                'Correct 32-bit millisecond-counter rollover
                    If Seconds < 0# Then Seconds = Seconds + RolloverSeconds(cPM_MethodSystemTime)
                Seconds = Elapsed_Validate(Seconds, StatusOut)

        Case cPM_MethodQPC
                'Validate the read BEFORE any arithmetic touches it.
                'The status is returned rather than raised, so the caller can
                'publish it before deciding whether to fail.
                If Not QPC_TryReadTick(QpcNow) Then
                    StatusOut = cPM_ReadQpcFailed
                    Elapsed_ComputeSeconds = 0#
                    Exit Function
                End If

                'Currency scaling cancels between numerator and denominator
                Seconds = CDbl((QpcNow - m_qpcStart) / m_qpcFreq)
                EndStamp = CDbl(QpcNow)
                QpcEndOut = QpcNow

        Case cPM_MethodNow
                EndStamp = Now() * 86400#
                Seconds = EndStamp - m_T1
                Seconds = Elapsed_Validate(Seconds, StatusOut)

        Case Else
                If m_StrictMode Then
                    Err.Raise ERR_CPM_RESOLVED_METHOD_INVALID, _
                              "cPerformanceManager.Elapsed_ComputeSeconds", _
                              "Resolved timer method is invalid."
                End If
                Elapsed_ComputeSeconds = 0#
                Exit Function

    End Select

'------------------------------------------------------------------------------
' PUBLISH ENDPOINT AND RESULT
'------------------------------------------------------------------------------
    'A clamped negative elapsed is not a measurement either. Report it the way a
    'failed read is reported, so every caller already handles it
        If StatusOut <> cPM_ReadOK Then
            Elapsed_ComputeSeconds = 0#
            Exit Function
        End If

    'Only a validated read reaches this point
        EndStampOut = EndStamp
        Elapsed_ComputeSeconds = Seconds

End Function


Private Sub Stats_RequireSamples( _
    ByRef Samples() As Double, _
    ByVal SourceProc As String)
'
'==============================================================================
'                           STATS REQUIRE SAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Rejects a sample vector that lies outside the supported domain
'
' WHY THIS EXISTS
'   Every statistics routine needs the same guard, so it lives here once.
'
' SUPPORTED DOMAIN
'   Finite, non-negative timing observations.
'
'   The statistics exist to describe elapsed times, which cannot be negative.
'   A negative value means the vector is not a timing vector, and every figure
'   computed from it would then describe something the caller did not intend.
'
'   The objection that Stats_Min works perfectly well on negative numbers is
'   true and beside the point: returning a confident answer over data that
'   should not have reached the routine is the failure this rejects.
'
' INPUTS
'   Samples
'     Sample vector to validate
'
'   SourceProc
'     Name of the calling procedure, used in the raised error source
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises ERR_CPM_STATS_NO_SAMPLES on an empty or uninitialized vector
'   - Raises ERR_CPM_STATS_OUT_OF_DOMAIN on a negative or non-finite value,
'     naming the offending index and value
'
' DEPENDENCIES
'   - Stats_Count
'
' NOTES
'   The scan is O(n) and runs once per public statistics call, so Stats_Text
'   scans several times over. For a diagnostic surface over sample vectors that
'   number in the thousands at most, that cost is not worth optimizing away with
'   a validity cache that could go stale.
'
'   The non-finite test is a proxy rather than a rigorous IEEE classification.
'   VBA raises on arithmetic overflow instead of producing infinities, and offers
'   no IsFinite, so NaN and Inf can only arrive through deliberate construction
'   or an API round-trip. The check costs nothing inside a scan that is already
'   happening, and is worth having for that reason alone.
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long     'Sample index
    Dim v       As Double   'Sample under inspection

'------------------------------------------------------------------------------
' VALIDATE SIZE
'------------------------------------------------------------------------------
    'Reject an empty or uninitialized sample vector
        If Stats_Count(Samples) < 1 Then
            Err.Raise ERR_CPM_STATS_NO_SAMPLES, _
                      "cPerformanceManager." & SourceProc, _
                      "Statistics require at least one sample."
        End If

'------------------------------------------------------------------------------
' VALIDATE DOMAIN
'------------------------------------------------------------------------------
    'Reject anything that is not a finite, non-negative timing observation
        For i = LBound(Samples) To UBound(Samples)
            v = Samples(i)

            'NaN is the only value that is not equal to itself
                If v <> v Then
                    Err.Raise ERR_CPM_STATS_OUT_OF_DOMAIN, _
                              "cPerformanceManager." & SourceProc, _
                              "Sample " & CStr(i) & " is not a number. Statistics " & _
                              "accept finite, non-negative timing observations."
                End If

            'A negative elapsed time is not a timing observation
                If v < 0# Then
                    Err.Raise ERR_CPM_STATS_OUT_OF_DOMAIN, _
                              "cPerformanceManager." & SourceProc, _
                              "Sample " & CStr(i) & " is negative (" & CStr(v) & "). " & _
                              "Statistics accept finite, non-negative timing observations."
                End If

            'Proxy for infinity: no real elapsed time approaches this magnitude
                If v > 1E+300 Then
                    Err.Raise ERR_CPM_STATS_OUT_OF_DOMAIN, _
                              "cPerformanceManager." & SourceProc, _
                              "Sample " & CStr(i) & " is not finite. Statistics " & _
                              "accept finite, non-negative timing observations."
                End If
        Next i

End Sub

Private Function Stats_TryCoefficientOfVariation( _
    ByRef Samples() As Double, _
    ByRef CvOut As Double) _
    As Boolean
'
'==============================================================================
'                  STATS TRY COEFFICIENT OF VARIATION
'------------------------------------------------------------------------------
' PURPOSE
'   Computes the coefficient of variation and reports whether it is defined
'
' WHY THIS EXISTS
'   Three public members need to know whether the coefficient of variation
'   exists, and each responds differently: Stats_CoefficientOfVariation raises,
'   Stats_IsContaminated fails safe, and Stats_Text reports it diagnostically.
'
'   Duplicating the validity decision across all three would let those policies
'   drift apart. This helper is the single definition; the callers hold only the
'   policy
'
' INPUTS
'   Samples
'     Sample vector
'
'   CvOut
'     Receives the coefficient of variation. Written ONLY when it is defined,
'     so a caller cannot consume a value that does not exist
'
' RETURNS
'   Boolean
'     TRUE  => the mean is strictly positive and CvOut holds StdDev / Mean
'     FALSE => the mean is not strictly positive and CvOut is untouched
'
' BEHAVIOR
'   - Computes the arithmetic mean
'   - Reports FALSE when that mean is zero or negative
'   - Otherwise divides the sample standard deviation by the mean
'
' ERROR POLICY
'   Does not raise. The caller decides what an undefined value means
'
' NOTES
'   For non-negative timing data a mean of zero implies every observation was
'   zero, which in practice means the reads failed rather than that the work was
'   instantaneous. A negative mean cannot arise from valid timing data at all;
'   it is handled here for completeness rather than because it is expected
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Avg     As Double   'Arithmetic mean

'------------------------------------------------------------------------------
' VALIDATE THE DENOMINATOR
'------------------------------------------------------------------------------
    'The quantity is only defined for a strictly positive mean.
    '
    'Only the zero case is reachable through the public API: Stats_RequireSamples
    'rejects negative samples, so a negative mean cannot arrive. The <= test is
    'retained as a defensive bound rather than narrowed to = 0
        Avg = Stats_Mean(Samples)

        If Avg <= 0# Then
            Stats_TryCoefficientOfVariation = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Publish the value only now that it is known to be defined
        CvOut = Stats_StdDev(Samples) / Avg
        Stats_TryCoefficientOfVariation = True

End Function

Private Function Stats_SortedCopy( _
    ByRef Samples() As Double) _
    As Double()
'
'==============================================================================
'                              STATS SORTED COPY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an ascending-sorted 1-based copy of a sample vector
'
' WHY THIS EXISTS
'   Median and percentile both need sorted data, and neither should reorder the
'   caller's array. Copying once here keeps that contract in one place
'
' INPUTS
'   Samples
'     Sample vector to copy and sort
'
' RETURNS
'   Double()
'     Ascending-sorted 1-based copy
'
' BEHAVIOR
'   - Copies the samples into a fresh 1-based array
'   - Sorts the copy in place
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Stats_QuickSort
'
' NOTES
'   The returned array is always 1-based regardless of the input bounds, so
'   callers can index it directly
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Copy()  As Double   'Working copy
    Dim i       As Long     'Source index
    Dim N       As Long     'Element count

'------------------------------------------------------------------------------
' COPY
'------------------------------------------------------------------------------
    'Rebase the samples into a fresh 1-based array
        N = UBound(Samples) - LBound(Samples) + 1
        ReDim Copy(1 To N)

        For i = 1 To N
            Copy(i) = Samples(LBound(Samples) + i - 1)
        Next i

'------------------------------------------------------------------------------
' SORT
'------------------------------------------------------------------------------
    'Sort the copy ascending
        If N > 1 Then
            Stats_QuickSort Copy, 1, N
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the sorted copy
        Stats_SortedCopy = Copy

End Function

Private Sub Stats_QuickSort( _
    ByRef A() As Double, _
    ByVal FirstIdx As Long, _
    ByVal LastIdx As Long)
'
'==============================================================================
'                              STATS QUICKSORT
'------------------------------------------------------------------------------
' PURPOSE
'   Sorts a Double array ascending, in place, between the supplied bounds
'
' WHY THIS EXISTS
'   Median and percentile need sorted data. A quicksort keeps the cost near
'   N log N, which matters because overhead sampling routinely produces
'   thousands of values
'
' INPUTS
'   A
'     Array to sort in place
'
'   FirstIdx
'     Lower bound of the partition
'
'   LastIdx
'     Upper bound of the partition
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Partitions around the middle element
'   - Recurses into each partition that still holds more than one element
'
' ERROR POLICY
'   Raises errors normally
'
' NOTES
'   Recursion depth stays near log2(N), so even a million samples reaches only
'   about twenty frames
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long     'Ascending scan index
    Dim j       As Long     'Descending scan index
    Dim Pivot   As Double   'Partition pivot value
    Dim Swap    As Double   'Temporary used during exchange

'------------------------------------------------------------------------------
' PARTITION
'------------------------------------------------------------------------------
    'Seed the scan indices and pivot
        i = FirstIdx
        j = LastIdx
        Pivot = A((FirstIdx + LastIdx) \ 2)

    'Exchange out-of-place pairs until the scans cross
        Do While i <= j
            Do While A(i) < Pivot
                i = i + 1
            Loop
            Do While A(j) > Pivot
                j = j - 1
            Loop
            If i <= j Then
                Swap = A(i)
                A(i) = A(j)
                A(j) = Swap
                i = i + 1
                j = j - 1
            End If
        Loop

'------------------------------------------------------------------------------
' RECURSE
'------------------------------------------------------------------------------
    'Sort the lower partition
        If FirstIdx < j Then
            Stats_QuickSort A, FirstIdx, j
        End If
    'Sort the upper partition
        If i < LastIdx Then
            Stats_QuickSort A, i, LastIdx
        End If

End Sub

Public Property Get LastReadStatus() As cPM_ReadStatus
'
'==============================================================================
'                              LAST READ STATUS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the outcome of the most recent native timing read.
'
' WHY THIS EXISTS
'   In non-strict mode a failed read returns 0 rather than raising. Without a
'   status, a caller cannot distinguish "zero seconds elapsed" from "no reading
'   was obtained" - and the second silently poisons any aggregate built from it.
'
' INPUTS
'   None.
'
' RETURNS
'   cPM_ReadStatus
'     cPM_ReadOK                      - the last read succeeded
'     cPM_ReadQpcFailed               - QueryPerformanceCounter failed
'     cPM_ReadSystemTimeFailed        - timeGetSystemTime failed
'     cPM_ReadSystemTimeFormatInvalid - timeGetSystemTime returned a bad format
'     cPM_ReadFallbackToMethod2       - the requested backend was unreadable and
'                                       method 2 was bound instead
'
' BEHAVIOR
'   Reset to cPM_ReadOK at the start of StartTimer, ElapsedSeconds and
'   Checkpoint, so it always describes the most recent relevant operation rather
'   than an older failure.
'
' ERROR POLICY
'   Does not raise.
'
' NOTES
'   cPM_ReadFallbackToMethod2 is not a failure. The session is valid and the
'   measurement is usable - the status records that the requested backend was
'   coerced, so the coercion is visible rather than silent. ActiveMethodID
'   reports the backend that was actually bound.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the outcome of the most recent native read
        LastReadStatus = m_LastReadStatus

End Property


Private Function Read_StatusToErrorNumber( _
    ByVal Status As cPM_ReadStatus) _
    As Long
'
'==============================================================================
'                        READ STATUS TO ERROR NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Maps a read status onto the error number raised for it in strict mode.
'
' WHY THIS EXISTS
'   Keeps the status-to-error mapping in one place, so a status can never be
'   reported under two different error numbers from two different call sites.
'
' INPUTS
'   Status - the read status to map
'
' RETURNS
'   Long - the corresponding error number
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' MAP
'------------------------------------------------------------------------------
    'Resolve the error number for this status
        Select Case Status
            Case cPM_ReadQpcFailed
                Read_StatusToErrorNumber = ERR_CPM_QPC_READ_FAILED
            Case cPM_ReadSystemTimeFailed
                Read_StatusToErrorNumber = ERR_CPM_TIMEGETSYSTEMTIME_FAILED
            Case cPM_ReadSystemTimeFormatInvalid
                Read_StatusToErrorNumber = ERR_CPM_TIMEGETSYSTEMTIME_FORMAT
            Case cPM_ReadElapsedInvalid
                Read_StatusToErrorNumber = ERR_CPM_NEGATIVE_ELAPSED
            Case Else
                Read_StatusToErrorNumber = ERR_CPM_START_CAPTURE_FAILED
        End Select

End Function


Private Function Read_StatusToText( _
    ByVal Status As cPM_ReadStatus) _
    As String
'
'==============================================================================
'                            READ STATUS TO TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a human-readable description of a read status, for error messages.
'
' INPUTS
'   Status - the read status to describe
'
' RETURNS
'   String - a readable description
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' MAP
'------------------------------------------------------------------------------
    'Resolve the description for this status
        Select Case Status
            Case cPM_ReadOK
                Read_StatusToText = "The timing read succeeded."
            Case cPM_ReadQpcFailed
                Read_StatusToText = "QueryPerformanceCounter failed."
            Case cPM_ReadSystemTimeFailed
                Read_StatusToText = "timeGetSystemTime failed."
            Case cPM_ReadSystemTimeFormatInvalid
                Read_StatusToText = "timeGetSystemTime did not return TIME_MS format."
            Case cPM_ReadFallbackToMethod2
                Read_StatusToText = "The requested backend was unreadable; method 2 was bound."
            Case cPM_ReadElapsedInvalid
                Read_StatusToText = "Elapsed time was negative after rollover correction; " & _
                                    "the timing source moved backwards."
            Case Else
                Read_StatusToText = "The timing read failed."
        End Select

End Function


Private Sub Measure_ArmWorkerStartFault( _
    ByRef WorkerInstance As cPerformanceManager, _
    ByVal ExpectedMethod As cPM_TimerMethod)
'
'==============================================================================
'                       MEASURE ARM WORKER START FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Arms one forced native read failure on the worker's next start capture, if
'   any remain armed
'
' WHY THIS EXISTS
'   Rejecting a fallback cycle can only be tested if a fallback can be produced
'   on demand, and a backend does not fail because a test would like it to
'
'   The start and endpoint seams are separate because they arm at different
'   moments and prove different things. A start failure produces a fallback the
'   harness must reject; an endpoint failure produces a read the harness must
'   exclude. One counter could not express both
'
' INPUTS
'   WorkerInstance
'     The isolated worker whose next start capture should fail
'   ExpectedMethod
'     The backend the cycle intends to bind, which selects the fault
'
' RETURNS
'   None
'
' BEHAVIOR
'   Consumes one armed start failure and forwards it to the matching low-level
'   seam. Method 5 arms a QPC read failure, method 4 a system-time read failure.
'   Every other backend is a no-op, because backends 1, 2, 3 and 6 have no start
'   failure mode and so cannot fall back
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   The low-level seams are one-shot, so the forced failure is consumed by the
'   requested backend's capture and the method-2 fallback capture then succeeds.
'   That is exactly the state this seam exists to produce
'
' UPDATED
'   2026-08-30
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Nothing armed, nothing to do
        If m_TestWorkerStartFailures <= 0 Then Exit Sub

    'Consume one, whether or not the backend can act on it, so an armed count
    'against a backend with no start failure mode still drains
        m_TestWorkerStartFailures = m_TestWorkerStartFailures - 1

    'Select the fault the requested backend can actually suffer
        Select Case ExpectedMethod
            Case cPM_MethodQPC
                WorkerInstance.Test_ForceNextQPCReadFailure
            Case cPM_MethodSystemTime
                WorkerInstance.Test_ForceNextSystemTimeReadFailure
        End Select

End Sub


Private Sub Measure_ArmWorkerEndpointFault( _
    ByRef WorkerInstance As cPerformanceManager, _
    ByVal ExpectedMethod As cPM_TimerMethod)
'
'==============================================================================
'                      MEASURE ARM WORKER ENDPOINT FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Arms one forced native read failure on the worker's next endpoint read, if
'   any remain armed
'
' WHY THIS EXISTS
'   The counterpart to Measure_ArmWorkerStartFault. Kept separate so a test can
'   target start validation or endpoint handling precisely, rather than arming
'   something that lands wherever the loop happens to consume it
'
' INPUTS
'   WorkerInstance
'     The isolated worker whose next endpoint read should fail
'   ExpectedMethod
'     The backend bound by this cycle, which selects the fault
'
' RETURNS
'   None
'
' BEHAVIOR
'   Consumes one armed endpoint failure and forwards it to the matching
'   low-level seam. Method 5 arms a QPC read failure, method 4 a system-time
'   read failure. Every other backend is a no-op
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   Called after Application.Run returns, so the injected fault cannot land on
'   the workload or on a warm-up cycle
'
' UPDATED
'   2026-08-30
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Nothing armed, nothing to do
        If m_TestWorkerEndpointFailures <= 0 Then Exit Sub

    'Consume one, whether or not the backend can act on it
        m_TestWorkerEndpointFailures = m_TestWorkerEndpointFailures - 1

    'Select the fault the bound backend can actually suffer
        Select Case ExpectedMethod
            Case cPM_MethodQPC
                WorkerInstance.Test_ForceNextQPCReadFailure
            Case cPM_MethodSystemTime
                WorkerInstance.Test_ForceNextSystemTimeReadFailure
        End Select

End Sub


Private Function Measure_NoValidSamplesText( _
    ByVal Iterations As Long, _
    ByVal RejectedSamples As Long, _
    ByVal ExpectedMethod As cPM_TimerMethod, _
    ByVal LastFailure As cPM_ReadStatus) _
    As String
'
'==============================================================================
'                       MEASURE NO VALID SAMPLES TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the description for ERR_CPM_MEASURE_NO_VALID_SAMPLES, naming the
'   reason the vector is empty
'
' WHY THIS EXISTS
'   Two very different conditions produce an empty vector, and one message
'   cannot describe both honestly. If every cycle fell back, the reads did not
'   fail at all: the host could not start the requested backend. Telling that
'   caller "no measured read produced a value" sends them to look at a clock
'   that is working perfectly
'
' INPUTS
'   Iterations        - measured iterations requested
'   RejectedSamples   - measured cycles rejected for backend fallback
'   ExpectedMethod    - backend the caller asked for
'   LastFailure       - most recent failure or rejection status
'
' RETURNS
'   String
'     A description naming the dominant cause
'
' BEHAVIOR
'   - All cycles rejected: names the fallback and the backend that was requested
'   - No cycle rejected: reports failed reads, as before
'   - Mixed: reports both counts, then the most recent status
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   The error number is unchanged. At the aggregate level it remains accurate:
'   there is no valid requested-backend sample to return. Only the explanation
'   is status-aware
'
' UPDATED
'   2026-08-30
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Every cycle fell back: the reads worked, the backend did not
        If RejectedSamples = Iterations Then
            Measure_NoValidSamplesText = _
                "No requested-backend sample was retained. All " & _
                CStr(Iterations) & " measured iteration(s) fell back from " & _
                MethodName(ExpectedMethod) & " to " & _
                MethodName(cPM_MethodTickCount) & _
                "; mixed-backend observations were rejected."
            Exit Function
        End If

    'Nothing was rejected: this is the pre-v1.4.0 condition
        If RejectedSamples = 0 Then
            Measure_NoValidSamplesText = _
                "No measured read produced a value. " & _
                Read_StatusToText(LastFailure) & _
                " All " & CStr(Iterations) & " iteration(s) failed."
            Exit Function
        End If

    'Both causes contributed, so report both rather than picking one
        Measure_NoValidSamplesText = _
            "No measured sample was retained across " & CStr(Iterations) & _
            " iteration(s): " & CStr(RejectedSamples) & _
            " rejected after falling back from " & MethodName(ExpectedMethod) & _
            ", " & CStr(Iterations - RejectedSamples) & " failed to read. " & _
            Read_StatusToText(LastFailure)

End Function


Private Function Measure_BuildQualifiedTarget( _
    ByVal WorkbookName As String, _
    ByVal ProcedureName As String) _
    As String
'
'==============================================================================
'                      MEASURE BUILD QUALIFIED TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the string Application.Run should be given, from an explicit
'   workbook name
'
' WHY THIS EXISTS
'   Measure_QualifyProcedureName reads ThisWorkbook.Name, which a test cannot
'   substitute. The host workbook has whatever name it has, so the apostrophe
'   case could only ever be exercised by renaming the project or by trusting a
'   comment
'
'   Separating the decision from the source of the workbook name makes the rule
'   itself testable with any name, while production still supplies the real one
'
' INPUTS
'   WorkbookName
'     Name of the workbook to bind an unqualified procedure to
'   ProcedureName
'     Name as supplied by the caller, qualified or not
'
' RETURNS
'   String
'     The name Application.Run should be given
'
' BEHAVIOR
'   - Leading and trailing spaces are removed from ProcedureName first, so
'     both branches below see the same normalized value
'   - A normalized name already containing "!" is returned as-is, so an explicit
'     qualification always wins
'   - Otherwise the name is bound to WorkbookName, single-quoted so that a name
'     containing spaces still resolves, with embedded apostrophes doubled so
'     that a name such as O'Brien.xlsm resolves too
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   The quoting matters. "PERFORMANCE MANAGER.xlsm!Proc" fails to resolve;
'   "'PERFORMANCE MANAGER.xlsm'!Proc" succeeds. So does the escaping:
'   "'O'Brien.xlsm'!Proc" terminates the quoted section at the apostrophe and
'   resolves to nothing useful, while "'O''Brien.xlsm'!Proc" is correct
'
'   Only leading and trailing spaces are removed. Trim$ removes space
'   characters only, so a tab, carriage return or line feed at either end
'   survives and reaches Application.Run
'
'   That boundary is deliberate. MeasureProcedure rejects a blank name with
'   Len(Trim$(ProcedureName)) = 0, and using the same function in both places is
'   what stops validation and qualification drifting apart. A name that passes
'   the blank check is the same name this function normalizes
'
'   Interior spaces are preserved. A caller who supplied them meant something
'   this function should not silently rewrite
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NormalizedName  As String   'Procedure name with outer spaces removed
    Dim SafeWorkbook    As String   'Workbook name with apostrophes escaped

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Normalize once, so both branches below agree on what the caller asked for
        NormalizedName = Trim$(ProcedureName)

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'An explicit qualification always wins; the caller owns its own quoting
        If InStr(1, NormalizedName, "!") > 0 Then
            Measure_BuildQualifiedTarget = NormalizedName
            Exit Function
        End If

    'Double any embedded apostrophe so the quoted section is not terminated early
        SafeWorkbook = Replace(WorkbookName, "'", "''")

    'Bind to the supplied workbook, quoted for names containing spaces
        Measure_BuildQualifiedTarget = "'" & SafeWorkbook & "'!" & NormalizedName

End Function


Private Function Measure_QualifyProcedureName( _
    ByVal ProcedureName As String) _
    As String
'
'==============================================================================
'                      MEASURE QUALIFY PROCEDURE NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a workbook-qualified name for Application.Run
'
' WHY THIS EXISTS
'   Application.Run resolves an unqualified name against the active workbook
'   first. When several open workbooks define a procedure of the same name, the
'   harness could silently measure the wrong one, and which one would depend on
'   what the user happened to have in front
'
'   Qualifying with the workbook that hosts this class makes resolution
'   deterministic
'
' INPUTS
'   ProcedureName
'     Name as supplied by the caller, qualified or not
'
' RETURNS
'   String
'     The name Application.Run should be given
'
' BEHAVIOR
'   Supplies ThisWorkbook.Name to Measure_BuildQualifiedTarget, which holds the
'   qualification rule. See that function for the spacing, quoting and escaping
'   contract
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   To measure a procedure in a different workbook, pass the qualified form
'   yourself
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Apply the qualification rule to the workbook that hosts this class
        Measure_QualifyProcedureName = _
            Measure_BuildQualifiedTarget(ThisWorkbook.Name, ProcedureName)

End Function

Friend Function Test_QualifyProcedureNameFor( _
    ByVal WorkbookName As String, _
    ByVal ProcedureName As String) _
    As String
'
'==============================================================================
'                     TEST QUALIFY PROCEDURE NAME FOR
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the qualification rule to a caller-supplied workbook name
'
' WHY THIS EXISTS
'   The rule depends on the host workbook name, and a test cannot rename the
'   host. Without this seam the apostrophe and space cases could only be
'   asserted about a name the suite never actually produces
'
'   This exposes the decision, not a way to change it. Production dispatch still
'   reads ThisWorkbook.Name and is unaffected by anything a test passes here
'
' INPUTS
'   WorkbookName
'     Workbook name to bind an unqualified procedure to
'   ProcedureName
'     Name as supplied by the caller, qualified or not
'
' RETURNS
'   String
'     The name Application.Run would be given for that workbook
'
' BEHAVIOR
'   Returns Measure_BuildQualifiedTarget unchanged. Holds no state, arms
'   nothing, and clears nothing
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   String construction only. That the constructed target actually dispatches
'   from a workbook whose name contains an apostrophe is verified by hand and
'   recorded in the release evidence, because it needs a second workbook on disk
'
' UPDATED
'   2026-08-27
'==============================================================================

    'Expose the rule for inspection with an arbitrary workbook name
        Test_QualifyProcedureNameFor = _
            Measure_BuildQualifiedTarget(WorkbookName, ProcedureName)

End Function


Friend Sub Test_ForceNextNegativeElapsed()
'
'==============================================================================
'                      TEST FORCE NEXT NEGATIVE ELAPSED
'------------------------------------------------------------------------------
' PURPOSE
'   Arms a ONE-SHOT forced backward-clock condition on the next validated
'   elapsed computation
'
' WHY THIS EXISTS
'   A timing source that moves backwards cannot be arranged on demand. Without a
'   seam, the clamp and its status could only be tested by calling
'   Elapsed_Validate directly, leaving the end-to-end consequences permanently
'   unverified: what ElapsedSeconds reports, whether Checkpoint records a row,
'   and whether the harness keeps the sample.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   The next call to Elapsed_Validate treats its input as negative, and clears
'   the flag.
'
' ERROR POLICY
'   Does not raise.
'
' NOTES
'   Method 5 does not call Elapsed_Validate, so this seam has no effect on a QPC
'   session. Use methods 1, 2, 3, 4 or 6 to exercise it.
'
'   One-shot and self-clearing, for the same reasons as the other seams.
'
' UPDATED
'   2026-08-16
'==============================================================================

    'Arm a single forced backward-clock condition
        m_TestForceNextNegativeElapsed = True

End Sub

Friend Sub Test_ForceWorkerReadFailures( _
    ByVal FailureCount As Long)
'
'==============================================================================
'                     TEST FORCE WORKER READ FAILURES
'------------------------------------------------------------------------------
' PURPOSE
'   Arms a number of forced read failures inside the isolated worker instance
'   that the measurement harness creates
'
' WHY THIS EXISTS
'   The other fault-injection seams arm a flag on the instance they are called
'   on. The harness measures on a private worker, so those seams cannot reach
'   the reads that actually produce the sample vector
'
'   Without this, the failure reporting added for the harness could only ever be
'   tested on a healthy run, which proves nothing
'
' INPUTS
'   FailureCount
'     Number of measured reads to fail. Negative values are treated as zero
'
' RETURNS
'   None
'
' BEHAVIOR
'   The next FailureCount measured reads inside MeasureProcedure or
'   MeasureOverhead_Samples return a failure. The counter is consumed as the
'   reads occur and is cleared when the harness finishes
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   Warm-up iterations are not affected, so a forced failure lands on a measured
'   sample rather than a discarded one
'
' UPDATED
'   2026-08-16
'==============================================================================

    'Arm the requested number of forced worker endpoint read failures
        If FailureCount < 0 Then FailureCount = 0
        m_TestWorkerEndpointFailures = FailureCount

End Sub


Friend Sub Test_ForceWorkerStartReadFailures( _
    ByVal FailureCount As Long)
'
'==============================================================================
'                    TEST FORCE WORKER START READ FAILURES
'------------------------------------------------------------------------------
' PURPOSE
'   Arms forced native read failures on the START capture of the next measured
'   cycles run by MeasureProcedure
'
' WHY THIS EXISTS
'   A cycle whose requested backend fails to start falls back to method 2, and
'   v1.4.0 rejects the resulting observation rather than mixing it into a vector
'   the caller believes is homogeneous. That rejection cannot be tested unless a
'   start failure can be produced deliberately
'
'   This is separate from the endpoint seam because the two arm at different
'   moments and prove different things. Arming a single ambiguous counter, as
'   v1.3.0 did, could only ever reach the endpoint read
'
' INPUTS
'   FailureCount
'     Number of measured start captures to fail. Negative values are treated
'     as zero
'
' RETURNS
'   None
'
' BEHAVIOR
'   Each armed count is consumed by one measured cycle, before StartTimer. The
'   fault matches the requested backend: method 5 fails a QPC read, method 4 a
'   system-time read. Backends 1, 2, 3 and 6 have no start failure mode, so an
'   armed count against them drains without effect
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   Warm-up cycles are not affected, so a forced failure lands on a measured
'   cycle rather than a discarded one
'
'   The counter is cleared on every exit from MeasureProcedure, normal or
'   raised, so an armed failure cannot leak into a later call
'
' UPDATED
'   2026-08-30
'==============================================================================

    'Arm the requested number of forced worker start read failures
        If FailureCount < 0 Then FailureCount = 0
        m_TestWorkerStartFailures = FailureCount

End Sub


Friend Sub Test_ForceWorkerEndpointReadFailures( _
    ByVal FailureCount As Long)
'
'==============================================================================
'                  TEST FORCE WORKER ENDPOINT READ FAILURES
'------------------------------------------------------------------------------
' PURPOSE
'   Arms forced native read failures on the ENDPOINT read of the next measured
'   cycles run by MeasureProcedure
'
' WHY THIS EXISTS
'   The counterpart to Test_ForceWorkerStartReadFailures. Named explicitly so a
'   test states which half of the cycle it is exercising rather than relying on
'   where an ambiguous counter happened to be consumed
'
' INPUTS
'   FailureCount
'     Number of measured endpoint reads to fail. Negative values are treated
'     as zero
'
' RETURNS
'   None
'
' BEHAVIOR
'   Each armed count is consumed by one measured cycle, after Application.Run
'   returns and before ElapsedSeconds. The fault matches the bound backend:
'   method 5 fails a QPC read, method 4 a system-time read
'
' ERROR POLICY
'   Does not raise
'
' NOTES
'   Test_ForceWorkerReadFailures is retained as a compatibility alias for this
'   seam, since the endpoint read is the only thing it could ever reach
'
'   Cleared on every exit from MeasureProcedure, normal or raised
'
' UPDATED
'   2026-08-30
'==============================================================================

    'Arm the requested number of forced worker endpoint read failures
        If FailureCount < 0 Then FailureCount = 0
        m_TestWorkerEndpointFailures = FailureCount

End Sub

Friend Sub Test_ForceNextQPCReadFailure()
'
'==============================================================================
'                       TEST FORCE NEXT QPC READ FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Arms a ONE-SHOT forced failure of the next QueryPerformanceCounter read.
'
' WHY THIS EXISTS
'   The failure paths this class guards cannot be reached by ordinary calling
'   patterns - QPC does not fail on demand. Without a seam, the most important
'   correctness behaviour in the class would be permanently untested.
'
'   The seam sits at the raw-read abstraction, below all public timing
'   semantics, so every higher-level behaviour can be proven naturally:
'   fallback-before-commit, cache preservation, checkpoint abandonment.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   The next call to QPC_TryReadTick returns FALSE and clears the flag.
'
' ERROR POLICY
'   Does not raise.
'
' NOTES
'   The flag is deliberately ONE-SHOT and self-clearing. A sticky flag would
'   make test ordering significant and would leak a forced failure into later
'   cases - precisely the hidden state this class exists to eliminate.
'
'   Declared Friend, so it is reachable from the regression suite but invisible
'   outside the project.
'
' UPDATED
'   2026-08-15
'==============================================================================

    'Arm a single forced QPC read failure
        m_TestFailNextQpcRead = True

End Sub


Friend Sub Test_ForceNextSystemTimeReadFailure()
'
'==============================================================================
'                  TEST FORCE NEXT SYSTEM TIME READ FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Arms a ONE-SHOT forced failure of the next timeGetSystemTime read.
'
' WHY THIS EXISTS
'   Backend 4 previously converted a failed read into a large positive duration
'   through rollover correction. Proving that can no longer happen requires the
'   ability to force the read to fail.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   The next call to SystemTime_TryReadMs returns FALSE with
'   cPM_ReadSystemTimeFailed, and clears the flag.
'
' ERROR POLICY
'   Does not raise.
'
' NOTES
'   One-shot and self-clearing, for the same reasons as the QPC seam.
'
' UPDATED
'   2026-08-15
'==============================================================================

    'Arm a single forced timeGetSystemTime read failure
        m_TestFailNextSystemTimeRead = True

End Sub


Friend Sub Test_ForceNextSystemTimeFormatInvalid()
'
'==============================================================================
'              TEST FORCE NEXT SYSTEM TIME FORMAT INVALID
'------------------------------------------------------------------------------
' PURPOSE
'   Arms a ONE-SHOT forced wrong-format result from the next timeGetSystemTime
'   read.
'
' WHY THIS EXISTS
'   The runtime distinguishes an outright API failure from a successful call
'   that returns a format other than TIME_MS, and reports them as
'   cPM_ReadSystemTimeFailed and cPM_ReadSystemTimeFormatInvalid respectively.
'
'   Without a seam, only the first of those two paths could be exercised, so
'   the distinction would be asserted by the code but never by a test.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   The next call to SystemTime_TryReadMs returns FALSE with
'   cPM_ReadSystemTimeFormatInvalid, and clears the flag.
'
' ERROR POLICY
'   Does not raise.
'
' NOTES
'   One-shot and self-clearing, for the same reasons as the other seams: a
'   sticky flag would make test ordering significant.
'
'   Armed independently of the read-failure flag. If both are armed, the
'   read-failure flag is consumed first, matching the order in which the
'   runtime checks them.
'
' UPDATED
'   2026-08-16
'==============================================================================

    'Arm a single forced wrong-format result
        m_TestForceNextSystemTimeFormatInvalid = True

End Sub

Public Property Get TW_CalculationExempted() As Boolean
'
'==============================================================================
'                          TW CALCULATION EXEMPTED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when Calculation control could not be honoured on this host and
'   has been exempted from the shared effective state
'
' WHY THIS EXISTS
'   Application.Calculation cannot be read or written when no workbook is open.
'   In non-strict mode the shared manager exempts the flag rather than guessing
'   a baseline, and the caller needs a way to discover that rather than assuming
'   its suppression request was applied
'
' INPUTS
'   None.
'
' RETURNS
'   Boolean
'     TRUE  => Calculation was requested or captured but could not be honoured
'     FALSE => Calculation control is being applied normally, or was never wanted
'
' BEHAVIOR
'   Delegates to PM_TW_CalculationExempted in the shared manager.
'
' ERROR POLICY
'   Does not raise.
'
' NOTES
'   Calculation control requires the open-workbook set to remain stable for the
'   life of the shared scope. A baseline captured with no workbook open is
'   unknown, and the manager will not lazily capture one later, because that
'   would apply a baseline the scope never observed.
'
'   In strict mode an unhonourable Calculation request raises from TW_Turn_OFF
'   instead, so this property reports FALSE on that path - the session was
'   never started.
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report whether the shared manager exempted Calculation control
        TW_CalculationExempted = PM_TW_CalculationExempted()

End Property





