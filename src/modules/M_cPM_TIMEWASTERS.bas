Attribute VB_Name = "M_cPM_TimeWasters"
'==============================================================================
' MODULE: M_cPM_TIMEWASTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Shared, process-wide manager for Excel "time-waster" suppression used by
'   cPerformanceManager
'
'   The module also hosts a worksheet-output helper for structured checkpoint
'   reports (cPM_Report_WriteToRange)
'
' WHY THIS EXISTS
'   The Excel Application properties controlled here are global process state,
'   not instance-local state:
'
'     - ScreenUpdating
'     - EnableEvents
'     - DisplayAlerts
'     - Calculation
'     - Cursor
'
'   Therefore:
'     - multiple cPerformanceManager instances can overlap
'     - each instance can request a different exemption mask
'     - restore logic must be coordinated globally
'     - the final restore must happen exactly once when the last active session
'       ends
'
' REQUIRED BY
'   cPerformanceManager
'
' COMPILE-TIME CONTRACT
'   This module is required by cPerformanceManager because the class directly
'   references the following procedures exposed here:
'
'     - PM_TW_BeginSession
'     - PM_TW_EndSession
'     - PM_TW_ActiveCount
'     - PM_TW_NewInstanceKey
'
'   Additional diagnostic / recovery procedures exposed here are:
'
'     - PM_TW_IsInstanceActive
'     - PM_TW_EndAllSessions
'     - PM_TW_CalculationExempted
'
'   Worksheet-output helper exposed here:
'
'     - cPM_Report_WriteToRange
'
' IMPORT REQUIREMENT
'   Import this module together with cPerformanceManager.cls
'
' DESIGN
'   - The first active session captures the original Application baseline
'   - Each active instance registers its own disable-mask
'   - The effective disable-mask is the OR of all active instance masks
'   - Whenever a session begins, updates, or ends, the effective state is
'     recomputed
'   - When the final session ends, the original baseline is restored exactly
'     once and the shared store is released
'
' INSTANCE KEY POLICY
'   Instance keys are issued by PM_TW_NewInstanceKey from a module-level counter
'
'   A counter is sufficient and collision-proof here because g_TW_KeySeed and
'   g_TW_Sessions share the same module-global lifetime. A VBA project reset
'   clears both together, so the seed can never issue a key that a surviving
'   dictionary entry still holds
'
'   Object addresses (ObjPtr) must NOT be used as keys. VBA reuses heap
'   addresses, so a session that outlived its instance could be silently
'   inherited by an unrelated instance allocated at the same address
'
' HOST-STATE POLICY
'   Application.Calculation cannot be read or written when no workbook is open
'
'   Both the baseline capture and the effective-state apply paths therefore
'   guard on Workbooks.Count. All other supported flags remain safe to access
'   in a workbook-less host
'
' CALCULATION BASELINE VALIDITY
'   An unknown Calculation baseline and a baseline that happens to be Automatic
'   are different states, and must never share a representation. Recording a
'   synthetic Automatic when none could be read is what allowed a real baseline
'   to be overwritten on a workbook that opened later
'
'   g_TW_CALCULATION_VALID therefore tracks whether a baseline was genuinely
'   captured, separately from its value
'
' STABLE-HOST INVARIANT
'   Calculation control requires the open-workbook set to remain stable for the
'   life of the shared scope
'
'   When that invariant does not hold:
'     - the flag is exempted rather than guessed
'     - PM_TW_CalculationExempted reports the exemption
'     - a strict participant causes PM_TW_BeginSession to raise
'
'   The manager deliberately does NOT lazily capture a baseline once a workbook
'   opens mid-scope. Doing so would apply a baseline the scope never observed,
'   and no event reapplies the aggregate state on workbook lifecycle changes
'
' DEPENDENCIES
'   - Excel Application object model
'   - Late-bound Scripting.Dictionary via CreateObject("Scripting.Dictionary")
'
' REFERENCE POLICY
'   No manual reference to "Microsoft Scripting Runtime" is required
'
' ERROR POLICY
'   - Public begin/end operations raise errors normally for invalid calling
'     flow, invalid inputs, or Application-state failures
'   - Idempotent no-op paths remain intentional, for example when ending an
'     already-idle shared store or an inactive instance key
'   - Internal helpers assume a valid Excel host and valid calling flow
'   - A workbook-less host is NOT treated as an error. The Calculation flag is
'     skipped in that case rather than raising
'   - PM_TW_EndAllSessions is an emergency / recovery helper and also raises
'     errors normally
'
' NOTES
'   - This module should not appear as a user-runnable macro surface.
'     Therefore Option Private Module is used. Public procedures here remain
'     callable from anywhere inside the host VBA project
'   - Cursor suppression uses xlWait while active to force a deterministic
'     benchmark-time cursor state and avoid ordinary cursor-state churn during
'     benchmark runs
'
' VERSION
'   1.4.0
'
' UPDATED
'   2026-09-03 - Made Calculation exemption sticky for the complete scope.
'   2026-08-30
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit declaration of all variables
    Option Private Module   'Hide internal support procedures from Macro dialog

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Known TW mask bits
        Private Const PM_TW_MASK_NONE               As Long = 0
        Private Const PM_TW_MASK_SCREENUPDATING     As Long = 1
        Private Const PM_TW_MASK_ENABLEEVENTS       As Long = 2
        Private Const PM_TW_MASK_DISPLAYALERTS      As Long = 4
        Private Const PM_TW_MASK_CALCULATION        As Long = 8
        Private Const PM_TW_MASK_CURSOR             As Long = 16
        Private Const PM_TW_MASK_ALL                As Long = 31

    'Error numbers raised by this module.
    '
    'Public so callers can trap a specific condition by name rather than by a
    'magic offset. The values are unchanged from earlier releases.
    Public Enum cPM_TWError
        'Shared session registration
        ERR_TW_BEGIN_BLANK_KEY = vbObjectError + 2200
        ERR_TW_END_BLANK_KEY = vbObjectError + 2201
        ERR_TW_ISACTIVE_BLANK_KEY = vbObjectError + 2202
        ERR_TW_CALCULATION_UNAVAILABLE = vbObjectError + 2203

        'Structured report output
        ERR_TW_REPORT_NO_INSTANCE = vbObjectError + 2600
        ERR_TW_REPORT_NO_TARGET = vbObjectError + 2601
    End Enum

    'Structured report column positions
    '
    'These must track the column order built by cPerformanceManager.ReportAsArray
        Private Const PM_RPT_COL_DELTA              As Long = 7
        Private Const PM_RPT_COL_CUMULATIVE         As Long = 8

    'Number format applied to the numeric timing columns
        Private Const PM_RPT_TIMING_FORMAT          As String = "0.000000000"

'------------------------------------------------------------------------------
' PRIVATE SHARED STATE
'------------------------------------------------------------------------------
    'Dictionary:
    '   key   = instance key (String)
    '   item  = disable-mask (Long)
        Private g_TW_Sessions               As Object

    'Dictionary:
    '   key   = instance key (String)
    '   item  = TRUE when that instance requested strict host semantics
        Private g_TW_StrictKeys             As Object

    'Monotonic seed used to issue collision-proof instance keys
        Private g_TW_KeySeed                As Long

    'TRUE once the baseline Application state has been captured
        Private g_TW_BaselineSaved          As Boolean

    'TRUE only when a REAL Calculation baseline was captured from a live
    'workbook. An unknown baseline and a baseline that happens to be Automatic
    'are different states and must never share a representation.
        Private g_TW_CALCULATION_VALID      As Boolean

    'TRUE when Calculation control could not be honoured on this host and has
    'been exempted from the effective state for the life of the scope
        Private g_TW_CalcExempted           As Boolean

    'Saved baseline Application state
        Private g_TW_SCREENUPDATING         As Boolean
        Private g_TW_ENABLEEVENTS           As Boolean
        Private g_TW_DISPLAYALERTS          As Boolean
        Private g_TW_CALCULATION            As XlCalculation
        Private g_TW_CURSOR                 As XlMousePointer

'
'==============================================================================
'
'                                  PUBLIC API
'
'==============================================================================

Public Function PM_TW_NewInstanceKey() As String
'
'==============================================================================
'                           PM_TW_NEWINSTANCEKEY
'------------------------------------------------------------------------------
' PURPOSE
'   Issues a unique shared-TW registration key for one class instance
'
' WHY THIS EXISTS
'   Earlier revisions derived the instance key from ObjPtr(Me). That is unsafe:
'   VBA reuses heap addresses, so a session left behind by a destroyed instance
'   could be silently inherited by an unrelated instance later allocated at the
'   same address
'
'   A module-level counter removes that hazard entirely. g_TW_KeySeed and
'   g_TW_Sessions share the same module-global lifetime, so a project reset
'   clears both at once and the seed can never reissue a key that a surviving
'   dictionary entry still holds
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     Unique instance key for the lifetime of the current VBA project session
'
' BEHAVIOR
'   - Increments the shared key seed
'   - Returns a prefixed string form of the new seed value
'
' ERROR POLICY
'   Does not raise errors
'
' DEPENDENCIES
'   - g_TW_KeySeed
'
' NOTES
'   The caller is expected to cache the returned key for the lifetime of the
'   instance. Calling this repeatedly for the same instance would register that
'   instance more than once
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ISSUE KEY
'------------------------------------------------------------------------------
    'Advance the shared seed
        g_TW_KeySeed = g_TW_KeySeed + 1

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the prefixed key
        PM_TW_NewInstanceKey = "cPM#" & CStr(g_TW_KeySeed)

End Function

Public Sub PM_TW_BeginSession( _
    ByVal InstanceKey As String, _
    Optional ByVal ExceptMask As Long = 0, _
    Optional ByVal StrictHost As Boolean = False)
'
'==============================================================================
'                             PM_TW_BEGINSESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Starts or updates a shared TW suppression session for one class instance
'
' WHY THIS EXISTS
'   The calling class cannot safely manage Excel Application TW state
'   independently because that state is global across the Excel process.
'   Therefore the class registers its request here and this manager computes the
'   aggregate effective state across all active instances
'
' INPUTS
'   InstanceKey
'     Unique key identifying the calling class instance
'     Obtain it from PM_TW_NewInstanceKey
'
'   ExceptMask (optional)
'     Bitmask of TW flags to EXEMPT
'     Any flag present in ExceptMask remains at the original baseline state for
'     this instance's request
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Ensures the shared session store exists
'   - Captures the original Application baseline when the first active session
'     begins
'   - Converts the supplied exemption mask into a disable-mask
'   - Registers or updates this instance in the shared session store
'   - Recomputes and applies the aggregate effective state
'   - Rolls back the registration/update if effective-state application fails
'   - Best-effort reapplies the previous effective state after rollback
'
' ERROR POLICY
'   Raises errors normally
'
'   On failure the original error is preserved and re-raised after rollback, so
'   the caller always sees the true cause rather than a rollback side effect
'
' DEPENDENCIES
'   - PM_TW_EnsureStore
'   - PM_TW_SaveBaseline
'   - PM_TW_DisableMaskFromExcept
'   - PM_TW_AggregateDisableMask
'   - PM_TW_ApplyEffectiveState
'   - PM_TW_ResetSharedState
'
' NOTES
'   This routine is instance-idempotent:
'     - first call for an instance => begin/register
'     - later calls for same instance => update requested mask
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HadKeyBefore            As Boolean   'TRUE when the instance was already registered
    Dim PrevDisableMask         As Long      'Previously stored disable-mask for this instance
    Dim PrevStrictHost          As Boolean   'Previously stored strictness for this instance
    Dim WasFirstSession         As Boolean   'TRUE when this call began the first shared session

    Dim SavedErrNumber          As Long      'Captured original error number
    Dim SavedErrSource          As String    'Captured original error source
    Dim SavedErrDescription     As String    'Captured original error description

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank instance key
        If Len(Trim$(InstanceKey)) = 0 Then
            Err.Raise ERR_TW_BEGIN_BLANK_KEY, _
                      "M_cPM_TimeWasters.PM_TW_BeginSession", _
                      "InstanceKey cannot be blank."
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists
        PM_TW_EnsureStore
    'Capture whether the key already exists before modification
        HadKeyBefore = g_TW_Sessions.Exists(InstanceKey)
    'Capture whether this call is opening the first shared session
        WasFirstSession = (g_TW_Sessions.Count = 0)
    'Capture the previous disable-mask when this is an update
        If HadKeyBefore Then
            PrevDisableMask = CLng(g_TW_Sessions(InstanceKey))
            If g_TW_StrictKeys.Exists(InstanceKey) Then
                PrevStrictHost = CBool(g_TW_StrictKeys(InstanceKey))
            End If
        End If

'------------------------------------------------------------------------------
' CAPTURE BASELINE (FIRST ACTIVE SESSION ONLY)
'------------------------------------------------------------------------------
    'Capture the original Application state only when the first shared session begins
        If WasFirstSession Then
            PM_TW_SaveBaseline
        End If

'------------------------------------------------------------------------------
' REGISTER / UPDATE INSTANCE MASK
'------------------------------------------------------------------------------
    'Store this instance's requested disable-mask
        g_TW_Sessions(InstanceKey) = PM_TW_DisableMaskFromExcept(ExceptMask)
    'Store this instance's host-strictness preference alongside it
        g_TW_StrictKeys(InstanceKey) = StrictHost

'------------------------------------------------------------------------------
' APPLY EFFECTIVE SHARED STATE
'------------------------------------------------------------------------------
    'Recompute the aggregate disable-mask and apply it
        On Error GoTo ApplyFail
        PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()

    'Reject the session when Calculation control was requested but could not be
    'honoured and at least one active instance asked for strict host semantics.
    'Raising here routes through ApplyFail, which unwinds the registration.
        If g_TW_CalcExempted And PM_TW_AnyStrict() Then
            If (PM_TW_AggregateDisableMask() And PM_TW_MASK_CALCULATION) <> 0 Then
                Err.Raise ERR_TW_CALCULATION_UNAVAILABLE, _
                          "M_cPM_TimeWasters.PM_TW_BeginSession", _
                          "Calculation suppression was requested but no Calculation " & _
                          "baseline could be captured. Open a workbook before starting " & _
                          "the scope, exempt TW_Enum.Calculation, or use non-strict mode."
            End If
        End If
        On Error GoTo 0

    Exit Sub

ApplyFail:
'------------------------------------------------------------------------------
' CAPTURE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Capture the original error before rollback
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

'------------------------------------------------------------------------------
' ROLLBACK REGISTRATION / UPDATE
'------------------------------------------------------------------------------
    'Rollback must be best-effort so that we can preserve the original error
        On Error Resume Next

    'Restore the prior registration state
        If HadKeyBefore Then
            g_TW_Sessions(InstanceKey) = PrevDisableMask
            g_TW_StrictKeys(InstanceKey) = PrevStrictHost
        ElseIf g_TW_Sessions.Exists(InstanceKey) Then
            g_TW_Sessions.Remove InstanceKey
            If g_TW_StrictKeys.Exists(InstanceKey) Then
                g_TW_StrictKeys.Remove InstanceKey
            End If
        End If

'------------------------------------------------------------------------------
' BEST-EFFORT REAPPLY PREVIOUS EFFECTIVE STATE
'------------------------------------------------------------------------------
    'If rollback returned the manager to a true idle state, try to restore the
    'baseline and then clear shared state
        If g_TW_Sessions.Count = 0 Then
            If g_TW_BaselineSaved Then
                Err.Clear
                PM_TW_ApplyEffectiveState PM_TW_MASK_NONE

                If Err.Number = 0 Then
                    PM_TW_ResetSharedState
                End If
            Else
                PM_TW_ResetSharedState
            End If
    'Otherwise best-effort reapply the prior aggregate effective state
        Else
            Err.Clear
            PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
        End If

        On Error GoTo 0

'------------------------------------------------------------------------------
' RE-RAISE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Re-raise the original error
        Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription

End Sub

Public Sub PM_TW_EndSession( _
    ByVal InstanceKey As String)
'
'==============================================================================
'                              PM_TW_ENDSESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Ends a shared TW suppression session for one class instance
'
' WHY THIS EXISTS
'   TW suppression must be removed by deregistering the calling instance from
'   the shared manager, not by blindly restoring Excel state locally. This keeps
'   overlapping instances safe and ensures the effective state is recomputed
'   correctly from the remaining active sessions
'
' INPUTS
'   InstanceKey
'     Unique key identifying the calling class instance
'
' RETURNS
'   None
'
' BEHAVIOR
'   - If no shared store exists, exits immediately
'   - Removes the specified instance if present
'   - If no sessions remain:
'       * restores the original baseline exactly once when available
'       * clears the baseline-saved flag
'       * releases the shared dictionary
'   - Otherwise recomputes and reapplies the remaining effective state
'   - Rolls back the removal if effective-state application fails
'
' ERROR POLICY
'   Raises errors normally
'
'   On failure the original error is preserved and re-raised after rollback
'
' DEPENDENCIES
'   - PM_TW_ApplyEffectiveState
'   - PM_TW_AggregateDisableMask
'   - PM_TW_ResetSharedState
'
' NOTES
'   This routine is idempotent for unknown / inactive instance keys:
'     - if the instance is not present, no removal occurs
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HadKeyBefore            As Boolean   'TRUE when the instance was registered before removal
    Dim PrevDisableMask         As Long      'Previously stored disable-mask for this instance
    Dim PrevStrictHost          As Boolean   'Previously stored strictness for this instance

    Dim SavedErrNumber          As Long      'Captured original error number
    Dim SavedErrSource          As String    'Captured original error source
    Dim SavedErrDescription     As String    'Captured original error description

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank instance key
        If Len(Trim$(InstanceKey)) = 0 Then
            Err.Raise ERR_TW_END_BLANK_KEY, _
                      "M_cPM_TimeWasters.PM_TW_EndSession", _
                      "InstanceKey cannot be blank."
        End If

'------------------------------------------------------------------------------
' VALIDATE STORE
'------------------------------------------------------------------------------
    'If no shared store exists yet, there is nothing to end
        If g_TW_Sessions Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE PRE-REMOVAL STATE
'------------------------------------------------------------------------------
    'Capture whether the instance is currently registered
        HadKeyBefore = g_TW_Sessions.Exists(InstanceKey)

    'Capture the previous disable-mask when present
        If HadKeyBefore Then
            PrevDisableMask = CLng(g_TW_Sessions(InstanceKey))
            If Not g_TW_StrictKeys Is Nothing Then
                If g_TW_StrictKeys.Exists(InstanceKey) Then
                    PrevStrictHost = CBool(g_TW_StrictKeys(InstanceKey))
                End If
            End If
        End If

'------------------------------------------------------------------------------
' REMOVE INSTANCE (IF PRESENT)
'------------------------------------------------------------------------------
    'Remove the calling instance from both shared stores
        If HadKeyBefore Then
            g_TW_Sessions.Remove InstanceKey
            If Not g_TW_StrictKeys Is Nothing Then
                If g_TW_StrictKeys.Exists(InstanceKey) Then
                    g_TW_StrictKeys.Remove InstanceKey
                End If
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE OR REAPPLY
'------------------------------------------------------------------------------
        On Error GoTo ApplyFail

    'If no sessions remain, restore the original baseline
        If g_TW_Sessions.Count = 0 Then
            'Restore the original Application state only if a baseline was actually captured
                If g_TW_BaselineSaved Then
                    PM_TW_ApplyEffectiveState PM_TW_MASK_NONE
                End If

            'Return to a clean idle shared-state baseline
                PM_TW_ResetSharedState

            On Error GoTo 0
            Exit Sub
        End If

    'Otherwise recompute and apply the remaining aggregate disable-mask
        PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
        On Error GoTo 0

    Exit Sub

ApplyFail:
'------------------------------------------------------------------------------
' CAPTURE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Capture the original error before rollback
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

'------------------------------------------------------------------------------
' ROLLBACK REMOVAL
'------------------------------------------------------------------------------
    'Rollback must be best-effort so that we can preserve the original error
        On Error Resume Next
    'Restore the removed instance registration when it existed before the call
        If HadKeyBefore Then
            g_TW_Sessions(InstanceKey) = PrevDisableMask
            If Not g_TW_StrictKeys Is Nothing Then
                g_TW_StrictKeys(InstanceKey) = PrevStrictHost
            End If
        End If

'------------------------------------------------------------------------------
' BEST-EFFORT REAPPLY PREVIOUS EFFECTIVE STATE
'------------------------------------------------------------------------------
    'Best-effort reapply the prior aggregate effective state
        If Not g_TW_Sessions Is Nothing Then
            If g_TW_Sessions.Count > 0 Then
                Err.Clear
                PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
            End If
        End If

        On Error GoTo 0

'------------------------------------------------------------------------------
' RE-RAISE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Re-raise the original error
        Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription

End Sub

Public Function PM_TW_ActiveCount() As Long
'
'==============================================================================
'                             PM_TW_ACTIVECOUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of currently active shared TW sessions
'
' WHY THIS EXISTS
'   Useful for diagnostics, assertions, and visibility when verifying shared TW
'   lifecycle behavior across multiple cPerformanceManager instances
'
' INPUTS
'   None.
'
' RETURNS
'   Long
'     Number of currently active shared TW sessions
'
' BEHAVIOR
'   - Returns 0 when no shared store currently exists
'   - Otherwise returns the dictionary count
'
' ERROR POLICY
'   Does not raise errors
'
' DEPENDENCIES
'   - g_TW_Sessions
'
' NOTES
'   This routine does not create the shared store on an idle read path
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE STORE
'------------------------------------------------------------------------------
    'Return 0 when no shared store currently exists
        If g_TW_Sessions Is Nothing Then
            PM_TW_ActiveCount = 0
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the active session count
        PM_TW_ActiveCount = g_TW_Sessions.Count

End Function

Public Function PM_TW_IsInstanceActive( _
    ByVal InstanceKey As String) _
    As Boolean
'
'==============================================================================
'                          PM_TW_ISINSTANCEACTIVE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if the specified class instance currently has an active shared
'   TW session registered in the global manager
'
' WHY THIS EXISTS
'   Useful for diagnostics, regression testing, and troubleshooting of shared
'   TW registration / update / end behavior
'
' INPUTS
'   InstanceKey
'     Unique key identifying the class instance to inspect
'
' RETURNS
'   Boolean
'     TRUE  => the instance is currently registered
'     FALSE => the instance is not currently registered
'
' BEHAVIOR
'   - Returns FALSE when no shared store currently exists
'   - Otherwise queries the dictionary for the supplied key
'
' ERROR POLICY
'   Raises on a blank instance key
'
' DEPENDENCIES
'   - g_TW_Sessions
'
' NOTES
'   This routine does not create the shared store on an idle read path
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank instance key
        If Len(Trim$(InstanceKey)) = 0 Then
            Err.Raise ERR_TW_ISACTIVE_BLANK_KEY, _
                      "M_cPM_TimeWasters.PM_TW_IsInstanceActive", _
                      "InstanceKey cannot be blank."
        End If

'------------------------------------------------------------------------------
' VALIDATE STORE
'------------------------------------------------------------------------------
    'Return FALSE when no shared store currently exists
        If g_TW_Sessions Is Nothing Then
            PM_TW_IsInstanceActive = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return instance activity state
        PM_TW_IsInstanceActive = g_TW_Sessions.Exists(InstanceKey)

End Function

Public Sub PM_TW_EndAllSessions()
'
'==============================================================================
'                            PM_TW_ENDALLSESSIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Emergency / global reset for development, recovery, or test-cleanup
'   scenarios
'
' WHY THIS EXISTS
'   In normal operation each instance should end only its own session through
'   PM_TW_EndSession. However, during development, test teardown, or recovery
'   from interrupted flows it can be useful to force a full shared reset
'
'   A hard End statement clears module globals without running Class_Terminate,
'   which can leave Excel visibly suppressed with no active session to end.
'   This routine is the documented recovery path for that situation
'
' INPUTS
'   None.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Restores the original Application baseline when available
'   - Clears the baseline-saved flag
'   - Releases all shared session bookkeeping
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - PM_TW_ApplyEffectiveState
'   - PM_TW_ResetSharedState
'
' NOTES
'   This is not the normal lifecycle path.
'   Normal callers should use PM_TW_EndSession for the specific active instance
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' RESTORE BASELINE (IF AVAILABLE)
'------------------------------------------------------------------------------
    'Restore the original Application baseline if available
        If g_TW_BaselineSaved Then
            PM_TW_ApplyEffectiveState PM_TW_MASK_NONE
        End If

'------------------------------------------------------------------------------
' CLEAR SHARED STATE
'------------------------------------------------------------------------------
    'Release all active-session bookkeeping and baseline state
        PM_TW_ResetSharedState

End Sub

'
'==============================================================================
'
'                                PRIVATE HELPERS
'
'==============================================================================

Private Sub PM_TW_EnsureStore()
'
'==============================================================================
'                             PM_TW_ENSURESTORE
'------------------------------------------------------------------------------
' PURPOSE
'   Lazily creates the shared session dictionary
'
' WHY THIS EXISTS
'   The shared TW store should exist only when needed. This helper centralizes
'   the lazy-creation logic and keeps begin/update paths simple
'
' INPUTS
'   None.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates the shared dictionary only when it does not already exist
'   - Uses late binding so no external reference is required
'   - Sets binary key comparison explicitly
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - CreateObject("Scripting.Dictionary")
'
' NOTES
'   Read-only status helpers intentionally avoid calling this routine so that an
'   idle project remains in a true idle state
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create the shared dictionary only once
        If g_TW_Sessions Is Nothing Then
            Set g_TW_Sessions = CreateObject("Scripting.Dictionary")
            g_TW_Sessions.CompareMode = vbBinaryCompare
        End If

    'Create the parallel strictness store on the same lifetime
        If g_TW_StrictKeys Is Nothing Then
            Set g_TW_StrictKeys = CreateObject("Scripting.Dictionary")
            g_TW_StrictKeys.CompareMode = vbBinaryCompare
        End If

End Sub

Private Sub PM_TW_SaveBaseline()
'
'==============================================================================
'                             PM_TW_SAVEBASELINE
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the original Application state before any TW suppression is
'   applied
'
' WHY THIS EXISTS
'   Shared TW suppression must restore the exact original Excel baseline when
'   the final active session ends. Therefore that baseline must be captured once
'   at the beginning of the first shared session
'
' INPUTS
'   None.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the current Application state into shared baseline variables
'   - Skips Calculation when no workbook is open and records a safe default
'   - Marks the baseline as captured
'
' ERROR POLICY
'   Raises errors normally
'
'   A workbook-less host is not an error. Application.Calculation cannot be read
'   in that state, so the read is guarded rather than allowed to raise
'
' DEPENDENCIES
'   - Excel Application object model
'
' NOTES
'   This routine should only be called when the first shared session begins
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' CAPTURE BASELINE
'------------------------------------------------------------------------------
    'Read the current Application state into shared baseline variables
        With Application
            g_TW_SCREENUPDATING = .ScreenUpdating
            g_TW_ENABLEEVENTS = .EnableEvents
            g_TW_DISPLAYALERTS = .DisplayAlerts
            g_TW_CURSOR = .Cursor

            'Calculation is only readable while at least one workbook is open.
            'When it cannot be read the baseline is UNKNOWN, which is a distinct
            'state from "the baseline happened to be Automatic". Recording a
            'synthetic value here is what allowed a real baseline to be
            'overwritten later.
                If .Workbooks.Count > 0 Then
                    g_TW_CALCULATION = .Calculation
                    g_TW_CALCULATION_VALID = True
                Else
                    g_TW_CALCULATION = xlCalculationAutomatic
                    g_TW_CALCULATION_VALID = False
                End If
        End With

'------------------------------------------------------------------------------
' UPDATE STATE
'------------------------------------------------------------------------------
    'Mark the baseline as captured
        g_TW_BaselineSaved = True

End Sub

Private Sub PM_TW_ResetSharedState()
'
'==============================================================================
'                           PM_TW_RESETSHAREDSTATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the shared TW manager to a clean idle state
'
' WHY THIS EXISTS
'   Several paths need to clear the same shared-state variables:
'     - successful end of the final active session
'     - explicit emergency reset
'     - rollback after a failed first-session begin
'
'   Centralizing that reset logic avoids partial cleanup drift
'
' INPUTS
'   None.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Releases the shared session dictionary
'   - Clears the baseline-saved flag
'   - Clears the cached baseline values
'
' ERROR POLICY
'   Raises errors normally
'
' NOTES
'   The key seed is deliberately NOT reset here
'
'   Resetting it would allow a key to be reissued while a stale registration
'   could still exist, which is exactly the collision hazard this design avoids
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' CLEAR SHARED STORE
'------------------------------------------------------------------------------
    'Release the shared session dictionaries
        Set g_TW_Sessions = Nothing
        Set g_TW_StrictKeys = Nothing

'------------------------------------------------------------------------------
' CLEAR BASELINE FLAGS / VALUES
'------------------------------------------------------------------------------
    'Mark that no baseline is currently cached
        g_TW_BaselineSaved = False

    'Clear cached baseline values deterministically
        g_TW_SCREENUPDATING = False
        g_TW_ENABLEEVENTS = False
        g_TW_DISPLAYALERTS = False
        g_TW_CALCULATION = xlCalculationAutomatic
        g_TW_CALCULATION_VALID = False
        g_TW_CalcExempted = False
        g_TW_CURSOR = xlDefault

End Sub

Private Function PM_TW_DisableMaskFromExcept( _
    ByVal ExceptMask As Long) _
    As Long
'
'==============================================================================
'                        PM_TW_DISABLEMASKFROMEXCEPT
'------------------------------------------------------------------------------
' PURPOSE
'   Converts an exemption mask into a disable-mask
'
' WHY THIS EXISTS
'   Callers express their request as "leave these flags alone." Internally the
'   shared manager works more naturally with a disable-mask, because aggregate
'   shared state is computed by OR-ing together every active instance's disabled
'   flags
'
' INPUTS
'   ExceptMask
'     Bitmask of TW flags to exempt
'
' RETURNS
'   Long
'     Disable-mask representing the supported flags that should be forced into
'     their benchmark/performance state
'
' BEHAVIOR
'   - Inverts the supplied exemption bits
'   - Bounds the RESULT to the supported TW mask universe
'
' ERROR POLICY
'   Does not raise errors
'
' DEPENDENCIES
'   - PM_TW_MASK_ALL
'
' EXAMPLE
'   ExceptMask = ScreenUpdating Or EnableEvents
'   => disable all supported TW flags except those two
'
' NOTES
'   Unsupported high bits in ExceptMask are harmless. The AND against
'   PM_TW_MASK_ALL bounds the returned mask, so unknown bits can never reach the
'   apply path
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Invert the exemption bits and bound the result to the supported TW mask set
        PM_TW_DisableMaskFromExcept = (PM_TW_MASK_ALL And Not ExceptMask)

End Function

Private Function PM_TW_AggregateDisableMask() As Long
'
'==============================================================================
'                        PM_TW_AGGREGATEDISABLEMASK
'------------------------------------------------------------------------------
' PURPOSE
'   ORs together the disable-masks of all active shared TW sessions
'
' WHY THIS EXISTS
'   The effective TW state is the union of every currently active instance's
'   requested disable-mask. This helper centralizes that aggregation
'
' INPUTS
'   None.
'
' RETURNS
'   Long
'     Aggregate disable-mask across all active sessions
'
' BEHAVIOR
'   - Ensures the shared store exists
'   - ORs together every active instance's stored disable-mask
'   - Returns the final aggregate mask
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - PM_TW_EnsureStore
'   - g_TW_Sessions
'
' NOTES
'   Unlike the public read-only status helpers, this routine does create the
'   shared store. It is only reached from begin/end paths where the store is
'   expected to exist
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K    As Variant    'Dictionary key
    Dim Mask As Long       'Accumulated disable-mask

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists
        PM_TW_EnsureStore

'------------------------------------------------------------------------------
' AGGREGATE
'------------------------------------------------------------------------------
    'OR together every active instance's disable-mask
        For Each K In g_TW_Sessions.Keys
            Mask = (Mask Or CLng(g_TW_Sessions(K)))
        Next K

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the aggregate disable-mask
        PM_TW_AggregateDisableMask = Mask

End Function

Private Sub PM_TW_ApplyEffectiveState( _
    ByVal DisableMask As Long)
'
'==============================================================================
'                        PM_TW_APPLYEFFECTIVESTATE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the effective shared TW state using:
'     - the original saved baseline, and
'     - the aggregate disable-mask of all active sessions
'
' WHY THIS EXISTS
'   Shared TW control should never restore or force individual flags in an
'   ad-hoc per-instance way. Instead, each flag must be derived from:
'
'     - the original baseline captured at the first session, and
'     - whether any currently active instance wants that flag disabled
'
' INPUTS
'   DisableMask
'     Aggregate disable-mask to apply
'
' RETURNS
'   None
'
' BEHAVIOR
'   For each supported flag:
'     - if disabled by any active session => force benchmark/performance state
'     - otherwise => restore original baseline state
'
'   Calculation is skipped entirely when no workbook is open
'
' ERROR POLICY
'   Raises errors normally
'
'   A workbook-less host is not an error. Application.Calculation cannot be
'   written in that state, so the whole Calculation branch is guarded
'
' DEPENDENCIES
'   - Excel Application object model
'   - g_TW_SCREENUPDATING
'   - g_TW_ENABLEEVENTS
'   - g_TW_DISPLAYALERTS
'   - g_TW_CALCULATION
'   - g_TW_CURSOR
'
' NOTES
'   - This routine assumes the baseline Application state has already been
'     captured when restoration semantics are required
'   - Cursor uses xlWait when disabled to force a deterministic benchmark-time
'     cursor state
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' APPLY EFFECTIVE STATE
'------------------------------------------------------------------------------
    'Apply the effective state flag-by-flag
        With Application
            'ScreenUpdating
                If (DisableMask And PM_TW_MASK_SCREENUPDATING) <> 0 Then
                    .ScreenUpdating = False
                Else
                    .ScreenUpdating = g_TW_SCREENUPDATING
                End If
            'EnableEvents
                If (DisableMask And PM_TW_MASK_ENABLEEVENTS) <> 0 Then
                    .EnableEvents = False
                Else
                    .EnableEvents = g_TW_ENABLEEVENTS
                End If
            'DisplayAlerts
                If (DisableMask And PM_TW_MASK_DISPLAYALERTS) <> 0 Then
                    .DisplayAlerts = False
                Else
                    .DisplayAlerts = g_TW_DISPLAYALERTS
                End If
            'Calculation requires BOTH a live workbook and a real captured
            'baseline. Without both, the flag is exempted rather than guessed.
                If Not g_TW_CalcExempted _
                   And g_TW_CALCULATION_VALID _
                   And .Workbooks.Count > 0 Then
                    If (DisableMask And PM_TW_MASK_CALCULATION) <> 0 Then
                        .Calculation = xlCalculationManual
                    Else
                        .Calculation = g_TW_CALCULATION
                    End If
                Else
                    'Record the exemption when Calculation control was actually
                    'wanted, or when a captured baseline can no longer be restored
                        If ((DisableMask And PM_TW_MASK_CALCULATION) <> 0) _
                           Or g_TW_CALCULATION_VALID Then
                            g_TW_CalcExempted = True
                        End If
                End If
            'Cursor
                If (DisableMask And PM_TW_MASK_CURSOR) <> 0 Then
                    .Cursor = xlWait
                Else
                    .Cursor = g_TW_CURSOR
                End If
        End With

End Sub

'
'------------------------------------------------------------------------------
'
'                       HELPERS FOR WORKSHEET OUTPUT
'
'------------------------------------------------------------------------------
'

Public Sub cPM_Report_WriteToRange( _
    ByVal cPM As cPerformanceManager, _
    ByVal TargetTopLeft As Range, _
    Optional ByVal ClearOutputArea As Boolean = False)
'
'==============================================================================
'                         REPORT WRITE TO RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the structured checkpoint report to a worksheet range
'
' WHY THIS EXISTS
'   The class should own timing and structured capture, while worksheet output
'   belongs more naturally in a standard helper module
'
' INPUTS
'   cPM
'     Source cPerformanceManager instance
'
'   TargetTopLeft
'     Top-left output cell for the report
'
'   ClearOutputArea (optional)
'     TRUE  => clears the CurrentRegion of TargetTopLeft first
'     FALSE => writes only into the resized target report block
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the report as a 2D array
'   - Writes it to the target worksheet
'   - Applies lightweight formatting to headers and numeric columns
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - cPerformanceManager.ReportAsArray
'   - Excel Range object model
'
' NOTES
'   PM_RPT_COL_DELTA and PM_RPT_COL_CUMULATIVE must track the column order
'   produced by cPerformanceManager.ReportAsArray. Changing that column order
'   without updating these constants would format the wrong columns silently
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Data                As Variant              'Structured report array
    Dim RowCount            As Long                 'Report row count
    Dim ColCount            As Long                 'Report column count
    Dim OutputRange         As Range                'Resolved output range

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing class instance
        If cPM Is Nothing Then
            Err.Raise ERR_TW_REPORT_NO_INSTANCE, _
                      "M_cPM_TimeWasters.cPM_Report_WriteToRange", _
                      "cPerformanceManager instance cannot be Nothing."
        End If
    'Reject a missing output anchor cell
        If TargetTopLeft Is Nothing Then
            Err.Raise ERR_TW_REPORT_NO_TARGET, _
                      "M_cPM_TimeWasters.cPM_Report_WriteToRange", _
                      "TargetTopLeft cannot be Nothing."
        End If

'------------------------------------------------------------------------------
' READ REPORT ARRAY
'------------------------------------------------------------------------------
    'Read the structured checkpoint report as a 2D array
        Data = cPM.ReportAsArray
    'Resolve the report dimensions
        RowCount = UBound(Data, 1)
        ColCount = UBound(Data, 2)

'------------------------------------------------------------------------------
' RESOLVE OUTPUT RANGE
'------------------------------------------------------------------------------
    'Resolve the resized output range
        Set OutputRange = TargetTopLeft.Resize(RowCount, ColCount)

'------------------------------------------------------------------------------
' OPTIONAL CLEAR
'------------------------------------------------------------------------------
    'Clear the surrounding output area when requested
        If ClearOutputArea Then
            TargetTopLeft.CurrentRegion.Clear
        End If

'------------------------------------------------------------------------------
' WRITE REPORT
'------------------------------------------------------------------------------
    'Write the structured report array into the worksheet
        OutputRange.Value = Data

'------------------------------------------------------------------------------
' FORMAT OUTPUT
'------------------------------------------------------------------------------
    'Format the header row
        With OutputRange.Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    'Format numeric timing columns
        OutputRange.Columns(PM_RPT_COL_DELTA).NumberFormat = PM_RPT_TIMING_FORMAT
        OutputRange.Columns(PM_RPT_COL_CUMULATIVE).NumberFormat = PM_RPT_TIMING_FORMAT
    'Apply lightweight autofit
        OutputRange.EntireColumn.AutoFit

End Sub

Private Function PM_TW_AnyStrict() As Boolean
'
'==============================================================================
'                                PM_TW_ANYSTRICT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when any currently active instance requested strict host
'   semantics
'
' WHY THIS EXISTS
'   Host strictness is a per-instance preference, but the Application state it
'   governs is shared. If any participant asked to be told when Calculation
'   control cannot be honoured, the manager must tell it rather than silently
'   degrading for everyone
'
' INPUTS
'   None.
'
' RETURNS
'   Boolean
'     TRUE  => at least one active instance requested strict host semantics
'     FALSE => no active instance did, or no shared store exists
'
' BEHAVIOR
'   - Returns FALSE when no strictness store currently exists
'   - Otherwise ORs together every registered strictness flag
'
' ERROR POLICY
'   Does not raise errors
'
' NOTES
'   OR semantics mirror the aggregate disable-mask: the strictest active
'   requirement wins, exactly as the broadest suppression request does
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K As Variant    'Dictionary key

'------------------------------------------------------------------------------
' VALIDATE STORE
'------------------------------------------------------------------------------
    'Return FALSE when no strictness store currently exists
        If g_TW_StrictKeys Is Nothing Then
            PM_TW_AnyStrict = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' AGGREGATE
'------------------------------------------------------------------------------
    'Any single strict participant makes the whole scope strict
        For Each K In g_TW_StrictKeys.Keys
            If CBool(g_TW_StrictKeys(K)) Then
                PM_TW_AnyStrict = True
                Exit Function
            End If
        Next K

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'No active instance requested strict host semantics
        PM_TW_AnyStrict = False

End Function

Public Function PM_TW_CalculationExempted() As Boolean
'
'==============================================================================
'                        PM_TW_CALCULATIONEXEMPTED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when Calculation control could not be honoured on this host and
'   has been exempted from the effective state
'
' WHY THIS EXISTS
'   Application.Calculation cannot be read or written when no workbook is open.
'   In non-strict mode the manager exempts the flag rather than guessing a
'   baseline, and the caller needs a way to discover that rather than assuming
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
' ERROR POLICY
'   Does not raise errors
'
' NOTES
'   The exemption persists for the life of the shared scope. Calculation control
'   requires a stable open-workbook set: a baseline captured with no workbook
'   open is unknown, and the manager will not lazily capture one later, because
'   doing so would apply a baseline the scope never actually observed
'
' UPDATED
'   2026-08-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report whether Calculation control has been exempted
        PM_TW_CalculationExempted = g_TW_CalcExempted

End Function

Public Sub PM_TW_InternalTest_SetCalculationExempted( _
    ByVal IsExempted As Boolean)
'
'==============================================================================
'             PM_TW_INTERNALTEST_SETCALCULATIONEXEMPTED
'------------------------------------------------------------------------------
' PURPOSE
'   One-shot regression seam for the workbook-lifecycle exemption boundary.
'
' CONTRACT
'   Internal test infrastructure only. Production code must never call it.
'   Option Private Module keeps it outside the user-runnable macro surface.
'
' UPDATED
'   2026-09-03
'==============================================================================

    'Expose only the otherwise-unreachable host transition needed by the test.
        g_TW_CalcExempted = IsExempted

End Sub




