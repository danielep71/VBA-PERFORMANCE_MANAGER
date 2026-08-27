Attribute VB_Name = "M_cPM_USAGE_EXAMPLES"
'==============================================================================
' MODULE: M_cPM_USAGE_EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a compact, user-facing set of meaningful usage examples for
'   cPerformanceManager
'
' WHY THIS EXISTS
'   Once the demo workbook and regression suite exist, a large example module
'   becomes partly redundant. What remains genuinely useful is a smaller set of
'   examples that show:
'
'     - the normal recommended timing pattern
'     - how to format an already measured elapsed value
'     - how strict mode behaves on invalid usage
'     - how non-strict mode behaves on the same invalid usage
'     - how to benchmark with shared TW suppression
'     - how to structure real-world cleanup safely
'     - how to capture and export structured checkpoints
'     - how to benchmark repeatedly and read the statistics honestly
'     - how to compare two implementations without fooling yourself
'     - how to subtract a dispatch-matched baseline
'     - how to tell a failed read from a genuinely fast operation
'     - what the named enums do and do not protect against
'
'   This module therefore keeps only the examples that still add real teaching
'   value beyond the demo sheets and tests
'
' INPUTS
'   None at module level
'
' RETURNS
'   None at module level
'
' BEHAVIOR
'   - Public launcher procedures run curated example groups
'   - Individual public procedures demonstrate one meaningful usage pattern
'   - Output is written primarily to the Immediate Window
'   - The checkpoint example can also write a structured report to a worksheet
'
' ERROR POLICY
'   Individual examples use local cleanup and re-raise normally unless the
'   example intentionally demonstrates expected invalid usage
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Excel Application object model
'
' MEASUREMENT TARGETS
'   The measurement examples call procedures by name through Application.Run,
'   which reaches only Public procedures in standard modules. The cPM_Usage_*
'   workload procedures at the foot of this module exist for that purpose and
'   are deliberately Public
'
' NOTES
'   - Place this code in a STANDARD MODULE
'   - Results are printed primarily to the Immediate Window
'   - Press Ctrl+G in the VBA editor to open the Immediate Window
'   - Run worksheet-writing examples from a safe workbook / worksheet
'
' UPDATED
'   2026-08-27
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit     'Force explicit declaration of all variables

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const cPM_USAGE_SHEET_DATA           As String = "DATA_cPM"
    Private Const cPM_USAGE_CHECKPOINT_TOPLEFT   As String = "D3"

    'Rows written by the comparison workloads.
    '
    'Both must write the same number or the comparison is meaningless. The value
    'is small deliberately: cell-by-cell costs roughly a millisecond per cell, so
    'ten thousand rows made a ten-iteration comparison take minutes and the
    'variance alone was enough to make the run report as contaminated.
        Private Const cPM_USAGE_WORKLOAD_ROWS        As Long = 500

'
'==============================================================================
'
'                           PUBLIC: MASTER LAUNCHERS
'
'==============================================================================

Public Sub Run_All_UsageExamples()
'
'==============================================================================
'                          RUN ALL USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes all retained usage examples in a logical sequence
'
' WHY THIS EXISTS
'   Provides one compact entry point for reviewing the most meaningful usage
'   patterns without running the full demo workbook or regression suite
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Prints a module-level start banner
'   - Runs all retained example groups in teaching order
'   - Prints a module-level completion banner
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Run_CoreUsageExamples
'   - Run_ValidationUsageExamples
'   - Run_TimeWasterUsageExamples
'   - Run_SafePatternUsageExamples
'   - Run_CheckpointUsageExamples
'   - Run_MeasurementUsageExamples
'
' NOTES
'   This is the best launcher when you want the full compact walkthrough
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT MODULE START BANNER
'------------------------------------------------------------------------------
    'Print the overall suite start banner
        PrintModuleBanner "RUN ALL cPerformanceManager USAGE EXAMPLES - START"

'------------------------------------------------------------------------------
' RUN EXAMPLE GROUPS
'------------------------------------------------------------------------------
    'Run the core usage examples
        Run_CoreUsageExamples
    'Run the validation-behavior examples
        Run_ValidationUsageExamples
    'Run the TW-related usage examples
        Run_TimeWasterUsageExamples
    'Run the structured safe-pattern example
        Run_SafePatternUsageExamples
    'Run the checkpoint/report example
        Run_CheckpointUsageExamples
    'Run the measurement and statistics examples
        Run_MeasurementUsageExamples

'------------------------------------------------------------------------------
' PRINT MODULE END BANNER
'------------------------------------------------------------------------------
    'Print the overall suite completion banner
        PrintModuleBanner "RUN ALL cPerformanceManager USAGE EXAMPLES - COMPLETE"

End Sub

Public Sub Run_CoreUsageExamples()
'
'==============================================================================
'                         RUN CORE USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the core, day-to-day usage demonstrations
'
' WHY THIS EXISTS
'   These are the most immediately useful examples for ordinary adoption of the
'   class in client code
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_BasicTiming_DefaultQPC
'     - Example_ElapsedTime_FromMeasuredSeconds
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_BasicTiming_DefaultQPC
'   - Example_ElapsedTime_FromMeasuredSeconds
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "CORE USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the basic default-QPC timing example
        PrintExampleBanner "Example_BasicTiming_DefaultQPC"
        Example_BasicTiming_DefaultQPC
    'Run the formatting-from-existing-seconds example
        PrintExampleBanner "Example_ElapsedTime_FromMeasuredSeconds"
        Example_ElapsedTime_FromMeasuredSeconds

End Sub

Public Sub Run_ValidationUsageExamples()
'
'==============================================================================
'                      RUN VALIDATION USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the examples that contrast strict and non-strict validation
'
' WHY THIS EXISTS
'   Correct session usage is one of the defining design choices of the class,
'   and these examples show how the caller can choose fail-fast or forgiving
'   behavior
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_StrictMode
'     - Example_NonStrictMode
'     - Example_EnumSemantics
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_StrictMode
'   - Example_NonStrictMode
'   - Example_EnumSemantics
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "VALIDATION USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the strict-mode example
        PrintExampleBanner "Example_StrictMode"
        Example_StrictMode
    'Run the non-strict-mode example
        PrintExampleBanner "Example_NonStrictMode"
        Example_NonStrictMode
    'Run the enum-semantics example
        PrintExampleBanner "Example_EnumSemantics"
        Example_EnumSemantics

End Sub

Public Sub Run_TimeWasterUsageExamples()
'
'==============================================================================
'                      RUN TIME-WASTER USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the TW-related benchmark usage example
'
' WHY THIS EXISTS
'   Shared Excel TW suppression is one of the most practically useful features
'   when benchmarking worksheet or application work
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_TimeWasters_Basic
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_TimeWasters_Basic
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "TIME-WASTER USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the basic TW-suppression example
        PrintExampleBanner "Example_TimeWasters_Basic"
        Example_TimeWasters_Basic

End Sub

Public Sub Run_SafePatternUsageExamples()
'
'==============================================================================
'                    RUN SAFE-PATTERN USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the recommended structured cleanup example
'
' WHY THIS EXISTS
'   This is the most important integration example when using the class in
'   real procedures
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_SafePattern
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_SafePattern
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "SAFE-PATTERN USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the structured cleanup pattern example
        PrintExampleBanner "Example_SafePattern"
        Example_SafePattern

End Sub

Public Sub Run_CheckpointUsageExamples()
'
'==============================================================================
'                     RUN CHECKPOINT USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the structured checkpoint/reporting example
'
' WHY THIS EXISTS
'   Checkpoints are one of the newer high-value surfaces of the class and
'   deserve an explicit standalone example
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_CheckpointReport
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_CheckpointReport
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "CHECKPOINT USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the checkpoint/report example
        PrintExampleBanner "Example_CheckpointReport"
        Example_CheckpointReport

End Sub

'
'==============================================================================
'
'                       PRIVATE: OUTPUT / BANNER HELPERS
'
'==============================================================================

Private Sub PrintModuleBanner( _
    ByVal Title As String)
'
'==============================================================================
'                           PRINT MODULE BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a visually distinct top-level banner to the Immediate Window
'
' WHY THIS EXISTS
'   Improves readability when running groups of examples
'
' INPUTS
'   Title
'     Banner title text
'
' RETURNS
'   None
'
' BEHAVIOR
'   Prints a blank line, a delimiter, the title, and a closing delimiter
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT BANNER
'------------------------------------------------------------------------------
    'Print a blank line before the banner
        Debug.Print vbNullString
    'Print the opening delimiter
        Debug.Print String$(78, "=")
    'Print the banner title
        Debug.Print Title
    'Print the closing delimiter
        Debug.Print String$(78, "=")

End Sub

Private Sub PrintSectionBanner( _
    ByVal Title As String)
'
'==============================================================================
'                           PRINT SECTION BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a section-level banner to the Immediate Window
'
' WHY THIS EXISTS
'   Makes grouped example runs easier to read and review
'
' INPUTS
'   Title
'     Section title text
'
' RETURNS
'   None
'
' BEHAVIOR
'   Prints a blank line, a delimiter, the section title, and a closing
'   delimiter
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT BANNER
'------------------------------------------------------------------------------
    'Print a blank line before the section banner
        Debug.Print vbNullString
    'Print the opening delimiter
        Debug.Print String$(78, "-")
    'Print the section title
        Debug.Print Title
    'Print the closing delimiter
        Debug.Print String$(78, "-")

End Sub

Private Sub PrintExampleBanner( _
    ByVal ProcName As String)
'
'==============================================================================
'                           PRINT EXAMPLE BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a small marker identifying the example about to run
'
' WHY THIS EXISTS
'   Helps the user associate Immediate Window output with the procedure that
'   produced it
'
' INPUTS
'   ProcName
'     Name of the example procedure being executed
'
' RETURNS
'   None
'
' BEHAVIOR
'   Prints one compact marker line
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT BANNER
'------------------------------------------------------------------------------
    'Print the example marker
        Debug.Print ">>> " & ProcName

End Sub

Private Function cPM_Usage_GetDataSheet() As Worksheet
'
'==============================================================================
'                           GET DATA SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the standard DATA_cPM worksheet used by the usage examples
'
' WHY THIS EXISTS
'   Centralizes the workbook-qualified worksheet lookup used by several
'   examples and avoids repeated unqualified Worksheets(...) calls
'
' INPUTS
'   None
'
' RETURNS
'   Worksheet
'     DATA_cPM worksheet from ThisWorkbook
'
' ERROR POLICY
'   Raises errors normally if the expected worksheet is missing
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the standard data sheet from the host workbook
        Set cPM_Usage_GetDataSheet = ThisWorkbook.Worksheets(cPM_USAGE_SHEET_DATA)

End Function

'
'==============================================================================
'
'                     PUBLIC: CORE TIMING USAGE EXAMPLES
'
'==============================================================================

Public Sub Example_BasicTiming_DefaultQPC()
'
'==============================================================================
'                       EXAMPLE: BASIC TIMING (DEFAULT QPC)
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the simplest recommended usage pattern for cPerformanceManager
'
' WHY THIS EXISTS
'   This is the canonical "getting started" example:
'     - instantiate the class
'     - start timing
'     - perform work
'     - read elapsed seconds
'     - clean up deterministically
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a new cPerformanceManager instance
'   - Starts timing with the default method, which is method 5 (QPC)
'   - Writes a constant into a worksheet range
'   - Reads numeric elapsed time in seconds
'   - Prints the result to the Immediate Window
'   - Restores environment state and releases the instance
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This is the preferred example to show the normal benchmark path
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim WS                  As Worksheet              'Target worksheet
    Dim ElapsedS            As Double                 'Elapsed time in seconds

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing with the default method (5 = QPC)
        cPM.StartTimer

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Write a constant value into a worksheet range
        WS.Range("I6:I10006").Value = 7

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Read the numeric elapsed time
        ElapsedS = cPM.ElapsedSeconds
    'Print the measured elapsed seconds
        Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

Public Sub Example_ElapsedTime_FromMeasuredSeconds()
'
'==============================================================================
'               EXAMPLE: ELAPSEDTIME FROM MEASURED SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how to format an already measured elapsed-seconds value
'
' WHY THIS EXISTS
'   In real code, callers often need both:
'     - a numeric elapsed value for comparisons / storage
'     - a formatted elapsed string for reporting
'
'   This example shows the recommended pattern:
'     - measure once with ElapsedSeconds
'     - format that same value through ElapsedTime(, ElapsedSecondsIn)
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Starts timing explicitly with method 5 (QPC)
'   - Performs a calculation workload
'   - Reads numeric elapsed seconds once
'   - Formats that existing value without taking a second timing sample
'   - Prints both results
'   - Cleans up and releases the instance
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Application.Calculate
'   - Debug.Print
'
' NOTES
'   This is the recommended pattern when you want both numeric and
'   display-oriented output without double measurement
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim ElapsedS            As Double                 'Elapsed time in seconds
    Dim ElapsedT            As String                 'Formatted elapsed time

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing explicitly with QPC
        cPM.StartTimer cPM_MethodQPC, False

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Force a workbook/application calculation pass
        Application.Calculate

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Read elapsed seconds once
        ElapsedS = cPM.ElapsedSeconds
    'Format the already measured elapsed-seconds value
        ElapsedT = cPM.ElapsedTime(, ElapsedS)
    'Print the numeric result
        Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")
    'Print the formatted result
        Debug.Print "Elapsed time   : " & ElapsedT

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

'
'==============================================================================
'
'                      PUBLIC: VALIDATION USAGE EXAMPLES
'
'==============================================================================

Public Sub Example_StrictMode()
'
'==============================================================================
'                           EXAMPLE: STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how StrictMode enforces correct method/session usage
'
' WHY THIS EXISTS
'   A key design feature of cPerformanceManager is that elapsed reads are
'   session-bound. In strict mode, an invalid elapsed-method request raises an
'   error rather than being silently coerced
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Enables strict mode
'   - Starts a session using method 5
'   - Intentionally requests elapsed time with method 2
'   - Captures and prints the raised error
'   - Cleans up and releases the instance
'
' ERROR POLICY
'   Uses local expected-error handling to demonstrate the raised error safely
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'   - Err
'
' NOTES
'   This example is intentionally invalid by design
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim Dummy               As Double                 'Throwaway target for the failing call
    Dim ExpectedErrNum      As Long                   'Captured expected error number
    Dim ExpectedErrDesc     As String                 'Captured expected error description

    Dim SavedErrNumber      As Long                   'Captured unexpected error number
    Dim SavedErrSource      As String                 'Captured unexpected error source
    Dim SavedErrDescription As String                 'Captured unexpected error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE
'------------------------------------------------------------------------------
    'Enable strict validation behavior
        cPM.StrictMode = True
    'Start timing with method 5
        cPM.StartTimer cPM_MethodQPC, False

'------------------------------------------------------------------------------
' TRIGGER INTENTIONAL INVALID USAGE
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next
    'This is intentionally invalid because the active session uses method 5
        Dummy = cPM.ElapsedSeconds(cPM_MethodTickCount)
    'Capture the expected error information
        ExpectedErrNum = Err.Number
        ExpectedErrDesc = Err.Description
    'Clear the local expected-error state
        Err.Clear
    'Restore normal error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' PRINT RESULT
'------------------------------------------------------------------------------
    'Print the captured expected error number
        Debug.Print "Error number: " & ExpectedErrNum
    'Print the captured expected error text
        Debug.Print "Error text  : " & ExpectedErrDesc

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original unexpected error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original unexpected error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

Public Sub Example_NonStrictMode()
'
'==============================================================================
'                         EXAMPLE: NON-STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how the class behaves in non-strict mode when the caller asks
'   for an elapsed method that does not match the active session
'
' WHY THIS EXISTS
'   Non-strict mode is the forgiving mode of the class. Instead of raising, the
'   class may coerce the request and continue using the active session method
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Disables strict mode
'   - Starts a session with method 5
'   - Waits briefly
'   - Requests elapsed time using the wrong method identifier
'   - Prints the returned value and active method
'   - Cleans up and releases the instance
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This example contrasts directly with Example_StrictMode
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim ElapsedS            As Double                 'Elapsed time returned after fallback

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE
'------------------------------------------------------------------------------
    'Disable strict validation behavior
        cPM.StrictMode = False
    'Start timing with method 5
        cPM.StartTimer cPM_MethodQPC, False

'------------------------------------------------------------------------------
' APPLY SMALL DELAY
'------------------------------------------------------------------------------
    'Pause briefly to make the elapsed reading visible
        cPM.Pause 0.03, cPM_PauseSleep

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'In non-strict mode this falls back to the active session method
        ElapsedS = cPM.ElapsedSeconds(cPM_MethodTickCount)
    'Print the active method identifier
        Debug.Print "ActiveMethodID : " & cPM.ActiveMethodID
    'Print the active method name
        Debug.Print "MethodName     : " & cPM.MethodName(cPM.ActiveMethodID)
    'Print the returned elapsed time
        Debug.Print "ElapsedSeconds : " & Format$(ElapsedS, "0.000000000")

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

Public Sub Example_EnumSemantics()
'
'==============================================================================
'                          EXAMPLE: ENUM SEMANTICS
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates what the named enums do and do not protect against
'
' WHY THIS EXISTS
'   cPM_TimerMethod and cPM_PauseMethod share the numbers 1 to 4 and mean
'   entirely different things. It is tempting to conclude that giving them
'   separate types stops one being passed where the other belongs
'
'   It does not. VBA enum members and enum-typed parameters have Long
'   semantics, so a constant from either enum is accepted wherever the other is
'   expected. The call compiles, runs, and quietly does something the reader did
'   not intend
'
'   This example makes that visible, because a boundary the caller cannot see is
'   a boundary the caller will eventually cross
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Shows the intended, readable call for each enum
'   - Passes a timer constant to Pause and reports which strategy actually ran
'   - Passes a pause constant to StartTimer and reports which backend was bound
'   - States what the caller can rely on instead
'
' ERROR POLICY
'   Reports and continues through centralized cleanup
'
' DEPENDENCIES
'   - cPerformanceManager
'
' NOTES
'   Nothing here is a defect in the class, and nothing here should be copied
'   into working code. The enums are still worth using: they make call sites
'   legible and drive IntelliSense. What they cannot do is enforce a role at
'   compile time
'
'   The class does validate at run time. An out-of-range value is normalized or
'   rejected. What it cannot detect is a value that is in range for the
'   parameter but came from the wrong enum, because by then the two are
'   indistinguishable
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Performance manager instance
    Dim BoundMethod As cPM_TimerMethod        'Backend actually bound by StartTimer

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a fresh performance manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' THE INTENDED CALLS
'------------------------------------------------------------------------------
    'Each enum used in its own role. This is how the API is meant to read
        cPM.StartTimer cPM_MethodQPC
        cPM.Pause 0.01, cPM_PauseSleep
        Debug.Print "Intended usage     : backend " & CLng(cPM.ActiveMethodID) & _
                    " (" & cPM.MethodName(cPM.ActiveMethodID) & "), Sleep pause"
        cPM.ResetEnvironment

'------------------------------------------------------------------------------
' A TIMER CONSTANT PASSED TO PAUSE
'------------------------------------------------------------------------------
    'cPM_MethodTickCount is 2. So is cPM_PauseTimerLoop. Pause therefore runs
    'the Timer-loop strategy, not anything to do with GetTickCount, and no
    'error is raised because the value was always in range for the parameter
        cPM.StartTimer cPM_MethodQPC
        cPM.Pause 0.01, cPM_MethodTickCount
        Debug.Print "Wrong-role pause   : compiled and ran. cPM_MethodTickCount is 2,"
        Debug.Print "                     so Pause used strategy 2 (Timer loop)."
        cPM.ResetEnvironment

'------------------------------------------------------------------------------
' A PAUSE CONSTANT PASSED TO STARTTIMER
'------------------------------------------------------------------------------
    'The same confusion in the other direction. cPM_PauseAppWait is 3, so the
    'session binds backend 3 (timeGetTime) rather than the default QPC the
    'reader of this line would probably assume
        cPM.StartTimer cPM_PauseAppWait
        BoundMethod = cPM.ActiveMethodID
        Debug.Print "Wrong-role start   : bound backend " & CLng(BoundMethod) & _
                    " (" & cPM.MethodName(BoundMethod) & "),"
        Debug.Print "                     because cPM_PauseAppWait is 3."
        cPM.ResetEnvironment

'------------------------------------------------------------------------------
' THE POINT
'------------------------------------------------------------------------------
    'Say plainly what the contract is, so nobody infers a stronger one
        Debug.Print "Note: the enums document intent and drive IntelliSense."
        Debug.Print "      They are not a compile-time type boundary."
        Debug.Print "      Pass the enum that matches the parameter, and read"
        Debug.Print "      ActiveMethodID back when the bound backend matters."

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Report and continue through centralized cleanup
        Debug.Print "Error " & Err.Number & " - " & Err.Description
        Resume CleanExit

End Sub

'
'==============================================================================
'
'               PUBLIC: SHARED TIME-WASTER SUPPRESSION EXAMPLES
'
'==============================================================================

Public Sub Example_TimeWasters_Basic()
'
'==============================================================================
'                     EXAMPLE: TIMEWASTERS (BASIC)
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates basic shared Excel "time-waster" suppression for a benchmark
'   run
'
' WHY THIS EXISTS
'   Excel application behaviors such as ScreenUpdating, events, alerts,
'   calculation mode, and cursor changes can add noise to benchmarks badly.
'   This example shows the normal suppression pattern
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Turns off all supported TW settings for this shared scope
'   - Starts QPC timing
'   - Performs a worksheet workload
'   - Prints elapsed seconds
'   - Ends the TW session
'   - Cleans up and releases the instance
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   TW control is shared/global in effect, so this relies on the shared
'   manager model rather than direct instance-local restoration
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim WS                  As Worksheet              'Target worksheet
    Dim ElapsedS            As Double                 'Measured elapsed seconds

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' SUPPRESS TIME-WASTERS
'------------------------------------------------------------------------------
    'Disable all supported TW settings for this instance's shared session
        cPM.TW_Turn_OFF

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing with QPC
        cPM.StartTimer cPM_MethodQPC, False

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Execute a worksheet workload for benchmarking
        WS.Range("A1:A50000").Formula = "=ROW()"

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Read the numeric elapsed seconds
        ElapsedS = cPM.ElapsedSeconds
    'Print elapsed seconds with TW suppression in effect
        Debug.Print "Elapsed seconds with TW off: " & Format$(ElapsedS, "0.000000000")

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

'
'==============================================================================
'
'                     PUBLIC: RECOMMENDED SAFETY PATTERN
'
'==============================================================================

Public Sub Example_SafePattern()
'
'==============================================================================
'                         EXAMPLE: SAFE CLEANUP PATTERN
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the recommended structured pattern for using
'   cPerformanceManager safely in real procedures
'
' WHY THIS EXISTS
'   Benchmarks and TW suppression can modify environment state. A structured
'   cleanup block ensures that environment restoration still happens when the
'   workload raises an error
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Starts shared TW suppression
'   - Starts QPC timing
'   - Performs a workload
'   - Prints the formatted elapsed time
'   - Uses cleanup labels to ensure ResetEnvironment and object release happen
'     even if an error occurs
'
' ERROR POLICY
'   Uses structured local error handling with CleanFail / CleanExit labels
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This is the best example to follow when integrating the class into real
'   project code
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim WS                  As Worksheet              'Target worksheet

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route runtime failures to the cleanup-aware failure block
        On Error GoTo CleanFail
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE BENCHMARK ENVIRONMENT
'------------------------------------------------------------------------------
    'Start shared TW suppression for this instance
        cPM.TW_Turn_OFF
    'Start timing with QPC
        cPM.StartTimer cPM_MethodQPC, False

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Execute a worksheet workload
        WS.UsedRange.Calculate

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Print the formatted elapsed-time report
        Debug.Print "Elapsed: " & cPM.ElapsedTime

CleanExit:
'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release environment changes and the instance if the object exists
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If
    'Exit normally after cleanup
        Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' FAILURE EXIT
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Print the error information for diagnostics
        Debug.Print "Error " & SavedErrNumber & " - " & SavedErrDescription
    'Always route through the normal cleanup block
        Resume CleanExit

End Sub

Public Sub Example_CheckpointReport()
'
'==============================================================================
'                       EXAMPLE: CHECKPOINT REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates structured checkpoint capture and export
'
' WHY THIS EXISTS
'   Checkpoints are useful when one elapsed value is not enough and the caller
'   wants structured sub-measurements such as load / write / recalculate
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a fresh performance manager instance
'   - Assigns a run label
'   - Starts a fresh QPC timing session
'   - Runs three simple benchmark phases
'   - Captures a checkpoint after each phase
'   - Prints the text report to the Immediate Window
'   - Writes the array report to DATA_cPM starting at D3
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'   - ReportAsText
'   - ReportAsArray
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim WS                  As Worksheet              'Target worksheet
    Dim ReportArr           As Variant                'Structured report array
    Dim TopLeft             As Range                  'Top-left output anchor

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Resolve the top-left report output anchor
        Set TopLeft = WS.Range(cPM_USAGE_CHECKPOINT_TOPLEFT)
    'Create a fresh performance manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' INITIALIZE TIMING SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session, labelling it in the same call.
    '
    'Setting the label BEFORE StartTimer does not work: starting a session
    'clears the run label, so the report would come out unlabelled.
        cPM.StartTimer cPM_MethodQPC, False, "ImportWorkflow" 

'------------------------------------------------------------------------------
' CHECKPOINTED WORKLOAD
'------------------------------------------------------------------------------
    'Simulate phase 1
        WS.Range("A1:A10000").Value = 1
    'Capture the first checkpoint
        cPM.Checkpoint "LoadValues"
    'Simulate phase 2
        WS.Range("B1:B10000").Formula = "=ROW()"
    'Capture the second checkpoint
        cPM.Checkpoint "WriteFormulas"
    'Simulate phase 3
        Application.Calculate
    'Capture the third checkpoint
        cPM.Checkpoint "Recalculate", "Full workbook calculation pass"

'------------------------------------------------------------------------------
' OUTPUT TO IMMEDIATE WINDOW
'------------------------------------------------------------------------------
    'Print the readable checkpoint report
        Debug.Print cPM.ReportAsText

'------------------------------------------------------------------------------
' OUTPUT TO WORKSHEET
'------------------------------------------------------------------------------
    'Resolve the structured report array
        ReportArr = cPM.ReportAsArray
    'Clear the target write area sized to the current report
        TopLeft.Resize(UBound(ReportArr, 1), UBound(ReportArr, 2)).ClearContents
    'Write the structured report array to the worksheet
        TopLeft.Resize(UBound(ReportArr, 1), UBound(ReportArr, 2)).Value = ReportArr

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release environment changes
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

Public Sub Run_MeasurementUsageExamples()
'
'==============================================================================
'                      RUN MEASUREMENT USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the repeated-measurement and statistics demonstrations
'
' WHY THIS EXISTS
'   A single timing sample tells you what happened once. It cannot tell you
'   whether the number is stable, whether the machine was busy, or whether you
'   would get the same answer again
'
'   These examples show the surface that answers those questions, and how to
'   read its output without fooling yourself
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_Benchmark_WithStatistics
'     - Example_Compare_TwoImplementations
'     - Example_DispatchMatchedBaseline
'     - Example_ReadStatus_Diagnostics
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_Benchmark_WithStatistics
'   - Example_Compare_TwoImplementations
'   - Example_DispatchMatchedBaseline
'   - Example_ReadStatus_Diagnostics
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "MEASUREMENT AND STATISTICS USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the repeated-measurement example
        PrintExampleBanner "Example_Benchmark_WithStatistics"
        Example_Benchmark_WithStatistics
    'Run the comparison example
        PrintExampleBanner "Example_Compare_TwoImplementations"
        Example_Compare_TwoImplementations
    'Run the dispatch-matched baseline example
        PrintExampleBanner "Example_DispatchMatchedBaseline"
        Example_DispatchMatchedBaseline
    'Run the read-status diagnostics example
        PrintExampleBanner "Example_ReadStatus_Diagnostics"
        Example_ReadStatus_Diagnostics

End Sub

Public Sub Example_Benchmark_WithStatistics()
'
'==============================================================================
'                     EXAMPLE BENCHMARK WITH STATISTICS
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates repeated measurement of a procedure and how to read the result
'
' WHY THIS EXISTS
'   Timing distributions are right-skewed. Work has a floor - the time it
'   genuinely takes - but no ceiling, because the operating system can interrupt
'   at any moment
'
'   That shape has a practical consequence: the arithmetic mean is the worst
'   summary statistic for timing work, because a single scheduler preemption
'   shifts it while barely moving the median
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Measures a workload procedure thirty times, discarding three warm-up runs
'   - Prints the full statistics summary
'   - Checks explicitly whether the run is trustworthy
'
' ERROR POLICY
'   Raises errors normally after cleanup
'
' DEPENDENCIES
'   - cPerformanceManager
'   - cPM_Usage_Workload_BulkArray
'
' NOTES
'   Read the median and the minimum. The minimum is the run least disturbed by
'   scheduling and is the closest estimate of the true cost
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Performance manager instance
    Dim Samples()   As Double                 'Per-run elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a fresh performance manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Thirty measured runs after three discarded warm-up runs
        Samples = cPM.MeasureProcedure("cPM_Usage_Workload_BulkArray", 30, 3)

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'The summary reports median and minimum first, and the mean last
        Debug.Print cPM.Stats_Text(Samples, "Bulk array write")

    'Ask explicitly whether the result can be trusted
        If cPM.Stats_IsContaminated(Samples) Then
            Debug.Print ">>> Variance is high. Close background work and run again."
        Else
            Debug.Print ">>> Run looks clean. Use the median."
        End If

'------------------------------------------------------------------------------
' USE THE NUMBERS
'------------------------------------------------------------------------------
    'The raw vector is yours, so any further analysis is possible
        Debug.Print "Typical cost : " & Format$(cPM.Stats_Median(Samples), "0.000000000") & " s"
        Debug.Print "Best case    : " & Format$(cPM.Stats_Min(Samples), "0.000000000") & " s"
        Debug.Print "Tail (P95)   : " & Format$(cPM.Stats_Percentile(Samples, 95#), "0.000000000") & " s"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Report and continue through centralized cleanup
        Debug.Print "Error " & Err.Number & " - " & Err.Description
        Resume CleanExit

End Sub

Public Sub Example_Compare_TwoImplementations()
'
'==============================================================================
'                    EXAMPLE COMPARE TWO IMPLEMENTATIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates a defensible comparison between two implementations
'
' WHY THIS EXISTS
'   Comparing two single measurements proves nothing. Comparing two means proves
'   slightly less than nothing, because it compounds the contamination in each
'
'   This example shows the discipline that makes a speedup claim survive
'   scrutiny: repeated runs, medians, and a contamination check on both sides
'   before drawing any conclusion
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Measures two implementations of the same task, back to back
'   - Reports both summaries
'   - Refuses to state a speedup if either run looks contaminated
'
' ERROR POLICY
'   Raises errors normally after cleanup
'
' DEPENDENCIES
'   - cPerformanceManager
'   - cPM_Usage_Workload_CellByCell
'   - cPM_Usage_Workload_BulkArray
'
' NOTES
'   Both runs must happen on the same machine in the same session. Yesterday's
'   number compared against today's is not a comparison
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Performance manager instance
    Dim Slow()      As Double                 'Samples for the cell-by-cell approach
    Dim Fast()      As Double                 'Samples for the bulk-array approach

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a fresh performance manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' MEASURE BOTH, BACK TO BACK
'------------------------------------------------------------------------------
    'Measure the two implementations under identical conditions
        Slow = cPM.MeasureProcedure("cPM_Usage_Workload_CellByCell", 10, 2)
        Fast = cPM.MeasureProcedure("cPM_Usage_Workload_BulkArray", 10, 2)

'------------------------------------------------------------------------------
' REPORT BOTH
'------------------------------------------------------------------------------
    'Show the full picture for each rather than a single number
        Debug.Print cPM.Stats_Text(Slow, "Cell by cell")
        Debug.Print cPM.Stats_Text(Fast, "Bulk array")

'------------------------------------------------------------------------------
' CONCLUDE ONLY IF BOTH RUNS ARE TRUSTWORTHY
'------------------------------------------------------------------------------
    'A speedup computed from a contaminated run is not evidence of anything
        If cPM.Stats_IsContaminated(Slow) Or cPM.Stats_IsContaminated(Fast) Then
            Debug.Print ">>> One or both runs are noisy. No conclusion drawn."
            GoTo CleanExit
        End If

    'Compare medians, never means
        Debug.Print "Median speedup: " & _
                    Format$(cPM.Stats_Median(Slow) / cPM.Stats_Median(Fast), "0.00") & "x"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Report and continue through centralized cleanup
        Debug.Print "Error " & Err.Number & " - " & Err.Description
        Resume CleanExit

End Sub

Public Sub Example_DispatchMatchedBaseline()
'
'==============================================================================
'                    EXAMPLE DISPATCH MATCHED BASELINE
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates subtracting the harness cost from a measured workload
'
' WHY THIS EXISTS
'   MeasureProcedure calls its target through Application.Run, and that dispatch
'   has a cost which is included in every sample. For work measured in
'   milliseconds it is negligible; for very fast work it is not
'
'   MeasureOverhead_Samples cannot supply the correction, because it measures
'   the timing cycle only and never dispatches. Subtracting it removes the wrong
'   quantity and leaves the dispatch cost in place
'
'   MeasureBaseline runs an empty procedure through the identical path, so its
'   median is a baseline that can legitimately be subtracted
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Measures an empty procedure and a real workload through the same path
'   - Reports both, then the net cost
'
' ERROR POLICY
'   Raises errors normally after cleanup
'
' DEPENDENCIES
'   - cPerformanceManager
'   - cPM_Usage_BaselineEmpty
'   - cPM_Usage_Workload_BulkArray
'
' NOTES
'   Subtract medians, not means. Both vectors are right-skewed, and subtracting
'   two means compounds the contamination in each
'
'   The baseline is machine-specific and worth re-measuring on each host
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Performance manager instance
    Dim Baseline()  As Double                 'Dispatch-only cost
    Dim Workload()  As Double                 'Dispatch plus real work
    Dim NetSeconds  As Double                 'Workload cost with dispatch removed

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a fresh performance manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' MEASURE THE BASELINE AND THE WORKLOAD
'------------------------------------------------------------------------------
    'An empty procedure through the same dispatch path
        Baseline = cPM.MeasureBaseline("cPM_Usage_BaselineEmpty", 20, 3)
    'The real workload through that same path
        Workload = cPM.MeasureProcedure("cPM_Usage_Workload_BulkArray", 20, 3)

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Show what the harness itself costs on this machine
        Debug.Print "Dispatch baseline : " & _
                    Format$(cPM.Stats_Median(Baseline), "0.000000000") & " s"
        Debug.Print "Measured workload : " & _
                    Format$(cPM.Stats_Median(Workload), "0.000000000") & " s"

    'Subtract medians to get the net cost of the work itself
        NetSeconds = cPM.Stats_Median(Workload) - cPM.Stats_Median(Baseline)
        Debug.Print "Net workload cost : " & Format$(NetSeconds, "0.000000000") & " s"

'------------------------------------------------------------------------------
' INTERPRET
'------------------------------------------------------------------------------
    'Say plainly when the correction does not matter
        If cPM.Stats_Median(Baseline) < (cPM.Stats_Median(Workload) / 100#) Then
            Debug.Print ">>> Dispatch is under 1% of the measurement. Correction optional."
        Else
            Debug.Print ">>> Dispatch is a material share of the measurement. Subtract it."
        End If

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Report and continue through centralized cleanup
        Debug.Print "Error " & Err.Number & " - " & Err.Description
        Resume CleanExit

End Sub

Public Sub Example_ReadStatus_Diagnostics()
'
'==============================================================================
'                     EXAMPLE READ STATUS DIAGNOSTICS
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates telling a failed read apart from a genuinely fast operation
'
' WHY THIS EXISTS
'   In non-strict mode a failed timing read returns zero. Without a status, a
'   caller cannot distinguish "this took no measurable time" from "no reading was
'   obtained at all"
'
'   The second silently poisons any aggregate built from it, which is why the
'   distinction is exposed rather than left to inference
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Takes an ordinary measurement and reports its status
'   - Runs the harness and reports how many of its reads failed
'   - Explains why the two use different surfaces
'
' ERROR POLICY
'   Raises errors normally after cleanup
'
' DEPENDENCIES
'   - cPerformanceManager
'   - cPM_Usage_Workload_BulkArray
'
' NOTES
'   LastReadStatus describes the reads made by THIS instance. The harness
'   measures on an isolated worker that is released before the sample vector is
'   returned, so its outcome arrives through the FailedReadsOut argument instead
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Performance manager instance
    Dim ElapsedS    As Double                 'Directly measured elapsed seconds
    Dim Samples()   As Double                 'Harness samples
    Dim FailedReads As Long                   'Harness reads that did not succeed
    Dim LastStatus  As cPM_ReadStatus         'Status of the last harness failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a fresh performance manager instance
        Set cPM = New cPerformanceManager
    'Non-strict mode returns zero instead of raising, which is the case worth showing
        cPM.StrictMode = False

'------------------------------------------------------------------------------
' A DIRECT MEASUREMENT
'------------------------------------------------------------------------------
    'Measure something on this instance
        cPM.StartTimer cPM_MethodQPC
        cPM.Pause 0.02, cPM_PauseSleep
        ElapsedS = cPM.ElapsedSeconds

    'A returned zero means nothing on its own; the status is what qualifies it
        If cPM.LastReadStatus = cPM_ReadOK Then
            Debug.Print "Direct measurement : " & Format$(ElapsedS, "0.000000000") & " s (valid)"
        Else
            Debug.Print "Direct measurement : no valid reading, status " & _
                        CLng(cPM.LastReadStatus)
        End If

'------------------------------------------------------------------------------
' A HARNESS RUN
'------------------------------------------------------------------------------
    'The harness reports its own outcome, because it measures on a worker
    'instance that no longer exists by the time the samples come back
        Samples = cPM.MeasureProcedure("cPM_Usage_Workload_BulkArray", 10, 2, _
                                       cPM_MethodQPC, FailedReads, LastStatus)

        Debug.Print "Harness samples    : " & cPM.Stats_Count(Samples)
        Debug.Print "Failed reads       : " & FailedReads

    'Only trust the statistics if every read produced a measurement
        If FailedReads = 0 Then
            Debug.Print "Median             : " & _
                        Format$(cPM.Stats_Median(Samples), "0.000000000") & " s"
        Else
            Debug.Print ">>> " & FailedReads & " read(s) failed, last status " & _
                        CLng(LastStatus) & ". The samples are not trustworthy."
        End If

'------------------------------------------------------------------------------
' THE POINT
'------------------------------------------------------------------------------
    'Make the distinction explicit rather than leaving it implied
        Debug.Print "Note: LastReadStatus covers this instance's own reads."
        Debug.Print "      Harness reads are reported through FailedReadsOut."

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Report and continue through centralized cleanup
        Debug.Print "Error " & Err.Number & " - " & Err.Description
        Resume CleanExit

End Sub

'
'==============================================================================
'
'                       PUBLIC: MEASUREMENT TARGETS
'
'==============================================================================
'
'   Application.Run reaches only Public procedures in standard modules, so these
'   are deliberately Public. They exist to be measured, not to be called
'   directly.
'

Public Sub cPM_Usage_BaselineEmpty()
'
'==============================================================================
'                        CPM USAGE BASELINE EMPTY
'------------------------------------------------------------------------------
' PURPOSE
'   Empty procedure used as the target of MeasureBaseline
'
' WHY THIS EXISTS
'   A dispatch-matched baseline needs a procedure that does nothing, reached
'   through the same Application.Run path as a real workload
'
'   The component does not ship one. Application.Run cannot reach a procedure in
'   a module declared Option Private Module, which the companion TW module uses
'   to keep its internals out of the Macro dialog, so a bundled empty procedure
'   would require a third file containing two lines
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-08-16
'==============================================================================

    'Deliberately empty

End Sub

Public Sub cPM_Usage_Workload_BulkArray()
'
'==============================================================================
'                      CPM USAGE WORKLOAD BULK ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Writes cPM_USAGE_WORKLOAD_ROWS values to a worksheet in a single assignment
'
' WHY THIS EXISTS
'   Serves as the faster half of the comparison example, and as a general
'   workload for the measurement examples
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' DEPENDENCIES
'   - cPM_Usage_GetDataSheet
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS      As Worksheet    'Target worksheet
    Dim Values  As Variant      'Values to write
    Dim i       As Long         'Loop index

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Build the whole block in memory first
        ReDim Values(1 To cPM_USAGE_WORKLOAD_ROWS, 1 To 1)
        For i = 1 To cPM_USAGE_WORKLOAD_ROWS
            Values(i, 1) = i
        Next i

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'One assignment for the entire range
        WS.Range("H1").Resize(cPM_USAGE_WORKLOAD_ROWS, 1).Value = Values

End Sub

Public Sub cPM_Usage_Workload_CellByCell()
'
'==============================================================================
'                     CPM USAGE WORKLOAD CELL BY CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Writes cPM_USAGE_WORKLOAD_ROWS values to a worksheet one cell at a time
'
' WHY THIS EXISTS
'   Serves as the slower half of the comparison example. The difference between
'   this and the bulk-array version is the classic VBA performance lesson, and
'   it is large enough to survive measurement noise
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' DEPENDENCIES
'   - cPM_Usage_GetDataSheet
'
' NOTES
'   Deliberately inefficient. Do not use this as a model for real code
'
' UPDATED
'   2026-08-16
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS      As Worksheet    'Target worksheet
    Dim i       As Long         'Loop index

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'One assignment per cell, which is what makes this slow
        For i = 1 To cPM_USAGE_WORKLOAD_ROWS
            WS.Cells(i, 9).Value = i
        Next i

End Sub
