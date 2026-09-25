Attribute VB_Name = "XLSprintTimer"
' XLSprint VBA timer, bounded trace buffer and drill-down driver.
' SPDX-License-Identifier: MIT
'
' The trace contract (schema, span kinds, attribute allow-list) is
' docs/DESIGN.md in the XLSprint repository. This module is imported into a
' disposable copy of the workbook by xlsprint/runner.py and is never saved.
'
' XSP_MicroTimer is adapted from Microsoft's MicroTimer in "Excel improving
' calculation performance" (MicrosoftDocs/VBA-Docs), sample code licensed
' under the MIT License, Copyright (c) Microsoft Corporation.
'
' Rules this module keeps:
'   * Never writes cell values, formula text or string constants. Names are
'     structural identifiers only (sheet, address, defined name, procedure).
'   * Non-interactive: no dialogs of any kind.
'   * Timestamps are raw QueryPerformanceCounter ticks captured as the last
'     statement of a begin and the first statement of an end; conversion to
'     nanoseconds and JSON serialisation happen only in XSP_Flush, outside
'     every timed region.
'   * A mismatched XSP_End is recorded as a fault and never repaired.
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function XSP_QPF Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function XSP_QPC Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
Private Declare Function XSP_QPF Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function XSP_QPC Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If

Public Const XSP_MAX_EVENTS As Long = 100000
Public Const XSP_VBA_VERSION As String = "xlsprint-vba/1"

Private Const XSP_MAX_DEPTH As Long = 64
Private Const XSP_MAX_ATTR_LEN As Long = 256
Private Const XSP_ERR_PLAN As Long = 513

Private Const EV_B As Byte = 1
Private Const EV_E As Byte = 2
Private Const EV_M As Byte = 3

Private Const ST_OK As Byte = 0
Private Const ST_ERROR As Byte = 1
Private Const ST_ABORTED As Byte = 2

' Excel constants, spelled out so the module does not depend on references.
Private Const XL_CALC_MANUAL As Long = -4135
Private Const XL_CALC_AUTOMATIC As Long = -4105
Private Const XL_CALC_SEMIAUTOMATIC As Long = 2

' ---- session state -------------------------------------------------------
Private m_active As Boolean
Private m_mode As String
Private m_outPath As String
Private m_pass As Long
Private m_nextId As Long
Private m_rootParent As Long
Private m_rootDepth As Long
Private m_max As Long
Private m_count As Long
Private m_dropped As Long
Private m_faults As Long
Private m_attrRejected As Long
Private m_firstFault As String
Private m_freq As Currency

' ---- event buffer (parallel arrays, 1-based) ------------------------------
Private m_evType() As Byte
Private m_evId() As Long
Private m_evParent() As Long
Private m_evDepth() As Long
Private m_evPass() As Long
Private m_evTicks() As Currency
Private m_evKind() As String
Private m_evName() As String
Private m_evStatus() As Byte
Private m_evAttrs() As String

' ---- open-span stack ------------------------------------------------------
Private m_stack(1 To XSP_MAX_DEPTH) As Long
Private m_sp As Long

' ---- saved application settings ------------------------------------------
Private m_saved As Boolean
Private m_savedCalc As Long
Private m_savedIter As Boolean
Private m_savedScreen As Boolean

' ===========================================================================
' Clock
' ===========================================================================

' Microsoft's MicroTimer: seconds from QueryPerformanceCounter.
Public Function XSP_MicroTimer() As Double
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    XSP_MicroTimer = 0
    ' Get frequency.
    If cyFrequency = 0 Then XSP_QPF cyFrequency
    ' Get ticks.
    XSP_QPC cyTicks1
    ' Seconds
    If cyFrequency Then XSP_MicroTimer = cyTicks1 / cyFrequency
End Function

' Current MicroTimer time in integer nanoseconds, returned as a decimal
' string so no precision is lost crossing COM. The host parses it with int().
Public Function XSP_ClockProbe() As String
    Dim t As Currency
    XSP_QPC t
    If m_freq = 0 Then XSP_QPF m_freq
    XSP_ClockProbe = TicksToNs(t)
End Function

' ns = floor(ticks * 1e9 / frequency), computed in Decimal. Both Currency
' values carry the same 1/10000 scale, which cancels. Monotonic in ticks.
Private Function TicksToNs(ByVal t As Currency) As String
    TicksToNs = CStr(Int(CDec(t) * CDec(1000000000) / CDec(m_freq)))
End Function

Public Function XSP_Version() As String
    XSP_Version = XSP_VBA_VERSION
End Function

Public Function XSP_Bitness() As Long
#If Win64 Then
    XSP_Bitness = 64
#Else
    XSP_Bitness = 32
#End If
End Function

' ===========================================================================
' Session
' ===========================================================================

' Starts a recording session that XSP_Flush later writes to outPath.
' mode: "on" records every span; "off" records only run.pass, calc.full and
' calc.recalc. VBA ids start at firstId. Root spans get parent parentId and
' depth parentDepth + 1 (depth 0 when parentId is 0). Returns "ok" or
' "error:<reason>".
Public Function XSP_Init(ByVal outPath As String, ByVal mode As String, ByVal firstId As Long, _
        ByVal parentId As Long, ByVal parentDepth As Long, _
        Optional ByVal maxEvents As Long = XSP_MAX_EVENTS) As String
    On Error GoTo Fail
    m_active = False
    If mode <> "on" And mode <> "off" Then XSP_Init = "error:mode": Exit Function
    If firstId < 1 Or parentId < 0 Or parentDepth < 0 Or maxEvents < 1 Then
        XSP_Init = "error:arguments": Exit Function
    End If
    If Len(outPath) = 0 Then XSP_Init = "error:path": Exit Function

    m_outPath = outPath
    m_mode = mode
    m_pass = 0
    m_nextId = firstId
    m_rootParent = parentId
    If parentId > 0 Then m_rootDepth = parentDepth + 1 Else m_rootDepth = 0
    m_max = maxEvents
    m_count = 0
    m_dropped = 0
    m_faults = 0
    m_attrRejected = 0
    m_firstFault = ""
    m_sp = 0

    ReDim m_evType(1 To m_max)
    ReDim m_evId(1 To m_max)
    ReDim m_evParent(1 To m_max)
    ReDim m_evDepth(1 To m_max)
    ReDim m_evPass(1 To m_max)
    ReDim m_evTicks(1 To m_max)
    ReDim m_evKind(1 To m_max)
    ReDim m_evName(1 To m_max)
    ReDim m_evStatus(1 To m_max)
    ReDim m_evAttrs(1 To m_max)

    m_freq = 0
    XSP_QPF m_freq
    If m_freq = 0 Then XSP_Init = "error:no_qpc": Exit Function

    m_active = True
    XSP_Init = "ok"
    Exit Function
Fail:
    m_active = False
    XSP_Init = "error:" & CStr(Err.Number)
End Function

' ===========================================================================
' Public instrumentation API for user procedures
' ===========================================================================

' Opens a span. Returns its id, or 0 when nothing is recorded: no session is
' active, the session is "off", or the kind is not allowed here. Pass the
' returned value to XSP_End, including 0. Only "vba.proc" is accepted from
' user code; calc.* spans come from the drill-down driver alone. Pass a
' procedure name as spanName, never data.
Public Function XSP_Begin(ByVal kind As String, ByVal spanName As String) As Long
    On Error GoTo Fail
    If Not m_active Then Exit Function
    If kind = "vba.proc" Then
        XSP_Begin = SpanBegin(kind, spanName, "")
    Else
        Fault "begin_kind_not_allowed"
    End If
    Exit Function
Fail:
    Fault "begin_error_" & CStr(Err.Number)
    XSP_Begin = 0
End Function

' Closes the innermost open span. A different id is a fault: it is counted,
' reported in the footer, and the stack is left as it is.
Public Sub XSP_End(ByVal id As Long)
    On Error GoTo Fail
    SpanEnd id, ST_OK, ""
    Exit Sub
Fail:
    Fault "end_error_" & CStr(Err.Number)
End Sub

' Records an instantaneous marker (trace-on only).
Public Sub XSP_Marker(ByVal markerName As String)
    Dim t As Currency, i As Long
    On Error GoTo Fail
    XSP_QPC t
    If Not m_active Then Exit Sub
    If m_mode <> "on" Then Exit Sub
    If m_count >= m_max Then m_dropped = m_dropped + 1: m_nextId = m_nextId + 1: Exit Sub
    m_count = m_count + 1
    i = m_count
    m_evType(i) = EV_M
    m_evId(i) = m_nextId
    m_nextId = m_nextId + 1
    If m_sp > 0 Then m_evParent(i) = m_stack(m_sp) Else m_evParent(i) = m_rootParent
    m_evDepth(i) = m_rootDepth + m_sp
    m_evPass(i) = m_pass
    m_evKind(i) = "marker"
    m_evName(i) = markerName
    m_evStatus(i) = ST_OK
    m_evAttrs(i) = ""
    m_evTicks(i) = t
    Exit Sub
Fail:
    Fault "marker_error_" & CStr(Err.Number)
End Sub

' ===========================================================================
' Span core
' ===========================================================================

Private Function ShouldRecord(ByVal kind As String) As Boolean
    If m_mode = "on" Then
        ShouldRecord = True
    Else
        ShouldRecord = (kind = "run.pass" Or kind = "calc.full" Or kind = "calc.recalc")
    End If
End Function

' Returns the span id, or 0 when the span is not tracked. When the buffer is
' full the span is still tracked on the stack (so ends keep matching) but its
' events are counted as dropped.
Private Function SpanBegin(ByVal kind As String, ByVal nm As String, ByVal attrs As String) As Long
    Dim id As Long, i As Long, t As Currency
    If Not m_active Then Exit Function
    If Not ShouldRecord(kind) Then Exit Function
    If m_sp >= XSP_MAX_DEPTH Then Fault "depth_limit": Exit Function
    id = m_nextId
    m_nextId = m_nextId + 1
    m_sp = m_sp + 1
    m_stack(m_sp) = id
    If m_count < m_max Then
        m_count = m_count + 1
        i = m_count
        m_evType(i) = EV_B
        m_evId(i) = id
        If m_sp > 1 Then m_evParent(i) = m_stack(m_sp - 1) Else m_evParent(i) = m_rootParent
        m_evDepth(i) = m_rootDepth + m_sp - 1
        m_evPass(i) = m_pass
        m_evKind(i) = kind
        m_evName(i) = nm
        m_evStatus(i) = ST_OK
        m_evAttrs(i) = attrs
        ' Timestamp last, so bookkeeping is outside the span.
        XSP_QPC t
        m_evTicks(i) = t
    Else
        m_dropped = m_dropped + 1
    End If
    SpanBegin = id
End Function

Private Sub SpanEnd(ByVal id As Long, ByVal status As Byte, ByVal attrs As String)
    Dim t As Currency, i As Long
    ' Timestamp first, so bookkeeping is outside the span.
    XSP_QPC t
    If id = 0 Or Not m_active Then Exit Sub
    If m_sp < 1 Then Fault "end_without_open_span": Exit Sub
    If m_stack(m_sp) <> id Then Fault "end_not_innermost": Exit Sub
    m_sp = m_sp - 1
    If m_count < m_max Then
        m_count = m_count + 1
        i = m_count
        m_evType(i) = EV_E
        m_evId(i) = id
        m_evStatus(i) = status
        m_evAttrs(i) = attrs
        m_evTicks(i) = t
    Else
        m_dropped = m_dropped + 1
    End If
End Sub

Private Sub Fault(ByVal what As String)
    m_faults = m_faults + 1
    If Len(m_firstFault) = 0 Then m_firstFault = what
End Sub

' ===========================================================================
' Attributes (allow-listed keys; values checked against DESIGN.md pattern)
' ===========================================================================

Private Function KeyAllowed(ByVal key As String) As Boolean
    Select Case key
        Case "method", "repeat", "cells", "areas", "address", "sheet", "group", _
             "calc_mode", "threads", "multithreaded", "iteration", "error_code", _
             "note", "count", "overhead_mode", "bytes", "events", "expected", _
             "found", "sha256_prefix"
            KeyAllowed = True
    End Select
End Function

'' Contract pattern ^[\w .:$!'#\-/()\[\],@]*$ (Python, Unicode \w), at most
' 256 characters. VBA has no Unicode letter class, so beyond ASCII this
' accepts letters of the common scripts (Latin, Greek, Cyrillic, Hebrew,
' Arabic, Thai, kana, CJK, Hangul): a subset of \w, never wider. Anything
' else makes the caller drop the attribute (counted in attr_rejected); the
' event is always kept.
Private Function ValueAllowed(ByVal v As String) As Boolean
    Dim i As Long, c As Long
    If Len(v) > XSP_MAX_ATTR_LEN Then Exit Function
    For i = 1 To Len(v)
        c = AscW(Mid$(v, i, 1)) And &HFFFF&
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122
            Case 95, 32, 46, 58, 36, 33, 39, 35, 45, 47, 40, 41, 91, 93, 44, 64
            Case &HAA&, &HB5&, &HBA&, &HC0& To &HD6&, &HD8& To &HF6&, &HF8& To &H24F&
            Case &H386&, &H388& To &H38A&, &H38C&, &H38E& To &H3A1&, &H3A3& To &H3CE&
            Case &H400& To &H481&, &H48A& To &H52F&, &H5D0& To &H5EA&, &H620& To &H64A&
            Case &HE01& To &HE30&, &H3041& To &H3096&, &H30A1& To &H30FA&
            Case &H4E00& To &H9FFF&, &HAC00& To &HD7A3&
            Case Else
                Exit Function
        End Select
    Next i
    ValueAllowed = True
End Function

Private Function AttrJoin(ByVal acc As String, ByVal pair As String) As String
    If Len(acc) = 0 Then AttrJoin = pair Else AttrJoin = acc & "," & pair
End Function

Private Function AttrS(ByVal acc As String, ByVal key As String, ByVal v As String) As String
    AttrS = acc
    If Not KeyAllowed(key) Or Not ValueAllowed(v) Then m_attrRejected = m_attrRejected + 1: Exit Function
    AttrS = AttrJoin(acc, JsonStr(key) & ":" & JsonStr(v))
End Function

' Integer-valued numbers only; anything else is rejected, never rounded.
Private Function AttrN(ByVal acc As String, ByVal key As String, ByVal v As Variant) As String
    AttrN = acc
    If Not KeyAllowed(key) Or Not IsNumeric(v) Then m_attrRejected = m_attrRejected + 1: Exit Function
    If CDec(v) <> Int(CDec(v)) Then m_attrRejected = m_attrRejected + 1: Exit Function
    AttrN = AttrJoin(acc, JsonStr(key) & ":" & CStr(CDec(v)))
End Function

Private Function AttrB(ByVal acc As String, ByVal key As String, ByVal v As Boolean) As String
    AttrB = acc
    If Not KeyAllowed(key) Then m_attrRejected = m_attrRejected + 1: Exit Function
    If v Then
        AttrB = AttrJoin(acc, JsonStr(key) & ":true")
    Else
        AttrB = AttrJoin(acc, JsonStr(key) & ":false")
    End If
End Function

' ===========================================================================
' JSON
' ===========================================================================

' JSON string literal. Output is pure ASCII: quote, backslash and control
' characters use their escapes and every code unit outside 32..126 becomes
' \uXXXX (UTF-16 surrogate pairs stay valid JSON).
Private Function JsonStr(ByVal s As String) As String
    Dim i As Long, c As Long, r As String, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch) And &HFFFF&
        Select Case c
            Case 34: r = r & "\"""
            Case 92: r = r & "\\"
            Case 8: r = r & "\b"
            Case 9: r = r & "\t"
            Case 10: r = r & "\n"
            Case 12: r = r & "\f"
            Case 13: r = r & "\r"
            Case 32 To 126: r = r & ch
            Case Else: r = r & "\u" & Right$("000" & Hex$(c), 4)
        End Select
    Next i
    JsonStr = """" & r & """"
End Function

Private Function StatusStr(ByVal st As Byte) As String
    Select Case st
        Case ST_OK: StatusStr = "ok"
        Case ST_ERROR: StatusStr = "error"
        Case Else: StatusStr = "aborted"
    End Select
End Function

Private Function EventJson(ByVal i As Long) As String
    Dim s As String
    Select Case m_evType(i)
        Case EV_B
            s = "{""type"":""B"",""id"":" & CStr(m_evId(i)) & ",""parent"":" & CStr(m_evParent(i)) & _
                ",""pass"":" & CStr(m_evPass(i)) & ",""depth"":" & CStr(m_evDepth(i)) & _
                ",""ns"":" & TicksToNs(m_evTicks(i)) & ",""clock"":""vba"",""kind"":" & JsonStr(m_evKind(i)) & _
                ",""name"":" & JsonStr(m_evName(i)) & ",""attrs"":{" & m_evAttrs(i) & "}}"
        Case EV_E
            s = "{""type"":""E"",""id"":" & CStr(m_evId(i)) & ",""ns"":" & TicksToNs(m_evTicks(i)) & _
                ",""clock"":""vba"",""status"":" & JsonStr(StatusStr(m_evStatus(i)))
            If Len(m_evAttrs(i)) > 0 Then s = s & ",""attrs"":{" & m_evAttrs(i) & "}"
            s = s & "}"
        Case Else
            s = "{""type"":""M"",""id"":" & CStr(m_evId(i)) & ",""parent"":" & CStr(m_evParent(i)) & _
                ",""pass"":" & CStr(m_evPass(i)) & ",""ns"":" & TicksToNs(m_evTicks(i)) & _
                ",""clock"":""vba"",""kind"":""marker"",""name"":" & JsonStr(m_evName(i)) & ",""attrs"":{}}"
    End Select
    EventJson = s
End Function

Private Function FooterJson() As String
    Dim s As String
    s = "{""type"":""vba_footer"",""events"":" & CStr(m_count) & ",""dropped"":" & CStr(m_dropped) & ",""truncated"":"
    If m_dropped > 0 Then s = s & "true" Else s = s & "false"
    s = s & ",""open_spans"":" & CStr(m_sp) & ",""faults"":" & CStr(m_faults) & _
        ",""attr_rejected"":" & CStr(m_attrRejected) & ",""first_fault"":" & JsonStr(m_firstFault) & _
        ",""mode"":" & JsonStr(m_mode) & ",""pass"":" & CStr(m_pass) & ",""max_events"":" & CStr(m_max) & "}"
    FooterJson = s
End Function

' Writes every buffered event plus the vba_footer line to the Init path and
' ends the session. Returns "ok" or "error:<reason>".
Public Function XSP_Flush() As String
    Dim fso As Object, ts As Object, i As Long, f As Integer
    On Error GoTo Fail
    If Len(m_outPath) = 0 Then XSP_Flush = "error:not_initialized": Exit Function
    ' FileSystemObject handles non-ANSI paths; the content is ASCII either way.
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Fail
    If Not fso Is Nothing Then
        Set ts = fso.CreateTextFile(m_outPath, True, False)
        For i = 1 To m_count
            ts.WriteLine EventJson(i)
        Next i
        ts.WriteLine FooterJson()
        ts.Close
    Else
        f = FreeFile
        Open m_outPath For Output As #f
        For i = 1 To m_count
            Print #f, EventJson(i)
        Next i
        Print #f, FooterJson()
        Close #f
    End If
    m_active = False
    XSP_Flush = "ok"
    Exit Function
Fail:
    m_active = False
    XSP_Flush = "error:" & CStr(Err.Number)
End Function

' ===========================================================================
' Application settings
' ===========================================================================

' Saves Calculation, Iteration and ScreenUpdating, then sets manual
' calculation and ScreenUpdating off for the session. Returns the original
' calculation mode as "automatic" | "manual" | "semiautomatic".
Public Function XSP_SaveSettings() As String
    If Not m_saved Then
        m_savedCalc = Application.Calculation
        m_savedIter = Application.Iteration
        m_savedScreen = Application.ScreenUpdating
        m_saved = True
    End If
    Application.ScreenUpdating = False
    If Application.Calculation <> XL_CALC_MANUAL Then Application.Calculation = XL_CALC_MANUAL
    Select Case m_savedCalc
        Case XL_CALC_AUTOMATIC: XSP_SaveSettings = "automatic"
        Case XL_CALC_SEMIAUTOMATIC: XSP_SaveSettings = "semiautomatic"
        Case Else: XSP_SaveSettings = "manual"
    End Select
End Function

' Restores what XSP_SaveSettings saved. Iteration is restored before the
' calculation mode. Returns "ok", "not_saved" or "error:<n>".
' The runner does not call this: its Excel instance is private and the copy
' is closed without saving, so switching back to automatic would only force
' a pointless recalculation. It is kept for interactive use of the module.
Public Function XSP_RestoreSettings() As String
    On Error GoTo Fail
    If Not m_saved Then XSP_RestoreSettings = "not_saved": Exit Function
    Application.Iteration = m_savedIter
    If Application.Calculation <> m_savedCalc Then Application.Calculation = m_savedCalc
    Application.ScreenUpdating = m_savedScreen
    m_saved = False
    XSP_RestoreSettings = "ok"
    Exit Function
Fail:
    XSP_RestoreSettings = "error:" & CStr(Err.Number)
End Function

' ===========================================================================
' Drill-down driver
' ===========================================================================

' Runs one pass of the plan and flushes the session. planPath names a
' UTF-16LE text file (read once per pass, before any span opens) holding
' steps separated by ChrW(30); fields within a step are separated by ChrW(31):
'   full | fullrebuild | recalc | sheet<US>Sheet | macro<US>MacroName
'   range<US>Sheet<US>Addr1,Addr2,...<US>Key   (Key: R0001, N0001 or G0001)
'   name<US>Sheet<US>Addr1,Addr2,...<US>Key
'   group<US>Sheet<US>Addr1,Addr2,...<US>Key
' The host passes the plan as a file because Application.Run string
' arguments may be limited in length. For range/name/group steps the span
' name is the Key; the host maps it to the trace name, and has already
' extended the addresses to whole array formulas.
' Mode "off" executes exactly the same calls and records only run.pass,
' calc.full and calc.recalc. A failing step ends its span with status "error"
' and the pass continues; a pass-level failure (including a malformed plan)
' ends run.pass with status "aborted".
' Returns "ok", "aborted:<n>", or "<status>;flush:<error>".
Public Function XSP_RunPass(ByVal passId As Long, ByVal mode As String, ByVal planPath As String) As String
    Dim steps() As String, i As Long, passSpan As Long, began As Boolean, passEnded As Boolean
    Dim oldScreen As Boolean, oldIter As Boolean, spanAttributes As String, abortCode As Long
    Dim status As String, flushResult As String, plan As String

    If Not m_active Then XSP_RunPass = "aborted:not_initialized": Exit Function
    If mode <> m_mode Then
        Fault "mode_mismatch"
        flushResult = XSP_Flush()
        XSP_RunPass = "aborted:mode_mismatch"
        Exit Function
    End If

    On Error GoTo Fatal
    m_pass = passId
    plan = ReadTextFile(planPath)
    oldScreen = Application.ScreenUpdating
    oldIter = Application.Iteration
    Application.ScreenUpdating = False
    If Application.Calculation <> XL_CALC_MANUAL Then Application.Calculation = XL_CALC_MANUAL

    spanAttributes = AttrS("", "overhead_mode", mode)
    spanAttributes = AttrS(spanAttributes, "calc_mode", "manual")
    spanAttributes = AttrB(spanAttributes, "iteration", oldIter)
    spanAttributes = AttrB(spanAttributes, "multithreaded", Application.MultiThreadedCalculation.Enabled)
    spanAttributes = AttrN(spanAttributes, "threads", Application.MultiThreadedCalculation.ThreadCount)
    passSpan = SpanBegin("run.pass", "pass", spanAttributes)
    began = True

    If Len(plan) > 0 Then
        steps = Split(plan, ChrW$(30))
        For i = LBound(steps) To UBound(steps)
            If Len(steps(i)) = 0 Then Err.Raise vbObjectError + XSP_ERR_PLAN
            If RunStep(steps(i)) = XSP_ERR_PLAN Then Err.Raise vbObjectError + XSP_ERR_PLAN
        Next i
    End If

    SpanEnd passSpan, ST_OK, ""
    passEnded = True
    status = "ok"

Done:
    On Error Resume Next
    If began And Not passEnded Then SpanEnd passSpan, ST_ABORTED, AttrN("", "error_code", abortCode)
    Application.Iteration = oldIter
    Application.ScreenUpdating = oldScreen
    Err.Clear
    flushResult = XSP_Flush()
    If flushResult = "ok" Then
        XSP_RunPass = status
    Else
        XSP_RunPass = status & ";flush:" & flushResult
    End If
    Exit Function

Fatal:
    abortCode = Err.Number
    If abortCode = vbObjectError + XSP_ERR_PLAN Then abortCode = XSP_ERR_PLAN
    If abortCode = 0 Then abortCode = -1
    status = "aborted:" & CStr(abortCode)
    Resume Done
End Function

' Reads a UTF-16LE text file (BOM optional) as written by the host.
Private Function ReadTextFile(ByVal path As String) As String
    Dim fso As Object, ts As Object, f As Integer, b() As Byte, s As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If Not fso Is Nothing Then
        ' 1 = ForReading, -1 = TristateTrue (UTF-16LE)
        Set ts = fso.OpenTextFile(path, 1, False, -1)
        If Not ts.AtEndOfStream Then s = ts.ReadAll
        ts.Close
    Else
        f = FreeFile
        Open path For Binary Access Read As #f
        If LOF(f) > 0 Then
            ReDim b(0 To LOF(f) - 1)
            Get #f, , b
            s = b
        End If
        Close #f
    End If
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then s = Mid$(s, 2)
    End If
    ReadTextFile = s
End Function

' Returns 0 on success, a non-zero error code for a failed step (recorded in
' its span), or XSP_ERR_PLAN for a malformed step.
Private Function RunStep(ByVal stepText As String) As Long
    Dim f() As String
    f = Split(stepText, ChrW$(31))
    Select Case f(0)
        Case "full"
            If UBound(f) <> 0 Then RunStep = XSP_ERR_PLAN: Exit Function
            RunStep = StepApp("calc.full")
        Case "fullrebuild"
            If UBound(f) <> 0 Then RunStep = XSP_ERR_PLAN: Exit Function
            RunStep = StepApp("calc.fullrebuild")
        Case "recalc"
            If UBound(f) <> 0 Then RunStep = XSP_ERR_PLAN: Exit Function
            RunStep = StepApp("calc.recalc")
        Case "sheet"
            If UBound(f) <> 1 Then RunStep = XSP_ERR_PLAN: Exit Function
            RunStep = StepSheet(f(1))
        Case "range"
            If UBound(f) <> 3 Then RunStep = XSP_ERR_PLAN: Exit Function
            RunStep = StepRange("calc.range", f(1), f(2), f(3))
        Case "name"
            If UBound(f) <> 3 Then RunStep = XSP_ERR_PLAN: Exit Function
            RunStep = StepRange("calc.name", f(1), f(2), f(3))
        Case "group"
            If UBound(f) <> 3 Then RunStep = XSP_ERR_PLAN: Exit Function
            RunStep = StepRange("calc.group", f(1), f(2), f(3))
        Case "macro"
            If UBound(f) <> 1 Then RunStep = XSP_ERR_PLAN: Exit Function
            RunStep = StepMacro(f(1))
        Case Else
            RunStep = XSP_ERR_PLAN
    End Select
End Function

Private Function NonZero(ByVal code As Long) As Long
    If code = 0 Or code = XSP_ERR_PLAN Then NonZero = -1 Else NonZero = code
End Function

Private Function StepApp(ByVal kind As String) As Long
    Dim sid As Long, began As Boolean, code As Long, method As String
    On Error GoTo Fail
    Select Case kind
        Case "calc.full": method = "Application.CalculateFull"
        Case "calc.fullrebuild": method = "Application.CalculateFullRebuild"
        Case Else: method = "Application.Calculate"
    End Select
    sid = SpanBegin(kind, "workbook", AttrS("", "method", method))
    began = True
    Select Case kind
        Case "calc.full": Application.CalculateFull
        Case "calc.fullrebuild": Application.CalculateFullRebuild
        Case Else: Application.Calculate
    End Select
    SpanEnd sid, ST_OK, ""
    Exit Function
Fail:
    code = Err.Number
    Resume Failed
Failed:
    On Error Resume Next
    If Not began Then sid = SpanBegin(kind, "workbook", AttrS("", "method", method))
    SpanEnd sid, ST_ERROR, AttrN("", "error_code", code)
    StepApp = NonZero(code)
End Function

Private Function StepSheet(ByVal sheetName As String) As Long
    Dim ws As Worksheet, sid As Long, began As Boolean, code As Long, spanAttributes As String
    On Error GoTo Fail
    spanAttributes = AttrS("", "method", "Worksheet.Calculate")
    Set ws = ThisWorkbook.Worksheets(sheetName)
    sid = SpanBegin("calc.sheet", sheetName, spanAttributes)
    began = True
    ws.Calculate
    SpanEnd sid, ST_OK, ""
    Exit Function
Fail:
    code = Err.Number
    Resume Failed
Failed:
    On Error Resume Next
    If Not began Then sid = SpanBegin("calc.sheet", sheetName, spanAttributes)
    SpanEnd sid, ST_ERROR, AttrN("", "error_code", code)
    StepSheet = NonZero(code)
End Function

' Range.Calculate over one block, defined name or formula group.
' Range.Calculate is used rather than Range.CalculateRowMajorOrder (which
' Microsoft's RangeTimer uses on Excel 2007+): CalculateRowMajorOrder ignores
' dependencies inside the range, so its result and cost can differ from a
' dependency-correct calculation. The method is recorded in attrs.
' As in RangeTimer, the range covers whole array formulas (the host extends
' the addresses from its static inspection, so no per-cell scan runs here)
' and iteration is switched off for the timing, then restored.
Private Function StepRange(ByVal kind As String, ByVal sheetName As String, _
        ByVal addrList As String, ByVal key As String) As Long
    Dim ws As Worksheet, rng As Range, sid As Long, began As Boolean, code As Long
    Dim spanAttributes As String, addr As String, oldIter As Boolean, iterChanged As Boolean
    On Error GoTo Fail
    spanAttributes = AttrS("", "method", "Range.Calculate")

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set rng = BuildRange(ws, addrList)

    addr = rng.Address
    If Len(addr) <= XSP_MAX_ATTR_LEN Then spanAttributes = AttrS(spanAttributes, "address", addr)
    spanAttributes = AttrN(spanAttributes, "cells", rng.CountLarge)
    spanAttributes = AttrN(spanAttributes, "areas", rng.Areas.Count)

    oldIter = Application.Iteration
    spanAttributes = AttrB(spanAttributes, "iteration", oldIter)
    If oldIter Then
        Application.Iteration = False
        iterChanged = True
    End If

    sid = SpanBegin(kind, key, spanAttributes)
    began = True
    rng.Calculate
    SpanEnd sid, ST_OK, ""
    If iterChanged Then Application.Iteration = oldIter
    Exit Function
Fail:
    code = Err.Number
    Resume Failed
Failed:
    On Error Resume Next
    If Not began Then sid = SpanBegin(kind, key, spanAttributes)
    SpanEnd sid, ST_ERROR, AttrN("", "error_code", code)
    If iterChanged Then Application.Iteration = oldIter
    StepRange = NonZero(code)
End Function

' Union of comma-separated A1 areas on one sheet. Built area by area so the
' 255-character limit of Range("a,b,c") does not apply.
Private Function BuildRange(ByVal ws As Worksheet, ByVal addrList As String) As Range
    Dim parts() As String, i As Long, r As Range
    parts = Split(addrList, ",")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If r Is Nothing Then
                Set r = ws.Range(parts(i))
            Else
                Set r = Application.Union(r, ws.Range(parts(i)))
            End If
        End If
    Next i
    If r Is Nothing Then Err.Raise vbObjectError + XSP_ERR_PLAN
    Set BuildRange = r
End Function

' Runs a user macro in this workbook inside a vba.proc span. Its own
' XSP_Begin/XSP_End spans nest below. Session settings the macro may have
' changed (manual calculation, ScreenUpdating off) are re-applied afterwards,
' outside the span.
Private Function StepMacro(ByVal macroName As String) As Long
    Dim sid As Long, began As Boolean, code As Long, qual As String
    On Error GoTo Fail
    qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
    sid = SpanBegin("vba.proc", macroName, AttrS("", "method", "Application.Run"))
    began = True
    Application.Run qual
    SpanEnd sid, ST_OK, ""
    Application.ScreenUpdating = False
    If Application.Calculation <> XL_CALC_MANUAL Then Application.Calculation = XL_CALC_MANUAL
    Exit Function
Fail:
    code = Err.Number
    Resume Failed
Failed:
    On Error Resume Next
    If Not began Then sid = SpanBegin("vba.proc", macroName, AttrS("", "method", "Application.Run"))
    SpanEnd sid, ST_ERROR, AttrN("", "error_code", code)
    Application.ScreenUpdating = False
    If Application.Calculation <> XL_CALC_MANUAL Then Application.Calculation = XL_CALC_MANUAL
    StepMacro = NonZero(code)
End Function
