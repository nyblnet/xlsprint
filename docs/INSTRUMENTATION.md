# Adding VBA instrumentation safely

XLSprint times VBA only where you add explicit begin/end calls. It does not
hook arbitrary procedures, because no reliable, verifiable hook for that
exists in VBA.

## How it fits together

- `xlsprint profile` imports `XLSprintTimer.bas` into the **disposable copy**
  only. Your original workbook never contains the module.
- `--macro Name` runs `Name` inside the calc pass, wrapped in a `vba.proc`
  span. Spans that `Name` opens itself nest under that wrapper.
- If a requested `--macro` produces no `vba.proc` span, verification fails
  and the run fails with it.
- In trace-off passes, `XSP_Begin` returns 0 and records nothing, so the
  matched pairs measure instrumentation overhead.

## The pattern

Your workbook has to compile and run whether or not XLSprint is present. Call
the timer late-bound, through a guarded shim:

```vb
' Module: XspShim, in your workbook. Safe when XLSprint is absent.
Option Explicit

Public Function XspBegin(ByVal procName As String) As Long
    On Error Resume Next            ' XLSprint absent -> returns 0
    XspBegin = Application.Run("XSP_Begin", "vba.proc", procName)
End Function

Public Sub XspEnd(ByVal id As Long)
    If id = 0 Then Exit Sub
    On Error Resume Next
    Application.Run "XSP_End", id
End Sub
```

```vb
Public Sub RefreshModel()
    Dim t As Long
    t = XspBegin("RefreshModel")
    On Error GoTo Fail
    ' ... work ...
    XspEnd t
    Exit Sub
Fail:
    XspEnd t                        ' always close, including on error
    Err.Raise Err.Number, , "RefreshModel failed"
End Sub
```

## Rules

1. **Balance every begin with exactly one end**, on every exit path: normal
   return, `Exit Sub`, and error handlers.
2. **Close the innermost span first.** A mismatched `XSP_End` is recorded as a
   fault and is never repaired. A trace with faults, or with a span still open
   at flush, fails validation.
3. **Pass names, never data.** The span name is a procedure name. Do not put
   cell values, formula text, customer identifiers, or file paths in it.
   Names are hashed in the trace by default.
4. **Instrument coarse procedures, not tight loops.** Each begin/end pair
   costs a little, and more when it goes through `Application.Run`. The
   trace-off/trace-on overhead section shows the cost. If it reads "not
   resolvable", it is below run-to-run noise.
5. **Keep macros non-interactive.** A `MsgBox`, `InputBox`, or breakpoint hangs
   the invisible Excel instance. Guard them with a flag your macro can check.
6. **Mind the bound.** At most `--max-events` events are buffered, default
   100 000. An overflow truncates the trace, and a truncated trace fails
   validation.

## What a `vba.proc` time means

It is the direct wall time between the begin and end calls. That includes the
`Application.Run` call overhead at the span edges, and any calculation the
procedure triggers. The report shows self time (duration minus the union of
child spans) separately as a derived value.
