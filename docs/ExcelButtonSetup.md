# Excel "Combine Daily Report" Button Setup

This guide adds a single button in Excel that runs the Python combiner on the *currently open* workbook.

## What the Button Does
1. Saves the active workbook (must contain both sheets: `Timesheet` and `New Formula Job Sheet`).
2. Runs `python run_combine_workbook.py --workbook <this workbook> --open --print-final-only` in the project folder.
3. Python creates intermediate files and the final combined report inside `outputs/`.
4. Opens the resulting combined Excel file automatically.
5. Shows a confirmation message.

## Prerequisites (Once Per Machine)
1. Install Python from https://www.python.org (check "Add python.exe to PATH").
2. Install dependencies:
   ```powershell
   cd "c:\Users\RyanBastianelli\AMA-DailyReport"
   pip install pandas openpyxl python-dateutil
   ```

## Add the Macro
1. Open your source workbook (or a template you reuse).
2. Save it as macro-enabled: File > Save As > Excel Macro-Enabled Workbook (`.xlsm`).
3. Press `Alt+F11` (VBA editor).
4. Insert > Module.
5. Paste the code below.

```vba
Option Explicit

Const PYTHON_EXE As String = "python"   ' Or full path to python.exe
Const PROJECT_DIR As String = "c:\\Users\\RyanBastianelli\\AMA-DailyReport"
Const OUTPUT_DIR As String = "outputs"   ' Relative to PROJECT_DIR

Sub CombineDailyReport_CurrentWorkbook()
    Dim wbPath As String
    wbPath = ThisWorkbook.FullName
    If Len(Dir(PROJECT_DIR, vbDirectory)) = 0 Then
        MsgBox "Project folder not found: " & PROJECT_DIR, vbCritical
        Exit Sub
    End If
    ' Check sheet existence early
    If Not SheetExists("Timesheet") Then
        MsgBox "Sheet 'Timesheet' not found.", vbExclamation
        Exit Sub
    End If
    If Not SheetExists("New Formula Job Sheet") Then
        MsgBox "Sheet 'New Formula Job Sheet' not found.", vbExclamation
        Exit Sub
    End If

    ThisWorkbook.Save ' ensure latest changes

    Dim cmd As String
    cmd = "cmd /c cd /d " & Q(PROJECT_DIR) & " && " & _
          Q(PYTHON_EXE) & " run_combine_workbook.py --workbook " & Q(wbPath) & _
          " --outdir " & Q(OUTPUT_DIR) & " --open --print-final-only"

    Application.StatusBar = "Combining daily report..."
    Dim resultPath As String
    resultPath = RunAndCapture(cmd)

    Application.StatusBar = False

    If Len(resultPath) = 0 Then
        MsgBox "Failed to create report. Check that Python is installed and dependencies are present.", vbCritical
    Else
        MsgBox "Combined report created:" & vbCrLf & resultPath, vbInformation, "Daily Report"
    End If
End Sub

Private Function Q(s As String) As String: Q = Chr(34) & s & Chr(34): End Function

Private Function RunAndCapture(cmd As String) As String
    Dim wsh As Object, execObj As Object, outLine As String, allOut As String
    On Error GoTo fail
    Set wsh = CreateObject("WScript.Shell")
    Set execObj = wsh.Exec(cmd)
    Do While Not execObj.StdOut.AtEndOfStream
        outLine = execObj.StdOut.ReadLine
        If Trim(outLine) <> "" Then allOut = allOut & outLine & vbLf
    Loop
    If execObj.ExitCode <> 0 Then GoTo fail
    Dim parts() As String, i As Long
    parts = Split(allOut, vbLf)
    For i = UBound(parts) To 0 Step -1
        If Trim(parts(i)) <> "" Then
            RunAndCapture = Trim(parts(i))
            Exit For
        End If
    Next i
    Exit Function
fail:
    RunAndCapture = ""
End Function

Private Function SheetExists(name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, name, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function
```

## Add the Button
1. Enable Developer tab (File > Options > Customize Ribbon > check Developer).
2. Developer > Insert > Button (Form Control).
3. Drag on the sheet to place it.
4. Assign macro: `CombineDailyReport_CurrentWorkbook`.
5. Right-click button > Edit Text → “Combine Daily Report”.

## Test
1. Ensure the workbook has the two required sheets.
2. Click the button.
3. After success, the combined report opens; also check the `outputs` folder in the project directory.

## Troubleshooting
| Symptom | Fix |
|---------|-----|
| "python" not recognized | Use full path to python.exe in `PYTHON_EXE` constant |
| Macro not listed | Make sure workbook saved as `.xlsm` and code is in a Module (not a Sheet) |
| Empty result / error | Verify both sheet names and that dependencies are installed |
| Output not opening | Ensure `--open` flag present; if still no, open manually from `outputs` |
| Wrong project folder | Update `PROJECT_DIR` constant |

## Optional Tweaks
- Disable auto-open: remove `--open` in the command.
- Single sheet output: add `--single-sheet` after `--print-final-only`.
- Custom base name: append `--name "MyReport"`.

---
Need enhancements (ribbon button, per-job toggle dialog, logging)? Open an issue or ask. Enjoy!
