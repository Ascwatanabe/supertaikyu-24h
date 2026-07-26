Attribute VB_Name = "LapDataModule"
Option Explicit

Private isAutoUpdateOn As Boolean
Private CFG_SHEET      As String
Private DATA_SHEET     As String
Private ALL_CLASS      As String
Private m_Init         As Boolean
Private Const MAX_ROWS As Long = 50000

' Initialize Japanese sheet/class names via ChrW to avoid encoding issues
Private Sub EnsureInit()
    If m_Init Then Exit Sub
    CFG_SHEET  = ChrW(35373) & ChrW(23450)                                                    ' "Sett" = [setsutei]
    DATA_SHEET = ChrW(12521) & ChrW(12483) & ChrW(12503) & ChrW(19968) & ChrW(35239)          ' "Lap list"
    ALL_CLASS  = ChrW(20840) & ChrW(12463) & ChrW(12521) & ChrW(12473)                        ' "All classes"
    m_Init = True
End Sub

' ---------------------------------------------------------------------------
' Add buttons to the config sheet (called by Workbook_Open or manually)
' ---------------------------------------------------------------------------
Sub InitButtons()
    Call EnsureInit
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CFG_SHEET)

    Dim btn As Button
    On Error Resume Next
    For Each btn In wsConfig.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0

    Dim r As Range
    Set r = wsConfig.Range("D3")
    Dim btnU As Button
    Set btnU = wsConfig.Buttons.Add(r.Left, r.Top, 120, 22)
    btnU.Caption  = ChrW(20170) & ChrW(12377) & ChrW(12368) & ChrW(26356) & ChrW(26032)       ' "Update now"
    btnU.OnAction = "UpdateLapData"
    btnU.Name     = "btnUpdate"

    Set r = wsConfig.Range("D5")
    Dim btnA As Button
    Set btnA = wsConfig.Buttons.Add(r.Left, r.Top, 120, 22)
    btnA.Caption  = ChrW(33258) & ChrW(21205) & ChrW(26356) & ChrW(26032) & ": OFF " & ChrW(9654)  ' "Auto update: OFF"
    btnA.OnAction = "ToggleAutoUpdate"
    btnA.Name     = "btnAutoUpdate"

    ' Start auto-update immediately after buttons are initialized
    Call StartAutoUpdate
End Sub

' ---------------------------------------------------------------------------
' Start auto-update (call from Workbook_Open or InitButtons)
' ---------------------------------------------------------------------------
Sub StartAutoUpdate()
    Call EnsureInit
    If isAutoUpdateOn Then Exit Sub  ' already running
    isAutoUpdateOn = True

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CFG_SHEET)

    Dim captOn As String
    captOn = ChrW(33258) & ChrW(21205) & ChrW(26356) & ChrW(26032) & ": ON " & ChrW(9632)  ' "Auto update: ON"

    Dim btn As Button
    For Each btn In wsConfig.Buttons
        If btn.Name = "btnAutoUpdate" Then btn.Caption = captOn
    Next btn

    Dim interval As Long
    interval = 30
    On Error Resume Next
    interval = CLng(wsConfig.Range("B7").Value)
    On Error GoTo 0
    If interval < 10 Then interval = 10

    Application.OnTime Now() + TimeSerial(0, 0, interval), _
        "'" & ThisWorkbook.Name & "'!AutoUpdateLapData"
End Sub

' ---------------------------------------------------------------------------
' Toggle auto-update on/off
' ---------------------------------------------------------------------------
Sub ToggleAutoUpdate()
    Call EnsureInit
    isAutoUpdateOn = Not isAutoUpdateOn

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CFG_SHEET)

    Dim captOn  As String
    Dim captOff As String
    captOn  = ChrW(33258) & ChrW(21205) & ChrW(26356) & ChrW(26032) & ": ON "  & ChrW(9632)  ' "Auto update: ON"
    captOff = ChrW(33258) & ChrW(21205) & ChrW(26356) & ChrW(26032) & ": OFF " & ChrW(9654)  ' "Auto update: OFF"

    Dim btn As Button
    For Each btn In wsConfig.Buttons
        If btn.Name = "btnAutoUpdate" Then
            If isAutoUpdateOn Then
                btn.Caption = captOn
                Call UpdateLapData
            Else
                btn.Caption = captOff
                Call CancelAutoUpdate
            End If
        End If
    Next btn
End Sub

Sub CancelAutoUpdate()
    On Error Resume Next
    Application.OnTime EarliestTime:=Now(), _
        Procedure:="'" & ThisWorkbook.Name & "'!AutoUpdateLapData", _
        Schedule:=False
    On Error GoTo 0
End Sub

Sub AutoUpdateLapData()
    If Not isAutoUpdateOn Then Exit Sub
    Call EnsureInit
    Call UpdateLapData

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CFG_SHEET)

    Dim interval As Long
    interval = 30
    On Error Resume Next
    interval = CLng(wsConfig.Range("B7").Value)
    On Error GoTo 0
    If interval < 10 Then interval = 10

    Application.OnTime Now() + TimeSerial(0, 0, interval), _
        "'" & ThisWorkbook.Name & "'!AutoUpdateLapData"
End Sub

' ---------------------------------------------------------------------------
' Main update: generate PS script -> convert JSON to CSV -> write to sheet
' ---------------------------------------------------------------------------
Sub UpdateLapData()
    Call EnsureInit
    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    On Error GoTo ErrHandler

    Dim wsConfig As Worksheet
    Dim wsData   As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CFG_SHEET)
    Set wsData   = ThisWorkbook.Sheets(DATA_SHEET)

    Dim jsonPath As String
    jsonPath = Trim(wsConfig.Range("B3").Value)
    If jsonPath = "" Then
        MsgBox "Please set the JSON path (URL or file path).", vbExclamation
        GoTo Cleanup
    End If

    Dim selectedClass As String
    selectedClass = Trim(wsConfig.Range("B5").Value)
    If selectedClass = "" Then selectedClass = ALL_CLASS

    Dim csvPath As String : csvPath = Environ("TEMP") & "\supertaikyu_lap.csv"
    Dim psPath  As String : psPath  = Environ("TEMP") & "\supertaikyu_conv.ps1"
    Dim errPath As String : errPath = Environ("TEMP") & "\supertaikyu_err.txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Write PowerShell conversion script
    Dim ts As Object
    Set ts = fso.CreateTextFile(psPath, True, False)
    ts.WriteLine "$ErrorActionPreference = 'Stop'"
    ts.WriteLine "try {"
    If LCase(Left(jsonPath, 4)) = "http" Then
        ts.WriteLine "    $r = Invoke-WebRequest -Uri '" & Replace(jsonPath, "'", "''") & "' -UseBasicParsing -TimeoutSec 30"
        ts.WriteLine "    $j = [System.Text.Encoding]::UTF8.GetString($r.Content) | ConvertFrom-Json"
    Else
        ts.WriteLine "    $j = Get-Content -Path '" & Replace(jsonPath, "'", "''") & "' -Raw -Encoding UTF8 | ConvertFrom-Json"
    End If
    ts.WriteLine "    $s = $j.laps | Sort-Object @{E={[int]$_.car_no}}, @{E={[int]$_.lap_no}}"
    ts.WriteLine "    $c = $s | Select-Object car_no,car_class,driver_slot,driver_name,lap_no,lap_time_ms,lap_time,recorded_at | ConvertTo-Csv -NoTypeInformation"
    ts.WriteLine "    [System.IO.File]::WriteAllLines('" & Replace(csvPath, "'", "''") & "', $c, [System.Text.Encoding]::Unicode)"
    ts.WriteLine "} catch {"
    ts.WriteLine "    $_.Exception.Message | Out-File '" & Replace(errPath, "'", "''") & "' -Encoding UTF8"
    ts.WriteLine "    exit 1"
    ts.WriteLine "}"
    ts.Close

    On Error Resume Next
    fso.DeleteFile csvPath
    fso.DeleteFile errPath
    On Error GoTo 0

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = wsh.Run("powershell -NoProfile -ExecutionPolicy Bypass -File """ & psPath & """", 0, True)

    If Not fso.FileExists(csvPath) Then
        Dim em As String
        em = "Data conversion failed (exit code: " & rc & ")"
        If fso.FileExists(errPath) Then
            Dim et As Object
            Set et = fso.OpenTextFile(errPath, 1, False, -1)
            em = em & vbNewLine & et.ReadAll
            et.Close
        End If
        MsgBox em, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Call ReadCsvToSheet(csvPath, wsData, selectedClass)

    wsConfig.Range("B9").Value = Format(Now(), "yyyy/mm/dd hh:nn:ss")

    On Error Resume Next
    fso.DeleteFile csvPath
    fso.DeleteFile psPath
    On Error GoTo 0

    GoTo Cleanup

ErrHandler:
    MsgBox "Unexpected error: " & Err.Description, vbCritical

Cleanup:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

' ---------------------------------------------------------------------------
' Read CSV and write to the data sheet
' ---------------------------------------------------------------------------
Sub ReadCsvToSheet(csvPath As String, ws As Worksheet, filterClass As String)
    Call EnsureInit
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(csvPath, 1, False, -1)

    If Not ts.AtEndOfStream Then ts.ReadLine  ' skip header

    Dim data(0 To MAX_ROWS, 1 To 8) As Variant
    Dim rc As Long
    rc = 0

    Dim bestMs     As Object
    Dim bestRowIdx As Object
    Set bestMs     = CreateObject("Scripting.Dictionary")
    Set bestRowIdx = CreateObject("Scripting.Dictionary")

    Do While Not ts.AtEndOfStream And rc < MAX_ROWS
        Dim ln As String
        ln = ts.ReadLine
        If Trim(ln) = "" Then GoTo Skip

        Dim f() As String
        f = ParseCsvLine(ln)
        If UBound(f) < 7 Then GoTo Skip

        ' Class filter
        If filterClass <> ALL_CLASS And filterClass <> "" Then
            If Trim(f(1)) <> filterClass Then GoTo Skip
        End If

        data(rc, 1) = Trim(f(0))                                  ' car_no
        data(rc, 2) = Trim(f(1))                                  ' car_class
        data(rc, 3) = Trim(f(3))                                  ' driver_name
        data(rc, 4) = Trim(f(2))                                  ' driver_slot
        data(rc, 5) = IIf(IsNumeric(f(4)), CLng(f(4)), f(4))     ' lap_no
        data(rc, 6) = Trim(f(6))                                  ' lap_time
        Dim ms As Long
        ms = 0
        If IsNumeric(f(5)) Then ms = CLng(f(5))
        data(rc, 7) = ms                                           ' lap_time_ms
        data(rc, 8) = Trim(f(7))                                  ' recorded_at

        ' Track best lap per car
        If ms > 0 Then
            Dim cn As String
            cn = Trim(f(0))
            If Not bestMs.Exists(cn) Then
                bestMs(cn) = ms
                bestRowIdx(cn) = rc
            ElseIf ms < bestMs(cn) Then
                bestMs(cn) = ms
                bestRowIdx(cn) = rc
            End If
        End If

        rc = rc + 1
Skip:
    Loop

    ts.Close

    ws.Range("A2:H" & ws.Rows.Count).ClearContents
    ws.Range("A2:H" & ws.Rows.Count).Interior.ColorIndex = xlNone

    If rc = 0 Then
        MsgBox "No data to display. Check the class filter.", vbInformation
        Exit Sub
    End If

    Dim out() As Variant
    ReDim out(1 To rc, 1 To 8)
    Dim i As Long
    Dim j As Integer
    For i = 0 To rc - 1
        For j = 1 To 8
            out(i + 1, j) = data(i, j)
        Next j
    Next i

    ws.Range(ws.Cells(2, 1), ws.Cells(rc + 1, 8)).Value = out

    ' Highlight best lap row per car in yellow
    Dim k As Variant
    For Each k In bestRowIdx.Keys
        Dim hi As Long
        hi = bestRowIdx(k) + 2
        ws.Range(ws.Cells(hi, 1), ws.Cells(hi, 8)).Interior.Color = RGB(255, 255, 0)
    Next k

    ws.Columns("A:H").AutoFit
End Sub

' ---------------------------------------------------------------------------
' Parse a single CSV line (handles double-quoted fields)
' ---------------------------------------------------------------------------
Function ParseCsvLine(ln As String) As String()
    Dim f(0 To 30) As String
    Dim fc  As Long : fc = 0
    Dim cur As String : cur = ""
    Dim inQ As Boolean : inQ = False
    Dim i As Long
    Dim c As String

    For i = 1 To Len(ln)
        c = Mid(ln, i, 1)
        If c = Chr(34) Then
            If inQ And i < Len(ln) And Mid(ln, i + 1, 1) = Chr(34) Then
                cur = cur & Chr(34)
                i = i + 1
            Else
                inQ = Not inQ
            End If
        ElseIf c = "," And Not inQ Then
            f(fc) = cur : fc = fc + 1 : cur = ""
        Else
            cur = cur & c
        End If
    Next i
    f(fc) = cur

    Dim r() As String
    ReDim r(0 To fc)
    Dim m As Long
    For m = 0 To fc
        r(m) = f(m)
    Next m
    ParseCsvLine = r
End Function