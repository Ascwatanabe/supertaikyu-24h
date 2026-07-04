# =============================================================================
# setup_lap_viewer.ps1
# スーパー耐久 ラップタイム一覧 Excel ファイル生成スクリプト
# 実行すると tools\lap_viewer.xlsm が作成される
# =============================================================================
param(
    [string]$DataPath = ""
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot  = Split-Path -Parent $scriptDir

if ($DataPath -eq "") {
    $DataPath = Join-Path $repoRoot "data\lap_history.json"
    $DataPath = [System.IO.Path]::GetFullPath($DataPath)
}

$outputPath = Join-Path $scriptDir "lap_viewer.xlsm"

Write-Host "======================================================"
Write-Host " スーパー耐久 ラップビューア セットアップ"
Write-Host "======================================================"
Write-Host "データ: $DataPath"
Write-Host "出力先: $outputPath"
Write-Host ""

# ---- VBA メインモジュール ---------------------------------------------------
$vbaMain = @'
Option Explicit

Private isAutoUpdateOn As Boolean

Private Const CFG_SHEET  As String = "設定"
Private Const DATA_SHEET As String = "ラップ一覧"
Private Const MAX_ROWS   As Long   = 50000

' ===========================================================================
' ボタン初期化（Workbook_Open から呼ばれる）
' ===========================================================================
Sub InitButtons()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CFG_SHEET)

    ' 既存ボタン削除
    Dim btn As Button
    On Error Resume Next
    For Each btn In wsConfig.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0

    ' 今すぐ更新ボタン (D3 の左に配置)
    Dim r As Range
    Set r = wsConfig.Range("D3")
    Dim btnU As Button
    Set btnU = wsConfig.Buttons.Add(r.Left, r.Top, 120, 22)
    btnU.Caption = "今すぐ更新"
    btnU.OnAction = "UpdateLapData"
    btnU.Name = "btnUpdate"

    ' 自動更新トグルボタン (D5 の左に配置)
    Set r = wsConfig.Range("D5")
    Dim btnA As Button
    Set btnA = wsConfig.Buttons.Add(r.Left, r.Top, 120, 22)
    btnA.Caption = "自動更新: OFF ▶"
    btnA.OnAction = "ToggleAutoUpdate"
    btnA.Name = "btnAutoUpdate"
End Sub

' ===========================================================================
' 自動更新トグル
' ===========================================================================
Sub ToggleAutoUpdate()
    isAutoUpdateOn = Not isAutoUpdateOn

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CFG_SHEET)

    Dim btn As Button
    For Each btn In wsConfig.Buttons
        If btn.Name = "btnAutoUpdate" Then
            If isAutoUpdateOn Then
                btn.Caption = "自動更新: ON ■"
                Call UpdateLapData
            Else
                btn.Caption = "自動更新: OFF ▶"
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

' ===========================================================================
' メイン更新処理
' ===========================================================================
Sub UpdateLapData()
    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    On Error GoTo ErrHandler

    Dim wsConfig As Worksheet
    Dim wsData   As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(CFG_SHEET)
    Set wsData   = ThisWorkbook.Sheets(DATA_SHEET)

    ' --- 設定読み込み ---
    Dim jsonPath As String
    jsonPath = Trim(wsConfig.Range("B3").Value)
    If jsonPath = "" Then
        MsgBox "JSONファイルパスを設定してください。", vbExclamation
        GoTo Cleanup
    End If

    Dim selectedClass As String
    selectedClass = Trim(wsConfig.Range("B5").Value)
    If selectedClass = "" Then selectedClass = "全クラス"

    ' --- 一時ファイルパス ---
    Dim csvPath As String
    csvPath = Environ("TEMP") & "\supertaikyu_lap.csv"

    Dim psPath As String
    psPath = Environ("TEMP") & "\supertaikyu_conv.ps1"

    Dim errPath As String
    errPath = Environ("TEMP") & "\supertaikyu_err.txt"

    ' --- PowerShell スクリプト生成（JSON→CSV変換）---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.CreateTextFile(psPath, True, False)
    ts.WriteLine "$ErrorActionPreference = 'Stop'"
    ts.WriteLine "try {"
    ts.WriteLine "    $j = Get-Content -Path '" & Replace(jsonPath, "'", "''") & "' -Raw -Encoding UTF8 | ConvertFrom-Json"
    ts.WriteLine "    $s = $j.laps | Sort-Object @{E={[int]$_.car_no}}, @{E={[int]$_.lap_no}}"
    ts.WriteLine "    $c = $s | Select-Object car_no,car_class,driver_slot,driver_name,lap_no,lap_time_ms,lap_time,recorded_at | ConvertTo-Csv -NoTypeInformation"
    ts.WriteLine "    [System.IO.File]::WriteAllLines('" & Replace(csvPath, "'", "''") & "', $c, [System.Text.Encoding]::UTF8)"
    ts.WriteLine "} catch {"
    ts.WriteLine "    $_.Exception.Message | Out-File '" & Replace(errPath, "'", "''") & "' -Encoding UTF8"
    ts.WriteLine "    exit 1"
    ts.WriteLine "}"
    ts.Close

    ' 古い CSV 削除
    On Error Resume Next
    fso.DeleteFile csvPath
    fso.DeleteFile errPath
    On Error GoTo 0

    ' --- PowerShell 実行（完了まで同期待機）---
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = wsh.Run("powershell -NoProfile -ExecutionPolicy Bypass -File """ & psPath & """", 0, True)

    ' --- 変換結果確認 ---
    If Not fso.FileExists(csvPath) Then
        Dim em As String
        em = "データ変換に失敗しました（終了コード: " & rc & "）"
        If fso.FileExists(errPath) Then
            Dim et As Object
            Set et = fso.OpenTextFile(errPath, 1, False, -1)
            em = em & vbNewLine & et.ReadAll
            et.Close
        End If
        MsgBox em, vbExclamation, "エラー"
        GoTo Cleanup
    End If

    ' --- シートへ書き込み ---
    Call ReadCsvToSheet(csvPath, wsData, selectedClass)

    ' --- 最終更新時刻 ---
    wsConfig.Range("B9").Value = Format(Now(), "yyyy/mm/dd hh:nn:ss")

    ' --- 一時ファイル削除 ---
    On Error Resume Next
    fso.DeleteFile csvPath
    fso.DeleteFile psPath
    On Error GoTo 0

    GoTo Cleanup

ErrHandler:
    MsgBox "予期しないエラー: " & Err.Description, vbCritical

Cleanup:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

' ===========================================================================
' CSV → シート書き込み
' ===========================================================================
Sub ReadCsvToSheet(csvPath As String, ws As Worksheet, filterClass As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(csvPath, 1, False, -1)  ' -1 = TristateMixed (UTF-8 BOM 対応)

    ' ヘッダー行スキップ
    If Not ts.AtEndOfStream Then ts.ReadLine

    ' データ格納配列（0 ベース、大きめに確保）
    Dim data(0 To MAX_ROWS, 1 To 8) As Variant
    Dim rc As Long
    rc = 0

    ' ベストラップ追跡用 Dictionary
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

        ' クラスフィルタ
        If filterClass <> "全クラス" And filterClass <> "" Then
            If Trim(f(1)) <> filterClass Then GoTo Skip
        End If

        data(rc, 1) = Trim(f(0))                                              ' 車番
        data(rc, 2) = Trim(f(1))                                              ' クラス
        data(rc, 3) = Trim(f(3))                                              ' ドライバー名
        data(rc, 4) = Trim(f(2))                                              ' スロット
        data(rc, 5) = IIf(IsNumeric(f(4)), CLng(f(4)), f(4))                 ' 周番号
        data(rc, 6) = Trim(f(6))                                              ' ラップタイム
        Dim ms As Long
        ms = 0
        If IsNumeric(f(5)) Then ms = CLng(f(5))
        data(rc, 7) = ms                                                       ' ラップタイム(ms)
        data(rc, 8) = Trim(f(7))                                              ' 記録日時

        ' ベストラップ追跡
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

    ' シートクリア（ヘッダー行は残す）
    ws.Range("A2:H" & ws.Rows.Count).ClearContents
    ws.Range("A2:H" & ws.Rows.Count).Interior.ColorIndex = xlNone

    If rc = 0 Then
        MsgBox "表示できるデータがありません。クラスフィルタを確認してください。", vbInformation
        Exit Sub
    End If

    ' 出力用配列（1 ベース）
    Dim out() As Variant
    ReDim out(1 To rc, 1 To 8)
    Dim i As Long
    Dim j As Integer
    For i = 0 To rc - 1
        For j = 1 To 8
            out(i + 1, j) = data(i, j)
        Next j
    Next i

    ' 一括書き込み（高速）
    ws.Range(ws.Cells(2, 1), ws.Cells(rc + 1, 8)).Value = out

    ' ベストラップ行を黄色でハイライト
    Dim k As Variant
    For Each k In bestRowIdx.Keys
        Dim hi As Long
        hi = bestRowIdx(k) + 2   ' 0始まり + ヘッダー行 = シート行
        ws.Range(ws.Cells(hi, 1), ws.Cells(hi, 8)).Interior.Color = RGB(255, 255, 0)
    Next k

    ' 列幅自動調整
    ws.Columns("A:H").AutoFit
End Sub

' ===========================================================================
' CSV 1 行パース（ダブルクォート対応）
' ===========================================================================
Function ParseCsvLine(ln As String) As String()
    Dim f(0 To 30) As String
    Dim fc  As Long
    fc = 0
    Dim cur As String
    cur = ""
    Dim inQ As Boolean
    inQ = False
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
            f(fc) = cur
            fc = fc + 1
            cur = ""
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
'@

# ---- VBA ThisWorkbook モジュール -------------------------------------------
$vbaWorkbook = @'
Private Sub Workbook_Open()
    Call InitButtons
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.OnTime EarliestTime:=Now(), _
        Procedure:="'" & ThisWorkbook.Name & "'!AutoUpdateLapData", _
        Schedule:=False
    On Error GoTo 0
End Sub
'@

# ---- Excel COM でブック作成 -------------------------------------------------
# セットアップ前に起動済みの Excel PID を記録（終了時に自分のものだけ落とすため）
$preExcelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

Write-Host "Excel を起動しています..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()

    # --- シート準備 ---
    # 既存シート名を変更
    $wb.Sheets(1).Name = "設定"

    # "ラップ一覧" シートを先頭の後に追加
    $wsData   = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets("設定"))
    $wsData.Name = "ラップ一覧"

    # 不要な余分シートを削除
    while ($wb.Sheets.Count -gt 2) {
        $wb.Sheets($wb.Sheets.Count).Delete()
    }

    $wsConfig = $wb.Sheets("設定")
    $wsData   = $wb.Sheets("ラップ一覧")

    # ---- 設定シート レイアウト -----------------------------------------------
    # タイトル
    $t = $wsConfig.Range("A1")
    $t.Value = "スーパー耐久 ラップタイム一覧 設定"
    $t.Font.Bold = $true
    $t.Font.Size = 14

    # JSONパス
    $wsConfig.Range("A3").Value = "JSON ファイルパス:"
    $wsConfig.Range("A3").Font.Bold = $true
    $wsConfig.Range("B3").Value = $DataPath
    $wsConfig.Columns("B").ColumnWidth = 70

    # クラスフィルタ
    $wsConfig.Range("A5").Value = "クラスフィルタ:"
    $wsConfig.Range("A5").Font.Bold = $true
    $wsConfig.Range("B5").Value = "全クラス"

    # クラスのドロップダウン検証
    $dv = $wsConfig.Range("B5").Validation
    $dv.Delete()
    # 3 = xlValidateList / 1 = xlValidAlertStop / 1 = xlBetween
    $dv.Add(3, 1, 1, "全クラス,ST-X,ST-Z,ST-Q,ST-1,ST-2,ST-3,ST-4,ST-TCR,ST-USA,ST-5R,ST-5F")
    $dv.ShowInput  = $false
    $dv.ShowError  = $false

    # 更新間隔
    $wsConfig.Range("A7").Value = "更新間隔（秒）:"
    $wsConfig.Range("A7").Font.Bold = $true
    $wsConfig.Range("B7").Value = 30

    # 最終更新
    $wsConfig.Range("A9").Value = "最終更新:"
    $wsConfig.Range("A9").Font.Bold = $true
    $wsConfig.Range("B9").Value = "未更新"

    # ---- ラップ一覧シート ヘッダー -------------------------------------------
    $headers = @("車番","クラス","ドライバー名","スロット","周番号","ラップタイム","ラップタイム(ms)","記録日時")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $wsData.Cells(1, $i + 1).Value = $headers[$i]
    }
    $hdr = $wsData.Range("A1:H1")
    $hdr.Font.Bold     = $true
    $hdr.Interior.Color = 11829830   # Steel Blue: RGB(70,130,180)
    $hdr.Font.Color    = 16777215   # White

    # ---- VBA コード埋め込み -------------------------------------------------
    try {
        $vbaProj = $wb.VBProject

        # 標準モジュール追加
        $mod = $vbaProj.VBComponents.Add(1)   # 1 = vbext_ct_StdModule
        $mod.Name = "LapDataModule"
        $mod.CodeModule.AddFromString($vbaMain)

        # ThisWorkbook にイベントコード追加
        $wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString($vbaWorkbook)

        Write-Host "VBA コードを埋め込みました。"
        $vbaOk = $true
    }
    catch {
        Write-Warning "VBA の自動埋め込みに失敗しました。"
        Write-Warning "原因: $_"
        Write-Warning ""
        Write-Warning "【手動対応が必要です】"
        Write-Warning "1. Excel のオプション > セキュリティ センター > セキュリティ センターの設定"
        Write-Warning "   > マクロの設定 > 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック"
        Write-Warning "2. このスクリプトを再実行してください。"
        Write-Warning ""
        Write-Warning "または: tools\LapDataModule.bas を Visual Basic Editor でインポートしてください。"
        $vbaOk = $false
    }

    # ---- .xlsm として保存 ---------------------------------------------------
    # 既存ファイルがあれば削除
    if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

    $wb.SaveAs($outputPath, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    Write-Host ""
    Write-Host "保存完了: $outputPath"

    if (-not $vbaOk) {
        Write-Host ""
        Write-Host ".bas ファイルも出力しておきます..."
        $basHeader = "Attribute VB_Name = ""LapDataModule""`r`n"
        [System.IO.File]::WriteAllText(
            (Join-Path $scriptDir "LapDataModule.bas"),
            $basHeader + $vbaMain,
            [System.Text.Encoding]::UTF8
        )
        Write-Host "tools\LapDataModule.bas を VBE でインポートしてください。"
    }

    Write-Host ""
    Write-Host "======================================================"
    Write-Host " セットアップ完了！"
    Write-Host " lap_viewer.xlsm を開いてください。"
    Write-Host " 初回はマクロを有効化してください。"
    Write-Host "======================================================"
}
finally {
    # Excel プロセスを確実に終了させる
    $excelPid = $null
    try { $excelPid = $excel.Hwnd } catch {}

    # COM 参照解放 → スクリプト起動後に増えた EXCEL プロセスのみ強制終了
    # （Quit() / Close() はハングするため使わない）
    try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    [System.GC]::Collect()
    Start-Sleep -Milliseconds 800
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $preExcelPids -notcontains $_.Id } |
        Stop-Process -Force -ErrorAction SilentlyContinue
}
