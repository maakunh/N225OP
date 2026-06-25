'==========================================================
' 日経225ミニOP ログ取得 + data.js自動更新
' OneDrive対応・固定フォルダ版
'==========================================================
' 使い方:
'   1. VBエディタ(Alt+F11)で「ファイル」→「ファイルのインポート」からこの.basを取り込む
'   2. DASH_FOLDER を任意の出力先(OneDrive外推奨)に設定
'   3. 3つのHTMLをDASH_FOLDERに置く(data.jsは自動生成)
'   4. StartLogging で取得開始 / OpenDashboards でブラウザ表示 / StopLogging で停止
'==========================================================

Dim NextRun As Date
Dim Running As Boolean
Dim ExportCounter As Long
Dim PrevG1Val As String      ' 前回のG1値(変化検知用)
Dim LastAliveTime As Date    ' G1が最後に変化したPC実時刻

Const INTERVAL As String = "00:01:00"           ' ログ間隔
Const FIRST_DATA_ROW As Long = 4                ' Liveのデータ開始行
Const EXPORT_EVERY As Long = 5                  ' ログ何回ごとにdata.jsを更新するか
Const DASH_FOLDER As String = "C:\OptionDash"   ' ★ダッシュボード一式の固定フォルダ(OneDrive外)

' 出力フォルダを確保して返す(無ければ作成)
Function DashFolder() As String
    If Dir(DASH_FOLDER, vbDirectory) = "" Then
        On Error Resume Next
        MkDir DASH_FOLDER
        On Error GoTo 0
    End If
    DashFolder = DASH_FOLDER
End Function

Sub StartLogging()
    Running = True
    ExportCounter = 0
    LogSnapshot
    MsgBox "ログ取得を開始しました(間隔:" & INTERVAL & ")" & vbCrLf & _
           "出力先:" & DASH_FOLDER
End Sub

Sub StopLogging()
    Running = False
    On Error Resume Next
    Application.OnTime NextRun, "LogSnapshot", , False
    On Error GoTo 0
    MsgBox "ログ取得を停止しました"
End Sub

Sub LogSnapshot()
    If Not Running Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanExit

    Dim wsLive As Worksheet, wsLog As Worksheet
    Set wsLive = ThisWorkbook.Sheets("Live")
    Set wsLog = ThisWorkbook.Sheets("Log")

    ' G1(MS2取得の時刻)が動いているか監視。変化したら実時刻を記録
    Dim curG1 As String
    If IsError(wsLive.Range("G1").Value) Then
        curG1 = "(err)"
    Else
        curG1 = CStr(wsLive.Range("G1").Value)
    End If
    If curG1 <> PrevG1Val Then
        LastAliveTime = Now
        PrevG1Val = curG1
    End If

    Dim n As Long
    n = Application.WorksheetFunction.Count(wsLive.Range("A" & FIRST_DATA_ROW & ":A100000"))
    If n <= 0 Then GoTo CleanExit
    Dim lastLiveRow As Long
    lastLiveRow = FIRST_DATA_ROW + n - 1

    Dim outRow As Long
    outRow = wsLog.Cells(wsLog.Rows.Count, 2).End(xlUp).Row + 1
    If outRow < 2 Then outRow = 2

    Dim tStr As String
    tStr = Format(Now, "yyyy/mm/dd hh:mm:ss")

    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 6)
    Dim i As Long, k As Long
    k = 0
    For i = FIRST_DATA_ROW To lastLiveRow
        If IsNumeric(wsLive.Cells(i, 1).Value) And wsLive.Cells(i, 1).Value <> "" Then
            k = k + 1
            outArr(k, 1) = tStr
            outArr(k, 2) = wsLive.Cells(i, 1).Value
            outArr(k, 3) = wsLive.Cells(i, 3).Value
            outArr(k, 4) = wsLive.Cells(i, 4).Value
            outArr(k, 5) = wsLive.Cells(i, 6).Value
            outArr(k, 6) = wsLive.Cells(i, 7).Value
        End If
    Next i
    If k > 0 Then
        wsLog.Range(wsLog.Cells(outRow, 1), wsLog.Cells(outRow + k - 1, 6)).Value = outArr
    End If

    ' data.js を間引いて自動更新
    ExportCounter = ExportCounter + 1
    If ExportCounter >= EXPORT_EVERY Then
        ExportCounter = 0
        On Error Resume Next
        ExportDataJs
        On Error GoTo CleanExit
    End If

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Running Then
        NextRun = Now + TimeValue(INTERVAL)
        Application.OnTime NextRun, "LogSnapshot"
    End If
End Sub

' data.js を書き出す(固定フォルダへ・リトライ付き)
Sub ExportDataJs()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Log")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim folder As String
    folder = DashFolder()
    If folder = "" Then Exit Sub
    Dim jsPath As String
    jsPath = folder & Application.PathSeparator & "data.js"

    Dim data As Variant
    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6)).Value
    Dim nRows As Long
    nRows = UBound(data, 1)
    Dim parts() As String
    ReDim parts(1 To nRows)
    Dim i As Long
    For i = 1 To nRows
        parts(i) = "[""" & CStr(data(i, 1)) & """," & _
                   NzV(data(i, 2)) & "," & NzV(data(i, 3)) & "," & _
                   NzV(data(i, 4)) & "," & NzV(data(i, 5)) & "," & NzV(data(i, 6)) & "]"
    Next i
    Dim sb As String
    sb = "window.OPTION_DATA = [" & vbCrLf & Join(parts, "," & vbCrLf) & vbCrLf & "];"

    ' ストール検知用メタ情報(G1が最後に動いた実時刻)。未取得時は空文字
    Dim aliveStr As String
    If LastAliveTime = 0 Then
        aliveStr = ""
    Else
        aliveStr = Format(LastAliveTime, "yyyy-mm-dd hh:mm:ss")
    End If
    sb = sb & vbCrLf & "window.OPTION_META = {lastAlive:""" & aliveStr & """};"

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText sb

    ' 書き込みリトライ(同期ロック対策)
    On Error Resume Next
    Dim attempt As Long
    For attempt = 1 To 3
        Err.Clear
        stream.SaveToFile jsPath, 2     ' 2=上書き
        If Err.Number = 0 Then Exit For
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Next attempt
    On Error GoTo 0
    stream.Close
End Sub

Function NzV(v As Variant) As String
    If IsEmpty(v) Or v = "" Then
        NzV = "0"
    ElseIf IsNumeric(v) Then
        NzV = CStr(v)
    Else
        NzV = "0"
    End If
End Function

' ダッシュボードを開く(固定フォルダから)
Sub OpenDashboards()
    Dim folder As String
    folder = DashFolder()
    ExportDataJs   ' 最新化してから開く
    Dim p As String
    p = folder & Application.PathSeparator & "dashboard.html"
    If Dir(p) <> "" Then
        ThisWorkbook.FollowHyperlink p
    Else
        MsgBox "dashboard.html が見つかりません。" & vbCrLf & _
               "3つのHTMLを次のフォルダに置いてください:" & vbCrLf & folder
    End If
End Sub

' 緊急停止(NextRunが失われた場合の総当たり解除)
Sub ForceStopLogging()
    Running = False
    Dim t As Date, i As Long
    On Error Resume Next
    For i = 0 To 480
        t = (Now - TimeValue("00:01:00")) + TimeValue("00:01:00") * i
        Application.OnTime EarliestTime:=t, Procedure:="LogSnapshot", Schedule:=False
    Next i
    On Error GoTo 0
    MsgBox "強制停止を試行しました"
End Sub
