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
Dim LastUnderlying As String ' 先物ミニ現在値(Live!F1)の退避
Dim FutNextRun As Date       ' 先物自動更新の次回実行時刻
Dim FutRunning As Boolean    ' 先物自動更新が動作中か

Const INTERVAL As String = "00:01:00"           ' ログ間隔
Const FIRST_DATA_ROW As Long = 4                ' Liveのデータ開始行
Const EXPORT_EVERY As Long = 5                  ' ログ何回ごとにdata.jsを更新するか
Const DASH_FOLDER As String = "C:\OptionDash"   ' ★ダッシュボード一式の固定フォルダ(OneDrive外)

' --- 先物チャート(ExportFutures)用 ---
Const FUT_SHEET As String = "Chart"             ' RssChart式を配置したシート名
Const FUT_FORMULA_ROW As Long = 2               ' 各RssChart式を置く行(データはこの1行下から展開)
Const COL_MINI5 As Long = 1      ' A  : ミニ5分足の式起点列
Const COL_LGD As Long = 12       ' L  : ラージ日足
Const COL_LG60 As Long = 23      ' W  : ラージ60分足
Const COL_TPD As Long = 34       ' AH : TOPIX日足
Const COL_TP60 As Long = 45      ' AS : TOPIX60分足
Const FUT_INTERVAL As String = "01:00:00"  ' 先物自動更新の間隔(60分)
Const RUNLOG_SHEET As String = "RunLog"    ' 実行ログのシート名
Const RUNLOG_MAX As Long = 10000           ' 実行ログの最大行数(超えたら古い行から削除)

' 出力フォルダを確保して返す(無ければ作成)
Function DashFolder() As String
    If Dir(DASH_FOLDER, vbDirectory) = "" Then
        On Error Resume Next
        MkDir DASH_FOLDER
        On Error GoTo 0
    End If
    DashFolder = DASH_FOLDER
End Function

'==========================================================
' 実行ログ(RunLogシート)
' 主要マクロの実行を「日時・機能名・備考」で記録する。
' RUNLOG_MAX を超えたら古い行(先頭)から削除してローテーション。
'==========================================================
Sub WriteRunLog(ByVal funcName As String, Optional ByVal note As String = "")
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(RUNLOG_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = RUNLOG_SHEET
        ws.Range("A1").Value = "日時"
        ws.Range("B1").Value = "機能"
        ws.Range("C1").Value = "備考"
    End If

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
    ws.Cells(r, 2).Value = funcName
    ws.Cells(r, 3).Value = note

    ' ローテーション: データ行数(ヘッダー除く)がRUNLOG_MAXを超えたら古い行から削除
    Dim dataRows As Long
    dataRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
    If dataRows > RUNLOG_MAX Then
        Dim del As Long
        del = dataRows - RUNLOG_MAX
        ws.Rows("2:" & (1 + del)).Delete Shift:=xlUp
    End If
    On Error GoTo 0
End Sub

Sub StartLogging()
    WriteRunLog "StartLogging", "ログ取得開始"
    Running = True
    ExportCounter = 0
    LogSnapshot
    MsgBox "ログ取得を開始しました(間隔:" & INTERVAL & ")" & vbCrLf & _
           "出力先:" & DASH_FOLDER
End Sub

Sub StopLogging()
    WriteRunLog "StopLogging", "ログ取得停止"
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

    ' 先物ミニ現在値(F1)を退避。ATM判定の原資産に使う
    If IsError(wsLive.Range("F1").Value) Then
        LastUnderlying = ""
    Else
        LastUnderlying = NzV(wsLive.Range("F1").Value)
    End If

    ' B1/C1の限月数式(TODAY())を最新化。開きっぱなしでも日付またぎで繰り上がるよう毎分再計算。
    ' C1はB1を参照するため B1→C1 の順で再計算する。
    On Error Resume Next
    wsLive.Range("B1").Calculate
    wsLive.Range("C1").Calculate
    On Error GoTo CleanExit

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
    Dim uOut As String
    If LastUnderlying = "" Or LastUnderlying = "0" Then
        uOut = "null"
    Else
        uOut = LastUnderlying
    End If
    sb = sb & vbCrLf & "window.OPTION_META = {lastAlive:""" & aliveStr & _
         """,underlying:" & uOut & "};"

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

    ' 1時間足(最大30日)も書き出す
    On Error Resume Next
    ExportDataJsHourly
    On Error GoTo 0
End Sub

' data_hourly.js を書き出す(1時間足・最大30日・各時間の最終スナップショットを代表値とする)
' スマイル/ヒートマップ用。data.js(1分足)とは独立。
Sub ExportDataJsHourly()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Log")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim folder As String
    folder = DashFolder()
    If folder = "" Then Exit Sub
    Dim jsPath As String
    jsPath = folder & Application.PathSeparator & "data_hourly.js"

    Dim data As Variant
    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6)).Value
    Dim nRows As Long
    nRows = UBound(data, 1)

    ' 30日より前を除外する基準時刻
    Dim cutoff As Date
    cutoff = Now - 30

    ' 各「日付+時(13文字)+行使価格」の最終スナップショットを採用。
    ' Logは時系列追記なので後勝ち(後の行が新しい)で上書きすれば各時間の最終値が残る。
    ' Dictionaryは初出時の挿入順を保持するため、出力順は時系列・行使価格順に保たれる。
    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")   ' 行キー("yyyy/mm/dd hh|strike")→JS配列文字列

    Dim i As Long
    Dim tFull As String, tHour As String, tJs As String
    Dim rowDate As Date
    Dim rowKey As String
    For i = 1 To nRows
        tFull = CStr(data(i, 1))
        If Len(tFull) >= 13 Then
            ' 30日フィルタ(時刻として解釈できる場合のみ)
            If IsDate(tFull) Then
                rowDate = CDate(tFull)
                If rowDate < cutoff Then GoTo ContinueLoop
            End If
            tHour = Left(tFull, 13)                       ' "yyyy/mm/dd hh"
            tJs = Replace(tHour, "/", "-") & ":00"        ' "yyyy-mm-dd hh:00"
            rowKey = tHour & "|" & NzV(data(i, 2))
            vals(rowKey) = "[""" & tJs & """," & _
                           NzV(data(i, 2)) & "," & NzV(data(i, 3)) & "," & _
                           NzV(data(i, 4)) & "," & NzV(data(i, 5)) & "," & NzV(data(i, 6)) & "]"
        End If
ContinueLoop:
    Next i

    If vals.Count = 0 Then Exit Sub

    ' 出力(Dictionaryは挿入順を保つ。後勝ち上書きでも順序は初出位置に保たれる)
    Dim parts() As String
    ReDim parts(1 To vals.Count)
    Dim kk As Variant, idx As Long
    idx = 0
    For Each kk In vals.Keys
        idx = idx + 1
        parts(idx) = vals(kk)
    Next kk

    Dim sb As String
    sb = "window.OPTION_DATA_H = [" & vbCrLf & Join(parts, "," & vbCrLf) & vbCrLf & "];"

    ' メタ情報(data.jsと同じ)
    Dim aliveStr As String
    If LastAliveTime = 0 Then
        aliveStr = ""
    Else
        aliveStr = Format(LastAliveTime, "yyyy-mm-dd hh:mm:ss")
    End If
    Dim uOut As String
    If LastUnderlying = "" Or LastUnderlying = "0" Then
        uOut = "null"
    Else
        uOut = LastUnderlying
    End If
    sb = sb & vbCrLf & "window.OPTION_META = {lastAlive:""" & aliveStr & _
         """,underlying:" & uOut & "};"

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText sb
    On Error Resume Next
    Dim attempt As Long
    For attempt = 1 To 3
        Err.Clear
        stream.SaveToFile jsPath, 2
        If Err.Number = 0 Then Exit For
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Next attempt
    On Error GoTo 0
    stream.Close
End Sub

Function NzV(ByVal v As Variant) As String
    If IsError(v) Then
        NzV = "0"
    ElseIf IsEmpty(v) Then
        NzV = "0"
    ElseIf Len(Trim(CStr(v))) = 0 Then
        NzV = "0"
    ElseIf IsNumeric(v) Then
        NzV = CStr(v)
    Else
        NzV = "0"
    End If
End Function

' ダッシュボードを開く(固定フォルダから)
Sub OpenDashboards()
    WriteRunLog "OpenDashboards", "ダッシュボードを開く"
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
    WriteRunLog "ForceStopLogging", "緊急停止"
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

'==========================================================
' 先物自動更新タイマー(60分ごとにExportFuturesを実行)
' 1分ループ(LogSnapshot)とは独立した別タイマー。
' StartFutures で開始 / StopFutures で停止。
'==========================================================
Sub StartFutures()
    WriteRunLog "StartFutures", "先物自動更新開始"
    FutRunning = True
    FuturesTick   ' すぐ1回実行し、以降は60分ごと
    MsgBox "先物の自動更新を開始しました(間隔:" & FUT_INTERVAL & ")"
End Sub

Sub StopFutures()
    WriteRunLog "StopFutures", "先物自動更新停止"
    FutRunning = False
    On Error Resume Next
    Application.OnTime FutNextRun, "FuturesTick", , False
    On Error GoTo 0
    MsgBox "先物の自動更新を停止しました"
End Sub

Sub FuturesTick()
    If Not FutRunning Then Exit Sub
    On Error Resume Next
    ExportFutures True   ' silent=True(自動時はMsgBoxを出さない)
    On Error GoTo 0
    If FutRunning Then
        FutNextRun = Now + TimeValue(FUT_INTERVAL)
        Application.OnTime FutNextRun, "FuturesTick"
    End If
End Sub

'==========================================================
' 先物チャート(data_futures.js)書き出し
' RssChartで6系列を取得しdata_futures.jsへ。NT倍率はラージ日足/TOPIX日足で算出。
' 別系統・毎回取り直し。手動マクロ ExportFutures を実行する。
' ※ 定数(FUT_SHEET/FUT_FORMULA_ROW/COL_*)はモジュール先頭で宣言。
' RssChartは省略ヘッダー時 [銘柄名称,市場名称,足種,日付,時刻,始値,高値,安値,終値,出来高] を
' 式セルの下方向に展開する。各系列は11列間隔で配置(10列+余白1)。
'==========================================================
Sub ExportFutures(Optional ByVal silent As Boolean = False)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        If Not silent Then _
            MsgBox "「" & FUT_SHEET & "」シートが見つかりません。" & vbCrLf & _
                   "RssChart式を配置したシートを用意してください。"
        Exit Sub
    End If

    WriteRunLog "ExportFutures", IIf(silent, "先物更新(自動60分)", "先物更新(手動)")

    ' RSSの再計算を待つ(任意で延長可)
    Application.Calculate
    Application.Wait Now + TimeValue("00:00:02")

    Dim folder As String
    folder = DashFolder()
    If folder = "" Then Exit Sub

    ' 各系列を読み出してJS配列文字列化
    Dim sMini As String, sLgD As String, sLg60 As String, sTpD As String, sTp60 As String
    sMini = SeriesJs(ws, FUT_FORMULA_ROW, COL_MINI5)
    sLgD = SeriesJs(ws, FUT_FORMULA_ROW, COL_LGD)
    sLg60 = SeriesJs(ws, FUT_FORMULA_ROW, COL_LG60)
    sTpD = SeriesJs(ws, FUT_FORMULA_ROW, COL_TPD)
    sTp60 = SeriesJs(ws, FUT_FORMULA_ROW, COL_TP60)

    ' NT倍率(ラージ日足終値 / TOPIX日足終値)を日付で対応させて算出
    Dim sNT As String
    sNT = NtRatioJs(ws, FUT_FORMULA_ROW, COL_LGD, COL_TPD)

    Dim sb As String
    sb = "window.FUTURES_DATA = {" & vbCrLf & _
         "  mini5: [" & sMini & "]," & vbCrLf & _
         "  lgD: [" & sLgD & "]," & vbCrLf & _
         "  lg60: [" & sLg60 & "]," & vbCrLf & _
         "  tpD: [" & sTpD & "]," & vbCrLf & _
         "  tp60: [" & sTp60 & "]," & vbCrLf & _
         "  nt: [" & sNT & "]" & vbCrLf & _
         "};" & vbCrLf & _
         "window.FUTURES_META = {generated:""" & Format(Now, "yyyy-mm-dd hh:mm:ss") & """};"

    Dim jsPath As String
    jsPath = folder & Application.PathSeparator & "data_futures.js"
    WriteUtf8 jsPath, sb

    If Not silent Then MsgBox "data_futures.js を書き出しました。" & vbCrLf & folder
End Sub

' 1系列を読み、{t,o,h,l,c}のJSオブジェクト配列文字列にする。
' RssChart展開: 列+3=日付, +4=時刻, +5=始値, +6=高値, +7=安値, +8=終値 (0始まりの相対)
Private Function SeriesJs(ByVal ws As Worksheet, ByVal formulaRow As Long, ByVal baseCol As Long) As String
    Dim out As String, cnt As Long
    Dim r As Long
    Dim dCol As Long, tCol As Long, oCol As Long, hCol As Long, lCol As Long, cCol As Long
    dCol = baseCol + 3: tCol = baseCol + 4
    oCol = baseCol + 5: hCol = baseCol + 6: lCol = baseCol + 7: cCol = baseCol + 8

    ' データは式セルの1行下(ヘッダー)を飛ばし、さらに下から。
    ' 実運用では式セル行の次行以降に値が展開される。空セルが出たら終端とみなす。
    Dim startRow As Long
    startRow = formulaRow + 1
    Dim maxScan As Long
    maxScan = 2000
    For r = startRow To startRow + maxScan
        Dim dv As Variant, cv As Variant
        dv = ws.Cells(r, dCol).Value
        cv = ws.Cells(r, cCol).Value
        If IsError(dv) Then Exit For
        If IsEmpty(dv) Then Exit For
        If Len(Trim(CStr(dv))) = 0 Then Exit For
        If IsDashMarker(dv) Then Exit For   ' --------(RSSの終端マーカー)以降は古い残骸
        ' 終値が数値の行だけ採用(ヘッダー等はスキップ)
        Dim cOk As Boolean
        cOk = False
        If Not IsError(cv) Then
            If IsNumeric(cv) Then cOk = True
        End If
        If cOk Then
            Dim tlabel As String
            tlabel = FmtDate(dv) & FmtTime(ws.Cells(r, tCol).Value)
            If out <> "" Then out = out & ","
            out = out & "{t:""" & tlabel & """,o:" & NzV(ws.Cells(r, oCol).Value) & _
                  ",h:" & NzV(ws.Cells(r, hCol).Value) & ",l:" & NzV(ws.Cells(r, lCol).Value) & _
                  ",c:" & NzV(ws.Cells(r, cCol).Value) & "}"
            cnt = cnt + 1
        End If
    Next r
    SeriesJs = out
End Function

' NT倍率: ラージ日足とTOPIX日足の終値を日付キーで突き合わせて比を算出
Private Function NtRatioJs(ByVal ws As Worksheet, ByVal formulaRow As Long, ByVal lgCol As Long, ByVal tpCol As Long) As String
    ' TOPIX日足の(日付→終値)辞書を作る
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long, startRow As Long
    startRow = formulaRow + 1
    Dim tpD As Long, tpC As Long
    tpD = tpCol + 3: tpC = tpCol + 8
    Dim dv As Variant, cv As Variant, key As String
    For r = startRow To startRow + 2000
        dv = ws.Cells(r, tpD).Value
        cv = ws.Cells(r, tpC).Value
        If IsError(dv) Then Exit For
        If IsEmpty(dv) Then Exit For
        If Len(Trim(CStr(dv))) = 0 Then Exit For
        If IsDashMarker(dv) Then Exit For   ' --------(RSSの終端マーカー)以降は古い残骸
        If Not IsError(cv) Then
            If IsNumeric(cv) Then
                If CDbl(cv) <> 0 Then
                    key = FmtDate(dv)
                    dic(key) = CDbl(cv)
                End If
            End If
        End If
    Next r

    ' ラージ日足を走査し、同日付のTOPIXがあれば比を出す
    Dim lgD As Long, lgC As Long
    lgD = lgCol + 3: lgC = lgCol + 8
    Dim out As String
    For r = startRow To startRow + 2000
        dv = ws.Cells(r, lgD).Value
        cv = ws.Cells(r, lgC).Value
        If IsError(dv) Then Exit For
        If IsEmpty(dv) Then Exit For
        If Len(Trim(CStr(dv))) = 0 Then Exit For
        If IsDashMarker(dv) Then Exit For   ' --------(RSSの終端マーカー)以降は古い残骸
        Dim lcOk As Boolean
        lcOk = False
        If Not IsError(cv) Then
            If IsNumeric(cv) Then lcOk = True
        End If
        If lcOk Then
            key = FmtDate(dv)
            If dic.Exists(key) Then
                Dim tp As Double
                tp = CDbl(dic(key))
                If tp <> 0 Then
                    Dim nt As Double
                    nt = CDbl(cv) / tp
                    If out <> "" Then out = out & ","
                    out = out & "{t:""" & key & """,v:" & Format(nt, "0.0000") & "}"
                End If
            End If
        End If
    Next r
    NtRatioJs = out
End Function


' RSSチャートの終端マーカー判定。
' 日付/時刻セルが "-" のみ(--------等)なら、それ以降は前回取得時の古い残骸とみなす。
Private Function IsDashMarker(ByVal v As Variant) As Boolean
    IsDashMarker = False
    On Error Resume Next
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Dim s As String
    s = Trim(CStr(v))
    If Len(s) = 0 Then Exit Function
    If Len(Replace(s, "-", "")) = 0 Then IsDashMarker = True
    On Error GoTo 0
End Function

' 日付値を yyyy-mm-dd 文字列へ(RssChartの日付はExcel日付 or 文字列のことがある)
Private Function FmtDate(ByVal v As Variant) As String
    On Error Resume Next
    If IsDate(v) Then
        FmtDate = Format(v, "yyyy-mm-dd")
    Else
        ' "2026/06/25" 等の文字列を正規化
        FmtDate = Replace(Trim(CStr(v)), "/", "-")
    End If
    On Error GoTo 0
End Function

' 時刻値を " hh:mm" 文字列へ(空なら空文字=日足)
Private Function FmtTime(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        FmtTime = ""
    ElseIf IsEmpty(v) Then
        FmtTime = ""
    ElseIf Len(Trim(CStr(v))) = 0 Then
        FmtTime = ""
    ElseIf IsDate(v) Then
        FmtTime = " " & Format(v, "hh:mm")
    Else
        Dim s As String
        s = Trim(CStr(v))
        If Len(s) >= 5 Then FmtTime = " " & Left(s, 5) Else FmtTime = " " & s
    End If
    On Error GoTo 0
End Function

' UTF-8でファイル書き出し(リトライ付き) — data.js書き出しと同方式
Private Sub WriteUtf8(ByVal path As String, ByVal content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText content
    On Error Resume Next
    Dim attempt As Long
    For attempt = 1 To 3
        Err.Clear
        stream.SaveToFile path, 2
        If Err.Number = 0 Then Exit For
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Next attempt
    On Error GoTo 0
    stream.Close
End Sub
