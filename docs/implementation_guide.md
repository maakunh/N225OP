# 日経225ミニオプション IV監視ダッシュボード 実装手順書

楽天証券「MarketSpeed II RSS」でリアルタイム取得したオプションデータを Excel に時系列で蓄積し、ブラウザ上の 3 つのダッシュボード（IVスマイル・価格マルチチャート・IVヒートマップ）で可視化する仕組みを、ゼロから構築するための手順書です。

---

## 0. 完成イメージと全体構成

### この仕組みでできること

- 日経225ミニオプションの全権利行使価格について、コール／プットの「現在値」と「IV（インプライド・ボラティリティ）」を 1 分ごとに自動取得・記録する
- 記録したデータをブラウザの 3 ダッシュボードで可視化する
  - **IVスマイル**：ある時点の権利行使価格×IVの曲線。時間変化も再生可能
  - **価格マルチチャート**：行使価格ごとの価格推移を小チャートで一覧
  - **IVヒートマップ**：行使価格×時刻のIVを色で表現
- データ更新はExcelが自動で行い、ブラウザは自動リロードで最新を反映

### データの流れ

```
MarketSpeed II（起動・ログイン）
        │  RSS関数でリアルタイム取得
        ▼
Excel「Live」シート（現在値・IVの一覧スナップショット）
        │  VBAが1分ごとに値をコピー
        ▼
Excel「Log」シート（時系列で縦に蓄積）
        │  VBAが数分ごとに書き出し
        ▼
data.js（ブラウザが読めるJavaScript形式のデータ）
        │  各HTMLが読み込み
        ▼
3つのダッシュボード（ブラウザで表示）
```

### 用意するファイル（すべて同じフォルダに置く）

| ファイル名 | 役割 | 作り方 |
|---|---|---|
| `（任意）.xlsm` | RSS関数とVBAを含むExcelブック本体 | 手順2〜5で作成 |
| `data.js` | 時系列データ（自動生成） | VBAが書き出す |
| `dashboard.html` | IVスマイル | 手順6で作成 |
| `price_dashboard.html` | 価格マルチチャート | 手順6で作成 |
| `heatmap_dashboard.html` | IVヒートマップ | 手順6で作成 |
| `guide.html` | 使い方ガイド（任意） | 別途配布 |

---

## 1. 事前準備

### 1-1. 必要なもの

- **楽天証券の口座**と **MarketSpeed II** のインストール（RSS機能を含む）
- **Microsoft Excel**（デスクトップ版。Microsoft 365 / 2019 以降を推奨）
- **Webブラウザ**（Microsoft Edge または Google Chrome）
- インターネット接続（ダッシュボードがChart.jsをCDNから読み込むため）

### 1-2. MarketSpeed II RSS を有効にする

1. MarketSpeed II を起動し、ログインする
2. メニューから RSS 機能を有効化する（Excelアドインとして登録される）
3. Excel を起動し、リボンに RSS 関連のタブ／関数が使える状態になっていることを確認する

> **重要**：RSS関数はMarketSpeed IIが起動・ログインしている間のみ値を返します。ログオフ中や未起動時は値が取得できません。

---

## 2. Excelブックの作成とシート構成

### 2-1. ブックを作成して保存する

1. Excel で新規ブックを作成する
2. **必ず一度「名前を付けて保存」する**。ファイルの種類は **「Excel マクロ有効ブック（*.xlsm）」** を選ぶ
3. 保存先フォルダを決める（このフォルダが後で全ファイルの置き場所になる）

> ブックを保存しておかないと、VBA が出力先フォルダを特定できず data.js を書き出せません。

### 2-2. シートを2枚用意する

ブック内に次の 2 つのシートを作る（シート名は正確に合わせる）。

- **Live**：RSS関数を並べてリアルタイムのスナップショットを表示するシート
- **Log**：1分ごとのスナップショットを時系列で蓄積するシート

---

## 3. 「Live」シートの構築（RSS関数）

### 3-1. レイアウト

「Live」シートを次の列構成にする。**データは4行目から**始める（1〜3行目はタイトル・見出し用）。

| 列 | 内容 |
|---|---|
| A | 権利行使価格 |
| B | コール銘柄コード |
| C | コール現在値 |
| D | コールIV |
| E | プット銘柄コード |
| F | プット現在値 |
| G | プットIV |

見出しの例（1〜3行目）：

```
1行目:           コール                    プット
2行目:
3行目: 権利行使価格 銘柄コード 現在値 IV 銘柄コード 現在値 IV
4行目: 75375      （関数）   …
```

### 3-2. 銘柄コードを関数で生成する

オプションの銘柄コードは一覧から探す必要はなく、`RssFOPCode` 関数で「銘柄種類・限月・C/P区分・行使価格」から自動生成できる。

```
=RssFOPCode(銘柄種類, 限月, C/P区分, 行使価格)
```

- **銘柄種類**：日経225ミニオプションは **`"N225MOP"`** を指定する
- **限月**：`"202507"` のように年月6桁。ウィークリーオプションは語尾に `#n`（例 `"202507#2"`）
- **C/P区分**：**コール = `1`、プット = `2`**
- **行使価格**：`A4` などのセル参照

> 銘柄種類 `N225MOP` は日経225「ミニ」オプション専用の指定。従来のラージのオプションとは別商品なので混同しないこと。IV項目名など他の引数表記は、楽天証券公式の RSS 関数リファレンス（`https://marketspeed.jp/guide/manual/ms2rss_function.pdf`）で確認すること。バージョンにより異なる場合がある。

### 3-3. 現在値・IVを取得する

生成した銘柄コードを使い、`RssFOPMarket` で各値を取得する。

```
=RssFOPMarket(銘柄コード, "取得項目")
```

各行（例：4行目）の数式イメージ：

| セル | 数式 |
|---|---|
| B4（C銘柄コード） | `=RssFOPCode("N225MOP","202507",1,A4)` |
| C4（C現在値） | `=RssFOPMarket(B4,"現在値")` |
| D4（C_IV） | `=RssFOPMarket(B4,"IV")` |
| E4（P銘柄コード） | `=RssFOPCode("N225MOP","202507",2,A4)` |
| F4（P現在値） | `=RssFOPMarket(E4,"現在値")` |
| G4（P_IV） | `=RssFOPMarket(E4,"IV")` |

> `RssFOPCode` の第1引数が銘柄種類（`N225MOP`）、第2引数が限月、第3引数がC/P区分（コール=1／プット=2）、第4引数が行使価格。限月 `"202507"` の部分は対象限月に合わせて変更する。

### 3-4. 全行使価格に展開する

1. A列に対象の権利行使価格を縦に並べる（ATM中心に上下へ。例：125円刻みで上下数十本）
2. 4行目の数式を、対象行の最終行までフィルコピーする
3. コール・プット両方の現在値とIVが表示されることを確認する

> この時点で「Live」シートに、全権利行使価格のリアルタイムなスナップショットが表示される状態になる。

---

## 4. 「Log」シートの構築

### 4-1. レイアウト

「Log」シートの 1 行目に、次のヘッダーを入れる。

| 列 | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| 1行目 | 時刻 | 権利行使価格 | C現在値 | C_IV | P現在値 | P_IV |

データは2行目以降に、VBAが自動で縦に追記していく（**縦持ち＝1行が1行使価格の1スナップショット**）。最初は空でよい。

---

## 5. VBAマクロの実装

### 5-1. VBエディタを開く

1. Excel で `Alt + F11` を押して VBエディタを開く
2. メニュー「挿入」→「標準モジュール」で新しいモジュールを追加する
3. 下記コードを**すべて**貼り付ける

### 5-2. マクロ全文

```vba
'==========================================================
' 日経225ミニOP ログ取得 + data.js自動更新
'==========================================================
Dim NextRun As Date
Dim Running As Boolean
Dim ExportCounter As Long

Const INTERVAL As String = "00:01:00"   ' ログ間隔
Const FIRST_DATA_ROW As Long = 4        ' Liveのデータ開始行
Const EXPORT_EVERY As Long = 5          ' ログ何回ごとにdata.jsを更新するか

Sub StartLogging()
    Running = True
    ExportCounter = 0
    LogSnapshot
    MsgBox "ログ取得を開始しました（間隔：" & INTERVAL & "）"
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

' data.js を書き出す（ブラウザは開かない）
Sub ExportDataJs()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Log")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Dim folder As String
    folder = ThisWorkbook.Path
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

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText sb
    stream.SaveToFile jsPath, 2
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

' ダッシュボードを開く
Sub OpenDashboards()
    Dim folder As String
    folder = ThisWorkbook.Path
    If folder = "" Then MsgBox "先にブックを保存してください": Exit Sub
    ExportDataJs   ' 最新化してから開く
    Dim p As String
    p = folder & Application.PathSeparator & "dashboard.html"
    If Dir(p) <> "" Then ThisWorkbook.FollowHyperlink p Else MsgBox "dashboard.html がありません: " & folder
End Sub

' 緊急停止（NextRunが失われた場合の総当たり解除）
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
```

### 5-3. 設定値の意味

| 定数 | 既定値 | 意味 |
|---|---|---|
| `INTERVAL` | `"00:01:00"` | ログ取得間隔。5分にするなら `"00:05:00"` |
| `FIRST_DATA_ROW` | `4` | Liveシートのデータ開始行。レイアウトに合わせる |
| `EXPORT_EVERY` | `5` | ログ何回ごとにdata.jsを書き出すか（1分間隔×5＝約5分ごと） |

### 5-4. 動作確認

1. MarketSpeed II が起動・ログイン済みで、Liveシートに値が出ていることを確認
2. VBエディタで `StartLogging` を実行（または Excel の「開発」タブ→マクロ→`StartLogging`）
3. 「Log」シートに、4行目以降ではなく**2行目以降**へ全行使価格分のデータが追記されることを確認
4. 数分待ち、同じフォルダに `data.js` が生成されることを確認
5. 止めるときは `StopLogging` を実行

> **画面が固まる場合**：`Esc` または `Ctrl + Break` で中断。`FIRST_DATA_ROW` がLiveの実際のデータ開始行と一致しているか確認する（不一致だと余計な行まで読んで重くなる）。それでも止まらないときは `ForceStopLogging` を実行する。

---

## 6. ダッシュボードHTMLの設置

3つのHTMLファイルを、Excelブックと**同じフォルダ**に作成する。各ファイルは `data.js` を読み込み、ブラウザ上で描画する。文字コードは **UTF-8** で保存すること。

> 各HTMLの全文は付録（このファイル末尾の「付録A〜C」）に掲載。テキストエディタ（メモ帳等）に貼り付け、指定のファイル名で保存する。

### 6-1. ファイル名

| ファイル名 | 内容 |
|---|---|
| `dashboard.html` | IVスマイル（時刻スライダー・再生・OTM側IV・スキュー表示） |
| `price_dashboard.html` | 価格マルチチャート（ATM中心に表示本数可変） |
| `heatmap_dashboard.html` | IVヒートマップ（コール/プット/OTM切替・色分け） |

### 6-2. 共通の仕様

- 3画面は上部ナビで相互に行き来できる
- 各画面右上の「自動更新」にチェックを入れると、5分ごとにページが自動リロードされ最新の `data.js` を反映する（チェック状態は画面間で共有）
- 「0」のデータ（板なし）は線を繋がず飛ばす設計

---

## 7. 起動と運用

### 7-1. 毎回の起動手順

1. MarketSpeed II を起動・ログインする
2. Excelブック（.xlsm）を開く（マクロを有効化する）
3. Liveシートに値が表示されていることを確認する
4. `StartLogging` を実行する
5. `OpenDashboards` を実行する（ブラウザでIVスマイルが開く）
6. ダッシュボード右上の「自動更新」にチェックを入れる

### 7-2. 終了手順

1. `StopLogging`（または `ForceStopLogging`）を実行する
2. 必要に応じてブックを保存する（Logの蓄積を残す場合）

### 7-3. 運用上の注意

- **ログ取得中はExcelを開いたままにする**。`Application.OnTime` はExcelが起動していないと発火しない
- 寄り付き前など板が薄い時間帯は現在値が「0」になることがあるが、ログ自体は正常に動く
- ブラウザに最新が反映されないときは `Ctrl + F5`（強制再読み込み）
- 監視する行使価格はATM中心に上下10〜15本程度に絞ると、描画も取得も安定する

---

## 8. トラブルシューティング

| 症状 | 主な原因 | 対処 |
|---|---|---|
| RSS関数が `#NAME?` や値なし | MarketSpeed II 未起動／RSS未有効 | MarketSpeed IIを起動・ログインし、RSSを有効化 |
| Liveのヘッダーがログに混入 | `FIRST_DATA_ROW` がデータ開始行と不一致 | 実際の開始行に合わせて修正（本例は4） |
| 実行した瞬間に固まる | A列の空セルを大量ループ | `WorksheetFunction.Count` 版（本手順のコード）を使用 |
| data.jsが出ない | ブック未保存でフォルダ未確定 | ブックを.xlsmで保存してから再実行 |
| ダッシュボードが「data.js が読み込めません」 | HTMLに `<script src="data.js">` が無い／別フォルダ | HTMLの`<head>`にscriptタグを追加、同一フォルダに配置 |
| グラフが更新されない | ブラウザキャッシュ | `Ctrl + F5` で強制再読み込み |
| 文字化け | data.jsの文字コード | 本手順の `ADODB.Stream`（UTF-8指定）版で書き出す |

---

## 9. データ構造リファレンス

### 9-1. Logシート（縦持ち）

| 列 | 内容 | 例 |
|---|---|---|
| A | 時刻（文字列） | `2026/06/24 10:00:00` |
| B | 権利行使価格 | `70000` |
| C | コール現在値 | `1845` |
| D | コールIV | `38.27` |
| E | プット現在値 | `2350` |
| F | プットIV | `36.63` |

### 9-2. data.js

```javascript
window.OPTION_DATA = [
["2026/06/24 10:00:00",75375,340,32.28,0,32.2],
["2026/06/24 10:00:00",75250,350,34.2,0,34.13],
...
];
```

各要素は `[時刻, 権利行使価格, C現在値, C_IV, P現在値, P_IV]` の配列。HTMLはこれを `window.OPTION_DATA` として読み込み、時刻や行使価格でグループ化して描画する。

---

## 付録A. dashboard.html（IVスマイル）

> ※本文は別添の `dashboard.html` を参照。`<head>` 内で Chart.js（CDN）と `data.js` を読み込み、時刻スライダーでスマイルの時間変化を表示する。

## 付録B. price_dashboard.html（価格マルチチャート）

> ※本文は別添の `price_dashboard.html` を参照。行使価格ごとの価格推移を小チャートで格子状に表示する。

## 付録C. heatmap_dashboard.html（IVヒートマップ）

> ※本文は別添の `heatmap_dashboard.html` を参照。行使価格×時刻のIVを色分け表示する。

---

## 改訂メモ

- 本手順書は MarketSpeed II RSS の関数仕様に依存する。RSS関数の項目名・引数はバージョンにより変わることがあるため、公式リファレンスで最新を確認すること。
- ダッシュボードはChart.jsをCDN（`cdn.jsdelivr.net`）から読み込むため、オフライン環境では別途ローカルにライブラリを配置する改修が必要。
