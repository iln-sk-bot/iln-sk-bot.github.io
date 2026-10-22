# Excel Search Popup POC 0.1.0

Windowsデスクトップ版Excel向けの、**独立した非モーダル検索ウィンドウ**の初期検証実装です。GitHub Pagesで静的ファイルをホストし、XMLマニフェストをWindows Excelへサイドロードします。

**目的:** ①リボンからポップアップを開く → ②現在のブックを検索する → ③検索結果のクリックでExcelのセルへジャンプする、という一連の流れが実機で成立するか検証します。

> **実機未検証です。** Linux環境で単体テスト・Excel APIモックを使った通信テスト・Chromiumによる画面操作テストは通過しています。Windows版Excel固有のダイアログ起動、非表示の共有ランタイム継続、ダイアログ通信はこの配布物を用いて確認してください。

## 1. 今回実装した機能

- Excelのホームタブに「検索ポップアップ」ボタン（通常UI）、「検索ツール診断」ボタン（フォールバックUI）を追加。
- Office Dialog APIで移動・サイズ変更可能な検索ウィンドウを開く。
- 長寿命Shared Runtimeから `Excel.run()` を実行し、Officeダイアログとはメッセージで双方向通信。
- 今開いているアドインのExcelブックから、表示シート一覧を取得。初期選択はアクティブシート。
- 最大5シートをまとめて読み込み。シートごとに先頭 **1,500行 × 32列**、セル文字列500文字までに制限（検証版）。読み取りは40行単位、Officeダイアログへの送信は約60,000文字以内のチャンク＋受信確認（ACK）。
- 読み込んだデータをウィンドウ側メモリにキャッシュして、入力と同時に絞り込み。
- 部分一致検索、正規表現検索、大文字小文字区別、対象列指定（`A,C` 等）。
- 先頭行を見出し扱いするか切り替え。検索結果は最大300行表示。結果数は上限を超えても集計。
- 結果のセルをクリックしてExcelの該当シート・セルを選択。
- 選択セルの表示文字列をコピー、表示結果をTSVでコピー。Clipboard APIを利用できない場合は従来方式にフォールバック。
- デモ画面（Officeなしでも一部のUIと検索処理を動作確認可能）。

今回はJexl、計算列、条件付き書式、保存済みRead Unit、履歴、キャッシュ永続化、更新イベント監視、任意範囲設定、複雑な型処理は**未実装**です。これらはProject View Excel Search v0.5.0から段階的に移植する対象です。

## 2. GitHub Pagesに公開する

1. 公開可能なGitHubリポジトリを用意します（例: `excel-search-office-popup-poc`）。
2. このZIPを展開し、`docs/` フォルダごとリポジトリへコミットします。`docs/` の中に `host.html`、`dialog.html`、各JS/CSS、アイコンPNGがある構造を保ってください。
3. GitHubリポジトリの **Settings → Pages → Build and deployment → Deploy from a branch** を選択し、`main` ブランチの `/docs` フォルダを公開対象にします。
4. 例えば以下のアドレスにアクセスできることを確認します。

   - `https://YOURNAME.github.io/excel-search-office-popup-poc/host.html`
   - `https://YOURNAME.github.io/excel-search-office-popup-poc/dialog.html?demo=1`（デモ）

> GitHub Pagesの内容は原則公開です。**機密データをJavaScriptやデモデータとしてアップロードしないでください。** 本実装にワークブック内容を外部APIへ送る処理はありませんが、利用するWebコードは公開されています。

## 3. Windows PCでマニフェストを生成する

PowerShellから、ZIPを展開したフォルダで実行します。

```powershell
.\configure.ps1 -PagesUrl "https://YOURNAME.github.io/excel-search-office-popup-poc/"
```

生成される `manifest.xml` はローカルで使います。`manifest.template.xml` を手動編集する必要はありません。

Windows PowerShellの実行ポリシーによって拒否された場合、ファイルを確認したうえで次の一回限りの起動形式も使えます。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\configure.ps1 -PagesUrl "https://YOURNAME.github.io/excel-search-office-popup-poc/"
```

**URLは`https://`で始め、リポジトリの公開ルート（`/docs/`は付けない）を指定します。** ユーザー・リポジトリ名を独自のものに置き換えてください。

`manifest.xml` の `Host.Url` は `.../host.html`、ポップアップはそこから同一オリジンの `.../dialog.html` を開きます。ドメインが異なるとOffice Dialog APIが拒否するため、両方とも同じGitHub Pagesサイトでホストしてください。

## 4. Excelにサイドロードする

### 方法A（Node.js不要）: 信頼済み共有フォルダカタログ

1. `C:\OfficeAddins` のようなローカルフォルダを作り、Windowsの共有設定で読み取り可能にします。自分のPCから `\\localhost\OfficeAddins` でアクセスできることを確認してください。
2. 生成した **`manifest.xml` だけ**をその共有フォルダへコピーします。WebファイルはGitHub Pagesにあります。
3. Excelで「ファイル → オプション → トラストセンター（セキュリティセンター） → トラストセンターの設定 → 信頼されているアドインカタログ」を開き、`\\localhost\OfficeAddins` を追加して **「メニューに表示」** を有効にします。
4. Excelを再起動します。
5. Excelの「ホーム → アドイン → 詳細設定 → 共有フォルダー」（表記はExcelのビルドによって異なる）から **Excel Search Popup POC** を追加します。

このネットワーク共有による登録は、Microsoftの**開発・テスト用**方式です。実運用配布としてはサポート対象外です。

### 方法B: Office開発者ツール（Node.jsをインストール済みの場合）

Microsoft公式の `office-addin-debugging` でもサイドロードできます。GitHub Pagesが公開済みで `manifest.xml` のURLが正しい状態で、リポジトリのルートから実行します。

```powershell
npx office-addin-debugging start .\manifest.xml desktop --app excel --no-debug --no-live-reload --prod
```

テスト終了時は登録解除してください。

```powershell
npx office-addin-debugging stop .\manifest.xml --prod
```

この方法はnpmへのアクセスが必要です。ネットワークポリシーで使えない環境では方法Aを使用してください。実機での起動結果は未確認です。

## 5. 動作確認手順

1. Excelで**機密情報を含まないテストブック**を開き、A1:D10程度の表にサンプルデータを入力します。
2. 「ホーム → 検索ポップアップ」を押し、独立した検索ウィンドウが開くことを確認します。
3. アクティブシートがチェック済みになっていることを確認し、**「選択したシートを読み込む」**を押します。
4. 検索語句を入力すると行数が絞られること、正規表現・列指定が反映されることを確認します。
5. 結果表のセルをクリックすると、Excel上で対応セルが選択されることを確認します。
6. 検索ウィンドウを開いたままExcelを操作し、ウィンドウに戻れることを確認します。
7. ウィンドウを**6分以上**開いたままにして、再読み込みやセル移動が引き続き動作するか確認します。
8. 「選択セルをコピー」「表示結果をTSVコピー」が機能するか確認します。
9. ブックのセル内容を変更してもキャッシュは自動更新されません。**「選択したシートを読み込む」**で再取得できることを確認します。

表示できない場合は「ホーム → 検索ツール診断」を開き、診断ログと「ポップアップを開く」ボタンを使ってください。通常は診断ペインを開く必要はありません。

## 6. 検証版の制約と留意点

- 現在のブックから`getUsedRangeOrNullObject(true)`で**値のある使用範囲**を取得し、その左上から指定上限だけ読みます。空行や書式のみの広大な範囲を不要に読み込まないための仕様です。
- セルは`Range.text`（Excelでの表示文字列）として読みます。数値/日付の元の型情報、数式そのもの、書式定義は保持しません。列幅によっては`####`が返る可能性があります。
- **先頭行を見出しとして扱う**は各シートの使用範囲の一行目に適用されます。Excelの行1に固定されません。
- 大規模ブックや複数ブック・大量シート向けの負荷最適化はありません。1回の検索は最大5シート×1,500行×32列に制限しています。
- クリップボードの書き込みはWebView2/Officeの許可状態に依存するため、PCによって失敗する可能性があります。
- ブック編集に合わせて画面を自動更新するイベント監視はありません。更新時は手動再読み込みしてください。
- Office Dialog APIはOffice UI上の非モーダルな子ウィンドウです。Excel APIの呼び出しは**子ウィンドウから直接はできず**、親の共有ランタイム経由です。親と子のランタイム継続やメッセージ通信を今回の最重要確認項目にしています。
- ダイアログがすでに開いている場合に再度リボンボタンを押しても別窓は増やしません。必要なら診断ペインから「開き直す」を使います。
- 管理された会社PCではサイドロードやGitHub Pagesへの通信がブロックされる場合があります。

## 7. 主要ファイル

```text
excel-search-office-popup-poc/
  README.md
  configure.ps1              URL差し替えとmanifest.xml生成
  manifest.template.xml      SharedRuntime + リボン定義
  docs/                      GitHub Pagesで公開するディレクトリ
    host.html                Excel側の共有ランタイムホスト/診断画面
    host.js                  Office.js、シート読み込み、通信ブリッジ
    dialog.html              検索ウィンドウ
    dialog.js                ダイアログ制御、テーブル描画、コピー
    search-core.js           Office非依存の検索ロジック
    styles.css
    icon-16.png, icon-32.png, icon-80.png
  tests/
    search-core.test.cjs     検索エンジン単体テスト
    host.test.cjs            Office/Excel APIモックで通信テスト
    browser_smoke.py         ChromiumでUI操作スモークテスト
    demo-screenshot.png      画面テストのスクリーンショット
```

### 開発環境でのテスト

```bash
node --test tests/*.test.cjs
python tests/browser_smoke.py
```

`browser_smoke.py` はPython PlaywrightとChrome/Chromiumが必要です。Node.js単体テストは外部依存なしで実行できます。

## 8. Microsoft公式参考資料

- [共有ランタイムの構成](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Office Dialog API](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins)
- [Windowsからネットワーク共有でサイドロード](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
- [Excel APIのパフォーマンス・サイズ制限](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/performance)
