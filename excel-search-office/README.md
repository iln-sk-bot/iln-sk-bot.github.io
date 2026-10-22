# Excel Search Taskpane v0.2.0（Windows デスクトップ版Excel）

VS Code / Project View版 **Excel Search v0.5.0** を基盤として作成した、個人サイドロード向けOfficeアドインの検証版です。

- 検索画面は **Excelのタスクペイン（ShowTaskpane）**。ドッキング・解除はExcelのUIで行います。
- データアクセスは **Office.js**。ダイアログ間通信、Function Commandによる起動・通信は使用しません。
- **現在このアドインを開いているブックのみ**を読み込みます。他のファイルは読み込みません。
- 静的WebファイルはGitHub Pagesの `/docs` にホストできます。サーバー処理・ビルド工程は不要です。
- v0.1.0とは別のアドインIDであり、以前のポップアップ検証版と共存できます。

## この版で動作確認した機能

- シート全体の読み込み（Office.js、範囲の分割読み込み）
- Excelテーブルの認識、シート・テーブル・任意範囲のRead Unit設定
- 検索文字列のインクリメンタル検索、正規表現、列指定
- 同梱 `@jbrowse/jexl 6.0.0` による条件式検索、型付き数値・ブール値
- 検索結果の通常クリックによる、Excelの該当シート・セルの選択
- 元UIのセル複数選択、TSVコピー、検索履歴、保存条件
- 狭いドッキングタスクペイン向けレイアウト
- 読み込みの手動更新（↻ボタン）

VS Code版v0.5.0の計算列、条件付き書式、ユーザー定義関数、グループ表示、診断モードのUI・JavaScriptも同梱しています。ただし、**これらの機能のExcel実機での互換検証はv0.2.0の対象外**です。未検証部分はプレビューとして扱ってください。

## 前提条件

- Microsoft 365などの**Windowsデスクトップ版Excel**でOffice.jsタスクペインが利用できること
- Officeアドインのサイドロードが管理ポリシーでブロックされていないこと
- GitHub PagesへHTTPS接続できること
- Windows PowerShell 5.1以降（設定スクリプト用）

## GitHub Pagesへの配置

1. GitHubにリポジトリを用意します（例：`excel-search-office`）。公開リポジトリで構いません。**GitHub Pagesに置くJavaScriptは公開される**ので、認証情報を記載しないでください。
2. ZIPを展開し、`docs/` フォルダーがリポジトリのルート直下にあるよう配置します。`manifest.template.xml` や `configure.ps1` は公開不要です。
3. GitHubの **Settings → Pages → Deploy from a branch → main /docs** を選択します。
4. 公開後、URLが例えば `https://USER.github.io/excel-search-office/` であることを確認します。トップページには説明だけが表示されます。

## マニフェストの生成

ZIP展開先のPowerShellで実行します。

```powershell
.\configure.ps1 -PagesUrl 'https://USER.github.io/excel-search-office/'
```

同じフォルダーに `manifest.xml` が生成されます。URLの末尾には `/` をつけることを推奨します（スクリプトでも補完します）。

**注意：** `manifest.template.xml` をそのままExcelにサイドロードしても、ダミーURLを参照するため動きません。

## Excelへのサイドロード（個人開発用）

これまでのポップアップ版と同様に、登録済みの方法で `manifest.xml` を読み込めます。

ネットワーク共有の信頼済みアドインカタログを使う場合は、生成した `manifest.xml` を共有フォルダーに置き、Excelの「ファイル → オプション → トラストセンター → 信頼されているアドインカタログ」でUNCパスを登録します。Excelを再起動後、「ホーム → アドイン → 詳細設定 → 共有フォルダー」から追加してください。

個人の開発用にレジストリ方式を利用している場合は、次の補助スクリプトも使えます（**手動で実行した場合のみ登録されます**）。

```powershell
.\register-dev.ps1
```

開発用の `HKCU:\Software\Microsoft\Office\16.0\WEF\Developer` にmanifest IDとXMLパスを登録します。Excelを完全終了・再起動し、アドイン一覧に反映させてください。解除は `unregister-dev.ps1` です。いずれも**テスト用方式**であり、会社管理PCではポリシーに従ってください。

リボンの **ホーム → 検索ビュー → Excel検索ビュー** から開きます。通常は画面右側にタスクペインが開きます。Windows版Excelでのドッキング解除はExcelのタスクペインのUIを使ってください。

## 操作

1. 現在のブックを初回起動時に読み込みます。大きいブックでは、読込進捗をステータスに表示します。
2. 「対象」を開き、シートやRead Unitを有効・無効にします。Excelテーブルの単位は初期状態で無効です。シート全体は初期状態で有効です。
3. 検索欄で文字列、正規表現、またはJexl式を入力します。Jexlの例：`names["数量"] >= 5`。
4. 検索結果のセルをクリックすると、そのブックの対応セルが選択されます。計算列（物理セルがない列）はジャンプ対象になりません。
5. 複数選択してCtrl+C、右クリックからTSVコピー、または各Read Unitの「全件TSV」を使います。
6. Excelを編集した後は、検索UIの **↻** ボタンで再読み込みしてください。

## 保存・プライバシー

- 設定とキャッシュはOffice WebViewの **IndexedDB** 内へ、ブックのURL/名前をキーとして保存します。元のExcelブックには検索設定を書き込みません。
- 登録するブックは常に現在の1ブックで、別ファイルの検索はできません。
- GitHub PagesからHTML/JavaScriptが読み込まれます。検索中のセル値をサーバーへ送信するコードは含めていません。ただしOffice.js自体はMicrosoftの配信先から読み込まれます。
- **未保存ブック**には安定したファイルURLがないため、保存先の識別子をセッションごとに作り、設定は次回起動に引き継がれません。
- ブック名やパスが変わると、前の設定とは別の保存先として扱われます。

## v0.2.0の制限事項

- **実際のWindows版Excelでの動作確認は未実施**。Office.jsを模擬したブラウザテストとソース検証を行っています。
- 自動変更検出はありません。Excel側でセル編集、計算、行・列追加、テーブル構成変更を行った後は手動更新してください。
- 全シートの使用範囲を読み込みます（描画は元UIの上限最大1,000行）。Office APIからの取得は最大8,000セル単位で分割、ブック合計の使用範囲は初期版では **最大2,000,000セル** です。超過時は明確なエラーを表示します。
- VS Code版の「Excel内蔵日付書式をISO 8601化」の互換実装は未完了。v0.2.0ではOffice.jsの`Range.text`を表示に利用し、数値の生値をJexlへ渡します。特に**1904日付方式**のシリアル正規化は未実装です。
- Office.jsのテーブル・セルの操作はWindows版Excelで追加検証が必要です。マクロ・外部リンク・セル書式を変更するコードはありません。
- IndexedDBがブラウザ/WebViewのセキュリティ設定でブロックされる環境では、現在は初期化できません。
- ブックが非常に大きい場合、2,000,000セル以下でもメモリ不足や処理の遅延が起こり得ます。
- Clipboard APIの利用可否はWebViewの権限に依存します。使えない場合は従来のコピー処理にフォールバックします。

## ソース構成

```text
excel-search-office-taskpane-0.2.0/
  manifest.template.xml    Excel用XMLマニフェスト（雛形）
  configure.ps1            GitHub Pages URLの反映
  register-dev.ps1         開発用サイドロード補助（任意）
  unregister-dev.ps1       登録解除
  docs/
    index.html             GitHub Pagesトップページ
    taskpane.html          VS Code版v0.5.0のUIと検索エンジンを移植
    office-adapter.mjs     Excel.run によるデータ・セル操作
    xlsx.full.min.js       SheetJS CE（既存のSSF関数等に使用）
    vendor/                Jexl v6・date-and-time とライセンス
    icon-*.png             リボンアイコン
  tests/
    adapter.test.mjs       Excel.runモック単体テスト
    browser_smoke.py       ChromiumのUI自動検証（Excel APIを模擬）
    taskpane-preview.png   狭幅プレビュー画像
```

## 開発テスト

Node.js環境：

```sh
node tests/adapter.test.mjs
```

Playwright + Chromiumがある環境：

```sh
python tests/browser_smoke.py
```

ブラウザスモークテストはネットワークなしでHTML・Office API・IndexedDBを模擬するもので、**実際のExcelの保証にはなりません**。

## 既存版からの主な変更点

- Office Dialog APIを廃止し、タスクペイン内の検索画面から直接Excel.runを実行
- VS Code版v0.5.0のUI・検索エンジン・Read Unit構造をそのまま保持
- ファイル解析用 `view.workspace.readBinary` → `Excel.Range.values/text` を用いた分割読み込み
- `view.storage` → ブック別IndexedDBに変更
- Excelファイル追加・削除機能を非表示にし、現在のブック1件へ固定
- 通常セルクリックでExcelのセル選択を実行

### 参照仕様

- [Show or hide an Office Add-in task pane](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/show-hide-add-in)
- [Configure a shared runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Excel Range API](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range)

