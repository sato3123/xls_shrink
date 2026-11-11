# XlsImageShrink

フォルダ（サブフォルダ含む）内の `.xls` を `.xlsx` / `.xlsm` に一括変換し、貼り付けられた画像を圧縮する Windows 用 GUI ツールです。

- マクロ（VBA・Excel 4.0 マクロシート・名前定義の XLM 関数）があれば `.xlsm`、なければ `.xlsx`
- 画像は表示サイズに対する解像度（ppi）・長辺上限・JPEG 品質で再圧縮。位置や表示サイズ（drawing XML）は変更しない
- 1 ファイルあたりのサイズ上限を指定でき、上限に収まる中で最も高い画質を選ぶ
- 既存の `.xlsx` / `.xlsm` を「画像圧縮のみ」の対象に含めるか選べる
- 元ファイルは「残す / 削除 / 移動」から選択。移動したものは履歴から元に戻せる
- 変換前後のレイアウトを比較し、差異があれば既定で元ファイルを削除・移動しない
- オフラインで動作（ネットワーク通信なし）

## 動作環境

- Windows 7 SP1 / 8.1 / 10 / 11
- .NET Framework 4.8
- Microsoft Excel 2010 以降（.xls の変換とレイアウト検証に使用）

## ビルド

Visual Studio は不要です。Windows 標準の .NET Framework コンパイラでビルドします。

```
build.bat          ビルドして、実機の Excel を使う自己テストを実行
build.bat notest   ビルドのみ
```

GitHub Actions でも push のたびにビルドし、Excel を使わないテスト（`--citest`）を実行します。
ビルド済みの exe は Actions の実行結果（Artifacts）から取得できます。

## テスト

| コマンド | 内容 | Excel |
| --- | --- | --- |
| `XlsImageShrink.exe --citest <フォルダ>` | 画像圧縮・サイズ制御・スキャン・圧縮のみのジョブ・移動と復元・設定 | 不要 |
| `XlsImageShrink.exe --selftest <フォルダ>` | 上記に加え、xls の変換・マクロ判定・レイアウト検証 | 必要 |

結果は指定フォルダの `citest_report.txt` / `selftest_report.txt` に出力されます。

## ダウンロード

[Releases](https://github.com/sato3123/xls_shrink/releases) から `XlsImageShrink.exe` をダウンロードしてください。インストールは不要です。

## Code signing policy

Windows 向けのリリース（`XlsImageShrink.exe`）は、コード署名を付けて配布する予定です（SignPath Foundation に申請中）。

Free code signing provided by [SignPath.io](https://about.signpath.io/), certificate by [SignPath Foundation](https://signpath.org/).

- 署名するのは、このリポジトリのソースから GitHub Actions（GitHub ホストのランナー）でビルドした成果物だけです。
- チームの役割
  - 作成者・コミッター: [sato3123](https://github.com/sato3123)
  - レビュアー: [sato3123](https://github.com/sato3123)
  - 署名の承認者: [sato3123](https://github.com/sato3123)
- すべてのメンバーは GitHub と SignPath で多要素認証を使用します。

## Privacy policy

このプログラムは、利用者や作業内容の情報をどこにも送信しません。ネットワーク通信は一切行いません（更新確認・利用状況の送信もありません）。
設定・ログ・レポート・履歴は、実行ファイルと同じフォルダ（書き込めない場合は `%APPDATA%\XlsImageShrink`）にだけ保存されます。

## ライセンス

[MIT License](LICENSE)

## 注意

- 署名のない exe は、Windows 11 のスマート アプリ コントロールが有効な PC では起動がブロックされます。配布にはコード署名が必要です。
- ダウンロードした exe は「ブロックの解除」（ファイルのプロパティ）が必要な場合があります。
