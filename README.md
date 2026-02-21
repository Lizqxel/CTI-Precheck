# CTI-Precheck

## デスクトップCSVアップロードUI（MVP）

以下に対応しています。

- CSVファイルのアップロード
- A列（郵便番号）/B列（住所）の基本チェック
- 提供判定の一括実行
- 停止ボタンによる途中キャンセル
- 監視設定（ブラウザ表示ON/OFF）
- 進捗ログ表示
- 結果CSV保存（A=郵便番号, B=住所, C=判定結果）
- 失敗行のポップアップ表示

## 起動方法

1. 依存関係をインストール

```bash
pip install -r requirements.txt
```

2. デスクトップアプリを起動

```bash
python app.py
```

## 使い方

1. 「CSVファイルを選択」で入力CSVを読み込む
2. 必要に応じて設定を変更
	- ブラウザ表示で監視する（OFFでヘッドレス）
3. 「提供判定開始」を押す
4. 途中で止める場合は「停止」を押す
5. 完了後「結果CSV保存」で出力する

## EXE化（Windows）

1. 依存関係をインストール

```bash
pip install -r requirements.txt
```

2. `version.py` の値を更新

- `VERSION`
- `GITHUB_OWNER`
- `GITHUB_REPO`

3. Releaseビルドを実行

```bash
python build_release.py
```

- `pyinstaller build_release.spec` 実行後に `dist/checksums.txt` も自動生成されます。

4. 生成されたEXEを確認

- 出力先: `dist/CTI-Precheck-{VERSION}.exe`
- チェックサム: `dist/checksums.txt`

5. GitHub Release の assets に以下を添付

- `CTI-Precheck-{VERSION}.exe`
- `checksums.txt`

## GitHub Release 運用

- タグ名は `vX.Y.Z` 形式（例: `v0.1.1`）
- Release assets に `CTI-Precheck-X.Y.Z.exe` を添付
- SHA256 検証用に以下どちらかを必ず含める
	- `checksums.txt`（推奨）
	- Release本文に `ファイル名 = SHA256` または `SHA256  ファイル名` 形式

`checksums.txt` 例:

```text
9f8e5f5b9f9b95d1d3f2827f44d57dfe96f90b7ec3b0e1002c3f27111e08f3bb  CTI-Precheck-0.1.1.exe
```

## アップデート機能

- 起動時に自動更新チェック（既定: 24時間ごと）
- 「設定」内の「更新チェック」、またはヘルプメニューから手動チェック
- バージョン比較は `packaging.version` による semver 比較
- ダウンロードはリトライ付き（HTTP 429/5xx対応）
- ダウンロード後に SHA256 検証を実施
- 実行中EXE更新は `.bat` による差し替え適用

更新状態は `settings.json` に保存されます。

- `update_settings`: チャンネル、ETag、最終チェック時刻、最終結果
- `update_history`: 更新処理履歴（最大30件）