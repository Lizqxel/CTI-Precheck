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

1. PyInstaller をインストール

```bash
pip install pyinstaller
```

2. 単一EXEを作成

```bash
pyinstaller --noconsole --onefile --name CTI-Precheck app.py
```

生成物は `dist/CTI-Precheck.exe` に出力されます。