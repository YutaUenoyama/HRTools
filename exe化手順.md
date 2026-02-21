# HRTool exe化手順

## 概要
PythonスクリプトをWindows実行ファイル（.exe）に変換する手順です。
**この作業は一度だけ、Windows環境で実行する必要があります。**

## 前提条件
- Windows環境（Windows 10/11）
- インターネット接続

---

## 手順

### ステップ1: Pythonのインストール（一度だけ）

1. Python公式サイトにアクセス
   - https://www.python.org/downloads/

2. **Download Python 3.12.x**（最新版）をクリック

3. インストーラーを実行
   - **重要**: 「Add Python to PATH」に必ずチェックを入れる
   - 「Install Now」をクリック

4. インストール完了後、コマンドプロンプトで確認
   ```cmd
   python --version
   ```
   → `Python 3.12.x` と表示されればOK

---

### ステップ2: 必要なパッケージのインストール（一度だけ）

1. コマンドプロンプトを開く
   - スタートメニュー → 「cmd」と入力 → Enter

2. プロジェクトフォルダに移動
   ```cmd
   cd /d "C:\Users\<ユーザー名>\Downloads\人事管理ツール"
   ```
   ※ 実際のフォルダパスに置き換えてください

3. 必要なパッケージをインストール
   ```cmd
   pip install -r requirements.txt
   ```

---

### ステップ3: exe化の実行

1. コマンドプロンプトでプロジェクトフォルダにいることを確認

2. 以下のコマンドを実行
   ```cmd
   pyinstaller --onefile --noconsole --name HRTool --icon=NONE HRTool.py
   ```

   **オプション説明:**
   - `--onefile`: 単一のexeファイルに全てを含める
   - `--noconsole`: コンソールウィンドウを表示しない
   - `--name HRTool`: 出力ファイル名を指定
   - `--icon=NONE`: アイコンなし（必要ならicoファイルを指定可能）

3. ビルド完了まで待つ（1〜3分程度）

4. 完成
   - `dist/HRTool.exe` が生成される

---

### ステップ4: 配布パッケージの作成

1. 以下のフォルダ構成で配布パッケージを作成
   ```
   HRTool/
   ├── HRTool.exe          ← dist/HRTool.exe をコピー
   ├── input/              ← 空フォルダを作成
   └── output/             ← 空フォルダを作成
   ```

2. このフォルダをZIPで圧縮して配布

---

## エンドユーザーの使い方

1. ZIPファイルを解凍
2. `HRTool.exe` をダブルクリック
3. ダイアログに従って操作
   - 「はい」: inputフォルダ内の全Excelファイルを統合
   - 「いいえ」: 終了
4. 処理完了後、`output/統合ファイル_YYYYMMDD_HHMMSS.xlsx` が生成される

**Pythonのインストールは不要です！**

---

## トラブルシューティング

### 「python は、内部コマンドまたは外部コマンド...」エラー
→ Pythonのインストール時に「Add Python to PATH」にチェックを入れ忘れた可能性
→ Pythonを再インストール

### 「pip: command not found」エラー
→ 以下を実行
```cmd
python -m pip install -r requirements.txt
```

### exeファイルが起動しない
→ ログファイル（処理ログ_*.txt）を確認してエラー内容をチェック

### ウイルス対策ソフトに検出される
→ PyInstallerで作成されたexeは誤検出されることがあります
→ ウイルス対策ソフトの除外設定に追加

---

## 備考

- exe化は**一度だけ**実行すればOK
- コードを修正した場合は、再度ステップ3を実行してexeを再生成
- 生成されたexeファイルは約30〜50MB程度のサイズ
