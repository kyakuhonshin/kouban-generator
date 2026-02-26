# 香盤ジェネレーター デプロイ完全ガイド

## ■ 1. GitHubへpushする手順

### 1-1. GitHubで新規リポジトリ作成
1. https://github.com にアクセス
2. 右上の「+」→「New repository」
3. 設定：
   - Repository name: `kouban-generator`（任意）
   - Description: 空欄でOK
   - Public（チェック済みのまま）
   - 「Add a README file」はチェックしない
   - 「Add .gitignore」はチェックしない
4. 緑の「Create repository」ボタンをクリック
5. 作成後、画面に表示されるURLをコピー（後で使う）
   - 例: `https://github.com/あなたのユーザー名/kouban-generator.git`

### 1-2. ローカルでgit初期化
```bash
cd /Users/minishin/.openclaw/workspace/kouban_web
git init
```

### 1-3. ファイルをステージング
```bash
git add .
```

### 1-4. コミット
```bash
git commit -m "Initial commit"
```

### 1-5. mainブランチに設定
```bash
git branch -M main
```

### 1-6. GitHubと連携
```bash
git remote add origin https://github.com/あなたのユーザー名/kouban-generator.git
```
※URLは1-1でコピーしたものに置き換え

### 1-7. push実行
```bash
git push -u origin main
```

### 1-8. 成功確認
- GitHubのリポジトリページを開く
- ファイル一覧に `main.py`, `parser.py`, `requirements.txt`, `templates/` が表示されていればOK

---

## ■ 2. Renderへのデプロイ手順

### 2-1. Renderアカウント作成（未登録の場合）
1. https://dashboard.render.com にアクセス
2. 「Get Started for Free」→ GitHubでログイン
3. 認可画面で「Authorize render」をクリック

### 2-2. New Web Service作成
1. Dashboardで「New +」ボタンをクリック
2. 「Web Service」を選択
3. GitHubリポジトリ一覧が表示される
4. `kouban-generator` を選択（見つからなければ「Configure account」で権限設定）

### 2-3. 設定入力
| 項目 | 入力値 |
|------|--------|
| Name | `kouban-generator`（任意） |
| Region | Singapore（推奨） |
| Branch | main |
| Runtime | Python 3 |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `uvicorn main:app --host 0.0.0.0 --port 10000` |
| Instance Type | Free（無料プラン） |

### 2-4. デプロイ開始
1. 画面下部の「Create Web Service」をクリック
2. ビルドログが流れる（2〜3分）
3. 「Your service is live」が表示されたら完了

### 2-5. URL確認
- 画面上部に `https://kouban-generator.onrender.com` のようなURLが表示される
- このURLをブラウザで開く

---

## ■ 3. デプロイ後の確認チェック

### 3-1. ルートURLが開くか
```
https://あなたのサービス名.onrender.com/
```
- 黒背景の「香盤ジェネレーター」ページが表示されることを確認

### 3-2. docxアップロードテスト
1. 「ファイルを選択」ボタンをクリック
2. 既存のdocx脚本を選択
3. 「変換してダウンロード」をクリック
4. ブラウザにkouban.csvがダウンロードされることを確認

### 3-3. CSV内容確認
- ExcelまたはNumbersで開く
- scene_no, location, D/N, summary, キャスト名... の列が正しく入っているか確認

### 3-4. エラーハンドリング確認
- docx以外のファイルを選択 → エラーメッセージが表示されるか
- 空のファイルをアップロード → エラーが出るか

---

## ■ 4. よくあるエラーと対処法

### エラー1: Module not found（ビルド時）
**症状**: ビルドログに `ModuleNotFoundError: No module named 'xxx'`

**対処**:
```bash
# ローカルでrequirements.txtを修正
cd /Users/minishin/.openclaw/workspace/kouban_web
```

必要なパッケージを追加：
```
fastapi
uvicorn
python-docx
jinja2
python-multipart
```

追加後：
```bash
git add requirements.txt
git commit -m "Add missing dependencies"
git push origin main
```

Renderは自動で再デプロイされる

---

### エラー2: Portエラー（起動時）
**症状**: `Error: Port 8000 is already in use` または `Application failed to start`

**対処**:
Start Commandが正しいか確認：
```
uvicorn main:app --host 0.0.0.0 --port 10000
```

※Renderでは `--port 10000` が必須（他のポートは使えない）

---

### エラー3: 500エラー（アクセス時）
**症状**: ページを開くと「Internal Server Error」

**対処**:
1. Renderダッシュボードで「Logs」タブを開く
2. エラーメッセージを確認
3. 多くの場合、以下のいずれか：
   - `templates` フォルダがGitHubにpushされていない → フォルダ構造を確認
   - ファイルパスの問題 → main.pyの `directory="templates"` が正しいか確認

---

### エラー4: ビルド成功するがページが真っ白
**症状**: URLを開くと何も表示されない

**対処**:
ブラウザの開発者ツール（F12）→ Consoleタブでエラー確認

多くは：
- 静的ファイル（CSS/JS）のパス問題 → index.htmlのパスを確認
- これは通常発生しない（今回の構成では問題なし）

---

### エラー5: GitHub連携でリポジトリが表示されない
**症状**: Renderで「New Web Service」→ GitHubリポジトリ一覧に対象がない

**対処**:
1. Renderダッシュボード → 左メニュー「Settings」
2. 「GitHub」セクションで「Configure」
3. リポジトリアクセス権限を「All repositories」または該当リポジトリに設定

---

## ■ 補足: Pythonバージョン指定

もしPython 3.11以上を明示的に指定したい場合：

1. ルートに `runtime.txt` を作成：
```
python-3.11.0
```

2. pushする：
```bash
git add runtime.txt
git commit -m "Specify Python version"
git push origin main
```

ただし、指定なしでもPython 3.10以上が使われるため必須ではない。

---

**準備完了。上記手順に従って進めてくれ。**
