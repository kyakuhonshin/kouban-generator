import os
import tempfile
from datetime import datetime, timezone, timedelta
from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import parser

app = FastAPI(title="香盤ジェネレーター")

# アプリバージョン（PATCH更新ごとに+0.0.1）
APP_VERSION = "v0.9.3"

templates = Jinja2Templates(directory="templates")

# 最大ファイルサイズ：20MB
MAX_FILE_SIZE = 20 * 1024 * 1024

@app.get("/version")
def version():
    return {"version": APP_VERSION}

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "version": APP_VERSION})

@app.get("/terms", response_class=HTMLResponse)
async def terms(request: Request):
    html_content = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>利用規約 - 香盤ジェネレーター</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            background-color: #000;
            color: #ccc;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Hiragino Sans", sans-serif;
            line-height: 1.8;
            padding: 60px 20px;
        }
        .container {
            max-width: 680px;
            margin: 0 auto;
        }
        h1 {
            color: #fff;
            font-size: 1.5rem;
            margin-bottom: 40px;
            letter-spacing: 0.05em;
        }
        h2 {
            color: #aaa;
            font-size: 1.1rem;
            margin: 32px 0 16px;
            font-weight: 500;
        }
        p {
            margin-bottom: 16px;
            font-size: 0.95rem;
        }
        a {
            color: #888;
            text-decoration: none;
        }
        a:hover {
            color: #aaa;
        }
        .back-link {
            margin-top: 48px;
            padding-top: 24px;
            border-top: 1px solid #333;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>利用規約</h1>
        
        <h2>第1条（サービス内容）</h2>
        <p>本サービスは、脚本データをCSV形式へ変換するツールを提供するものです。</p>
        
        <h2>第2条（データの取扱い）</h2>
        <p>アップロードされたファイルは変換処理のみに使用され、処理完了後に削除されます。当社はデータの保存・蓄積を行いません。</p>
        
        <h2>第3条（利用者責任）</h2>
        <p>利用者は、自らの責任において脚本データをアップロードするものとします。</p>
        
        <h2>第4条（免責）</h2>
        <p>本サービスの利用により生じた損害について、当社は一切の責任を負いません。変換精度を保証するものではありません。</p>
        
        <h2>第5条（サービスの停止・変更）</h2>
        <p>当社は予告なくサービスを変更・停止する場合があります。</p>
        
        <h2>第6条（禁止事項）</h2>
        <p>違法コンテンツのアップロードは禁止します。</p>
        
        <div class="back-link">
            <a href="/">← トップページに戻る</a>
        </div>
    </div>
</body>
</html>
"""
    return HTMLResponse(content=html_content)

@app.get("/privacy", response_class=HTMLResponse)
async def privacy(request: Request):
    html_content = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>プライバシーポリシー - 香盤ジェネレーター</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            background-color: #000;
            color: #ccc;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Hiragino Sans", sans-serif;
            line-height: 1.8;
            padding: 60px 20px;
        }
        .container {
            max-width: 680px;
            margin: 0 auto;
        }
        h1 {
            color: #fff;
            font-size: 1.5rem;
            margin-bottom: 40px;
            letter-spacing: 0.05em;
        }
        h2 {
            color: #aaa;
            font-size: 1.1rem;
            margin: 32px 0 16px;
            font-weight: 500;
        }
        p {
            margin-bottom: 16px;
            font-size: 0.95rem;
        }
        a {
            color: #888;
            text-decoration: none;
        }
        a:hover {
            color: #aaa;
        }
        .back-link {
            margin-top: 48px;
            padding-top: 24px;
            border-top: 1px solid #333;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>プライバシーポリシー</h1>
        
        <h2>個人情報の収集について</h2>
        <p>本サービスは個人情報の収集を目的としません。アップロードされたファイルは変換処理のみに使用され、サーバー上に保存されることはありません。</p>
        
        <h2>アクセス解析</h2>
        <p>現在、個人を特定できるアクセス解析は導入していません。</p>
        
        <h2>お問い合わせ</h2>
        <p>ご質問がございましたら、MEW Creatorsまでお問い合わせください。</p>
        
        <div class="back-link">
            <a href="/">← トップページに戻る</a>
        </div>
    </div>
</body>
</html>
"""
    return HTMLResponse(content=html_content)

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    # ファイルサイズチェック
    content = await file.read()
    if len(content) > MAX_FILE_SIZE:
        return {"error": "ファイルサイズは20MBまでです"}
    
    # docx以外は拒否
    if not file.filename.endswith('.docx'):
        return {"error": "docxファイルのみ対応しています"}
    
    # 一時ファイルに保存
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        tmp.write(content)
        tmp_path = tmp.name
    
    try:
        # パース処理
        csv_content = parser.parse_docx_to_csv(tmp_path)
        
        # CSVを返却
        from io import BytesIO
        csv_bytes = csv_content.encode('utf-8-sig')
        
        # ファイル名生成（JST: kouban_YYYYMMDD_HHMM.csv）
        jst = timezone(timedelta(hours=9))
        timestamp = datetime.now(jst).strftime("%Y%m%d_%H%M")
        filename = f"kouban_{timestamp}.csv"
        
        return StreamingResponse(
            BytesIO(csv_bytes),
            media_type="text/csv",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    except Exception as e:
        return {"error": f"解析エラー: {str(e)}"}
    finally:
        # 一時ファイル削除
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
