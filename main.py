import os
import tempfile
from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import parser

app = FastAPI(title="香盤ジェネレーター")

templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    # docx以外は拒否
    if not file.filename.endswith('.docx'):
        return {"error": "docxファイルのみ対応しています"}
    
    # 一時ファイルに保存
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name
    
    try:
        # パース処理
        csv_content = parser.parse_docx_to_csv(tmp_path)
        
        # CSVを返却
        from io import BytesIO
        csv_bytes = csv_content.encode('utf-8-sig')
        
        return StreamingResponse(
            BytesIO(csv_bytes),
            media_type="text/csv",
            headers={"Content-Disposition": "attachment; filename=kouban.csv"}
        )
    except Exception as e:
        return {"error": f"処理エラー: {str(e)}"}
    finally:
        # 一時ファイル削除
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
