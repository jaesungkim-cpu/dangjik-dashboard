from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse
import pandas as pd

app = FastAPI()

data_store = []

@app.get("/")
def upload_page():
    return HTMLResponse("""
    <h2>당직 대시보드 업로드</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file">
        <button type="submit">업로드</button>
    </form>
    <a href="/dashboard">대시보드 보기</a>
    """)

@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    df = pd.read_excel(file.file)
    global data_store
    data_store = df.to_dict(orient="records")
    return {"message": "업로드 완료"}

@app.get("/dashboard")
def dashboard():
    total = len(data_store)
    people = len(set([d.get("이름","") for d in data_store]))

    return {
        "총 건수": total,
        "참여 인원": people
    }
