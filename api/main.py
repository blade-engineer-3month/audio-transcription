from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import whisper
import uuid
import os
import subprocess

app = FastAPI()

# CORS（GitHub Pages からのアクセス許可）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

model = whisper.load_model("small")

JOBS = {}
TEMP_DIR = "temp"
os.makedirs(TEMP_DIR, exist_ok=True)

@app.post("/api/transcribe")
async def transcribe(file: UploadFile = File(...)):
    job_id = str(uuid.uuid4())
    input_path = f"{TEMP_DIR}/{job_id}_{file.filename}"

    with open(input_path, "wb") as f:
        f.write(await file.read())

    JOBS[job_id] = {"status": "processing", "progress": 10}

    # 動画 → 音声
    if file.filename and file.filename.lower().endswith((".mp4", ".avi", ".mov", ".wmv")):
        audio_path = input_path + ".wav"
        subprocess.run(["ffmpeg", "-i", input_path, "-vn", audio_path, "-y"])
    else:
        audio_path = input_path

    result = model.transcribe(audio_path, language="ja")
    JOBS[job_id]["text"] = str(result["text"]).replace("。", "。\n")
    JOBS[job_id]["status"] = "completed"
    JOBS[job_id]["progress"] = 100

    return {"job_id": job_id, "status": "queued"}

@app.get("/api/status/{job_id}")
def status(job_id: str):
    return JOBS.get(job_id, {"status": "error"})

@app.get("/api/result/{job_id}")
def result(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        return {"error": "not found"}
    return {
        "job_id": job_id,
        "text": job.get("text", "")
    }