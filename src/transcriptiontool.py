import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import whisper
from docx import Document
from docx.shared import Pt, Inches
import subprocess
import noisereduce as nr
import soundfile as sf
import librosa
import threading
import re

os.environ["PATH"] += os.pathsep + r"C:\ffmpeg\bin"


class AudioTranscriberApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("音声文字起こし")
        self.geometry("600x350")
        self.configure(bg="#f0f0f0")
        self.cancel_flag = False
        self.chunk_seconds = 300  # 5分ごと分割
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="音声/動画ファイルを選択またはドラッグ＆ドロップ",
                 font=("Arial", 12, "bold"), bg="#f0f0f0").pack(pady=12)

        self.entry_file = tk.Entry(self, width=60, font=("Arial", 10))
        self.entry_file.pack(pady=5)
        self.entry_file.drop_target_register(DND_FILES)
        self.entry_file.dnd_bind('<<Drop>>', self.on_drop)

        frame = tk.Frame(self, bg="#f0f0f0")
        frame.pack(pady=5)
        tk.Button(frame, text="参照...", command=self.select_files).pack(side=tk.LEFT, padx=5)

        self.var_txt = tk.BooleanVar(value=True)
        self.var_docx = tk.BooleanVar(value=True)
        tk.Checkbutton(self, text="txt出力", variable=self.var_txt, bg="#f0f0f0").pack(anchor=tk.W, padx=30)
        tk.Checkbutton(self, text="Word出力", variable=self.var_docx, bg="#f0f0f0").pack(anchor=tk.W, padx=30)

        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=400)
        self.progress_bar.pack(pady=10)

        btn_frame = tk.Frame(self, bg="#f0f0f0")
        btn_frame.pack(pady=5)
        self.start_btn = tk.Button(btn_frame, text="文字起こし実行", command=self.start_transcription, bg="#0099cc", fg="white")
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.cancel_btn = tk.Button(btn_frame, text="中断", command=self.cancel_process, state="disabled")
        self.cancel_btn.pack(side=tk.LEFT, padx=5)

    def on_drop(self, event):
        path = event.data.strip('{}')
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, path)

    def select_files(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("Audio/Video Files", "*.mp3 *.m4a *.wav *.ogg *.mp4 *.avi *.mov *.wmv")])
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, " ".join(filepaths))

    def start_transcription(self):
        self.cancel_flag = False
        self.start_btn["state"] = "disabled"
        self.cancel_btn["state"] = "normal"
        threading.Thread(target=self.transcribe_files).start()

    def cancel_process(self):
        self.cancel_flag = True

    def split_paragraphs(self, text):
        for _ in ["。", "."]:
            text = re.sub(r'([。\.])\s*', r'\1\n\n', text)
        return text.strip()

    def split_audio(self, y, sr, chunk_seconds=300):
        total_seconds = int(len(y) / sr)
        chunks = []
        for start_sec in range(0, total_seconds, chunk_seconds):
            if self.cancel_flag:
                return []
            end_sec = min(start_sec + chunk_seconds, total_seconds)
            chunks.append(y[start_sec * sr:end_sec * sr])
        return chunks

    def transcribe_files(self):
        filepaths = self.entry_file.get().split()
        audio_files = [path for path in filepaths if path.lower().endswith((".mp3", ".wav", ".m4a", ".ogg", ".mp4", ".avi", ".mov", ".wmv"))]

        # 全体チャンク数カウント（進捗バー計算用）
        all_chunks = 0
        chunk_counts = {}
        for path in audio_files:
            y, sr = self.load_audio_file(path)
            if y is None:
                continue
            chunks = self.split_audio(y, sr, chunk_seconds=self.chunk_seconds)
            chunk_counts[path] = len(chunks)
            all_chunks += len(chunks)

        if all_chunks == 0:
            self.reset_buttons()
            return

        done_chunks = 0
        model = whisper.load_model("small")

        for filepath in audio_files:
            if self.cancel_flag:
                break
            done_chunks = self.transcribe_file(filepath, model, all_chunks, done_chunks, chunk_counts[filepath])

        self.reset_buttons()

    def load_audio_file(self, filepath):
        FFMPEG_PATH = r"C:\ffmpeg\bin\ffmpeg.exe"
        folder = os.path.dirname(filepath)
        filename_noext = os.path.splitext(os.path.basename(filepath))[0]
        mp3_file = os.path.join(folder, f"{filename_noext}_convert.mp3")

        if os.path.splitext(filepath)[1].lower() in [".m4a", ".mp4", ".avi", ".mov", ".wmv"]:
            # ★ コマンドを文字列にしてダブルクォーテーションで囲む
            command = f'"{FFMPEG_PATH}" -y -i "{filepath}" "{mp3_file}"'
            result = subprocess.run(command, shell=True, capture_output=True)
            if result.returncode != 0:
                messagebox.showerror("変換失敗", f"ffmpegエラー:\n{result.stderr.decode('utf-8')}")
                return None, None
            filepath = mp3_file

        if not os.path.exists(filepath):
            messagebox.showerror("エラー", f"ファイルが存在しません:\n{filepath}")
            return None, None

        y, sr = librosa.load(filepath, sr=16000)
        y_clean = nr.reduce_noise(y, sr=sr, stationary=True)

        # 一時ファイル削除
        if filepath.endswith("_convert.mp3"):
            os.remove(filepath)

        return y_clean, sr


    def transcribe_file(self, filepath, model, all_chunks, done_chunks, current_file_chunks):
        y, sr = self.load_audio_file(filepath)
        if y is None:
            return done_chunks
        chunks = self.split_audio(y, sr, chunk_seconds=self.chunk_seconds)
        if self.cancel_flag:
            return done_chunks

        all_text = ""
        for idx, chunk in enumerate(chunks):
            if self.cancel_flag:
                return done_chunks
            temp_wav = f"temp_chunk_{idx}.wav"
            try:
                sf.write(temp_wav, chunk, sr)
                result = model.transcribe(temp_wav, fp16=False)
                all_text += result["text"].strip() + "\n\n"
            finally:
                if os.path.exists(temp_wav):
                    os.remove(temp_wav)

            done_chunks += 1
            progress = 100 * done_chunks / all_chunks
            self.progress_bar["value"] = progress
            self.update_idletasks()

        formatted_text = self.split_paragraphs(all_text)
        folder = os.path.dirname(filepath)
        dt_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.splitext(os.path.basename(filepath))[0]
        base_output = f"{base_name}_{dt_str}"

        # ファイル名重複チェック
        def get_nonexistent_path(path, ext):
            i = 1
            candidate = f"{path}.{ext}"
            while os.path.exists(candidate):
                candidate = f"{path}_{i}.{ext}"
                i += 1
            return candidate

        if self.var_txt.get():
            txt_path = get_nonexistent_path(os.path.join(folder, base_output), "txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(formatted_text)

        if self.var_docx.get():
            docx_path = get_nonexistent_path(os.path.join(folder, base_output), "docx")
            doc = Document()
            section = doc.sections[0]
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Meiryo UI'
            font.size = Pt(11)
            for paragraph in formatted_text.split("\n\n"):
                para = doc.add_paragraph(paragraph)
                para.style = doc.styles["Normal"]
            doc.save(docx_path)
        messagebox.showinfo("完了", f"文字起こしが完了しました。\n保存先：\n{folder}")
        return done_chunks

    def reset_buttons(self):
        self.start_btn["state"] = "normal"
        self.cancel_btn["state"] = "disabled"
        self.progress_bar["value"] = 0


if __name__ == "__main__":
    app = AudioTranscriberApp()
    app.mainloop()
