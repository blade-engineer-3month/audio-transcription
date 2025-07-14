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
        self.geometry("600x300")
        self.configure(bg="#f0f0f0")
        self.cancel_flag = False
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="音声ファイルを選択またはドラッグ＆ドロップ",
                 font=("Arial", 12, "bold"), bg="#f0f0f0").pack(pady=10)

        self.entry_file = tk.Entry(self, width=60, font=("Arial", 10))
        self.entry_file.pack(pady=5)
        self.entry_file.drop_target_register(DND_FILES)
        self.entry_file.dnd_bind('<<Drop>>', self.on_drop)

        frame = tk.Frame(self, bg="#f0f0f0")
        frame.pack(pady=5)
        tk.Button(frame, text="参照...", command=self.select_file).pack(side=tk.LEFT, padx=5)

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

    def select_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Audio Files", "*.mp3 *.m4a *.wav *.ogg")])
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, filepath)

    def start_transcription(self):
        self.cancel_flag = False
        self.start_btn["state"] = "disabled"
        self.cancel_btn["state"] = "normal"
        threading.Thread(target=self.transcribe_file).start()

    def cancel_process(self):
        self.cancel_flag = True

    def split_paragraphs(self, text):
        for _ in ["。", "."]:
            text = re.sub(r'([。\.])\s*', r'\1\n\n', text)
        return text.strip()

    def split_audio(self, y, sr, chunk_seconds=1800):
        total_seconds = int(len(y) / sr)
        chunks = []
        for start_sec in range(0, total_seconds, chunk_seconds):
            if self.cancel_flag:
                return []
            end_sec = min(start_sec + chunk_seconds, total_seconds)
            chunks.append(y[start_sec * sr:end_sec * sr])
        return chunks

    def transcribe_file(self):
        filepath = self.entry_file.get()
        if not filepath or not os.path.isfile(filepath):
            messagebox.showerror("エラー", "音声ファイルを選択してください。")
            return

        self.progress_bar["value"] = 10
        self.update_idletasks()

        FFMPEG_PATH = r"C:\ffmpeg\bin\ffmpeg.exe"
        filename, ext = os.path.splitext(filepath)
        if ext.lower() == ".m4a":
            mp3_file = filename + ".mp3"
            subprocess.call([FFMPEG_PATH, '-y', '-i', filepath, mp3_file])
            filepath = mp3_file

        y, sr = librosa.load(filepath, sr=16000)
        y_clean = nr.reduce_noise(y, sr=sr, stationary=True)
        chunks = self.split_audio(y_clean, sr, chunk_seconds=1800)
        if self.cancel_flag:
            self.reset_buttons()
            return

        model = whisper.load_model("small")
        all_text = ""
        for idx, chunk in enumerate(chunks):
            if self.cancel_flag:
                self.reset_buttons()
                return
            temp_wav = f"temp_chunk_{idx}.wav"
            sf.write(temp_wav, chunk, sr)
            result = model.transcribe(temp_wav, fp16=False)
            all_text += result["text"].strip() + "\n\n"
            os.remove(temp_wav)
            self.progress_bar["value"] = 10 + (80 / len(chunks)) * (idx + 1)
            self.update_idletasks()

        formatted_text = self.split_paragraphs(all_text)
        folder = os.path.dirname(filepath)
        dt_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.splitext(os.path.basename(filepath))[0]
        base_output = f"{base_name}_{dt_str}"

        if self.var_txt.get():
            with open(os.path.join(folder, f"{base_output}.txt"), "w", encoding="utf-8") as f:
                f.write(formatted_text)

        if self.var_docx.get():
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
            doc.save(os.path.join(folder, f"{base_output}.docx"))

        self.progress_bar["value"] = 100
        self.reset_buttons()
        messagebox.showinfo("完了", f"文字起こしが完了しました。\n保存先：\n{folder}")

    def reset_buttons(self):
        self.start_btn["state"] = "normal"
        self.cancel_btn["state"] = "disabled"
        self.progress_bar["value"] = 0


app = AudioTranscriberApp()
app.mainloop()
