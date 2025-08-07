import tkinter as tk
from tkinter import filedialog, messagebox
import vlc
import re
from datetime import datetime
import ollama
import time
import subprocess
import tempfile
import os
import openpyxl
import threading

from subtitle_utils import parse_srt_file, srt_time_to_seconds
from translator import translate_word
from excel_logger import setup_excel_file, save_translation

class SubtitleTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Subtitle Translator with Embedded Subtitles")
        self.root.geometry("1000x700")
        self.theme = "light"

        # Caching and defaults
        self.translation_cache = {}
        self.target_lang = tk.StringVar(value="English")

        # VLC setup
        self.instance = vlc.Instance()
        self.player = self.instance.media_player_new()
        self.media = None
        self.video_path = None

        # GUI elements
        self.build_gui()
        self.setup_excel_file()

        # Subtitle tracking
        self.subtitles = []
        self.last_subtitle_text = None

        # Subtitle rendering
        self.event_manager = self.player.event_manager()
        self.event_manager.event_attach(vlc.EventType.MediaPlayerTimeChanged, self.update_subtitles)

    def build_gui(self):
        self.video_frame = tk.Frame(self.root, bg="black")
        self.video_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.video_frame, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.update()

        self.subtitle_overlay = tk.Frame(self.video_frame, bg="black", highlightthickness=0)
        self.subtitle_overlay.place(relx=0.5, rely=0.85, anchor="center")

        # Control bindings
        self.root.bind('<space>', lambda e: self.toggle_play_pause())
        self.root.bind('<Right>', lambda e: self.seek_relative(10))
        self.root.bind('<Left>', lambda e: self.seek_relative(-10))
        self.root.bind('f', self.toggle_fullscreen)
        self.root.bind('t', self.toggle_translation_overlay)

        # Controls
        self.controls_frame = tk.Frame(self.root)
        self.controls_frame.pack()

        tk.Button(self.controls_frame, text="Load Video", command=self.load_video).pack(side=tk.LEFT)
        tk.Button(self.controls_frame, text="Load Subtitles", command=self.load_subtitles).pack(side=tk.LEFT)
        tk.Button(self.controls_frame, text="Extract Embedded Subs", command=self.extract_embedded_subtitles).pack(side=tk.LEFT)
        tk.OptionMenu(self.controls_frame, self.target_lang, "English", "French", "German", "Arabic").pack(side=tk.LEFT)
        tk.Button(self.controls_frame, text="Toggle Theme", command=self.toggle_theme).pack(side=tk.LEFT)

        self.status_label = tk.Label(self.root, text="Ready")
        self.status_label.pack()

    def toggle_play_pause(self):
        if self.player.is_playing():
            self.player.pause()
        else:
            self.player.play()
            self.root.after(1000, self.extract_embedded_subtitles)

    def load_video(self):
        video_file = filedialog.askopenfilename(filetypes=[("Video files", "*.mp4 *.mkv *.avi")])
        if video_file:
            self.video_path = video_file
            self.media = self.instance.media_new(video_file)
            self.player.set_media(self.media)
            self.player.set_hwnd(self.canvas.winfo_id())
            self.player.play()

    def load_subtitles(self):
        srt_file = filedialog.askopenfilename(filetypes=[("Subtitle files", "*.srt")])
        if srt_file:
            self.subtitles = parse_srt_file(srt_file)
            self.status_label.config(text=f"Loaded external subtitles")

    def extract_embedded_subtitles(self):
        if not self.video_path:
            messagebox.showwarning("Warning", "Please load a video first.")
            return

        try:
            # Use ffprobe to get detailed subtitle stream info
            probe_cmd = [
                "ffprobe", "-v", "error", "-select_streams", "s", "-show_entries",
                "stream=index:stream_tags=language", "-of", "csv=p=0", self.video_path
            ]
            probe_result = subprocess.run(probe_cmd, capture_output=True, text=True)
            lines = probe_result.stdout.strip().split("
")
            tracks = [(line.split(',')[0], line.split(',')[1] if len(line.split(',')) > 1 else "Unknown")
                      for line in lines if line]

            if not tracks:
                self.status_label.config(text="No embedded subtitle tracks found.")
                return

            dialog = tk.Toplevel(self.root)
            dialog.title("Choose Subtitle Track")
            tk.Label(dialog, text="Select subtitle track to extract:").pack(pady=5)
            selected_idx = tk.StringVar(value=tracks[0][0])

            for index, lang in tracks:
                tk.Radiobutton(dialog, text=f"Stream {index} ({lang})", variable=selected_idx, value=index).pack(anchor="w")

            def on_select():
                stream_index = selected_idx.get()
                srt_path = os.path.join(tempfile.gettempdir(), f"extracted_subs_{os.getpid()}.srt")
                cmd = ["ffmpeg", "-y", "-i", self.video_path, "-map", f"0:{stream_index}", srt_path]
                result = subprocess.run(cmd, stderr=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
                if os.path.exists(srt_path) and os.path.getsize(srt_path) > 0:
                    self.subtitles = parse_srt_file(srt_path)
                    self.status_label.config(text=f"Loaded subtitle stream {stream_index}")
                else:
                    self.status_label.config(text="Failed to extract selected subtitle track.")
                dialog.destroy()

            tk.Button(dialog, text="Extract", command=on_select).pack(pady=10)

        except Exception as e:
            self.status_label.config(text=f"Subtitle extraction error: {e}")

    def update_subtitles(self, event):
        current_time = self.player.get_time() / 1000
        subtitle_line = ""
        for sub in self.subtitles:
            if sub['start'] <= current_time <= sub['end']:
                subtitle_line = sub['text']
                break

        if subtitle_line == self.last_subtitle_text:
            return

        self.last_subtitle_text = subtitle_line

        for widget in self.subtitle_overlay.winfo_children():
            widget.destroy()

        if subtitle_line:
            words = subtitle_line.split()
            for word in words:
                lbl = tk.Label(self.subtitle_overlay, text=word, font=("Segoe UI", self.get_scaled_font(), "bold"), fg="#fff", bg="#222", cursor="hand2")
                lbl.pack(side=tk.LEFT, padx=2)
                lbl.bind("<Button-1>", lambda e, w=word.lower(): self.translate_word_async(w, subtitle_line, e.widget))

    def get_scaled_font(self):
        width = self.root.winfo_width()
        return max(14, int(width / 60))

    def translate_word_async(self, word, sentence, widget):
        threading.Thread(target=self.translate_word_logic, args=(word, sentence, widget)).start()

    def translate_word_logic(self, word, sentence, widget):
        if word in self.translation_cache:
            translation = self.translation_cache[word]
        else:
            translation = translate_word(word, self.target_lang.get())
            self.translation_cache[word] = translation
            save_translation(word, translation, sentence)

        self.show_tooltip(widget, translation)

    def show_tooltip(self, widget, text):
        tooltip = tk.Toplevel(self.root)
        tooltip.wm_overrideredirect(True)
        tooltip.geometry(f"+{widget.winfo_rootx()}+{widget.winfo_rooty()-30}")
        tk.Label(tooltip, text=text, bg="#fffbe6", font=("Segoe UI", 12, "italic"), relief="solid", borderwidth=1).pack()
        self.root.after(3000, tooltip.destroy)

    def seek_relative(self, seconds):
        new_time = max(0, self.player.get_time() // 1000 + seconds)
        self.player.set_time(new_time * 1000)

    def toggle_fullscreen(self, event=None):
        self.root.attributes("-fullscreen", not self.root.attributes("-fullscreen"))

    def toggle_translation_overlay(self, event=None):
        if self.subtitle_overlay.winfo_viewable():
            self.subtitle_overlay.place_forget()
        else:
            self.subtitle_overlay.place(relx=0.5, rely=0.85, anchor="center")

    def toggle_theme(self):
        if self.theme == "light":
            self.root.configure(bg="#1e1e1e")
            self.status_label.config(bg="#1e1e1e", fg="white")
            self.theme = "dark"
        else:
            self.root.configure(bg="#f4f4f4")
            self.status_label.config(bg="#f4f4f4", fg="black")
            self.theme = "light"

    def setup_excel_file(self):
        setup_excel_file()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = SubtitleTranslatorApp(root)
    app.run()
