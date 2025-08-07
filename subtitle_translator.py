import tkinter as tk
from tkinter import filedialog
import os
import re
import time
from datetime import datetime
import ollama

# Add VLC to the DLL search path
vlc_path = r"C:\Program Files\VideoLAN\VLC"
if os.path.exists(vlc_path):
    os.add_dll_directory(vlc_path)

import vlc

class SubtitleTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Subtitle Translator with Ollama")
        self.root.geometry("800x600")

        # VLC instance
        self.instance = vlc.Instance()
        self.player = self.instance.media_player_new()
        self.media = None

        # GUI elements
        self.canvas = tk.Canvas(self.root, bg="black")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.subtitle_text = tk.Text(self.root, height=3, bg="white", font=("Arial", 12))
        self.subtitle_text.pack(fill=tk.X, padx=10, pady=5)
        self.subtitle_text.config(state="disabled")

        self.load_button = tk.Button(self.root, text="Load Video and Subtitles", command=self.load_files)
        self.load_button.pack(pady=5)

        self.status_label = tk.Label(self.root, text="", font=("Arial", 10))
        self.status_label.pack(pady=5)

        # Subtitle data
        self.subtitles = []
        self.current_subtitle_index = -1

        # Bind click event for subtitle words
        self.subtitle_text.tag_configure("word", foreground="blue", underline=True)
        self.subtitle_text.bind("<Button-1>", self.handle_word_click)

        # VLC event manager
        self.event_manager = self.player.event_manager()
        self.event_manager.event_attach(vlc.EventType.MediaPlayerTimeChanged, self.update_subtitles)

        # Ollama model (e.g., llama3.1 or mistral)
        self.ollama_model = "llama3.1"  # Change to your preferred model

    def load_files(self):
        video_file = filedialog.askopenfilename(filetypes=[("Video files", "*.mp4 *.avi *.mkv")])
        subtitle_file = filedialog.askopenfilename(filetypes=[("Subtitle files", "*.srt")])
        if video_file and subtitle_file:
            self.media = self.instance.media_new(video_file)
            self.player.set_media(self.media)
            self.player.set_hwnd(self.canvas.winfo_id())  # Set video to canvas
            self.player.play()
            self.load_subtitles(subtitle_file)
            self.status_label.config(text="Video and subtitles loaded.")
        else:
            self.status_label.config(text="Please select both video and subtitle files.")

    def load_subtitles(self, subtitle_file):
        self.subtitles = []
        try:
            with open(subtitle_file, "r", encoding="utf-8") as f:
                content = f.read()
                subtitle_blocks = content.strip().split("\n\n")
                for block in subtitle_blocks:
                    lines = block.strip().split("\n")
                    if len(lines) >= 3:
                        time_str = lines[1]
                        start, end = self.parse_srt_time(time_str)
                        text = " ".join(lines[2:])
                        self.subtitles.append({"start": start, "end": end, "text": text})
            self.status_label.config(text=f"Loaded {len(self.subtitles)} subtitles.")
        except Exception as e:
            self.status_label.config(text=f"Error loading subtitles: {str(e)}")

    def parse_srt_time(self, time_str):
        # Parse SRT time format (e.g., 00:00:01,000 --> 00:00:02,000)
        start, end = time_str.split(" --> ")
        start_ms = self.time_to_ms(start)
        end_ms = self.time_to_ms(end)
        return start_ms, end_ms

    def time_to_ms(self, time_str):
        # Convert SRT time (HH:MM:SS,mmm) to milliseconds
        time_obj = datetime.strptime(time_str.replace(",", "."), "%H:%M:%S.%f")
        return int(time_obj.hour * 3600000 + time_obj.minute * 60000 +
                   time_obj.second * 1000 + time_obj.microsecond / 1000)

    def update_subtitles(self, event):
        if not self.subtitles:
            return
        current_time = self.player.get_time()  # VLC time in milliseconds
        self.subtitle_text.config(state="normal")
        self.subtitle_text.delete("1.0", tk.END)

        for i, subtitle in enumerate(self.subtitles):
            if subtitle["start"] <= current_time <= subtitle["end"]:
                self.current_subtitle_index = i
                words = subtitle["text"].split()
                for word in words:
                    clean_word = re.sub(r"[.,!?]", "", word.lower())
                    # Assume all words are potentially translatable
                    self.subtitle_text.insert(tk.END, word + " ", ("word", clean_word))
                break
            else:
                self.current_subtitle_index = -1

        self.subtitle_text.config(state="disabled")

    def handle_word_click(self, event):
        self.subtitle_text.config(state="normal")
        try:
            tag = self.subtitle_text.tag_names(tk.CURRENT)[1]  # Get the clean word tag
            translation = self.translate_word(tag)
            self.status_label.config(text=f"{tag}: {translation}")
        except IndexError:
            pass
        self.subtitle_text.config(state="disabled")

    def translate_word(self, word):
        try:
            # Query Ollama for translation
            prompt = f"Translate the Spanish word '{word}' to English. Provide only the translated word or a short phrase."
            response = ollama.generate(model=self.ollama_model, prompt=prompt)
            return response["response"].strip()
        except Exception as e:
            return f"Error translating: {str(e)}"

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = SubtitleTranslatorApp(root)
    app.run()
