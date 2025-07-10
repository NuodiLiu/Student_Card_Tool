import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

# === Apple风格颜色 ===
PRIMARY_COLOR = "#007AFF"
SECONDARY_COLOR = "#8E8E93"
ERROR_COLOR = "#FF3B30"
BACKGROUND_COLOR = "#F2F2F7"
HIGHLIGHT_COLOR = "#D6E7FF"

INPUT_FOLDER = "input"
INPUT_FILENAME = "selected.xlsx"

class MainWindow(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Student Card Generator")
        self.geometry("600x450")
        self.minsize(600, 450)
        self._configure_styles()
        self.configure(bg=BACKGROUND_COLOR)

        os.makedirs(INPUT_FOLDER, exist_ok=True)
        self.selected_file = None

        self._create_widgets()

    def _configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # 主按钮样式（蓝色）
        style.configure("Accent.TButton",
                        foreground="white",
                        background=PRIMARY_COLOR,
                        font=("Helvetica", 12, "bold"),
                        padding=6)
        style.map("Accent.TButton",
                  background=[("active", "#005BB5")])

        # 辅助按钮样式（灰色）
        style.configure("Secondary.TButton",
                        foreground="black",
                        background=SECONDARY_COLOR,
                        font=("Helvetica", 12),
                        padding=6)
        style.map("Secondary.TButton",
                  background=[("active", "#6C6C70")])
        
        style.configure("Green.Horizontal.TProgressbar", 
                troughcolor="white",         # track color
                background="#4CAF50")       # bar color
    
    def _choose_output_dir(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir.set(folder)

    def _create_widgets(self):
        title_label = tk.Label(self, text="Student Card Generator", font=("Helvetica", 20, "bold"),
                               bg=BACKGROUND_COLOR, fg=PRIMARY_COLOR)
        title_label.pack(pady=20)

        # === 文件选择 + 拖拽区域 ===
        self.drop_frame = tk.Frame(self, width=500, height=80, bg="white", highlightthickness=2, highlightbackground=SECONDARY_COLOR)
        self.drop_frame.pack(pady=10)
        self.drop_frame.pack_propagate(False)
        self.drop_frame_label = tk.Label(self.drop_frame, text="Drag Excel file here or click to browse", fg=SECONDARY_COLOR, bg="white")
        self.drop_frame_label.pack(expand=True)
        self.drop_frame.bind("<Button-1>", lambda e: self._select_file())

        # 拖拽事件绑定
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self._on_drop)
        self.drop_frame.dnd_bind('<<DragEnter>>', lambda e: self._highlight_drop(True))
        self.drop_frame.dnd_bind('<<DragLeave>>', lambda e: self._highlight_drop(False))

        # Mode
        self.mode_var = tk.StringVar(value="English Program")
        ttk.Label(self, text="Select Mode:").pack()
        ttk.Combobox(self, textvariable=self.mode_var, values=["English Program", "FS"], state="readonly").pack(pady=5)

        # Theme
        self.theme_var = tk.StringVar(value="green-yellow")
        ttk.Label(self, text="Select Theme:").pack()
        ttk.Combobox(self, textvariable=self.theme_var, values=["green-yellow", "blue-white"], state="readonly").pack(pady=5)

        # Count
        self.count_var = tk.IntVar(value=15)
        ttk.Label(self, text="Students per Card:").pack()
        ttk.Entry(self, textvariable=self.count_var, width=5).pack(pady=5)

        # allow user to select output path
        self.output_dir = tk.StringVar(value=os.path.abspath("output"))
        ttk.Label(self, text="Output Folder:").pack()
        path_frame = tk.Frame(self, bg=BACKGROUND_COLOR)
        path_frame.pack(pady=5)

        path_entry = ttk.Entry(path_frame, textvariable=self.output_dir, width=50)
        path_entry.pack(side=tk.LEFT)

        ttk.Button(path_frame, text="Browse", command=self._choose_output_dir).pack(side=tk.LEFT, padx=5)

        # Generate
        # 创建一个水平容器 Frame
        button_row = tk.Frame(self, bg=None)
        button_row.pack(pady=10)  # 加点间距，别挨太近

        # 左侧按钮
        ttk.Button(button_row, text="Generate Cards", style="Accent.TButton", command=self._on_generate).pack(side="left", padx=5)

        # 右侧按钮
        ttk.Button(button_row, text="Open Output Folder", style="Secondary.TButton", command=self._open_output_folder).pack(side="left", padx=5)

        # Progress bar
        self.progress = ttk.Progressbar(self, 
                                length=400, 
                                mode='determinate', 
                                style="Green.Horizontal.TProgressbar")
        self.progress.pack(pady=10)
        self.progress["value"] = 0
        self.progress["maximum"] = 100
        self.progress.pack_forget()

    def _highlight_drop(self, entering: bool):
        if entering:
            self.drop_frame.config(highlightbackground=PRIMARY_COLOR, bg=HIGHLIGHT_COLOR)
        else:
            self.drop_frame.config(highlightbackground=SECONDARY_COLOR, bg="white")

    def _select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self._save_file_to_input(file_path)

    def _on_drop(self, event):
        file_path = event.data.strip().strip('{}')  # 防止 mac/windows 拖拽路径出现大括号
        if os.path.isfile(file_path) and file_path.endswith(".xlsx"):
            self._save_file_to_input(file_path)
        else:
            messagebox.showerror("Invalid File", "Please drop a valid .xlsx file.")
        self._highlight_drop(False)

    def _save_file_to_input(self, file_path: str):
        # 清理旧缓存
        target_path = os.path.join(INPUT_FOLDER, INPUT_FILENAME)
        if os.path.exists(target_path):
            os.remove(target_path)
        shutil.copy(file_path, target_path)
        self.selected_file = target_path
        self.drop_frame_label.config(text=os.path.basename(file_path), fg="black")
        messagebox.showinfo("File Ready", f"✅ Excel file saved to input:\n{target_path}")

    def _on_generate(self):
        if not self.selected_file:
            messagebox.showerror("Missing File", "Please select an Excel file first.")
            return

        mode = self.mode_var.get()
        theme = self.theme_var.get()
        students_per_card = self.count_var.get()

        try:
            from viewmodels.processor import load_classgroups_from_excel
            from views.chart_renderer import render_class_card

            class_groups = load_classgroups_from_excel(
                self.selected_file, mode, students_per_card
            )

            # 设置进度条
            self.progress["value"] = 0
            self.progress["maximum"] = len(class_groups)
            self.progress.pack()

            # render card
            output_path = self.output_dir.get()
            for i, group in enumerate(class_groups, 1):
                render_class_card(group, output_path, card_index=i, theme=theme)
                self.progress["value"] = i
                self.update_idletasks()

            self.progress.pack_forget()
            messagebox.showinfo("✅ Success", "All cards generated successfully!")
            self._open_output_folder()
        except Exception as e:
            self.progress.pack_forget()
            messagebox.showerror("❌ Error", f"Failed to generate cards:\n{e}")



    def _open_output_folder(self):
        output_path = os.path.abspath("output")
        os.makedirs(output_path, exist_ok=True)
        if os.name == "nt":
            os.startfile(output_path)
        elif os.uname().sysname == "Darwin":
            os.system(f"open '{output_path}'")
        else:
            os.system(f"xdg-open '{output_path}'")


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
