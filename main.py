import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE


class PPTXBuilderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PPTX Definition Builder")
        self.root.geometry("650x600")

        self.txt_path = ""
        self.pptx_path = ""

        self.raw_topics = []
        self.current_index = 0

        self.setup_ui()

    def setup_ui(self):
        # --- File Selection Section ---
        file_frame = tk.LabelFrame(self.root, text="Step 1: File Setup", padx=10, pady=10)
        file_frame.pack(pady=10, padx=20, fill="x")

        self.btn_input = tk.Button(file_frame, text="Select Input (.txt)", command=self.select_input, width=20)
        self.btn_input.grid(row=0, column=0, pady=5)
        self.lbl_input = tk.Label(file_frame, text="No file selected", fg="grey", anchor="w")
        self.lbl_input.grid(row=0, column=1, padx=10, sticky="w")

        self.btn_output = tk.Button(file_frame, text="Select/Create Output (.pptx)", command=self.select_output,
                                    width=20)
        self.btn_output.grid(row=1, column=0, pady=5)
        self.lbl_output = tk.Label(file_frame, text="No location selected", fg="grey", anchor="w")
        self.lbl_output.grid(row=1, column=1, padx=10, sticky="w")

        self.btn_start = tk.Button(file_frame, text="START DEFINING", command=self.start_session,
                                   bg="#3498db", fg="white", state="disabled", font=("Arial", 10, "bold"))
        self.btn_start.grid(row=2, column=0, columnspan=2, pady=10, sticky="we")

        # --- Definition Input Section ---
        self.input_frame = tk.LabelFrame(self.root, text="Step 2: Definitions", padx=10, pady=10)
        self.input_frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.progress_label = tk.Label(self.input_frame, text="Waiting for files...", fg="blue")
        self.progress_label.pack()

        self.cat_label = tk.Label(self.input_frame, text="Category: ", font=("Arial", 11, "bold"))
        self.cat_label.pack(pady=2)

        self.topic_label = tk.Label(self.input_frame, text="Topic: ", font=("Arial", 13, "bold"), fg="#2c3e50")
        self.topic_label.pack(pady=5)

        self.text_box = tk.Text(self.input_frame, height=8, width=60, font=("Arial", 11), wrap="word", state="disabled")
        self.text_box.pack(pady=10, padx=10)

        self.submit_btn = tk.Button(self.input_frame, text="Submit & Save Slide",
                                    command=self.save_current_step, bg="#27ae60", fg="white", state="disabled")
        self.submit_btn.pack(pady=5)

        self.text_box.bind("<Return>", self.handle_enter)

    def select_input(self):
        path = filedialog.askopenfilename(title="Select source .txt", filetypes=[("Text files", "*.txt")])
        if path:
            self.txt_path = path
            self.lbl_input.config(text=os.path.basename(path), fg="black")
            self.check_ready()

    def select_output(self):
        # We use asksaveasfilename so they can pick an existing file or name a new one
        path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self.pptx_path = path
            self.lbl_output.config(text=os.path.basename(path), fg="black")
            self.check_ready()

    def check_ready(self):
        if self.txt_path and self.pptx_path:
            self.btn_start.config(state="normal")

    def get_existing_topics(self):
        """Reads the PowerPoint and returns a set of topic titles already present."""
        existing_topics = set()
        if os.path.exists(self.pptx_path):
            try:
                prs = Presentation(self.pptx_path)
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            # We check every text box on the slide for a match
                            text = shape.text.strip()
                            existing_topics.add(text)
            except Exception as e:
                print(f"Note: Could not read existing PPTX (it might be empty or new): {e}")
        return existing_topics

    def start_session(self):
        # 1. Parse text file
        self.raw_topics = []
        try:
            with open(self.txt_path, "r", encoding="utf-8") as f:
                curr_cat = "General"
                for line in f:
                    line = line.strip()
                    if not line: continue
                    if line.startswith("●"):
                        self.raw_topics.append({"cat": curr_cat, "topic": line.lstrip("● ").strip()})
                    else:
                        curr_cat = line
        except Exception as e:
            messagebox.showerror("Error", f"Could not read text file: {e}")
            return

        # 2. Check PowerPoint for progress
        existing = self.get_existing_topics()

        # 3. Find first topic not in the PowerPoint
        self.current_index = 0
        while self.current_index < len(self.raw_topics):
            if self.raw_topics[self.current_index]['topic'] in existing:
                self.current_index += 1
            else:
                break

        # 4. Enable UI
        self.text_box.config(state="normal")
        self.submit_btn.config(state="normal")
        self.btn_start.config(state="disabled")
        self.btn_input.config(state="disabled")
        self.btn_output.config(state="disabled")

        self.update_display()
        self.text_box.focus_force()

    def update_display(self):
        if self.current_index < len(self.raw_topics):
            item = self.raw_topics[self.current_index]
            self.progress_label.config(text=f"Topic {self.current_index + 1} of {len(self.raw_topics)}")
            self.cat_label.config(text=f"Category: {item['cat']}")
            self.topic_label.config(text=f"Topic: {item['topic']}")
            self.text_box.delete("1.0", tk.END)
        else:
            messagebox.showinfo("Done!", "All topics from the text file are present in the PowerPoint!")
            self.root.destroy()

    def handle_enter(self, event):
        if not (event.state & 0x1):  # Shift not held
            self.save_current_step()
            return "break"

    def save_current_step(self):
        definition = self.text_box.get("1.0", tk.END).strip()
        if not definition: return

        item = self.raw_topics[self.current_index]
        self.add_slide_to_pptx(item['cat'], item['topic'], definition)

        # Move to next
        self.current_index += 1
        self.update_display()

    def add_slide_to_pptx(self, category, topic, definition):
        if os.path.exists(self.pptx_path):
            try:
                prs = Presentation(self.pptx_path)
            except:
                prs = Presentation()
        else:
            prs = Presentation()

        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Title/Category Bar
        cat_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
        cat_frame = cat_shape.text_frame
        cat_frame.text = category.upper()
        cat_frame.paragraphs[0].font.size = Pt(14)
        cat_frame.paragraphs[0].font.bold = True

        # Topic Bar
        top_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(9), Inches(0.6))
        top_frame = top_shape.text_frame
        top_frame.text = topic
        top_frame.paragraphs[0].font.size = Pt(28)
        top_frame.paragraphs[0].font.bold = True

        # Definition Box
        def_shape = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5.5))
        def_frame = def_shape.text_frame
        def_frame.word_wrap = True
        def_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        p = def_frame.paragraphs[0]
        p.text = definition
        p.font.size = Pt(20)
        def_frame.vertical_anchor = MSO_ANCHOR.TOP

        try:
            prs.save(self.pptx_path)
        except PermissionError:
            messagebox.showerror("Error",
                                 "Could not save! Please close the PowerPoint file if it is open in another program.")


if __name__ == "__main__":
    root = tk.Tk()
    app = PPTXBuilderApp(root)
    root.mainloop()