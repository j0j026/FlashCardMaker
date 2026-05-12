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
        self.root.geometry("650x650")

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

        self.cat_label = tk.Label(self.input_frame, text="Category: ", font=("Arial", 11, "bold"), wraplength=550)
        self.cat_label.pack(pady=2)

        self.topic_label = tk.Label(self.input_frame, text="Topic: ", font=("Arial", 13, "bold"), fg="#2c3e50",
                                    wraplength=550)
        self.topic_label.pack(pady=5)

        self.text_box = tk.Text(self.input_frame, height=8, width=60, font=("Arial", 11), wrap="word", state="disabled")
        self.text_box.pack(pady=10, padx=10)

        # --- Navigation Buttons ---
        nav_frame = tk.Frame(self.input_frame)
        nav_frame.pack(pady=10)

        self.back_btn = tk.Button(nav_frame, text="← Back", command=self.go_back, state="disabled", width=10)
        self.back_btn.grid(row=0, column=0, padx=5)

        self.submit_btn = tk.Button(nav_frame, text="Submit & Save",
                                    command=self.save_current_step, bg="#27ae60", fg="white", state="disabled",
                                    width=20)
        self.submit_btn.grid(row=0, column=1, padx=5)

        self.skip_btn = tk.Button(nav_frame, text="Skip →", command=self.skip_step, state="disabled", width=10)
        self.skip_btn.grid(row=0, column=2, padx=5)

        self.text_box.bind("<Return>", self.handle_enter)

    def select_input(self):
        path = filedialog.askopenfilename(title="Select source .txt", filetypes=[("Text files", "*.txt")])
        if path:
            self.txt_path = path
            self.lbl_input.config(text=os.path.basename(path), fg="black")
            self.check_ready()

    def select_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self.pptx_path = path
            self.lbl_output.config(text=os.path.basename(path), fg="black")
            self.check_ready()

    def check_ready(self):
        if self.txt_path and self.pptx_path:
            self.btn_start.config(state="normal")

    def get_existing_slide_data(self, topic_name):
        """Searches for a slide by its hidden name and returns its definition text."""
        if not os.path.exists(self.pptx_path):
            return None

        try:
            prs = Presentation(self.pptx_path)
            for slide in prs.slides:
                if slide.name == topic_name:
                    # Our definition is always the 3rd shape added (index 2)
                    # We check if it exists and has text
                    if len(slide.shapes) >= 3 and slide.shapes[2].has_text_frame:
                        return slide.shapes[2].text_frame.text
            return None
        except:
            return None

    def start_session(self):
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

        # Find progress: start at the first topic that doesn't have a slide yet
        self.current_index = 0
        if os.path.exists(self.pptx_path):
            try:
                prs = Presentation(self.pptx_path)
                existing_names = [s.name for s in prs.slides]
                while self.current_index < len(self.raw_topics):
                    if self.raw_topics[self.current_index]['topic'] in existing_names:
                        self.current_index += 1
                    else:
                        break
            except:
                pass

        self.text_box.config(state="normal")
        self.submit_btn.config(state="normal")
        self.skip_btn.config(state="normal")
        self.btn_start.config(state="disabled")
        self.btn_input.config(state="disabled")
        self.btn_output.config(state="disabled")

        self.update_display()
        self.text_box.focus_force()

    def update_display(self):
        self.back_btn.config(state="normal" if self.current_index > 0 else "disabled")

        if self.current_index < len(self.raw_topics):
            item = self.raw_topics[self.current_index]
            self.progress_label.config(text=f"Topic {self.current_index + 1} of {len(self.raw_topics)}")
            self.cat_label.config(text=f"Category: {item['cat']}")
            self.topic_label.config(text=f"Topic: {item['topic']}")

            # CLEAR and FETCH existing content
            self.text_box.delete("1.0", tk.END)
            existing_text = self.get_existing_slide_data(item['topic'])
            if existing_text:
                self.text_box.insert("1.0", existing_text)
        else:
            messagebox.showinfo("Done!", "You've reached the end of the list!")
            self.root.destroy()

    def handle_enter(self, event):
        if not (event.state & 0x1):
            self.save_current_step()
            return "break"

    def go_back(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.update_display()

    def skip_step(self):
        self.current_index += 1
        self.update_display()

    def save_current_step(self):
        definition = self.text_box.get("1.0", tk.END).strip()
        if not definition:
            messagebox.showwarning("Warning", "Definition cannot be empty.")
            return

        item = self.raw_topics[self.current_index]
        self.add_or_update_slide(item['cat'], item['topic'], definition)

        self.current_index += 1
        self.update_display()

    def add_or_update_slide(self, category, topic, definition):
        if os.path.exists(self.pptx_path):
            try:
                prs = Presentation(self.pptx_path)
            except:
                prs = Presentation()
        else:
            prs = Presentation()

        # Check if slide already exists to update it
        target_slide = None
        for slide in prs.slides:
            if slide.name == topic:
                target_slide = slide
                break

        if target_slide:
            # Update existing shapes (assuming standard order: 0:Cat, 1:Topic, 2:Def)
            if len(target_slide.shapes) >= 3:
                target_slide.shapes[0].text_frame.text = category.upper()
                target_slide.shapes[1].text_frame.text = topic
                target_slide.shapes[2].text_frame.text = definition
        else:
            # Create new slide
            blank_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_layout)
            slide.name = topic

            # Category
            cat_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
            cat_frame = cat_shape.text_frame
            cat_frame.word_wrap = True
            cat_frame.text = category.upper()
            cat_frame.paragraphs[0].font.size = Pt(14)
            cat_frame.paragraphs[0].font.bold = True

            # Topic
            top_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(9), Inches(0.8))
            top_frame = top_shape.text_frame
            top_frame.word_wrap = True
            top_frame.text = topic
            top_frame.paragraphs[0].font.size = Pt(28)
            top_frame.paragraphs[0].font.bold = True

            # Definition
            def_shape = slide.shapes.add_textbox(Inches(0.5), Inches(1.6), Inches(9), Inches(5.4))
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
            messagebox.showerror("Error", "Close the PowerPoint file before saving!")


if __name__ == "__main__":
    root = tk.Tk()
    app = PPTXBuilderApp(root)
    root.mainloop()