import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

# When running as a PyInstaller bundle, files are extracted to a temp dir.
# Change to that directory so relative paths (ppt_template.pptx) still work.
if getattr(sys, "frozen", False):
    os.chdir(sys._MEIPASS)

from main import generate_all


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Sales Presentation Generator")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        self.root.configure(bg="#1a1a1a")

        # Title
        tk.Label(
            root, text="Sales Presentation Generator",
            font=("Segoe UI", 18, "bold"), fg="#f9d605", bg="#1a1a1a",
        ).pack(pady=(30, 5))

        tk.Label(
            root, text="Select your Excel file to generate presentations",
            font=("Segoe UI", 10), fg="#999999", bg="#1a1a1a",
        ).pack(pady=(0, 20))

        # File selection row
        frame = tk.Frame(root, bg="#1a1a1a")
        frame.pack(pady=5, padx=30, fill="x")

        self.file_var = tk.StringVar(value="No file selected")
        tk.Label(
            frame, textvariable=self.file_var,
            font=("Segoe UI", 9), fg="#cccccc", bg="#242424",
            anchor="w", padx=10, pady=8, relief="flat",
        ).pack(side="left", fill="x", expand=True)

        tk.Button(
            frame, text="Browse", command=self.browse,
            font=("Segoe UI", 9, "bold"), fg="#1a1a1a", bg="#f9d605",
            relief="flat", padx=16, pady=6, cursor="hand2",
        ).pack(side="right", padx=(8, 0))

        # Generate button
        self.btn = tk.Button(
            root, text="Generate Presentations", command=self.generate,
            font=("Segoe UI", 12, "bold"), fg="#1a1a1a", bg="#f9d605",
            relief="flat", padx=30, pady=10, cursor="hand2",
        )
        self.btn.pack(pady=30)

        # Status label
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(
            root, textvariable=self.status_var,
            font=("Segoe UI", 9), fg="#999999", bg="#1a1a1a",
        )
        self.status_label.pack()

        self.input_path = None

    def browse(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if path:
            self.input_path = path
            self.file_var.set(os.path.basename(path))

    def generate(self):
        if not self.input_path:
            messagebox.showwarning("No file", "Please select an Excel file first.")
            return

        self.btn.config(state="disabled", text="Generating...")
        self.status_var.set("Working... please wait")
        self.root.update()

        thread = threading.Thread(target=self._run_generation, daemon=True)
        thread.start()

    def _run_generation(self):
        try:
            output_dir = os.path.join(os.path.dirname(self.input_path), "output")
            generated = generate_all(self.input_path, output_dir)

            if not generated:
                self.root.after(0, lambda: messagebox.showwarning(
                    "No output",
                    "No presentations generated.\nCheck that your Excel has partners with CreatePPT = True.",
                ))
            else:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Done!",
                    f"{len(generated)} presentation(s) saved to:\n{output_dir}",
                ))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, self._reset_button)

    def _reset_button(self):
        self.btn.config(state="normal", text="Generate Presentations")
        self.status_var.set("")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
