import os
import tempfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from pdf2docx import Converter
from docx2pdf import convert
import platform

# Define colors for dark mode
BG_COLOR = "#2E2E2E"          # Dark grey
FG_COLOR = "#FFFFFF"          # White text for labels and buttons
BUTTON_COLOR = "#1E90FF"      # Dodger Blue
HIGHLIGHT_COLOR = "#FFD700"   # Gold
ENTRY_BG_COLOR = "#FFFFFF"    # White background for entries
ENTRY_FG_COLOR = "#000000"    # Black text for entries

def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select File",
        filetypes=[("PDF and DOCX Files", "*.pdf *.docx")],
    )
    if file_path:
        file_path_var.set(file_path)
        # Set default output path based on input file type
        base, ext = os.path.splitext(file_path)
        if ext.lower() == ".pdf":
            default_docx = base + ".docx"
            output_path_var.set(default_docx)
            # Enable page range inputs for PDF to DOCX
            start_page_entry.config(state='normal')
            end_page_entry.config(state='normal')
            start_page_label.config(state='normal')
            end_page_label.config(state='normal')
        elif ext.lower() == ".docx":
            default_pdf = base + ".pdf"
            output_path_var.set(default_pdf)
            # Disable page range inputs for DOCX to PDF
            start_page_var.set('')
            end_page_var.set('')
            start_page_entry.config(state='disabled')
            end_page_entry.config(state='disabled')
            start_page_label.config(state='disabled')
            end_page_label.config(state='disabled')

def browse_output_dir():
    directory = filedialog.askdirectory(title="Select Output Directory")
    if directory:
        current_output = os.path.basename(output_path_var.get())
        new_output_path = os.path.join(directory, current_output)
        output_path_var.set(new_output_path)

def convert_file():
    input_path = file_path_var.get()
    output_path = output_path_var.get()
    start_page = start_page_var.get().strip()
    end_page = end_page_var.get().strip()

    if not input_path or not output_path:
        messagebox.showerror("Error", "Please select both input and output paths.")
        return

    input_ext = os.path.splitext(input_path)[1].lower()

    if input_ext == ".pdf":
        # PDF to DOCX Conversion
        # Validate page inputs
        start = None
        end = None

        if start_page:
            if not start_page.isdigit():
                messagebox.showerror("Error", "Start page must be a positive integer.")
                return
            start = int(start_page) - 1  # 0-based index
            if start < 0:
                messagebox.showerror("Error", "Start page must be at least 1.")
                return

        if end_page:
            if not end_page.isdigit():
                messagebox.showerror("Error", "End page must be a positive integer.")
                return
            end = int(end_page)
            if end < 1:
                messagebox.showerror("Error", "End page must be at least 1.")
                return

        if start is not None and end is not None:
            if end <= start:
                messagebox.showerror("Error", "End page must be greater than Start page.")
                return

        try:
            # Copy PDF to a temporary location to avoid file lock issues
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                shutil.copyfile(input_path, tmp_pdf.name)
                tmp_pdf_path = tmp_pdf.name

            cv = Converter(tmp_pdf_path)
            if start is not None or end is not None:
                cv.convert(output_path, start=start, end=end)
            else:
                cv.convert(output_path)
            cv.close()

            # Remove temporary file
            os.remove(tmp_pdf_path)

            messagebox.showinfo("Success", f"Conversion successful!\nSaved as '{output_path}'.")
        except Exception as e:
            messagebox.showerror("Conversion Error", f"An error occurred:\n{e}")

    elif input_ext == ".docx":
        # DOCX to PDF Conversion
        try:
            # Handle open DOCX files by copying to a temporary location
            temp_docx = None
            if platform.system() == "Windows":
                temp_dir = tempfile.gettempdir()
                temp_docx = os.path.join(temp_dir, os.path.basename(input_path))
                shutil.copyfile(input_path, temp_docx)
                convert(temp_docx, output_path)
                # Remove temporary copy
                os.remove(temp_docx)
            else:
                # For macOS and Linux, attempt direct conversion
                convert(input_path, output_path)

            messagebox.showinfo("Success", f"Conversion successful!\nSaved as '{output_path}'.")
        except Exception as e:
            messagebox.showerror("Conversion Error", f"An error occurred:\n{e}")
    else:
        messagebox.showerror("Unsupported File", "Please select a PDF or DOCX file.")

# Initialize the main window
root = tk.Tk()
root.title("Document Converter")  # Clean title without icons
root.geometry("800x550")  # Increased window size for better layout
root.configure(bg=BG_COLOR)
root.resizable(False, False)

# Remove default window icon (if any)
# Uncomment the following line if you had previously set an icon
# root.iconbitmap('')  # No icon

# Set up styles
style = ttk.Style()
style.theme_use('clam')  # Use 'clam' theme as base

# Configure styles for dark mode with consistent entry background and appropriate text color
style.configure("TLabel",
                background=BG_COLOR,
                foreground=FG_COLOR,
                font=("Segoe UI", 12))
style.configure("TEntry",
                fieldbackground=ENTRY_BG_COLOR,
                foreground=ENTRY_FG_COLOR,  # Set entry text to black
                font=("Segoe UI", 12))
style.configure("TButton",
                background=BUTTON_COLOR,
                foreground=FG_COLOR,
                font=("Segoe UI", 12),
                borderwidth=0)
style.map("TButton",
          background=[('active', BUTTON_COLOR)],
          foreground=[('active', FG_COLOR)])

# Gold Vertical Line on the Left Side
gold_line = tk.Frame(root, bg=HIGHLIGHT_COLOR, width=5)
gold_line.pack(side=tk.LEFT, fill=tk.Y)

# Main Content Frame
content_frame = tk.Frame(root, bg=BG_COLOR)
content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

# Configure grid layout for content_frame
content_frame.columnconfigure(0, weight=1)

# File Selection
file_label = ttk.Label(content_frame, text="Select PDF or DOCX File:")
file_label.grid(row=0, column=0, sticky='w', pady=(0, 5))

file_frame = tk.Frame(content_frame, bg=BG_COLOR)
file_frame.grid(row=1, column=0, sticky='ew')

file_path_var = tk.StringVar()
file_entry = ttk.Entry(file_frame, textvariable=file_path_var, width=80, state='readonly', style='TEntry')
file_entry.pack(side=tk.LEFT, padx=(0, 10), pady=5, fill='x', expand=True)

browse_btn = ttk.Button(file_frame, text="Browse", command=browse_file)
browse_btn.pack(side=tk.LEFT)

# Output Selection
output_label = ttk.Label(content_frame, text="Output File:")
output_label.grid(row=2, column=0, sticky='w', pady=(20, 5))

output_frame = tk.Frame(content_frame, bg=BG_COLOR)
output_frame.grid(row=3, column=0, sticky='ew')

output_path_var = tk.StringVar()
output_entry = ttk.Entry(output_frame, textvariable=output_path_var, width=80, state='readonly', style='TEntry')
output_entry.pack(side=tk.LEFT, padx=(0, 10), pady=5, fill='x', expand=True)

browse_output_btn = ttk.Button(output_frame, text="Browse Output Folder", command=browse_output_dir)
browse_output_btn.pack(side=tk.LEFT)

# Optional Page Selection (only relevant for PDF to DOCX)
pages_label = ttk.Label(content_frame, text="Page Range (optional, for PDF to DOCX):")
pages_label.grid(row=4, column=0, sticky='w', pady=(20, 5))

pages_frame = tk.Frame(content_frame, bg=BG_COLOR)
pages_frame.grid(row=5, column=0, sticky='w')

start_page_var = tk.StringVar()
start_page_label = ttk.Label(pages_frame, text="Start Page:")
start_page_label.pack(side=tk.LEFT, padx=(0, 5))
start_entry = ttk.Entry(pages_frame, textvariable=start_page_var, width=10, style='TEntry')
start_entry.pack(side=tk.LEFT, padx=(0, 20))

end_page_var = tk.StringVar()
end_page_label = ttk.Label(pages_frame, text="End Page:")
end_page_label.pack(side=tk.LEFT, padx=(0, 5))
end_entry = ttk.Entry(pages_frame, textvariable=end_page_var, width=10, style='TEntry')
end_entry.pack(side=tk.LEFT)

# Initially disable page range inputs until a PDF is selected
start_page_entry = start_entry
end_page_entry = end_entry
start_page_label.config(state='disabled')
end_page_label.config(state='disabled')
start_entry.config(state='disabled')
end_entry.config(state='disabled')

# Convert Button
convert_btn = ttk.Button(content_frame, text="Convert", command=convert_file)
convert_btn.grid(row=6, column=0, pady=(30, 10), sticky='ew')

# Gold Horizontal Accent Line at the Bottom
accent = tk.Frame(root, bg=HIGHLIGHT_COLOR, height=2)
accent.pack(fill=tk.X, side=tk.BOTTOM)

# Apply custom styles manually for better dark mode (since ttk may not cover everything)
def set_dark_theme():
    root.configure(bg=BG_COLOR)
    for child in root.winfo_children():
        try:
            child.configure(background=BG_COLOR)
        except:
            pass

    # Update specific widgets inside content_frame
    for widget in content_frame.winfo_children():
        try:
            widget.configure(background=BG_COLOR)
        except:
            pass

    # Ensure labels have correct foreground
    file_label.configure(background=BG_COLOR, foreground=FG_COLOR)
    output_label.configure(background=BG_COLOR, foreground=FG_COLOR)
    pages_label.configure(background=BG_COLOR, foreground=FG_COLOR)
    start_page_label.configure(background=BG_COLOR, foreground=FG_COLOR)
    end_page_label.configure(background=BG_COLOR, foreground=FG_COLOR)

    # Ensure frames have correct background
    file_frame.configure(background=BG_COLOR)
    output_frame.configure(background=BG_COLOR)
    pages_frame.configure(background=BG_COLOR)

set_dark_theme()

# Run the application
root.mainloop()
