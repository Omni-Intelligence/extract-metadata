import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox


def extract_model(pbix_path):
    output_dir = os.path.join(os.path.dirname(pbix_path), "pbi_output")
    os.makedirs(output_dir, exist_ok=True)
    pbi_tools_path = os.path.join(os.getcwd(), "bin", "pbi-tools.exe")

    cmd = [
        pbi_tools_path,
        "extract",
        "-path",
        pbix_path,
        "-format",
        "json",
        "--output",
        output_dir,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)

    if result.returncode == 0:
        json_file = os.path.join(output_dir, "model.json")
        if os.path.exists(json_file):
            messagebox.showinfo("Success", f"Model metadata extracted to:\n{json_file}")
        else:
            messagebox.showerror(
                "Error", "Extraction succeeded but JSON file not found"
            )
    else:
        messagebox.showerror("Error", result.stderr)


def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("PBIX Files", "*.pbix")])
    if filename:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, filename)


def process_file():
    filepath = file_entry.get()
    if filepath.endswith(".pbix"):
        extract_model(filepath)
    else:
        messagebox.showerror("Error", "Please select a valid .pbix file.")


# Create main window
root = tk.Tk()
root.title("PBIX Model Extractor")
root.geometry("400x150")

# Create and pack widgets
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill="both")

label = tk.Label(frame, text="Select your .pbix file:")
label.pack(anchor="w")

file_frame = tk.Frame(frame)
file_frame.pack(fill="x", pady=5)

file_entry = tk.Entry(file_frame)
file_entry.pack(side="left", expand=True, fill="x")

browse_button = tk.Button(file_frame, text="Browse", command=browse_file)
browse_button.pack(side="right", padx=5)

button_frame = tk.Frame(frame)
button_frame.pack(pady=10)

extract_button = tk.Button(button_frame, text="Extract Model", command=process_file)
extract_button.pack(side="left", padx=5)

exit_button = tk.Button(button_frame, text="Exit", command=root.quit)
exit_button.pack(side="left", padx=5)

root.mainloop()
