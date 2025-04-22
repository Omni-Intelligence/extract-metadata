import json
import os
import platform
import subprocess
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox


def get_base_path():
    """Get the base path for resources whether running as script or executable"""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.abspath(os.path.dirname(__file__))


class PBIX_Extractor:
    def __init__(self):
        # Initialize the root widget
        self.root = tk.Tk()
        self.root.title("PBIX Model Extractor")

        # Window configuration
        self.root.configure(bg="#f3f0ea")
        self.root.geometry("800x250")
        # self.root.resizable(False, False)

        # Center window on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (800 / 2)
        y = (screen_height / 2) - (250 / 2)
        self.root.geometry(f"800x250+{int(x)}+{int(y)}")

        # Keep window on top
        self.root.attributes("-topmost", True)

        self.create_widgets()

        # Disable OK button initially
        self.button_OK["state"] = "disabled"

        self.root.mainloop()

    def create_widgets(self):
        # Title
        self.message = tk.Label(
            self.root,
            text="Select a PBIX file to extract model metadata",
            font=("Segoe UI", 14, "bold"),
            fg="#3d4244",
            justify=tk.CENTER,
            bg="#f3f0ea",
        )

        self.message.place(relheight=0.15, relx=0.14, rely=0.07)

        # Input file section
        self.LabelInputPath = tk.Label(self.root, text="Input PBIX:", bg="#f3f0ea")
        self.LabelInputPath.place(relheight=0.2, relx=0.1, rely=0.3)

        self.LabelSelectedFilePath = tk.Label(
            self.root, text="Select a PBIX file", bg="#f3f0ea"
        )
        self.LabelSelectedFilePath.place(relheight=0.2, relx=0.25, rely=0.3)
        # Output path section
        self.LabelOutputPath = tk.Label(self.root, text="Output Path:", bg="#f3f0ea")
        self.LabelOutputPath.place(relheight=0.2, relx=0.1, rely=0.51)

        self.LabelSelectedOutputPath = tk.Label(
            self.root, text="Select an output path", bg="#f3f0ea"
        )
        self.LabelSelectedOutputPath.place(relheight=0.2, relx=0.25, rely=0.51)

        self.button_input_browse = ttk.Button(
            self.root, text="Browse...", command=self.get_file_path
        )
        self.button_input_browse.place(relx=0.65, rely=0.315)

        self.button_output_browse = ttk.Button(
            self.root, text="Browse...", command=self.get_output_path
        )
        self.button_output_browse.place(relx=0.65, rely=0.52)

        self.button_cancel = ttk.Button(self.root, text="Cancel", command=self.cancel)
        self.button_cancel.place(relx=0.75, rely=0.8)

        self.button_OK = ttk.Button(self.root, text="OK", command=self.extract_model)
        self.button_OK.place(relx=0.87, rely=0.8)

    def get_file_path(self):
        self.file_path = filedialog.askopenfilename(
            title="Select PBIX File", filetypes=[("PBIX Files", "*.pbix")]
        )
        if self.file_path:
            self.button_input_browse.place_forget()
            self.LabelSelectedFilePath.config(text=self.file_path)
            if hasattr(self, "output_path"):
                self.button_OK["state"] = "normal"

    def get_output_path(self):
        self.output_path = filedialog.askdirectory(title="Select Output Directory")
        if self.output_path:
            self.output_path = os.path.join(self.output_path, "pbi_output")
            self.button_output_browse.place_forget()
            self.LabelSelectedOutputPath.config(text=self.output_path)
            if hasattr(self, "file_path"):
                self.button_OK["state"] = "normal"

    def cancel(self):
        self.root.destroy()
        exit()

    def extract_model(self):
        try:
            os.makedirs(self.output_path, exist_ok=True)
            log_file = os.path.join(self.output_path, "extraction_log.txt")

            base_path = get_base_path()

            # Check operating system
            if platform.system() == "Windows":
                powershell = (
                    r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
                )
                pbi_tools_path = os.path.join(base_path, "bin", "pbi-tools.exe")

                if not os.path.exists(pbi_tools_path):
                    messagebox.showerror(
                        "Error",
                        "pbi-tools.exe not found. Please ensure the application is properly installed.",
                    )
                    return

                file_path = os.path.normpath(self.file_path)
                output_path = os.path.normpath(self.output_path)

                cmd = f'''"{powershell}" "{pbi_tools_path}" extract "{file_path}" -extractFolder "{output_path}"'''

                result = subprocess.run(cmd, capture_output=True, text=True, shell=True)

                if result.returncode == 0:
                    messagebox.showinfo(
                        "Success",
                        f"Files extracted to:\n{self.output_path}\n\n"
                        "PRIMARY MODEL METADATA:\n"
                        "- DataModelSchema/model.json (tables, relationships, measures)\n\n"
                        "Additional metadata files:\n"
                        "- ReportMetadata.json (report-level metadata)\n"
                        "- Connections.json (data source information)\n"
                        "- DiagramLayout.json (model diagram layout)",
                    )
                else:
                    with open(log_file, "a") as f:
                        f.write(str(result))

                    messagebox.showerror("Error", result.stderr)
            else:  # Linux/Mac
                mock_file = os.path.join(self.output_path, "model.json")

                # Check if test file already exists
                if os.path.exists(mock_file):
                    response = messagebox.askyesno(
                        "File Exists",
                        "Test file already exists. Do you want to overwrite it?",
                    )
                    if not response:
                        return

                mock_output = {
                    "tables": [
                        {"name": "Table1", "columns": ["Col1", "Col2"]},
                        {"name": "Table2", "columns": ["Col3", "Col4"]},
                    ],
                    "relationships": [{"fromTable": "Table1", "toTable": "Table2"}],
                }

                with open(mock_file, "w") as f:
                    json.dump(mock_output, f, indent=2)

                messagebox.showinfo(
                    "Success (Test Mode)",
                    f"Mock metadata created at:\n{mock_file}\n\nNote: This is a test mode for Linux.",
                )
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.root.destroy()


if __name__ == "__main__":
    try:
        PBIX_Extractor()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Application failed to start: {str(e)}")
        sys.exit(1)
