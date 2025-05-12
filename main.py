import json
import os
import platform
import requests
import socket
import subprocess
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import uuid
from extract_pbi_model_info import PowerBIModelExtractor


def get_base_path():
    """Get the base path for resources whether running as script or executable"""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.abspath(os.path.dirname(__file__))


def send_error_report(error_message, error_type, additional_info=None):
    """Send error report to the logging website"""
    try:
        system_info = {
            "os": platform.system(),
            "os_version": platform.version(),
            "hostname": socket.gethostname(),
            "python_version": sys.version,
            "app_version": "1.0.0",
            "incident_id": str(uuid.uuid4()),
        }

        payload = {
            "error_message": str(error_message),
            "error_type": error_type,
            "system_info": system_info,
            "additional_info": additional_info or {},
        }

        response = requests.post(
            "https://app.enterprisedna.co/api/v1/extractor-error",
            json=payload,
            timeout=5,
        )
        return response.status_code == 200
    except Exception as e:
        print(f"Failed to send error report: {str(e)}")
        return False


class PBIX_Extractor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("EDNAHQ PBIX Model Extractor")

        self.root.configure(bg="#f3f0ea")
        self.root.geometry("800x250")

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (800 / 2)
        y = (screen_height / 2) - (250 / 2)
        self.root.geometry(f"800x250+{int(x)}+{int(y)}")

        self.root.attributes("-topmost", True)

        self.file_path = None
        self.output_path = None
        self.output_path_manually_set = False

        self.create_widgets()

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

        self.LabelSelectedFilePath = tk.Label(self.root, text="Select a PBIX file", bg="#f3f0ea")
        self.LabelSelectedFilePath.place(relheight=0.2, relx=0.25, rely=0.3)
        # Output path section
        self.LabelOutputPath = tk.Label(self.root, text="Output Path:", bg="#f3f0ea")
        self.LabelOutputPath.place(relheight=0.2, relx=0.1, rely=0.51)

        self.LabelSelectedOutputPath = tk.Label(self.root, text="Default: New subfolder next to PBIX", bg="#f3f0ea")
        self.LabelSelectedOutputPath.place(relheight=0.2, relx=0.25, rely=0.51)

        # Add Info button for Output Path
        self.button_output_info = ttk.Button(self.root, text="?", width=2, command=self.show_output_path_info)
        self.button_output_info.place(relx=0.21, rely=0.52, height=20)

        self.button_input_browse = ttk.Button(self.root, text="Browse...", command=self.get_file_path)
        self.button_input_browse.place(relx=0.8, rely=0.315)

        self.button_output_browse = ttk.Button(self.root, text="Browse...", command=self.get_output_path)
        self.button_output_browse.place(relx=0.8, rely=0.52)

        self.button_cancel = ttk.Button(self.root, text="Cancel", command=self.cancel)
        self.button_cancel.place(relx=0.75, rely=0.8)

        self.button_OK = ttk.Button(self.root, text="OK", command=self.extract_model)
        self.button_OK.place(relx=0.87, rely=0.8)

        # Powered by PBI-tools badge
        self.powered_by_label = tk.Label(
            self.root,
            text="Powered by PBI-tools",
            font=("Segoe UI", 8),
            fg="#666666",
            bg="#f3f0ea",
        )
        self.powered_by_label.place(relx=0.02, rely=0.93)

    def show_output_path_info(self):
        info_text = (
            "Extraction works best for 'V3' PBIX files (created with Power BI Desktop from September 2020 or newer).\n\n"
            "Legacy PBIX files might have limited support.\n\n"
            "Default Behavior: If you don't select a specific output path, a sub-folder will be created in the same directory as the input PBIX file. This sub-folder will have the same name as the PBIX file (without the .pbix extension).\n\n"
            "Manual Selection: If you use 'Browse...' to select an output path, the extracted files will be placed directly into the selected folder."
        )
        messagebox.showinfo("Output Path Information", info_text, parent=self.root)

    def get_file_path(self):
        self.file_path = filedialog.askopenfilename(title="Select PBIX File", filetypes=[("PBIX Files", "*.pbix")])
        if self.file_path:
            self.button_input_browse.place_forget()
            self.LabelSelectedFilePath.config(text=self.file_path)
            if self.file_path:
                self.button_OK["state"] = "normal"

    def get_output_path(self):
        output_path_temp = filedialog.askdirectory(title="Select Output Directory")
        if output_path_temp:
            self.output_path = output_path_temp
            self.output_path_manually_set = True
            self.button_output_browse.place_forget()
            self.LabelSelectedOutputPath.config(text=self.output_path)
            if self.file_path:
                self.button_OK["state"] = "normal"

    def cancel(self):
        self.root.destroy()
        exit()

    def extract_model(self):
        try:
            # Check if running on Windows
            if platform.system() != "Windows":
                messagebox.showerror(
                    "Unsupported Operating System",
                    "This application is designed for Windows only.\n\n"
                    "Requirements:\n"
                    "- Windows OS\n"
                    "- Power BI Desktop x64\n"
                    "  (Must be installed in default location:\n"
                    "   C:\\Program Files\\Microsoft Power BI Desktop\\)",
                    parent=self.root
                )
                error_info = {
                    "file_path": self.file_path,
                    "output_path": self.output_path,
                    "operation": "extract_model",
                }
                self.root.quit()
                return 

            # Check if Power BI Desktop is installed
            pbi_folder = r"C:\Program Files\Microsoft Power BI Desktop"
            if not os.path.exists(pbi_folder):
                messagebox.showerror(
                    "Power BI Desktop Not Found",
                    "Power BI Desktop is not installed in the default location:\n"
                    "C:\\Program Files\\Microsoft Power BI Desktop\\\n\n"
                    "Please install Power BI Desktop x64 version.",
                    parent=self.root
                )
                return

            # Validate file paths
            if self.output_path_manually_set:
                output_dir = os.path.normpath(self.output_path)
                input_file = os.path.normpath(self.file_path)
                
                # Check if input file is within output directory
                if os.path.commonpath([output_dir]) == os.path.commonpath([output_dir, input_file]):
                    messagebox.showerror(
                        "Invalid Path",
                        "The PBIX file cannot be located within the output folder.\n"
                        "Please select a different output location.",
                        parent=self.root
                    )
                    return

            # Check file access
            try:
                with open(self.file_path, "rb") as _:
                    pass
            except PermissionError:
                messagebox.showerror(
                    "File Access Error",
                    "Cannot access the PBIX file. Please ensure:\n\n"
                    "1. The file is not opened in Power BI Desktop\n"
                    "2. No other program is using the file\n"
                    "3. You have permissions to access the file",
                    parent=self.root,
                )
                return
            except Exception as e:
                messagebox.showerror(
                    "File Access Error", 
                    f"Error accessing the PBIX file: {str(e)}", 
                    parent=self.root
                )
                return

            effective_output_path = (
                self.output_path if self.output_path_manually_set else os.path.dirname(self.file_path)
            )

            # Windows path length checks
            file_path = os.path.normpath(self.file_path)
            output_path = os.path.normpath(effective_output_path)

            if len(file_path) >= 260:
                messagebox.showerror(
                    "Path Too Long",
                    "Input file path exceeds Windows 260 character limit:\n"
                    f"Length: {len(file_path)}\n\n"
                    "Please use a shorter file path or move the file closer to the root directory.",
                )
                return

            if len(output_path) >= 260:
                messagebox.showerror(
                    "Path Too Long",
                    "Output path exceeds Windows 260 character limit:\n"
                    f"Length: {len(output_path)}\n\n"
                    "Please select an output location with a shorter path.",
                )
                return

            base_path = get_base_path()
            pbi_tools_path = os.path.join(base_path, "bin", "pbi-tools.exe")

            if not os.path.exists(pbi_tools_path):
                messagebox.showerror(
                    "Missing Component",
                    "Required component not found: pbi-tools.exe",
                    parent=self.root
                )
                return
            
            cmd = [pbi_tools_path, "extract", os.path.normpath(self.file_path)]
            if self.output_path_manually_set:
                cmd.extend(["-extractFolder", os.path.normpath(self.output_path)])

            pbi_tools_result = subprocess.run(cmd, capture_output=True, text=True)

            if pbi_tools_result.returncode == 0:
                extraction_dir = ""
                if self.output_path_manually_set and os.path.exists(os.path.join(effective_output_path, "Model")):
                    extraction_dir = os.path.normpath(effective_output_path)
                else:
                    pbix_filename_no_ext = os.path.splitext(os.path.basename(self.file_path))[0]
                    potential_path = os.path.join(os.path.dirname(self.file_path), pbix_filename_no_ext)
                    if os.path.exists(os.path.join(potential_path, "Model")):
                        extraction_dir = os.path.normpath(potential_path)
                    else:
                        messagebox.showerror(
                            "Extraction Error",
                            "Could not find Model folder in expected locations:\n"
                            f"- {effective_output_path}\n"
                            f"- {potential_path}",
                            parent=self.root,
                        )
                        return
                try:
                    extractor = PowerBIModelExtractor(extraction_dir)
                    model_info = extractor.extract_all()

                    save_dir = os.path.normpath(
                        effective_output_path if self.output_path_manually_set else extraction_dir
                    )
                    output_json_path = os.path.join(save_dir, "pbi_model_info.json")
                    with open(output_json_path, "w", encoding="utf-8") as f:
                        json.dump(model_info, f, indent=2)

                    messagebox.showinfo(
                        "Success",
                        f"PBIX files extracted to:\n{extraction_dir}\n\n"
                        f"Model metadata successfully processed and saved to:\n{output_json_path}",
                        parent=self.root,
                    )
                except Exception as e_meta:
                    messagebox.showerror(
                        "Metadata Processing Error",
                        f"PBIX files extracted to:\n{extraction_dir}\n\n"
                        f"An error occurred while processing the metadata:\n{str(e_meta)}",
                        parent=self.root,
                    )
            else:
                print(str(pbi_tools_result))

                error_output = pbi_tools_result.stderr if pbi_tools_result.stderr else pbi_tools_result.stdout
                messagebox.showerror(
                    "PBI-Tools Extraction Error", f"Failed to extract PBIX file:\n\n{error_output}", parent=self.root
                )

        except Exception as e:
            error_info = {
                "file_path": self.file_path,
                "output_path": self.output_path,
                "operation": "extract_model",
            }
            send_error_report(str(e), "extraction_error", error_info)
            messagebox.showerror("Error", str(e))
        finally:
            self.root.destroy()


if __name__ == "__main__":
    try:
        PBIX_Extractor()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Application failed to start: {str(e)}")
        sys.exit(1)
