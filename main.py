import json
import os
import platform
import re
import subprocess
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from extract_pbi_model_info import PowerBIModelExtractor


def get_base_path():
    """Get the base path for resources whether running as script or executable"""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.abspath(os.path.dirname(__file__))


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

        self.LabelSelectedOutputPath = tk.Label(self.root, text="Default: Subfolder next to PBIX", bg="#f3f0ea")
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
            # Add file access check for Windows
            if platform.system() == "Windows":
                try:
                    with open(self.file_path, "rb") as _:
                        pass
                except PermissionError:
                    messagebox.showerror(
                        "File Access Error",
                        "Cannot access the PBIX file. Please make sure it is not opened in Power BI Desktop or another program.",
                        parent=self.root,
                    )
                    return
                except Exception as e:
                    messagebox.showerror(
                        "File Access Error", f"Error accessing the PBIX file: {str(e)}", parent=self.root
                    )
                    return

            effective_output_path = (
                self.output_path if self.output_path_manually_set else os.path.dirname(self.file_path)
            )

            # Check for Windows path length limitation
            if platform.system() == "Windows":
                file_path = os.path.normpath(self.file_path)

                if len(file_path) >= 260:
                    messagebox.showerror(
                        "Path Too Long",
                        "Input file path exceeds Windows 260 character limit:\n"
                        f"Length: {len(file_path)}\n\n"
                        "Please use a shorter file path or move the file closer to the root directory.",
                    )
                    return

                output_path = os.path.normpath(effective_output_path)

                if len(output_path) >= 260:
                    messagebox.showerror(
                        "Path Too Long",
                        "Output path exceeds Windows 260 character limit:\n"
                        f"Length: {len(output_path)}\n\n"
                        "Please select an output location with a shorter path.",
                    )
                    return

            base_path = get_base_path()
            bin_path = os.path.join(base_path, "bin")
            selected_tool = None
            tool_type = "Unknown"
            cmd = []
            pbi_tools_result = None

            # Check operating system
            if platform.system() == "Windows":
                pbi_tools_win_path = os.path.join(bin_path, "win", "pbi-tools.exe")
                pbi_tools_core_path = os.path.join(bin_path, "core", "pbi-tools.core.exe")

                # Check if Power BI Desktop is installed
                pbi_folder = r"C:\Program Files\Microsoft Power BI Desktop"
                pbi_installed = os.path.exists(pbi_folder)

                # Check if .NET Framework is available
                dotnet_framework_ok = False
                try:
                    powershell = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
                    dotnet_check = subprocess.run(
                        f'"{powershell}" -Command "Get-ChildItem \'HKLM:\\SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full\' | Get-ItemProperty -Name Release"',
                        capture_output=True,
                        text=True,
                    )
                    if dotnet_check.returncode == 0:
                        # Release values: https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/versions-and-dependencies
                        # .NET 4.7.2 = 461808
                        match = re.search(r"Release\s+:\s+(\d+)", dotnet_check.stdout)
                        if match and int(match.group(1)) >= 461808:
                            dotnet_framework_ok = True
                except Exception:
                    dotnet_framework_ok = False

                # Check if .NET 8 runtime is available
                dotnet_8_ok = False
                try:
                    result = subprocess.run(
                        ["dotnet", "--list-runtimes"],
                        capture_output=True,
                        text=True,
                        check=True,
                    )
                    runtimes = result.stdout
                    dotnet_8_ok = any("Microsoft.NETCore.App 8." in line for line in runtimes.splitlines())
                except Exception:
                    dotnet_8_ok = False

                # Determine which version of pbi-tools to use
                if pbi_installed and dotnet_framework_ok and os.path.exists(pbi_tools_win_path):
                    selected_tool = pbi_tools_win_path
                    tool_type = "Power BI Desktop"
                elif dotnet_8_ok and os.path.exists(pbi_tools_core_path):
                    selected_tool = pbi_tools_core_path
                    tool_type = ".NET Runtime"
                else:
                    error_message = "Required environment not found:\n\n"
                    if not pbi_installed:
                        error_message += "- Power BI Desktop not installed\n"
                    if not dotnet_framework_ok:
                        error_message += "- .NET Framework 4.7.2 or higher not detected\n"
                    if not dotnet_8_ok:
                        error_message += "- .NET 8 Runtime not detected\n"
                    error_message += "\nPlease install the missing software"
                    messagebox.showerror("Environment Error", error_message)
                    return

                file_path = os.path.normpath(self.file_path)

                cmd = [selected_tool, "extract", file_path]

                if self.output_path_manually_set:
                    cmd.extend(["-extractFolder", os.path.normpath(self.output_path)])

                pbi_tools_result = subprocess.run(cmd, capture_output=True, text=True)

            else:  # Linux
                pbi_tools_linux_path = os.path.join(bin_path, "linux", "pbi-tools.core")

                if os.path.exists(pbi_tools_linux_path):
                    try:
                        os.chmod(pbi_tools_linux_path, 0o755)

                        file_path = os.path.normpath(self.file_path)

                        cmd = [pbi_tools_linux_path, "extract", file_path]

                        if self.output_path_manually_set and self.output_path:
                            output_path = os.path.normpath(self.output_path)
                            cmd.extend(["-extractFolder", output_path])

                        print(f"Executing command: {cmd}")

                        pbi_tools_result = subprocess.run(cmd, capture_output=True, text=True)

                    except Exception as e:
                        messagebox.showerror("Linux Error", f"Error running pbi-tools on Linux: {str(e)}")

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
                        f"PBIX files extracted using {tool_type} to:\n{extraction_dir}\n\n"
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
            messagebox.showerror("Error", str(e))
        finally:
            self.root.destroy()


if __name__ == "__main__":
    try:
        PBIX_Extractor()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Application failed to start: {str(e)}")
        sys.exit(1)
