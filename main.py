import json
import os
import platform
import re
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

            # Check for Windows path length limitation
            if platform.system() == "Windows":
                file_path = os.path.normpath(self.file_path)
                output_path = os.path.normpath(self.output_path)

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
            bin_path = os.path.join(base_path, "bin")

            # Check operating system
            if platform.system() == "Windows":
                powershell = (
                    r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
                )

                pbi_tools_win_path = os.path.join(bin_path, "win", "pbi-tools.exe")
                pbi_tools_core_path = os.path.join(
                    bin_path, "core", "pbi-tools.core.exe"
                )

                # Check if Power BI Desktop is installed
                pbi_installed = False
                try:
                    pbi_check = subprocess.run(
                        f'"{powershell}" -Command "Get-AppxPackage -Name Microsoft.MicrosoftPowerBIDesktop"',
                        capture_output=True,
                        text=True,
                        shell=True,
                    )
                    if "DisplayName" in pbi_check.stdout:
                        pbi_installed = True
                except Exception:
                    pbi_installed = False

                # Check if .NET Framework is available
                dotnet_framework_ok = False
                try:
                    dotnet_check = subprocess.run(
                        f'"{powershell}" -Command "Get-ChildItem \'HKLM:\\SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full\' | Get-ItemProperty -Name Release"',
                        capture_output=True,
                        text=True,
                        shell=True,
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
                    dotnet_runtime_check = subprocess.run(
                        f'"{powershell}" -Command "dotnet --list-runtimes"',
                        capture_output=True,
                        text=True,
                        shell=True,
                    )
                    if dotnet_runtime_check.returncode == 0:
                        if re.search(
                            r"Microsoft\.NETCore\.App\s+8\.",
                            dotnet_runtime_check.stdout,
                        ):
                            dotnet_8_ok = True
                except Exception:
                    dotnet_8_ok = False

                # Determine which version of pbi-tools to use
                if (
                    pbi_installed
                    and dotnet_framework_ok
                    and os.path.exists(pbi_tools_win_path)
                ):
                    # Use standard version
                    selected_tool = pbi_tools_win_path
                    tool_type = "Windows"
                elif dotnet_8_ok and os.path.exists(pbi_tools_core_path):
                    # Use .NET Core version
                    selected_tool = pbi_tools_core_path
                    tool_type = ".NET Core"
                else:
                    # No valid environment found
                    error_message = "Required environment not found:\n\n"
                    if not pbi_installed:
                        error_message += "- Power BI Desktop not installed\n"
                    if not dotnet_framework_ok:
                        error_message += (
                            "- .NET Framework 4.7.2 or higher not detected\n"
                        )
                    if not dotnet_8_ok:
                        error_message += "- .NET 8 Runtime not detected\n"

                    error_message += "\nPlease install the missing components:\n"
                    error_message += (
                        "- Power BI Desktop: https://powerbi.microsoft.com/desktop/\n"
                    )
                    error_message += "- .NET Framework: https://dotnet.microsoft.com/download/dotnet-framework\n"
                    error_message += "- .NET 8 Runtime: https://dotnet.microsoft.com/download/dotnet/8.0"

                    messagebox.showerror("Environment Error", error_message)
                    return

                file_path = os.path.normpath(self.file_path)
                output_path = os.path.normpath(self.output_path)

                # Run the appropriate version of pbi-tools
                cmd = f'''"{powershell}" "{selected_tool}" extract "{file_path}" -extractFolder "{output_path}"'''

                result = subprocess.run(cmd, capture_output=True, text=True, shell=True)

                if result.returncode == 0:
                    messagebox.showinfo(
                        "Success",
                        f"Files extracted to:\n{self.output_path}\n\n"
                        f"Tool used: {tool_type} version\n\n"
                        "PRIMARY MODEL METADATA:\n"
                        "Model/Definition/model.bim\n\n"
                        "Additional metadata files:\n"
                        "- DataModelSchema/model.json (tables, relationships, measures)\n"
                        "- ReportMetadata.json (report-level metadata)\n"
                        "- Connections.json (data source information)\n"
                        "- DiagramLayout.json (model diagram layout)",
                    )
                else:
                    with open(log_file, "a") as f:
                        f.write(str(result))

                    messagebox.showerror("Error", result.stderr)
            else:  # Linux
                pbi_tools_linux_path = os.path.join(bin_path, "linux", "pbi-tools.core")

                if os.path.exists(pbi_tools_linux_path):
                    try:
                        os.chmod(pbi_tools_linux_path, 0o755)

                        file_path = os.path.normpath(self.file_path)
                        output_path = os.path.normpath(self.output_path)

                        cmd = f'{pbi_tools_linux_path} extract "{file_path}" -extractFolder "{output_path}"'

                        print(f"Executing command: {cmd}")

                        result = subprocess.run(
                            cmd, capture_output=True, text=True, shell=True
                        )

                        print(result, "\ncmd: ", cmd)

                        if result.returncode == 0:
                            messagebox.showinfo(
                                "Success",
                                f"Files extracted to:\n{self.output_path}\n\n"
                                f"Tool used: Linux version\n\n"
                                "PRIMARY MODEL METADATA:\n"
                                "Model/Definition/model.bim\n\n"
                                "Additional metadata files:\n"
                                "- DataModelSchema/model.json (tables, relationships, measures)\n"
                                "- ReportMetadata.json (report-level metadata)\n"
                                "- Connections.json (data source information)\n"
                                "- DiagramLayout.json (model diagram layout)",
                            )
                        else:
                            with open(log_file, "a") as f:
                                f.write(f"Command: {cmd}\n")
                                f.write(f"Return code: {result.returncode}\n")
                                f.write(f"Stdout: {result.stdout}\n")
                                f.write(f"Stderr: {result.stderr}\n")

                            messagebox.showerror(
                                "Error", f"pbi-tools execution failed: {result.stderr}"
                            )
                        return
                    except Exception as e:
                        messagebox.showerror(
                            "Linux Error", f"Error running pbi-tools on Linux: {str(e)}"
                        )

                # Fall back to mock behavior if Linux tool isn't available or failed
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
