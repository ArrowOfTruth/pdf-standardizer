"""Tkinter desktop interface for the PDF standardizer."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from pathlib import Path


def _hide_standalone_console() -> None:
    """Hide the console created for a double-clicked Windows executable."""
    if sys.platform != "win32":
        return

    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        get_console_window = kernel32.GetConsoleWindow
        get_console_window.restype = wintypes.HWND
        get_console_process_list = kernel32.GetConsoleProcessList
        get_console_process_list.argtypes = [
            ctypes.POINTER(wintypes.DWORD),
            wintypes.DWORD,
        ]
        get_console_process_list.restype = wintypes.DWORD
        user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.ShowWindow.restype = wintypes.BOOL
        console_window = get_console_window()
        if not console_window:
            return

        process_ids = (wintypes.DWORD * 16)()
        process_count = get_console_process_list(process_ids, len(process_ids))
        if process_count == 1:
            user32.ShowWindow(console_window, 0)
    except (AttributeError, OSError):
        return


def launch_gui() -> None:
    """Show a file-and-folder interface and run conversions in the background."""
    _hide_standalone_console()

    import tkinter as tk
    from tkinter import filedialog, ttk

    root = tk.Tk()
    root.title("PDF Standardizer")
    root.minsize(620, 420)
    root.geometry("720x500")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(3, weight=1)

    source_path = tk.StringVar()
    line_spacing = tk.StringVar(value="49")
    recurse = tk.BooleanVar(value=False)
    overwrite = tk.BooleanVar(value=False)
    results: queue.Queue[tuple[int, str, str] | tuple[None, str, str]] = (
        queue.Queue()
    )

    source_frame = ttk.LabelFrame(root, text="Input", padding=12)
    source_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
    source_frame.columnconfigure(0, weight=1)
    ttk.Entry(
        source_frame,
        textvariable=source_path,
        state="readonly",
    ).grid(row=0, column=0, sticky="ew", padx=(0, 8))

    def choose_file() -> None:
        selected = filedialog.askopenfilename(
            title="Select a PDF or PowerPoint presentation",
            filetypes=(
                ("PDF and PowerPoint files", "*.pdf *.pptx"),
                ("PDF files", "*.pdf"),
                ("PowerPoint files", "*.pptx"),
                ("All files", "*.*"),
            ),
        )
        if selected:
            source_path.set(selected)

    def choose_folder() -> None:
        selected = filedialog.askdirectory(title="Select a folder to convert")
        if selected:
            source_path.set(selected)

    ttk.Button(source_frame, text="Choose file...", command=choose_file).grid(
        row=0, column=1, padx=(0, 8)
    )
    ttk.Button(source_frame, text="Choose folder...", command=choose_folder).grid(
        row=0, column=2
    )

    options_frame = ttk.LabelFrame(root, text="Options", padding=12)
    options_frame.grid(row=1, column=0, sticky="ew", padx=16, pady=8)
    ttk.Label(options_frame, text="Line spacing (pixels)").grid(
        row=0, column=0, sticky="w"
    )
    ttk.Entry(options_frame, textvariable=line_spacing, width=8).grid(
        row=0, column=1, sticky="w", padx=(8, 20)
    )
    ttk.Checkbutton(
        options_frame,
        text="Search subfolders",
        variable=recurse,
    ).grid(row=0, column=2, sticky="w", padx=(0, 16))
    ttk.Checkbutton(
        options_frame,
        text="Replace existing outputs",
        variable=overwrite,
    ).grid(row=0, column=3, sticky="w")
    options_frame.columnconfigure(4, weight=1)

    status = tk.StringVar(value="Choose a file or folder to begin.")
    ttk.Label(root, textvariable=status).grid(
        row=2, column=0, sticky="w", padx=18, pady=(4, 6)
    )

    output_frame = ttk.LabelFrame(root, text="Results", padding=8)
    output_frame.grid(row=3, column=0, sticky="nsew", padx=16, pady=8)
    output_frame.columnconfigure(0, weight=1)
    output_frame.rowconfigure(0, weight=1)
    output_text = tk.Text(output_frame, wrap="word", height=10, state="disabled")
    output_text.grid(row=0, column=0, sticky="nsew")
    scrollbar = ttk.Scrollbar(
        output_frame,
        orient="vertical",
        command=output_text.yview,
    )
    scrollbar.grid(row=0, column=1, sticky="ns")
    output_text.configure(yscrollcommand=scrollbar.set)

    convert_button = ttk.Button(root, text="Convert")
    convert_button.grid(row=4, column=0, sticky="e", padx=16, pady=(4, 16))

    def show_results(text: str) -> None:
        output_text.configure(state="normal")
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, text or "No output was produced.")
        output_text.configure(state="disabled")

    def poll_results() -> None:
        try:
            return_code, stdout, stderr = results.get_nowait()
        except queue.Empty:
            root.after(100, poll_results)
            return

        output = "\n".join(part for part in (stdout.strip(), stderr.strip()) if part)
        show_results(output)
        if return_code is None:
            status.set("Could not start the converter.")
        elif return_code == 0:
            status.set("Conversion completed.")
        else:
            status.set("Conversion finished with errors.")
        convert_button.configure(state="normal")

    def start_conversion() -> None:
        selected = source_path.get()
        if not selected:
            status.set("Choose a PDF, PowerPoint presentation, or folder first.")
            return

        try:
            spacing = float(line_spacing.get())
        except ValueError:
            status.set("Line spacing must be a number of at least 1 pixel.")
            return
        if spacing < 1:
            status.set("Line spacing must be at least 1 pixel.")
            return

        source = Path(selected)
        command = [sys.executable]
        if not getattr(sys, "frozen", False):
            command.append(str(Path(__file__).with_name("convert_pptx_to_letter_pdf.py")))
        command.extend([str(source), "--line-spacing-px", str(spacing)])
        if source.is_dir() and recurse.get():
            command.append("--recurse")
        if overwrite.get():
            command.append("--overwrite")

        status.set("Converting... You can keep this window open while it runs.")
        show_results("")
        convert_button.configure(state="disabled")

        def convert_in_background() -> None:
            environment = os.environ.copy()
            environment["PYTHONUTF8"] = "1"
            try:
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=False,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    env=environment,
                )
                results.put((result.returncode, result.stdout, result.stderr))
            except OSError as exc:
                results.put((None, "", str(exc)))

        threading.Thread(target=convert_in_background, daemon=True).start()
        root.after(100, poll_results)

    convert_button.configure(command=start_conversion)
    root.mainloop()
