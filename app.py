import csv
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from html import escape


APP_TITLE = "Bloxademy Meet Recording Zipper"
BRAND_NAME = "Bloxademy"
WINDOW_BG = "#0f172a"
CARD_BG = "#111827"
ACCENT = "#3b82f6"
ACCENT_2 = "#8b5cf6"
TEXT = "#f8fafc"
MUTED = "#cbd5e1"
SUCCESS = "#22c55e"
ERROR = "#ef4444"
WARN = "#f59e0b"


class MeetZipperApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x700")
        self.root.configure(bg=WINDOW_BG)
        self.root.minsize(780, 620)

        self.selected_directory = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready.")
        self.output_text = tk.StringVar(value="No zip created yet.")

        self._build_style()
        self._build_ui()

    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Blox.TFrame", background=WINDOW_BG)
        style.configure("Card.TFrame", background=CARD_BG)
        style.configure(
            "Header.TLabel",
            background=WINDOW_BG,
            foreground=TEXT,
            font=("Segoe UI", 22, "bold"),
        )
        style.configure(
            "SubHeader.TLabel",
            background=WINDOW_BG,
            foreground=MUTED,
            font=("Segoe UI", 10),
        )
        style.configure(
            "CardTitle.TLabel",
            background=CARD_BG,
            foreground=TEXT,
            font=("Segoe UI", 12, "bold"),
        )
        style.configure(
            "CardText.TLabel",
            background=CARD_BG,
            foreground=MUTED,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Status.TLabel",
            background=CARD_BG,
            foreground=TEXT,
            font=("Segoe UI", 10, "bold"),
        )
        style.configure(
            "Blox.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=10,
        )

    def _build_ui(self):
        outer = ttk.Frame(self.root, style="Blox.TFrame", padding=18)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer, style="Blox.TFrame")
        header.pack(fill="x", pady=(0, 14))

        ttk.Label(header, text=BRAND_NAME, style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Bundle requested Google Meet recordings into a single zip with a manifest.",
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="1) Paste CSV list of Google Meet IDs", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Example: gqj-hxpf-xpc or a CSV column/list containing multiple IDs.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(4, 10))

        self.csv_text = tk.Text(
            card,
            height=12,
            wrap="word",
            bg="#020617",
            fg=TEXT,
            insertbackground=TEXT,
            relief="flat",
            font=("Consolas", 11),
            padx=12,
            pady=12,
        )
        self.csv_text.pack(fill="both", expand=False)

        directory_wrap = ttk.Frame(card, style="Card.TFrame")
        directory_wrap.pack(fill="x", pady=(16, 0))

        ttk.Label(directory_wrap, text="2) Select recordings directory", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(
            directory_wrap,
            text="Choose Folder",
            command=self.choose_directory,
            style="Blox.TButton",
        ).grid(row=0, column=1, sticky="e", padx=(10, 0))

        self.directory_label = ttk.Label(
            directory_wrap,
            textvariable=self.selected_directory,
            style="CardText.TLabel",
        )
        self.directory_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        directory_wrap.columnconfigure(0, weight=1)

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.pack(fill="x", pady=(18, 0))

        ttk.Button(
            actions,
            text="Create ZIP",
            command=self.create_zip,
            style="Blox.TButton",
        ).pack(side="left")

        ttk.Button(
            actions,
            text="Clear",
            command=self.clear_inputs,
            style="Blox.TButton",
        ).pack(side="left", padx=(10, 0))

        status_card = ttk.Frame(card, style="Card.TFrame", padding=(0, 18, 0, 0))
        status_card.pack(fill="both", expand=True)

        ttk.Label(status_card, text="Status", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(status_card, textvariable=self.status_text, style="Status.TLabel").pack(anchor="w", pady=(6, 4))
        ttk.Label(status_card, textvariable=self.output_text, style="CardText.TLabel").pack(anchor="w")

        ttk.Label(
            status_card,
            text="Matched files are moved into the zip and removed from the folder after success.",
            style="CardText.TLabel",
        ).pack(anchor="w", pady=(16, 0))

    def choose_directory(self):
        directory = filedialog.askdirectory(title="Select recordings directory")
        if directory:
            self.selected_directory.set(directory)
            self.set_status("Directory selected.", "info")

    def clear_inputs(self):
        self.csv_text.delete("1.0", tk.END)
        self.selected_directory.set("")
        self.output_text.set("No zip created yet.")
        self.set_status("Cleared.", "info")

    def set_status(self, text: str, level: str = "info"):
        self.status_text.set(text)

    @staticmethod
    def parse_google_meet_ids(raw_text: str):
        ids = []
        seen = set()

        # Accept CSV, newline-separated, comma-separated, or pasted mixed text.
        rows = csv.reader(raw_text.splitlines())
        for row in rows:
            for cell in row:
                for match in re.findall(r"[a-z]{3}-[a-z]{4}-[a-z]{3}", cell.lower()):
                    if match not in seen:
                        seen.add(match)
                        ids.append(match)

        # Fallback in case csv.reader sees a single blob.
        if not ids:
            for match in re.findall(r"[a-z]{3}-[a-z]{4}-[a-z]{3}", raw_text.lower()):
                if match not in seen:
                    seen.add(match)
                    ids.append(match)

        return ids

    @staticmethod
    def normalise_meet_id(meet_id: str) -> str:
        return meet_id.strip().lower().replace("-", "_")

    @staticmethod
    def find_matching_files(directory: Path, requested_ids):
        matches = {}
        all_files = [p for p in directory.iterdir() if p.is_file()]

        for meet_id in requested_ids:
            target = MeetZipperApp.normalise_meet_id(meet_id)
            matched_file = None

            for file_path in all_files:
                name_lower = file_path.name.lower()
                if target in name_lower:
                    matched_file = file_path
                    break

            matches[meet_id] = matched_file

        return matches

    @staticmethod
    def build_manifest_html(requested_ids, matches, zip_name):
        rows = []
        for meet_id in requested_ids:
            matched = matches.get(meet_id)
            status_symbol = "✓" if matched else "✗"
            status_word = "Found" if matched else "Missing"
            file_name = matched.name if matched else "—"
            row_class = "ok" if matched else "missing"
            rows.append(
                f"""
                <tr class=\"{row_class}\">
                    <td>{escape(meet_id)}</td>
                    <td>{status_symbol}</td>
                    <td>{escape(status_word)}</td>
                    <td>{escape(file_name)}</td>
                </tr>
                """.strip()
            )

        generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
    <meta charset=\"utf-8\">
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
    <title>Bloxademy Manifest</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            background: #0f172a;
            color: #f8fafc;
            margin: 0;
            padding: 24px;
        }}
        .card {{
            max-width: 1000px;
            margin: 0 auto;
            background: #111827;
            border-radius: 16px;
            padding: 24px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.25);
        }}
        h1 {{ margin-top: 0; }}
        .muted {{ color: #cbd5e1; }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            background: #020617;
            border-radius: 12px;
            overflow: hidden;
        }}
        th, td {{
            text-align: left;
            padding: 12px;
            border-bottom: 1px solid #1e293b;
        }}
        th {{
            background: #1e293b;
        }}
        .ok td:nth-child(2), .ok td:nth-child(3) {{ color: #22c55e; font-weight: bold; }}
        .missing td:nth-child(2), .missing td:nth-child(3) {{ color: #ef4444; font-weight: bold; }}
        code {{
            background: #1e293b;
            padding: 2px 6px;
            border-radius: 6px;
        }}
    </style>
</head>
<body>
    <div class=\"card\">
        <h1>Bloxademy Recording Manifest</h1>
        <p class=\"muted\">Generated: {escape(generated_at)}</p>
        <p class=\"muted\">ZIP File: <code>{escape(zip_name)}</code></p>
        <p class=\"muted\">Requested Google Meet IDs: {len(requested_ids)}</p>

        <table>
            <thead>
                <tr>
                    <th>Google Meet ID</th>
                    <th>Status</th>
                    <th>Result</th>
                    <th>Matched File</th>
                </tr>
            </thead>
            <tbody>
                {''.join(rows)}
            </tbody>
        </table>
    </div>
</body>
</html>
"""

    def create_zip(self):
        raw_text = self.csv_text.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showerror(APP_TITLE, "Please paste at least one Google Meet ID or CSV list.")
            self.set_status("No Meet IDs supplied.", "error")
            return

        if not self.selected_directory.get():
            messagebox.showerror(APP_TITLE, "Please choose a recordings directory.")
            self.set_status("No directory selected.", "error")
            return

        directory = Path(self.selected_directory.get())
        if not directory.exists() or not directory.is_dir():
            messagebox.showerror(APP_TITLE, "The selected directory is invalid.")
            self.set_status("Directory invalid.", "error")
            return

        requested_ids = self.parse_google_meet_ids(raw_text)
        if not requested_ids:
            messagebox.showerror(
                APP_TITLE,
                "No valid Google Meet IDs were found. Expected format like gqj-hxpf-xpc.",
            )
            self.set_status("No valid Meet IDs detected.", "error")
            return

        self.root.update_idletasks()
        self.set_status("Scanning files...", "info")

        matches = self.find_matching_files(directory, requested_ids)
        found_files = [path for path in matches.values() if path is not None]

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_name = f"bloxademy_recordings_{timestamp}.zip"
        zip_path = directory / zip_name

        if zip_path.exists():
            messagebox.showerror(APP_TITLE, f"Output zip already exists: {zip_path.name}")
            self.set_status("Zip already exists.", "error")
            return

        temp_dir = None
        try:
            self.set_status("Creating zip and manifest...", "info")
            temp_dir = Path(tempfile.mkdtemp(prefix="bloxademy_manifest_"))
            manifest_path = temp_dir / "manifest.html"
            manifest_path.write_text(
                self.build_manifest_html(requested_ids, matches, zip_name),
                encoding="utf-8",
            )

            with zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.write(manifest_path, arcname="manifest.html")
                for file_path in found_files:
                    zf.write(file_path, arcname=file_path.name)

            deleted_count = 0
            for file_path in found_files:
                if file_path.exists():
                    file_path.unlink()
                    deleted_count += 1

            found_count = len(found_files)
            missing_count = len(requested_ids) - found_count
            self.output_text.set(f"Created: {zip_path}")
            self.set_status(
                f"Done. Found {found_count}, missing {missing_count}, removed {deleted_count} original file(s).",
                "success",
            )
            messagebox.showinfo(
                APP_TITLE,
                f"ZIP created successfully.\n\nFound: {found_count}\nMissing: {missing_count}\nSaved to: {zip_path}",
            )

        except Exception as exc:
            if zip_path.exists():
                try:
                    zip_path.unlink()
                except OSError:
                    pass
            self.set_status(f"Error: {exc}", "error")
            messagebox.showerror(APP_TITLE, f"Something went wrong:\n\n{exc}")
        finally:
            if temp_dir and temp_dir.exists():
                shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    root = tk.Tk()
    app = MeetZipperApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
