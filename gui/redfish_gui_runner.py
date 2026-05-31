import os
import re
import sys
import json
import queue
import threading
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl

# Allow importing redfish_agent.py from repository root.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from redfish_agent import execute_redfish  # noqa: E402


@dataclass
class Target:
    root_url: str
    username: str
    password: str
    status: str = "Unknown"


class RedfishGuiRunner:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.master.title("Redfish Multi-Target Test Runner")
        self.master.geometry("1100x700")

        self.targets = []
        self.test_excel_path = tk.StringVar()
        self.log_queue = queue.Queue()
        self.last_report_path = None

        self.output_dir = os.path.join(SCRIPT_DIR, "output")
        os.makedirs(self.output_dir, exist_ok=True)

        self._build_ui()
        self._start_log_poll()

    def _build_ui(self):
        root = ttk.Frame(self.master, padding=10)
        root.pack(fill=tk.BOTH, expand=True)

        connection_frame = ttk.LabelFrame(root, text="Target Connection")
        connection_frame.pack(fill=tk.X, padx=5, pady=5)

        left_panel = ttk.LabelFrame(connection_frame, text="Target Editor")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(5, 12), pady=5)

        right_panel = ttk.LabelFrame(connection_frame, text="Target Actions")
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(0, 5), pady=5)

        connection_frame.rowconfigure(0, weight=1)
        connection_frame.columnconfigure(0, weight=1)
        connection_frame.columnconfigure(1, weight=0)

        input_row = ttk.Frame(left_panel)
        input_row.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))

        action_row = ttk.Frame(left_panel)
        action_row.grid(row=1, column=0, sticky="w", padx=4, pady=(2, 6))

        # Left panel: input row (independent from action row)
        ttk.Label(input_row, text="Root URL").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.root_url_entry = ttk.Entry(input_row, width=58)
        self.root_url_entry.grid(row=0, column=1, sticky="we", padx=5, pady=5)
        self.root_url_entry.insert(0, "https://127.0.0.1:8000")

        ttk.Label(input_row, text="Username").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.username_entry = ttk.Entry(input_row, width=16)
        self.username_entry.grid(row=0, column=3, sticky="we", padx=5, pady=5)
        self.username_entry.insert(0, "admin")

        ttk.Label(input_row, text="Password").grid(row=0, column=4, sticky="w", padx=5, pady=5)
        self.password_entry = ttk.Entry(input_row, width=16, show="*")
        self.password_entry.grid(row=0, column=5, sticky="we", padx=5, pady=5)

        # Left panel: action row (layout does not depend on input row width)
        ttk.Button(action_row, text="Add", command=self.add_target).pack(side=tk.LEFT, padx=5, pady=4)
        ttk.Button(action_row, text="Modify", command=self.modify_selected_target).pack(side=tk.LEFT, padx=5, pady=4)
        ttk.Button(action_row, text="Remove", command=self.remove_selected_target).pack(side=tk.LEFT, padx=5, pady=4)

        left_panel.columnconfigure(0, weight=1)
        input_row.columnconfigure(1, weight=3, minsize=280)
        input_row.columnconfigure(3, weight=1, minsize=120)
        input_row.columnconfigure(5, weight=1, minsize=120)

        # Right panel: utility actions in two rows
        ttk.Button(right_panel, text="Check Status", command=self.check_all_status).grid(
            row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5
        )
        ttk.Button(right_panel, text="Import Targets", command=self.import_targets).grid(
            row=1, column=0, sticky="ew", padx=5, pady=5
        )
        ttk.Button(right_panel, text="Export Targets", command=self.export_targets).grid(
            row=1, column=1, sticky="ew", padx=5, pady=5
        )

        right_panel.columnconfigure(0, weight=1)
        right_panel.columnconfigure(1, weight=1)

        list_frame = ttk.LabelFrame(root, text="Target List")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(
            list_frame,
            columns=("root_url", "username", "status"),
            show="headings",
            height=12,
        )
        self.tree.heading("root_url", text="Connection URL")
        self.tree.heading("username", text="Username")
        self.tree.heading("status", text="Connection Status")
        self.tree.column("root_url", width=500)
        self.tree.column("username", width=140)
        self.tree.column("status", width=200)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        test_frame = ttk.LabelFrame(root, text="Test Input and Execution")
        test_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(test_frame, text="Excel File").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(test_frame, textvariable=self.test_excel_path, width=80).grid(
            row=0, column=1, sticky="we", padx=5, pady=5
        )
        ttk.Button(test_frame, text="Browse", command=self.select_excel).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(test_frame, text="Run Test", command=self.run_test).grid(row=0, column=3, padx=5, pady=5)
        ttk.Button(test_frame, text="Open Latest Report", command=self.open_latest_report).grid(row=0, column=4, padx=5, pady=5)

        test_frame.columnconfigure(1, weight=1)

        log_frame = ttk.LabelFrame(root, text="Execution Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def log(self, message: str):
        self.log_queue.put(message)

    def _start_log_poll(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, msg + "\n")
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        self.master.after(120, self._start_log_poll)

    def add_target(self):
        root_url = self.root_url_entry.get().strip()
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not root_url or not username or not password:
            messagebox.showerror("Input Error", "Root URL, Username, and Password are required.")
            return

        for target in self.targets:
            if target.root_url == root_url and target.username == username:
                messagebox.showwarning("Duplicate Target", "This target already exists.")
                return

        target = Target(root_url=root_url, username=username, password=password, status="Checking...")
        self.targets.append(target)
        self.tree.insert("", tk.END, values=(target.root_url, target.username, target.status))

        threading.Thread(target=self._check_status_single, args=(target,), daemon=True).start()

    def _check_status_single(self, target: Target):
        status = self._probe_connection(target)
        target.status = status
        self.master.after(0, self._refresh_target_list)

    def _probe_connection(self, target: Target) -> str:
        test_url = target.root_url.rstrip("/") + "/redfish/v1"
        try:
            response = requests.get(test_url, auth=(target.username, target.password), verify=False, timeout=8)
            if response.status_code == 200:
                return "Connected"
            return f"HTTP {response.status_code}"
        except requests.RequestException as exc:
            return f"Failed: {str(exc)}"

    def _refresh_target_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for target in self.targets:
            self.tree.insert("", tk.END, values=(target.root_url, target.username, target.status))

    def _get_selected_target_index(self):
        selected = self.tree.selection()
        if not selected:
            return None
        return self.tree.index(selected[0])

    def _on_tree_select(self, _event=None):
        idx = self._get_selected_target_index()
        if idx is None or idx >= len(self.targets):
            return

        target = self.targets[idx]
        self.root_url_entry.delete(0, tk.END)
        self.root_url_entry.insert(0, target.root_url)
        self.username_entry.delete(0, tk.END)
        self.username_entry.insert(0, target.username)
        self.password_entry.delete(0, tk.END)
        self.password_entry.insert(0, target.password)

    def modify_selected_target(self):
        idx = self._get_selected_target_index()
        if idx is None:
            messagebox.showwarning("No Selection", "Please select one target in the list first.")
            return

        root_url = self.root_url_entry.get().strip()
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        if not root_url or not username or not password:
            messagebox.showerror("Input Error", "Root URL, Username, and Password are required.")
            return

        for i, t in enumerate(self.targets):
            if i != idx and t.root_url == root_url and t.username == username:
                messagebox.showwarning("Duplicate Target", "Another target already uses the same URL and username.")
                return

        target = self.targets[idx]
        target.root_url = root_url
        target.username = username
        target.password = password
        target.status = "Checking..."
        self._refresh_target_list()
        self.log(f"Modified target: {root_url} ({username})")
        threading.Thread(target=self._check_status_single, args=(target,), daemon=True).start()

    def remove_selected_target(self):
        idx = self._get_selected_target_index()
        if idx is None:
            messagebox.showwarning("No Selection", "Please select one target in the list first.")
            return

        target = self.targets[idx]
        confirm = messagebox.askyesno(
            "Confirm Remove",
            f"Remove target?\n\nURL: {target.root_url}\nUsername: {target.username}",
        )
        if not confirm:
            return

        removed = self.targets.pop(idx)
        self._refresh_target_list()
        self.log(f"Removed target: {removed.root_url} ({removed.username})")

    def check_all_status(self):
        if not self.targets:
            messagebox.showwarning("No Targets", "Please add at least one target first.")
            return

        self.log("Checking connection status for all targets...")
        for target in self.targets:
            target.status = "Checking..."
        self._refresh_target_list()

        for target in self.targets:
            threading.Thread(target=self._check_status_single, args=(target,), daemon=True).start()

    def export_targets(self):
        if not self.targets:
            messagebox.showwarning("No Targets", "There are no targets to export.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Export targets to JSON",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=SCRIPT_DIR,
            initialfile="targets.json",
        )
        if not file_path:
            return

        data = [
            {"root_url": t.root_url, "username": t.username, "password": t.password}
            for t in self.targets
        ]
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self.log(f"Exported {len(data)} targets to {file_path}")
        except Exception as exc:
            messagebox.showerror("Export Error", f"Failed to export targets: {exc}")

    def import_targets(self):
        file_path = filedialog.askopenfilename(
            title="Import targets from JSON",
            filetypes=[("JSON files", "*.json")],
            initialdir=SCRIPT_DIR,
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            messagebox.showerror("Import Error", f"Failed to read JSON file: {exc}")
            return

        if not isinstance(data, list):
            messagebox.showerror("Import Error", "JSON root must be a list of target objects.")
            return

        existing = {(t.root_url, t.username) for t in self.targets}
        imported_count = 0
        skipped_count = 0
        for item in data:
            if not isinstance(item, dict):
                skipped_count += 1
                continue

            root_url = str(item.get("root_url", "")).strip()
            username = str(item.get("username", "")).strip()
            password = str(item.get("password", "")).strip()
            if not root_url or not username or not password:
                skipped_count += 1
                continue

            key = (root_url, username)
            if key in existing:
                skipped_count += 1
                continue

            self.targets.append(Target(root_url=root_url, username=username, password=password, status="Unknown"))
            existing.add(key)
            imported_count += 1

        self._refresh_target_list()
        self.log(f"Imported {imported_count} targets from {file_path} (skipped {skipped_count}).")

        if imported_count > 0:
            for target in self.targets:
                if target.status == "Unknown":
                    target.status = "Checking..."
            self._refresh_target_list()
            for target in self.targets:
                if target.status == "Checking...":
                    threading.Thread(target=self._check_status_single, args=(target,), daemon=True).start()

    def open_latest_report(self):
        report_path = None
        if self.last_report_path and os.path.isfile(self.last_report_path):
            report_path = self.last_report_path
        else:
            html_files = [
                os.path.join(self.output_dir, name)
                for name in os.listdir(self.output_dir)
                if name.lower().endswith(".html")
            ]
            if html_files:
                report_path = max(html_files, key=os.path.getmtime)

        if not report_path:
            messagebox.showwarning("No Report", "No HTML report found in gui/output.")
            return

        webbrowser.open_new_tab(Path(report_path).resolve().as_uri())
        self.log(f"Opened report: {report_path}")

    def select_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select command Excel file",
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=REPO_ROOT,
        )
        if file_path:
            self.test_excel_path.set(file_path)

    def run_test(self):
        excel_path = self.test_excel_path.get().strip()
        if not excel_path:
            messagebox.showerror("Input Error", "Please select an Excel file.")
            return
        if not os.path.isfile(excel_path):
            messagebox.showerror("Input Error", f"Excel file not found: {excel_path}")
            return
        if not self.targets:
            messagebox.showerror("Input Error", "Please add at least one target.")
            return

        self.log("Starting parallel tests for all targets...")
        test_thread = threading.Thread(target=self._run_tests_worker, args=(excel_path,), daemon=True)
        test_thread.start()

    def _run_tests_worker(self, excel_path: str):
        run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        results = []
        threads = []
        result_lock = threading.Lock()

        def worker(target: Target):
            host_label = self._build_host_label(target.root_url)
            output_name = f"output_{host_label}_{run_stamp}.xlsx"
            output_path = os.path.join(self.output_dir, output_name)

            self.log(f"[START] {target.root_url} -> {output_name}")
            status = "Done"
            error_msg = ""
            try:
                execute_redfish(
                    username=target.username,
                    password=target.password,
                    root_url=target.root_url,
                    excel_path=excel_path,
                    output_excel_path=output_path,
                )
            except Exception as exc:
                status = "Failed"
                error_msg = str(exc)
                self.log(f"[ERROR] {target.root_url}: {error_msg}")

            with result_lock:
                results.append(
                    {
                        "root_url": target.root_url,
                        "username": target.username,
                        "output_path": output_path,
                        "run_status": status,
                        "error": error_msg,
                    }
                )

            self.log(f"[END] {target.root_url} ({status})")

        for target in self.targets:
            t = threading.Thread(target=worker, args=(target,), daemon=True)
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        report_path = self._generate_html_report(results, run_stamp)
        self.last_report_path = report_path
        self.log(f"HTML report generated: {report_path}")
        self.master.after(0, lambda: messagebox.showinfo("Test Completed", f"All tests finished.\nReport: {report_path}"))

    def _build_host_label(self, root_url: str) -> str:
        parsed = urlparse(root_url)
        host = parsed.netloc or parsed.path
        host = host.replace(":", "_")
        return re.sub(r"[^A-Za-z0-9_.-]", "_", host)

    def _analyze_output_excel(self, output_path: str) -> dict:
        info = {
            "batch_success": 0,
            "batch_error": 0,
            "error_rows": 0,
            "total_rows": 0,
            "final_result": "PASS",
        }
        if not os.path.isfile(output_path):
            info["final_result"] = "FAILED_TO_GENERATE"
            return info

        wb = openpyxl.load_workbook(output_path, data_only=True)
        ws = wb.active
        if ws is None:
            info["final_result"] = "INVALID_OUTPUT"
            return info

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or all(cell is None for cell in row):
                continue
            info["total_rows"] += 1
            method = str(row[0]) if row[0] is not None else ""
            status = str(row[4]) if len(row) > 4 and row[4] is not None else ""

            if method.startswith("BATCH("):
                if status.upper() == "SUCCESS":
                    info["batch_success"] += 1
                elif status.upper() == "ERROR":
                    info["batch_error"] += 1

            if status.upper() == "ERROR":
                info["error_rows"] += 1

        if info["batch_error"] > 0 or info["error_rows"] > 0:
            info["final_result"] = "FAIL"
        return info

    def _generate_html_report(self, results: list, run_stamp: str) -> str:
        report_name = f"report_{run_stamp}.html"
        report_path = os.path.join(self.output_dir, report_name)

        rows_html = []
        for item in results:
            analysis = self._analyze_output_excel(item["output_path"])
            result_badge = analysis["final_result"]
            color = "#1b8f3f" if result_badge == "PASS" else "#b11f1f"
            rows_html.append(
                "".join(
                    [
                        "<tr>",
                        f"<td>{item['root_url']}</td>",
                        f"<td>{item['username']}</td>",
                        f"<td>{item['run_status']}</td>",
                        f"<td>{analysis['total_rows']}</td>",
                        f"<td>{analysis['batch_success']}</td>",
                        f"<td>{analysis['batch_error']}</td>",
                        f"<td>{analysis['error_rows']}</td>",
                        f"<td style='font-weight:bold;color:{color};'>{result_badge}</td>",
                        f"<td>{os.path.basename(item['output_path'])}</td>",
                        "</tr>",
                    ]
                )
            )

        html = f"""
<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <title>Redfish GUI Test Report</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; }}
    h1 {{ margin-bottom: 4px; }}
    .meta {{ color: #555; margin-bottom: 16px; }}
    table {{ border-collapse: collapse; width: 100%; }}
    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
    th {{ background: #f4f4f4; }}
    tr:nth-child(even) {{ background: #fafafa; }}
  </style>
</head>
<body>
  <h1>Redfish GUI Test Report</h1>
  <div class=\"meta\">Generated at: {datetime.now().isoformat(timespec='seconds')}</div>
  <table>
    <thead>
      <tr>
        <th>Root URL</th>
        <th>Username</th>
        <th>Run Status</th>
        <th>Total Output Rows</th>
        <th>Batch SUCCESS</th>
        <th>Batch ERROR</th>
        <th>Error Rows</th>
        <th>Final Result</th>
        <th>Output File</th>
      </tr>
    </thead>
    <tbody>
      {''.join(rows_html)}
    </tbody>
  </table>
  <p style=\"margin-top:14px;color:#666;\">Output folder: {self.output_dir}</p>
</body>
</html>
"""

        with open(report_path, "w", encoding="utf-8") as f:
            f.write(html)
        return report_path


def main():
    root = tk.Tk()
    app = RedfishGuiRunner(root)
    root.mainloop()


if __name__ == "__main__":
    main()
