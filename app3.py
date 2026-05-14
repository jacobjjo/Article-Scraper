import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from main import process_file


class ExcelProcessorApp:

    def __init__(self, root):

        self.root = root

        self.root.title("Excel Processor")

        self.root.geometry("500x350")

        # -----------------------------
        # Thread communication queue
        # -----------------------------

        self.queue = queue.Queue()

        self.stop_event = threading.Event()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

        # -----------------------------
        # Variables
        # -----------------------------

        self.file_path = tk.StringVar()

        self.output_excel = tk.StringVar()

        self.filter_years = tk.StringVar()

        self.preferences = check_preferences()

        if self.preferences:
            self.file_path.set(self.preferences[0])
            self.filter_years.set(",".join(self.preferences[1:-1]))
            self.output_excel.set(self.preferences[len(self.preferences)-1])

        print(self.filter_years.get())
        # -----------------------------
        # Layout
        # -----------------------------

        main = ttk.Frame(root, padding=15)
        main.pack(fill="both", expand=True)

        # ----------------------------
        # File Selection
        # ----------------------------

        ttk.Label(main, text="Excel File").grid(
            row=0, column=0, sticky="w"
        )

        ttk.Entry(
            main,
            textvariable=self.file_path,
            width=45
        ).grid(row=1, column=0, padx=(0, 10), sticky="ew")

        ttk.Button(
            main,
            text="Browse...",
            command=self.browse_file
        ).grid(row=1, column=1)

        # ----------------------------
        # Text Field Example
        # ----------------------------

        ttk.Label(main, text="Enter Publication Years").grid(
            row=2, column=0, pady=(20, 0), sticky="w"
        )

        ttk.Entry(
            main,
            textvariable=self.filter_years
        ).grid(row=3, column=0, columnspan=2, sticky="ew")


        ttk.Label(main, text="Enter Output Excel Filename (defaults to output.xlsx)").grid(
            row=7,
            column=0,
            pady=(20, 0),
            sticky="w"
        )

        ttk.Entry(
            main,
            textvariable=self.output_excel
        ).grid(
            row=8,
            column=0,
            columnspan=2,
            sticky="ew"
        )

        self.progress = ttk.Progressbar(
            main,
            mode="determinate",
            maximum=100,
            value=0
        )

        self.progress.grid(
            row=9,
            column=0,
            columnspan=2,
            sticky="ew"
        )

        # ----------------------------
        # Run Button
        # ----------------------------

        self.run_button = ttk.Button(
            main,
            text="Process File",
            command=self.start_processing
        )
        self.run_button.grid(
            row=11,
            column=0,
            columnspan=2,
            pady=30
        )

        self.status_label = ttk.Label(
            main,
            text="Ready"
        )

        self.status_label.grid(
            row=10,
            column=0,
            columnspan=2,
            sticky="w"
        )

        # Make columns resize nicely
        main.columnconfigure(0, weight=1)

        # -----------------------------
        # Start queue polling
        # -----------------------------

        self.poll_queue()

    # ---------------------------------
    # Browse for file
    # ---------------------------------

    def on_close(self):

        self.stop_event.set()

        self.root.destroy()

    def save_preferences(self):

        preferences = f"{self.file_path.get()},{self.filter_years.get()},{self.output_excel.get()}"
        open("preferences.txt", "w").write(preferences)

    def browse_file(self):

        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if filename:
            self.file_path.set(filename)

    # ---------------------------------
    # Start worker thread
    # ---------------------------------

    def start_processing(self):

        if not self.file_path.get():

            messagebox.showwarning(
                "Missing File",
                "Please select a file."
            )

            return

        self.run_button.config(state="disabled")

        self.progress["value"] = 0

        self.status_label.config(text="Starting...")

        worker = threading.Thread(
            target=self.worker_thread,
            daemon=True
        )

        worker.start()

    # ---------------------------------
    # Worker thread
    # ---------------------------------

    def worker_thread(self):

        try:

            output_path = process_file(
                self.file_path.get(),
                output_excel=self.output_excel.get() if self.output_excel.get() else "output.xlsx",
                filter_year=self.filter_years.get().split(",") if self.filter_years.get() else None,
                progress_callback=self.send_progress,
                stop_event=self.stop_event
            )

            self.queue.put({
                "type": "complete",
                "output": output_path
            })

        except Exception as e:

            self.queue.put({
                "type": "error",
                "message": str(e)
            })

    # ---------------------------------
    # Send progress from worker
    # ---------------------------------

    def send_progress(self, percent, message):

        self.queue.put({
            "type": "progress",
            "percent": percent,
            "message": message
        })

    # ---------------------------------
    # Poll queue safely on UI thread
    # ---------------------------------

    def poll_queue(self):

        try:

            while True:

                item = self.queue.get_nowait()

                if item["type"] == "progress":

                    self.progress["value"] = item["percent"]

                    self.status_label.config(
                        text=item["message"]
                    )

                elif item["type"] == "complete":

                    self.progress["value"] = 100

                    self.status_label.config(
                        text="Complete"
                    )

                    self.run_button.config(
                        state="normal"
                    )

                    self.save_preferences()


                elif item["type"] == "error":

                    self.run_button.config(
                        state="normal"
                    )

                    self.status_label.config(
                        text="Error"
                    )

                    messagebox.showerror(
                        "Error",
                        item["message"]
                    )

        except queue.Empty:
            pass

        # poll again after 100ms
        self.root.after(100, self.poll_queue)


# -------------------------------------
# Start app
# -------------------------------------
def check_preferences():

    try:
        print("Checking preferences...")
        preferences = open("preferences.txt").read().strip().split(",")
        print(f"Preferences found: {preferences}")
        if preferences and len(preferences) >= 3:
            print(preferences)
            return preferences
        return None
    except FileNotFoundError:
        return None
    except Exception as e:
        print(f"Error reading preferences: {e}")
        return None
    
root = tk.Tk()

app = ExcelProcessorApp(root)

root.mainloop()