import xlwings as xw
import json
import os
import time

WATCH_DIR = "C:\\shared"

def run_macro(filename, macro_name):
    app = xw.App(visible=False)
    try:
        wb = app.books.open(os.path.join(WATCH_DIR, filename))
        wb.macro(macro_name)()
        wb.save()
        wb.close()
    finally:
        app.quit()

while True:
    try:
        job_path = os.path.join(WATCH_DIR, "job_config.json")
        if os.path.exists(job_path):
            with open(job_path, "r") as f:
                job = json.load(f)

            print(f"Running {job['macro']} on {job['input']}")
            run_macro(job["input"], job["macro"])
            os.rename(job_path, job_path + ".done")

    except Exception as e:
        print(f"Error: {e}")

    time.sleep(5)
