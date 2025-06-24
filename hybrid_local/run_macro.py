import json
import time

job = {
    "input": "input.xlsx",
    "output": "output.xlsx",
    "macro": "main_macro"
}

with open("/shared/job_config.json", "w") as f:
    json.dump(job, f)

print("Job file written to shared directory.")
