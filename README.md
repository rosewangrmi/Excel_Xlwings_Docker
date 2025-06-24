## 🚀 Excel Macro Runner (Hybrid Local Architecture)

This setup allows you to run Excel macros on a local Windows machine using `xlwings`, triggered by Docker-based orchestration.

### 📁 Project Structure

```
project/
├── hybrid_local/
│   ├── watcher.py           # Host-side macro executor
│   ├── Dockerfile           # Containerized job sender
│   ├── run_macro.py         # Writes job instructions
│   ├── job_config.json      # Sample macro job config
│   ├── start_watcher.bat    # Launch the watcher on Windows
│   └── docker-compose.yml   # Docker orchestration setup
```

### 🧱 Architecture Diagram

```
graph TD
    A[Docker Container - Orchestrator] -->|Job File + Excel Input| B[Windows Host Shared Folder]
    B --> C[xlwings Watcher - Python]
    C --> D[Excel Macro Execution via COM]
    D --> B
```

### 🛠️ Setup Instructions

#### 1. Prepare Shared Directory
Create a local shared folder on your Windows host (e.g. `C:\shared`) and ensure Docker Desktop has access to it.

#### 2. Install Requirements on Host
- Excel (Office 365, LTSC, or equivalent)
- Python 3.x
- `pip install xlwings`

#### 3. Run Watcher
From the Windows host:
```bash
start start_watcher.bat
```

#### 4. Run the Docker Container
From your project root:
```bash
docker-compose up --build
```

#### 5. Results
- Job is written by Docker to `/shared/job_config.json`
- `watcher.py` picks it up, runs the macro, and renames the job file
- Output Excel file is updated in the shared folder

### ✅ Notes
- This approach supports **only one macro job at a time**.
- Parallel execution would require multiple user sessions or Windows VMs.

### 📦 Packaging
To package:
- Zip the `project/hybrid_local/` folder for distribution
- Include README as documentation
