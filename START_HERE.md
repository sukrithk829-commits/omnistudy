# OmniStudy — Setup & Migration Guide

## Quick Start on New Device

### Option A: Windows (Easiest — 1-Click)
1. Extract this zip file into any folder (e.g. `C:\OmniStudy` or `D:\BCA`).
2. Double-click `run_app.bat`.
   * It will automatically create a local virtual environment, install all dependencies from `requirements.txt`, and start the Flask server.
3. Open your browser and go to:
   **http://127.0.0.1:5000**

---

### Option B: macOS / Linux (1-Click)
1. Extract the zip file into your desired directory.
2. Open Terminal in the extracted folder and run:
   ```bash
   chmod +x run_app.sh
   ./run_app.sh
   ```
3. Open your browser at:
   **http://127.0.0.1:5000**

---

### Option C: Manual Setup (Any OS)
1. Open terminal / command prompt in the extracted directory.
2. Create and activate a virtual environment:
   - **Windows:**
     ```cmd
     python -m venv .venv
     .venv\Scripts\activate
     ```
   - **macOS / Linux:**
     ```bash
     python3 -m venv .venv
     source .venv/bin/activate
     ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Start the application:
   ```bash
   python app.py
   ```
5. Navigate to: **http://127.0.0.1:5000**

---

## Demo Accounts Included in `omnistudy.db`

| Role    | Email                    | Password     | Description |
|---------|--------------------------|--------------|-------------|
| Student | `student@omnistudy.test` | `student123` | Complete student workspace with quizzes, 3D flashcards, notes |
| Staff   | `staff@omnistudy.test`   | `staff123`   | Staff material creation, analytics, document controls |
| Admin   | `admin@omnistudy.test`   | `admin123`   | Full administrative panel, user & document management |

---

## Included in this Package
* **`app.py`**: Core Flask application with full REST API, document ingestion, AI/NLP quiz generators, cheat sheet engine, flashcards, concept maps, and security modules.
* **`omnistudy.db`**: Pre-seeded SQLite database with registered users, study materials, quiz questions, and scores.
* **`templates/`**: Full responsive liquid glass web interface templates (18 Jinja2 HTML templates).
* **`uploads/`**: Uploaded sample documents and study resources.
* **`presentation_assets/` & `framed_assets/`**: High-resolution visuals, UI previews, and design assets.
* **PowerPoint Presentations (`*.pptx`) & Reports (`*.docx`)**: Complete project presentation decks and documentation reports.
* **`.git/`**: Full Git repository history and commit log.
* **Test Suite**: Automated tests (`test_routes_and_security.py`, `test_all_features.py`, `test_profile_and_approval.py`).
