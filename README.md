# OmniStudy

A Flask prototype for an interconnected student–staff academic knowledge hub.

## Included modules

- Role-aware student, staff, and administrator workspaces.
- Registration requests with administrator verification and audit records.
- Public repository and creator-only private academic locker.
- PDF, TXT, and Markdown material uploads with text extraction and protected original-file downloads.
- Automatic quiz generation from uploaded notes and text-based PDFs.
- Student-controlled public/private sharing, selected-student access, and blocked-student access for public materials.
- Per-material uploader insight pages showing unique student viewers, total views, identities, and timestamps.
- Document branching plus side-by-side comparison of changes.
- Persisted background generation of extractive quick reads and MCQ self-assessments.
- Staff engagement metrics for views, distinct learners, quiz attempts, and average scores.
- Administrator activity/system-health views and graduate lifecycle controls.
- A daily inactivity watchdog that releases opted-in notes only after a student is marked graduated.
- CSRF protection on all state-changing forms and password hashing for accounts.

## Run locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`.

## Gemini-powered study packs

Set `GEMINI_API_KEY` before starting the app (copy `.env.example` to a local `.env`, or set it in your shell). Each note receives one cached Gemini study pack, shared by every authorised reader: a quick read plus a 30-question bank (Easy, Medium, and Hard). Learners choose a focus and 5–30 questions without causing any further Gemini request. Their results are saved to personal performance and staff analytics. The default model is `gemini-3.5-flash`; override it with `GEMINI_MODEL`. If the key or service is unavailable, OmniStudy automatically uses its local study generator.

## Demo accounts

| Role | Email | Password |
| --- | --- | --- |
| Student | `student@omnistudy.test` | `student123` |
| Staff | `staff@omnistudy.test` | `staff123` |
| Admin | `admin@omnistudy.test` | `admin123` |

This starter uses SQLite and local uploads. Before production, set a secure `OMNISTUDY_SECRET_KEY`, move file storage to managed object storage, configure institutional email delivery, add password-reset policies, and run the scheduler through a production process manager.
