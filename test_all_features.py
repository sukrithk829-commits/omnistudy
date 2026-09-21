import unittest
import os
import json
import sqlite3
import tempfile
import io
from werkzeug.security import generate_password_hash

import app

class ComprehensiveSystemScan(unittest.TestCase):
    def setUp(self):
        self.db_fd, self.db_path = tempfile.mkstemp(suffix=".db")
        app.DATABASE = self.db_path
        app.app.config['TESTING'] = True
        app.app.config['SECRET_KEY'] = 'test-secret-key-12345'
        self.upload_dir = tempfile.mkdtemp()
        app.app.config['UPLOAD_FOLDER'] = self.upload_dir
        
        with app.app.app_context():
            app.init_db()
            db = app.get_db()
            # Student 1 (Verified)
            db.execute("INSERT OR REPLACE INTO users (id, name, email, password_hash, role, department, semester, last_login, is_verified, graduation_status) VALUES (10, 'Student One', 's1@test.com', ?, 'student', 'CS', '6', ?, 1, 'active')", (generate_password_hash('pass12345'), app.now_iso()))
            # Student 2 (Verified)
            db.execute("INSERT OR REPLACE INTO users (id, name, email, password_hash, role, department, semester, last_login, is_verified, graduation_status) VALUES (11, 'Student Two', 's2@test.com', ?, 'student', 'CS', '6', ?, 1, 'active')", (generate_password_hash('pass12345'), app.now_iso()))
            # Student 3 (Unverified)
            db.execute("INSERT OR REPLACE INTO users (id, name, email, password_hash, role, department, semester, last_login, is_verified, graduation_status) VALUES (12, 'Student Unverified', 's3@test.com', ?, 'student', 'CS', '6', ?, 0, 'active')", (generate_password_hash('pass12345'), app.now_iso()))
            # Staff 1
            db.execute("INSERT OR REPLACE INTO users (id, name, email, password_hash, role, department, semester, last_login, is_verified, graduation_status) VALUES (20, 'Staff Asha', 'staff1@test.com', ?, 'staff', 'CS', 'All', ?, 1, 'active')", (generate_password_hash('pass12345'), app.now_iso()))
            # Staff 2 (New Staff with 0 notes)
            db.execute("INSERT OR REPLACE INTO users (id, name, email, password_hash, role, department, semester, last_login, is_verified, graduation_status) VALUES (21, 'Staff New', 'staff2@test.com', ?, 'staff', 'Math', 'All', ?, 1, 'active')", (generate_password_hash('pass12345'), app.now_iso()))
            # Admin
            db.execute("INSERT OR REPLACE INTO users (id, name, email, password_hash, role, department, semester, last_login, is_verified, graduation_status) VALUES (30, 'Admin Super', 'admin1@test.com', ?, 'admin', 'Admin', 'All', ?, 1, 'active')", (generate_password_hash('pass12345'), app.now_iso()))
            db.commit()

        self.client = app.app.test_client()

    def tearDown(self):
        os.close(self.db_fd)
        if os.path.exists(self.db_path):
            try: os.remove(self.db_path)
            except: pass
        if os.path.exists(self.upload_dir):
            try:
                import shutil
                shutil.rmtree(self.upload_dir)
            except: pass

    def login_as(self, user_id):
        with self.client.session_transaction() as sess:
            sess['user_id'] = user_id
            sess['_csrf_token'] = 'test-token-fixed'

    def get_csrf(self):
        return 'test-token-fixed'

    def test_unverified_login_flow(self):
        # Fetch login page first to initialize session
        self.client.get('/login')
        with self.client.session_transaction() as sess:
            sess['_csrf_token'] = self.get_csrf()
        res = self.client.post('/login', data={'email': 's3@test.com', 'password': 'pass12345', '_csrf_token': self.get_csrf()}, follow_redirects=True)
        self.assertIn(b"waiting for administrator verification", res.data)

    def test_verified_student_login_flow(self):
        self.client.get('/login')
        with self.client.session_transaction() as sess:
            sess['_csrf_token'] = self.get_csrf()
        res = self.client.post('/login', data={'email': 's1@test.com', 'password': 'pass12345', '_csrf_token': self.get_csrf()}, follow_redirects=True)
        self.assertEqual(res.status_code, 200)
        self.assertIn(b"Student Academic Workspace", res.data)

    def test_rbac_restrictions(self):
        self.login_as(10)
        self.assertEqual(self.client.get('/admin').status_code, 403)
        self.assertEqual(self.client.get('/staff/analytics').status_code, 403)

        self.login_as(20)
        self.assertEqual(self.client.get('/staff/analytics').status_code, 200)
        self.assertEqual(self.client.get('/admin').status_code, 403)

        self.login_as(30)
        self.assertEqual(self.client.get('/admin').status_code, 200)

    def test_staff_dashboard_and_export_with_zero_documents(self):
        self.login_as(21)
        res = self.client.get('/dashboard')
        self.assertEqual(res.status_code, 200)
        self.assertIn(b"Teaching Intelligence Hub", res.data)

        res_csv = self.client.get('/staff/export-roster')
        self.assertEqual(res_csv.status_code, 200)
        self.assertEqual(res_csv.mimetype, 'text/csv')

    def test_document_lifecycle_and_cascade_delete(self):
        self.login_as(10)
        res = self.client.post('/documents/new', data={
            'title': 'Test Document Lifecycle',
            'subject': 'Testing',
            'department': 'CS',
            'semester': '6',
            'content': 'A relational database stores structured tables. Primary keys enforce uniqueness.',
            'visibility': 'public',
            'access_mode': 'all',
            '_csrf_token': self.get_csrf()
        }, follow_redirects=True)
        self.assertEqual(res.status_code, 200)

        with app.app.app_context():
            doc = app.get_db().execute("SELECT id FROM documents WHERE title = 'Test Document Lifecycle'").fetchone()
            doc_id = doc['id']

        # Access 4 innovation engines
        self.assertEqual(self.client.get(f'/documents/{doc_id}/flashcards').status_code, 200)
        self.assertEqual(self.client.get(f'/documents/{doc_id}/concept-graph').status_code, 200)
        self.assertEqual(self.client.get(f'/documents/{doc_id}/cheat-sheet').status_code, 200)
        self.assertEqual(self.client.get(f'/documents/{doc_id}/viva').status_code, 200)
        self.assertEqual(self.client.get(f'/documents/{doc_id}/quiz').status_code, 200)

        # Record progress in flashcard and viva
        self.client.post(f'/documents/{doc_id}/flashcards/review', data={'card_id': 'c1', 'rating': 'good', '_csrf_token': self.get_csrf()})
        self.client.post(f'/documents/{doc_id}/viva/evaluate', data={'question_id': '1', 'student_answer': 'Primary keys ensure uniqueness', '_csrf_token': self.get_csrf()})

        # Admin delete document -> verify cascade
        self.login_as(30)
        del_res = self.client.post(f'/documents/{doc_id}/delete', data={'_csrf_token': self.get_csrf()}, follow_redirects=True)
        self.assertEqual(del_res.status_code, 200)

        with app.app.app_context():
            db = app.get_db()
            self.assertIsNone(db.execute("SELECT 1 FROM documents WHERE id = ?", (doc_id,)).fetchone())
            self.assertIsNone(db.execute("SELECT 1 FROM flashcard_progress WHERE document_id = ?", (doc_id,)).fetchone())
            self.assertIsNone(db.execute("SELECT 1 FROM viva_history WHERE document_id = ?", (doc_id,)).fetchone())

    def test_admin_user_verification(self):
        self.login_as(30)
        res = self.client.post('/admin/users/12/verify', data={'_csrf_token': self.get_csrf()}, follow_redirects=True)
        self.assertEqual(res.status_code, 200)

        with app.app.app_context():
            user = app.get_db().execute("SELECT is_verified FROM users WHERE id = 12").fetchone()
            self.assertEqual(user['is_verified'], 1)

    def test_inactivity_watchdog_execution(self):
        with app.app.app_context():
            released = app.run_inactivity_watchdog()
            self.assertIsInstance(released, int)

if __name__ == '__main__':
    unittest.main()
