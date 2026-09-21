import sqlite3
import unittest
from datetime import UTC, datetime
from werkzeug.security import generate_password_hash

import app as omnistudy_app


class ProfileApprovalWorkflowTest(unittest.TestCase):
    def setUp(self):
        omnistudy_app.app.config.update(
            TESTING=True,
            SECRET_KEY="test-secret-key",
            WTF_CSRF_ENABLED=False,
        )
        self.client = omnistudy_app.app.test_client()

        # Create isolated test database
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        self._orig_get_db = omnistudy_app.get_db
        omnistudy_app.get_db = lambda: self.db

        with omnistudy_app.app.app_context():
            omnistudy_app.init_db()

    def tearDown(self):
        omnistudy_app.get_db = self._orig_get_db
        self.db.close()

    def login(self, email, password):
        return self.client.post("/login", data={"email": email, "password": password}, follow_redirects=True)

    def test_student_profile_update_and_admin_approval_lifecycle(self):
        # 1. Log in as student
        resp = self.login("student@omnistudy.test", "student123")
        self.assertEqual(resp.status_code, 200)

        # 2. View profile page
        resp = self.client.get("/profile")
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Active Profile", resp.data)
        self.assertIn(b"Ravi Student", resp.data)

        # 3. Submit edit profile request
        resp = self.client.post(
            "/profile/edit",
            data={
                "name": "Ravi Kumar",
                "reg_no": "22BCA999",
                "department": "Computer Applications",
                "semester": "6",
            },
            follow_redirects=True,
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Profile update request submitted successfully", resp.data)

        # 4. Verify student's name in users table is NOT yet changed
        student = self.db.execute("SELECT * FROM users WHERE email = 'student@omnistudy.test'").fetchone()
        self.assertEqual(student["name"], "Ravi Student")
        self.assertNotEqual(student["name"], "Ravi Kumar")

        # 5. Verify pending request in profile_update_requests
        req = self.db.execute("SELECT * FROM profile_update_requests WHERE user_id = ? AND status = 'pending'", (student["id"],)).fetchone()
        self.assertIsNotNone(req)
        self.assertEqual(req["new_name"], "Ravi Kumar")
        self.assertEqual(req["new_reg_no"], "22BCA999")
        self.assertEqual(req["status"], "pending")

        # 6. Verify dashboard displays the pending banner
        resp = self.client.get("/dashboard")
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Profile Update Request Pending Review", resp.data)

        # 7. Student cannot approve their own request (Forbidden)
        resp = self.client.post(f"/admin/profile-requests/{req['id']}/approve")
        self.assertEqual(resp.status_code, 403)

        # 8. Log in as Admin
        self.client.post("/logout")
        self.login("admin@omnistudy.test", "admin123")

        # 9. View admin page - pending profile request appears
        resp = self.client.get("/admin")
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Student Profile Change Requests", resp.data)
        self.assertIn(b"Ravi Kumar", resp.data)
        self.assertIn(b"22BCA999", resp.data)

        # 10. Admin approves request
        resp = self.client.post(f"/admin/profile-requests/{req['id']}/approve", follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Profile update approved for Ravi Kumar", resp.data)

        # 11. Verify student profile in users table is now updated
        updated_student = self.db.execute("SELECT * FROM users WHERE email = 'student@omnistudy.test'").fetchone()
        self.assertEqual(updated_student["name"], "Ravi Kumar")
        self.assertEqual(updated_student["reg_no"], "22BCA999")
        self.assertEqual(updated_student["department"], "Computer Applications")

        # 12. Verify request status is now 'approved'
        updated_req = self.db.execute("SELECT * FROM profile_update_requests WHERE id = ?", (req["id"],)).fetchone()
        self.assertEqual(updated_req["status"], "approved")

    def test_student_profile_update_rejection(self):
        # Student submits request
        self.login("student@omnistudy.test", "student123")
        self.client.post(
            "/profile/edit",
            data={
                "name": "Invalid Fake Name",
                "reg_no": "FAKE123",
                "department": "Wrong Dept",
                "semester": "1",
            },
            follow_redirects=True,
        )
        req = self.db.execute("SELECT * FROM profile_update_requests WHERE status = 'pending'").fetchone()

        # Admin rejects
        self.client.post("/logout")
        self.login("admin@omnistudy.test", "admin123")
        resp = self.client.post(f"/admin/profile-requests/{req['id']}/reject", follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Profile update request rejected", resp.data)

        # Verify users table NOT modified
        student = self.db.execute("SELECT * FROM users WHERE email = 'student@omnistudy.test'").fetchone()
        self.assertEqual(student["name"], "Ravi Student")
        self.assertNotEqual(student["name"], "Invalid Fake Name")

    def test_cancel_pending_profile_request(self):
        self.login("student@omnistudy.test", "student123")
        self.client.post(
            "/profile/edit",
            data={
                "name": "Change Mind",
                "reg_no": "22BCA555",
                "department": "CS",
                "semester": "6",
            },
        )
        # Verify request exists
        req = self.db.execute("SELECT * FROM profile_update_requests WHERE status = 'pending'").fetchone()
        self.assertIsNotNone(req)

        # Cancel request
        resp = self.client.post("/profile/cancel-request", follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"pending profile update request has been cancelled", resp.data)

        # Verify request removed
        req_after = self.db.execute("SELECT * FROM profile_update_requests WHERE status = 'pending'").fetchone()
        self.assertIsNone(req_after)


if __name__ == "__main__":
    unittest.main()
