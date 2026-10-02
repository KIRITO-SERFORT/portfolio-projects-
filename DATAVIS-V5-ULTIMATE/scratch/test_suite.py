import os
import sys
import pandas as pd
import numpy as np

sys.path.insert(0, r"D:\python\PYTHON AUTOMATION\DATAVIS-V5-ULTIMATE")

import auth_db
import supabase_db
import app

print("--- TESTING UPDATED AUTH DB & UPI ---")
auth_db.init_db()

email = "upi_user@example.com"
success, msg, u = auth_db.register_user("Rahul Verma", email, "pass1234")
print("Registration test:", success, msg)

if not u:
    success, msg, u = auth_db.authenticate_user(email, "pass1234")

# Test UPI Credit Addition (₹20 for 2 turns)
success, new_bal, msg = auth_db.add_credits_upi(u['id'], 2, 20.0, "UTR123456789012")
print("UPI Rs. 20 Credit test:", success, "New balance:", new_bal, msg)

u_refreshed = auth_db.get_user_by_id(u['id'])
print("Is Paid Tier:", u_refreshed['is_paid_tier'] == 1)

# Test Supabase module schema generation
sql = supabase_db.get_supabase_sql_schema()
print("Supabase SQL Schema generated successfully, length:", len(sql))

print("\n--- TESTING ALL APP FUNCTIONS ---")
dummy_data = app.generate_sample_marks_data()
numeric_cols = ["Mathematics", "Physics", "Chemistry", "English", "Computer Science"]
non_numeric_cols = ["Student Name"]

graded_df = app.compute_student_grades(dummy_data, numeric_cols, "Student Name")
print("Computed student grades count:", len(graded_df), "Top rank grade:", graded_df.iloc[0]['Grade'])

print("\nAll tests executed successfully!")
