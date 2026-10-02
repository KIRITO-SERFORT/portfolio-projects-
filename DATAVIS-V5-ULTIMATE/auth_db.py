import sqlite3
import hashlib
import os
import secrets
from datetime import datetime

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "datavis_users.db")


def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Initializes SQLite database for user accounts, turn tracking, and UPI transaction history."""
    conn = get_connection()
    cursor = conn.cursor()

    # Users Table with is_paid_tier column
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        name TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        salt TEXT NOT NULL,
        turns_remaining INTEGER NOT NULL DEFAULT 3,
        total_analyses INTEGER NOT NULL DEFAULT 0,
        is_paid_tier BOOLEAN NOT NULL DEFAULT 0,
        role TEXT NOT NULL DEFAULT 'user',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # Check if is_paid_tier column exists for backward compatibility
    cursor.execute("PRAGMA table_info(users);")
    cols = [col['name'] for col in cursor.fetchall()]
    if 'is_paid_tier' not in cols:
        cursor.execute("ALTER TABLE users ADD COLUMN is_paid_tier BOOLEAN NOT NULL DEFAULT 0;")

    # Analyses History Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS analyses_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        analysis_type TEXT NOT NULL,
        dataset_info TEXT,
        files_count INTEGER DEFAULT 1,
        extra_files_fee REAL DEFAULT 0,
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users (id)
    );
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        amount_inr REAL NOT NULL DEFAULT 0,
        turns_added INTEGER NOT NULL DEFAULT 0,
        payment_method TEXT DEFAULT 'GPay / UPI',
        utr_number TEXT NOT NULL DEFAULT '',
        status TEXT DEFAULT 'APPROVED',
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users (id)
    );
    """)

    cursor.execute("PRAGMA table_info(transactions);")
    tx_cols = [col['name'] for col in cursor.fetchall()]
    if tx_cols and ('amount_usd' in tx_cols or 'amount_inr' not in tx_cols):
        cursor.execute("DROP TABLE IF EXISTS transactions;")
        cursor.execute("""
        CREATE TABLE transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            amount_inr REAL NOT NULL DEFAULT 0,
            turns_added INTEGER NOT NULL DEFAULT 0,
            payment_method TEXT DEFAULT 'GPay / UPI',
            utr_number TEXT NOT NULL DEFAULT '',
            status TEXT DEFAULT 'APPROVED',
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (id)
        );
        """)

    conn.commit()
    conn.close()


def generate_salt():
    return secrets.token_hex(16)


def hash_password(password: str, salt: str) -> str:
    """PBKDF2 SHA256 password hashing with salt."""
    return hashlib.pbkdf2_hmac(
        'sha256',
        password.encode('utf-8'),
        salt.encode('utf-8'),
        100000
    ).hex()


def register_user(name: str, email: str, password: str) -> tuple[bool, str, dict | None]:
    """Registers a new user with 3 FREE turns."""
    email_clean = email.strip().lower()
    name_clean = name.strip()

    if not name_clean or not email_clean or not password:
        return False, "All fields are required.", None

    if len(password) < 6:
        return False, "Password must be at least 6 characters long.", None

    conn = get_connection()
    cursor = conn.cursor()

    try:
        salt = generate_salt()
        pwd_hash = hash_password(password, salt)
        cursor.execute(
            """
            INSERT INTO users (email, name, password_hash, salt, turns_remaining, total_analyses, is_paid_tier)
            VALUES (?, ?, ?, ?, 3, 0, 0)
            """,
            (email_clean, name_clean, pwd_hash, salt)
        )
        conn.commit()
        user_id = cursor.lastrowid

        user_dict = {
            "id": user_id,
            "name": name_clean,
            "email": email_clean,
            "turns_remaining": 3,
            "total_analyses": 0,
            "is_paid_tier": 0,
            "role": "user"
        }
        return True, "Account registered successfully!", user_dict
    except sqlite3.IntegrityError:
        return False, f"An account with email '{email_clean}' already exists.", None
    finally:
        conn.close()


def authenticate_user(email: str, password: str) -> tuple[bool, str, dict | None]:
    """Authenticates user with email and password."""
    email_clean = email.strip().lower()
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM users WHERE email = ?", (email_clean,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        return False, "Invalid email or password.", None

    user = dict(row)
    computed_hash = hash_password(password, user['salt'])

    if secrets.compare_digest(computed_hash, user['password_hash']):
        return True, "Login successful!", {
            "id": user['id'],
            "name": user['name'],
            "email": user['email'],
            "turns_remaining": user['turns_remaining'],
            "total_analyses": user['total_analyses'],
            "is_paid_tier": user.get('is_paid_tier', 0),
            "role": user['role']
        }
    else:
        return False, "Invalid email or password.", None


def get_user_by_id(user_id: int) -> dict | None:
    """Refreshes and returns the user object from database."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()

    if row:
        user = dict(row)
        return {
            "id": user['id'],
            "name": user['name'],
            "email": user['email'],
            "turns_remaining": user['turns_remaining'],
            "total_analyses": user['total_analyses'],
            "is_paid_tier": user.get('is_paid_tier', 0),
            "role": user['role']
        }
    return None


def get_user_turns(user_id: int) -> int:
    user = get_user_by_id(user_id)
    return user['turns_remaining'] if user else 0


def consume_turn(user_id: int, analysis_type: str = "Data Analysis", dataset_info: str = "", files_count: int = 1, extra_fee: float = 0.0) -> tuple[bool, int, str]:
    """Decrements 1 turn from user balance if turns_remaining > 0."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT turns_remaining, total_analyses FROM users WHERE id = ?", (user_id,))
    row = cursor.fetchone()

    if not row:
        conn.close()
        return False, 0, "User not found."

    turns = row['turns_remaining']
    total = row['total_analyses']

    if turns <= 0:
        conn.close()
        return False, 0, "No analysis turns remaining! Please top up your account balance."

    new_turns = turns - 1
    new_total = total + 1

    cursor.execute(
        "UPDATE users SET turns_remaining = ?, total_analyses = ? WHERE id = ?",
        (new_turns, new_total, user_id)
    )

    cursor.execute(
        """
        INSERT INTO analyses_log (user_id, analysis_type, dataset_info, files_count, extra_files_fee)
        VALUES (?, ?, ?, ?, ?)
        """,
        (user_id, analysis_type, dataset_info, files_count, extra_fee)
    )

    conn.commit()
    conn.close()

    return True, new_turns, f"Analysis executed successfully."


def add_credits_upi(user_id: int, turns_to_add: int, amount_inr: float, utr_number: str = "UPI_DEMO") -> tuple[bool, int, str]:
    """Credits turns to user balance upon UPI / GPay payment (e.g. ₹20 for 2 turns). Mark user as paid tier."""
    if turns_to_add <= 0:
        return False, 0, "Invalid turn count."

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT turns_remaining FROM users WHERE id = ?", (user_id,))
    row = cursor.fetchone()

    if not row:
        conn.close()
        return False, 0, "User not found."

    new_turns = row['turns_remaining'] + turns_to_add

    cursor.execute(
        "UPDATE users SET turns_remaining = ?, is_paid_tier = 1 WHERE id = ?",
        (new_turns, user_id)
    )

    cursor.execute(
        """
        INSERT INTO transactions (user_id, amount_inr, turns_added, payment_method, utr_number, status)
        VALUES (?, ?, ?, 'GPay / PhonePe / UPI', ?, 'APPROVED')
        """,
        (user_id, amount_inr, turns_to_add, utr_number)
    )

    conn.commit()
    conn.close()

    return True, new_turns, f"Successfully credited {turns_to_add} analysis turns! Current balance: {new_turns} turns."


def get_user_history(user_id: int) -> tuple[list[dict], list[dict]]:
    """Returns (analyses_logs, transactions_logs) for a user."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM analyses_log WHERE user_id = ? ORDER BY timestamp DESC", (user_id,))
    analyses = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM transactions WHERE user_id = ? ORDER BY timestamp DESC", (user_id,))
    txns = [dict(r) for r in cursor.fetchall()]

    conn.close()
    return analyses, txns


# Initialize DB on module import
init_db()
