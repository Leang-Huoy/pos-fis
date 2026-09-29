"""
Migration script for MoEYS POS system:
1. Clean school names: strip 'សាលាបឋមសិក្សា ' prefix, keeping only the remaining name.
2. Resolve duplicates in schools and update references in daily_records, purchases, suppliers, products.
3. Delete school user accounts (length 8) from users table.
4. Update Admin password to 'huoy565754'.
"""
import sqlite3
import hashlib
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def run_migration():
    conn = sqlite3.connect("school_pos.db")
    cursor = conn.cursor()

    print("=== 1. CLEANING PRIMARY SCHOOL NAMES ===")
    cursor.execute("SELECT id, name, commune, school_code FROM schools")
    all_schools = cursor.fetchall()
    print(f"Total schools before cleaning: {len(all_schools)}")

    # Map existing clean schools to avoid duplicates
    # First, let's identify what new names will be
    updated_schools = []
    to_delete_ids = []
    seen = {}

    for sid, name, comm, scode in all_schools:
        clean_name = str(name).strip()
        if clean_name.startswith("សាលាបឋមសិក្សា"):
            rem = clean_name[len("សាលាបឋមសិក្សា"):].strip()
            if rem:
                clean_name = rem

        key = (clean_name, str(comm or "").strip())
        if key in seen:
            prev_sid, prev_scode = seen[key]
            # Duplicate found! We will merge: keep the one with school_code or lower id
            to_delete_ids.append(sid)
            # Update references to the previous school
            old_name = name
            cursor.execute("UPDATE daily_records SET school_name=? WHERE school_name=?", (clean_name, old_name))
            cursor.execute("UPDATE purchases SET school_name=? WHERE school_name=?", (clean_name, old_name))
            cursor.execute("UPDATE suppliers SET school_name=? WHERE school_name=?", (clean_name, old_name))
            cursor.execute("UPDATE products SET school_name=? WHERE school_name=?", (clean_name, old_name))
            if scode and not prev_scode:
                cursor.execute("UPDATE schools SET school_code=? WHERE id=?", (scode, prev_sid))
        else:
            seen[key] = (sid, scode)
            if clean_name != name:
                old_name = name
                cursor.execute("UPDATE schools SET name=? WHERE id=?", (clean_name, sid))
                cursor.execute("UPDATE daily_records SET school_name=? WHERE school_name=?", (clean_name, old_name))
                cursor.execute("UPDATE purchases SET school_name=? WHERE school_name=?", (clean_name, old_name))
                cursor.execute("UPDATE suppliers SET school_name=? WHERE school_name=?", (clean_name, old_name))
                cursor.execute("UPDATE products SET school_name=? WHERE school_name=?", (clean_name, old_name))

    if to_delete_ids:
        ph = ",".join(["?"] * len(to_delete_ids))
        cursor.execute(f"DELETE FROM schools WHERE id IN ({ph})", to_delete_ids)
        print(f"Removed {len(to_delete_ids)} duplicate school rows after cleaning names.")

    cursor.execute("SELECT count(*) FROM schools")
    total_clean_sch = cursor.fetchone()[0]
    print(f"Total schools after cleaning: {total_clean_sch}")

    # Verify sample cleaned school names
    sample_sch = cursor.execute("SELECT name FROM schools LIMIT 10").fetchall()
    print("Sample cleaned school names:", [r[0] for r in sample_sch])

    print("\n=== 2. DELETING SCHOOL USERS ===")
    cursor.execute("SELECT count(*) FROM users WHERE LENGTH(username) = 8")
    sch_users_cnt = cursor.fetchone()[0]
    print(f"School user accounts to delete: {sch_users_cnt}")
    cursor.execute("DELETE FROM users WHERE LENGTH(username) = 8")

    print("\n=== 3. UPDATING ADMIN PASSWORD ===")
    new_admin_pass = hash_password("huoy565754")
    cursor.execute("UPDATE users SET password=? WHERE username='admin'", (new_admin_pass,))
    if cursor.rowcount == 0:
        cursor.execute("""
            INSERT INTO users (username, password, full_name, role)
            VALUES ('admin', ?, 'Administrator', 'Admin')
        """, (new_admin_pass,))
        print("Inserted new admin account with password huoy565754.")
    else:
        print("Updated admin password to huoy565754.")

    conn.commit()

    # User breakdown verification
    cursor.execute("SELECT LENGTH(username), COUNT(*) FROM users GROUP BY LENGTH(username)")
    user_counts = cursor.fetchall()
    print("\nRemaining user accounts breakdown:")
    for length, count in sorted(user_counts):
        if length == 2:
            lbl = "មន្ទីរអប់រំខេត្ត (Provincial)"
        elif length == 4:
            lbl = "ការិយាល័យអប់រំស្រុក (District)"
        elif length == 5:
            lbl = "Admin (Administrator)"
        elif length == 6:
            lbl = "រដ្ឋបាលឃុំ/សង្កាត់ (Commune)"
        else:
            lbl = f"Other (len {length})"
        print(f" - {lbl}: {count} គណនី")

    cursor.execute("SELECT count(*) FROM users")
    total_users = cursor.fetchone()[0]
    print(f"Total active users: {total_users}")

    conn.close()
    print("\nMigration completed successfully!")

if __name__ == "__main__":
    run_migration()
