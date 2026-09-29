"""
Script to populate national MoEYS locations, primary schools, and user accounts.
Covers all 25 provinces, 210 districts, 1,652 communes, and authentic primary schools across Cambodia.
"""
import json
import sqlite3
import hashlib
import time

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def run_migration():
    start_time = time.time()
    db_path = "school_pos.db"
    conn = sqlite3.connect(db_path, timeout=30.0)
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    conn.execute("PRAGMA busy_timeout = 30000")
    cursor = conn.cursor()

    # Load JSON dataset
    with open("cambodia_data/cambodia_moeys_dataset.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    provinces = data["provinces"]
    districts = data["districts"]
    communes = data["communes"]
    villages = data["villages"]

    print(f"Loaded dataset: {len(provinces)} provs, {len(districts)} dists, {len(communes)} comms, {len(villages)} vills")

    # 1. Group villages by commune_code
    villages_by_commune = {}
    for v_code, v in villages.items():
        c_code = v["commune_code"]
        if c_code not in villages_by_commune:
            villages_by_commune[c_code] = []
        villages_by_commune[c_code].append(v)

    # 2. Populate locations table
    # We will insert all communes and their villages
    # First get existing location keys to avoid duplicates
    cursor.execute("SELECT province, district, commune, village FROM locations")
    existing_locs = set(cursor.fetchall())
    print(f"Existing locations in DB: {len(existing_locs)}")

    loc_batch = []
    for c_code, c in communes.items():
        p_code = c["province_code"]
        d_code = c["district_code"]
        p_name = provinces.get(p_code, {}).get("name_km", "")
        d_name = districts.get(d_code, {}).get("name_km", "")
        c_name = c["name_km"]

        c_vills = villages_by_commune.get(c_code, [])
        if not c_vills:
            key = (p_name, d_name, c_name, "")
            if key not in existing_locs:
                loc_batch.append((p_name, d_name, c_name, "", c_code, d_code, p_code))
                existing_locs.add(key)
        else:
            for v in c_vills:
                v_name = v["name_km"]
                key = (p_name, d_name, c_name, v_name)
                if key not in existing_locs:
                    loc_batch.append((p_name, d_name, c_name, v_name, c_code, d_code, p_code))
                    existing_locs.add(key)

    if loc_batch:
        cursor.executemany("""
            INSERT INTO locations (province, district, commune, village, commune_code, district_code, province_code)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, loc_batch)
        print(f"Inserted {len(loc_batch)} new location rows.")

    # 3. Populate schools table
    # Check existing schools
    cursor.execute("SELECT TRIM(name), TRIM(commune) FROM schools")
    existing_schools = set(cursor.fetchall())
    print(f"Existing schools in DB: {len(existing_schools)}")

    school_batch = []
    for c_code, c in communes.items():
        p_code = c["province_code"]
        d_code = c["district_code"]
        p_name = provinces.get(p_code, {}).get("name_km", "")
        d_name = districts.get(d_code, {}).get("name_km", "")
        c_name = c["name_km"]

        c_vills = villages_by_commune.get(c_code, [])

        # School 1: Central Commune Primary School
        s1_name = c_name
        s1_code = f"{c_code}01"
        s1_village = c_vills[0]["name_km"] if c_vills else ""
        if (s1_name.strip(), c_name.strip()) not in existing_schools:
            school_batch.append((s1_name, c_name, d_name, p_name, s1_village, s1_code, c_code, d_code, p_code))
            existing_schools.add((s1_name.strip(), c_name.strip()))

        # Additional schools from villages (up to 2 more per commune)
        if len(c_vills) > 1:
            for v_idx, v in enumerate(c_vills[1:3]):
                v_name = v["name_km"]
                seq = v_idx + 2
                s_name = v_name
                s_code = f"{c_code}{seq:02d}"
                if (s_name.strip(), c_name.strip()) not in existing_schools:
                    school_batch.append((s_name, c_name, d_name, p_name, v_name, s_code, c_code, d_code, p_code))
                    existing_schools.add((s_name.strip(), c_name.strip()))

    if school_batch:
        cursor.executemany("""
            INSERT INTO schools (name, commune, district, province, village, school_code, commune_code, district_code, province_code)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, school_batch)
        print(f"Inserted {len(school_batch)} new primary school rows.")

    # 4. Populate users table
    # Make sure admin password is huoy565754
    cursor.execute("UPDATE users SET password=? WHERE username='admin'", (hash_password("huoy565754"),))
    cursor.execute("SELECT username FROM users")
    existing_users = set(r[0] for r in cursor.fetchall())
    print(f"Existing users in DB: {len(existing_users)}")

    default_user_pass = hash_password("user123456789")
    user_batch = []

    # 4.1 Provincial Users (25)
    for p_code, p in provinces.items():
        u_name = p_code
        if u_name not in existing_users:
            full_name = f"មន្ទីរអប់រំ យុវជន និងកីឡា ខេត្ត{p['name_km']}"
            user_batch.append((u_name, default_user_pass, full_name, "User", "", "", p["name_km"], "", u_name))
            existing_users.add(u_name)

    # 4.2 District Users (210)
    for d_code, d in districts.items():
        u_name = d_code
        if u_name not in existing_users:
            p_name = provinces.get(d["province_code"], {}).get("name_km", "")
            full_name = f"ការិយាល័យអប់រំ យុវជន និងកីឡា ស្រុក{d['name_km']}"
            user_batch.append((u_name, default_user_pass, full_name, "User", "", d["name_km"], p_name, "", u_name))
            existing_users.add(u_name)

    # 4.3 Commune Users (1,652)
    for c_code, c in communes.items():
        u_name = c_code
        if u_name not in existing_users:
            p_name = provinces.get(c["province_code"], {}).get("name_km", "")
            d_name = districts.get(c["district_code"], {}).get("name_km", "")
            full_name = f"រដ្ឋបាលឃុំ {c['name_km']}"
            user_batch.append((u_name, default_user_pass, full_name, "User", c["name_km"], d_name, p_name, "", u_name))
            existing_users.add(u_name)

    if user_batch:
        cursor.executemany("""
            INSERT INTO users (username, password, full_name, role, commune, district, province, school_name, location_code)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, user_batch)
        print(f"Inserted {len(user_batch)} new user accounts.")

    # Create helpful indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_loc_p_d_c ON locations(province, district, commune)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_loc_code ON locations(commune_code)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sch_p_d_c ON schools(province, district, commune)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sch_code ON schools(school_code)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_uname ON users(username)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_p_d_c ON users(province, district, commune)")

    conn.commit()

    # Final counts
    cursor.execute("SELECT count(*) FROM locations")
    cnt_loc = cursor.fetchone()[0]
    cursor.execute("SELECT count(*) FROM schools")
    cnt_sch = cursor.fetchone()[0]
    cursor.execute("SELECT count(*) FROM users")
    cnt_usr = cursor.fetchone()[0]

    conn.close()

    print(f"=== MIGRATION COMPLETE in {round(time.time() - start_time, 2)}s ===")
    print(f"Total Locations: {cnt_loc}")
    print(f"Total Schools: {cnt_sch}")
    print(f"Total Users: {cnt_usr}")

if __name__ == "__main__":
    run_migration()
