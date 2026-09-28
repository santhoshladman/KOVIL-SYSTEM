# Kovil Kanakku — XAMPP + MySQL Setup

## What's in this folder
| File | Purpose |
|---|---|
| `index.php` | The app itself — open this in the browser |
| `config.php` | Database connection settings (edit if needed) |
| `db.php` | Internal DB connection helper — no need to edit |
| `load_state.php` | API the app calls to **load** data from MySQL |
| `save_state.php` | API the app calls to **save** data to MySQL |
| `schema.sql` | Run once in phpMyAdmin to create the database |

## Setup steps

1. **Open XAMPP Control Panel** → click **Start** next to both **Apache** and **MySQL**. Both rows must turn green.

2. **Copy this whole folder** into your XAMPP `htdocs` folder, e.g.:
   ```
   C:\xampp\htdocs\kovil-backend\
   ```
   (Windows) or `/opt/lampp/htdocs/kovil-backend/` (Linux XAMPP). Keep all 6 files together in the same folder — don't separate them.

3. **Create the database.** Open `http://localhost/phpmyadmin` in your browser →
   - Click the **Import** tab
   - Click **Choose File** → select `schema.sql` from this folder
   - Click **Go**

   This creates a database called `kovil-backend` with one table (`app_state`) that holds all your temple's data.

4. **Open the app.** Go to:
   ```
   http://localhost/kovil-backend/
   ```
   (use whatever folder name you actually used in step 2). The app should load normally.

5. **Test it** — add a donor or fund account, then refresh the page. If it's still there, the MySQL connection is working.

## If it shows "Could not reach server" / "Database connection failed"

- Make sure Apache **and** MySQL are both green/running in the XAMPP Control Panel.
- Make sure you ran step 3 (imported `schema.sql`) — the app needs the `kovil-backend` database to exist first.
- If your MySQL root user has a password (not the XAMPP default), open `config.php` and update `DB_USER` / `DB_PASS` to match.

## Good to know

- Data is now stored **centrally in MySQL**, not per-browser. Anyone on the same network who opens `http://<your-pc-ip>/kovil-backend/` will see and edit the **same shared data** — this is different from before, where each browser had its own separate copy.
- Bulk Import / Bulk Export (Excel) work exactly as before — they read/write the same MySQL-backed data.
