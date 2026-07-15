"""
PS Tracker Launcher
-------------------
Click RUN: authenticates via Midway, downloads CSVs, generates report.
"""

import tkinter as tk
from tkinter import scrolledtext
import subprocess, threading, os, sys, json
import requests
from pathlib import Path
from datetime import datetime, timedelta

# ── Paths (all relative to this script's directory) ────────────────────────
BASE        = Path(__file__).parent
SCRIPT      = BASE / "trailing_4weeks_sideline.py"
CSV_FOLDER  = BASE / "FCLM_CSVs"
COOKIE_FILE = BASE / "fclm_session.txt"

# ── Config (warehouse ID, process ID, labels) ─────────────────────────────
_cfg_file = BASE / "config.json"
_cfg = json.loads(_cfg_file.read_text()) if _cfg_file.exists() else {}
WAREHOUSE_ID = _cfg.get("warehouse_id", "CLE3")
PROCESS_ID   = _cfg.get("process_id",   "1002980")
PATH_NAME    = _cfg.get("path_name",    "IB Problem Solve")
REPORT_LABEL = _cfg.get("report_label", "Sideline & Damageland")
OUTPUT       = BASE / _cfg.get("report_filename", "PS_Tracker_Trailing4Weeks.xlsx")

# ── Find uv ────────────────────────────────────────────────────────────────
def _find_uv():
    for candidate in [
        Path.home() / ".aki" / "bin" / "uv.exe",
        Path.home() / ".cargo" / "bin" / "uv.exe",
        Path("C:/Program Files/uv/uv.exe"),
    ]:
        if candidate.exists():
            return candidate
    return Path("uv")  # hope it's on PATH

UVX = _find_uv()

FCLM_BASE = (
    "https://fclm-portal.amazon.com/reports/functionRollup"
    f"?reportFormat=CSV&warehouseId={WAREHOUSE_ID}&processId={PROCESS_ID}&spanType=Week"
    "&maxIntradayDays=1&startHourIntraday=0&startMinuteIntraday=0"
    "&endHourIntraday=0&endMinuteIntraday=0"
    "&startHourIntraday1=7&startMinuteIntraday1=0"
    "&startHourIntraday2=17&startMinuteIntraday2=2"
    "&startHourIntraday3=18&startMinuteIntraday3=0"
    "&startHourIntraday4=4&startMinuteIntraday4=2"
)

def csv_url(dt):
    wk  = dt.strftime("%Y/%m/%d").replace("/", "%2F")
    end = (dt + timedelta(days=6)).strftime("%Y/%m/%d").replace("/", "%2F")
    return f"{FCLM_BASE}&startDateWeek={wk}&startDateDay={end}"

def csv_filename(dt):
    utc_start = dt + timedelta(hours=4)
    utc_end   = utc_start + timedelta(days=7)
    return (f"functionRollupReport-{WAREHOUSE_ID}-{PATH_NAME}-Week-"
            f"{utc_start.strftime('%Y%m%d%H%M%S')}-{utc_end.strftime('%Y%m%d%H%M%S')}.csv")

def last_sunday():
    t = datetime.now()
    days_back = (t.weekday() + 1) % 7
    return (t - timedelta(days=days_back)).strftime("%Y/%m/%d")

def parse_cookie_string(cookie_str):
    """Parse 'key=val; key2=val2' or a bare value into a cookie dict."""
    if "=" not in cookie_str:
        return {"session-id": cookie_str}  # bare session-id value
    jar = {}
    for part in cookie_str.split(";"):
        part = part.strip()
        if "=" in part:
            name, _, value = part.partition("=")
            jar[name.strip()] = value.strip()
    return jar


def get_midway_cookies():
    """Load cookies from ~/.midway/cookie — created by mwinit.
    Reads ALL lines from index 4+ including #HttpOnly_ lines, same as internal FCLM tools."""
    import time
    midway_file = Path.home() / ".midway" / "cookie"
    if not midway_file.exists():
        return {}
    lines = midway_file.read_text().splitlines()
    cookies = {}
    now = time.time()
    for line in lines[4:]:
        parts = line.split("\t")
        if len(parts) >= 7:
            try:
                expiry = int(parts[4])
            except ValueError:
                continue
            if expiry > now:  # skip expired
                cookies[parts[5]] = parts[6]
    return cookies


def get_cookies_via_cdp(log_fn=None):
    """Pull FCLM cookies from a running Chrome via DevTools Protocol.
    Requires Chrome started with --remote-debugging-port=9222."""
    import json

    def log(msg):
        if log_fn: log_fn(msg, "gray")

    try:
        r = requests.get("http://localhost:9222/json/version", timeout=2)
        if r.status_code != 200:
            return {}
        log("       Connected to Chrome remote debugging port.")
    except Exception:
        return {}  # CDP not available — silent fallback

    try:
        targets = requests.get("http://localhost:9222/json", timeout=2).json()
    except Exception as e:
        log(f"       CDP targets error: {e}")
        return {}

    if not targets:
        return {}

    try:
        import websocket  # websocket-client
        ws_url = targets[0].get("webSocketDebuggerUrl", "")
        if not ws_url:
            return {}
        ws = websocket.create_connection(ws_url, timeout=5)
        ws.send(json.dumps({"id": 1, "method": "Network.getAllCookies"}))
        data = json.loads(ws.recv())
        ws.close()
    except Exception as e:
        log(f"       CDP websocket error: {e}")
        return {}

    jar = {}
    for ck in data.get("result", {}).get("cookies", []):
        if "amazon.com" in ck.get("domain", ""):
            jar[ck["name"]] = ck["value"]

    if jar:
        log(f"       Loaded {len(jar)} cookies via Chrome DevTools.")
    return jar


def get_browser_cookies(log_fn=None, prefer="chrome"):
    """Read FCLM session cookies from Chrome, Firefox, or Edge on Windows.
    prefer='chrome'  → Chrome → Firefox → Edge (default)
    prefer='firefox' → Firefox → Chrome → Edge
    """
    import sqlite3, shutil, tempfile
    import json

    def log(msg):
        if log_fn: log_fn(msg, "gray")

    # DPAPI decryption for Chrome/Edge cookie values
    def decrypt_value(encrypted_value, local_state_path):
        try:
            import base64
            with open(local_state_path, "r", encoding="utf-8") as f:
                local_state = json.load(f)
            enc_key = base64.b64decode(
                local_state["os_crypt"]["encrypted_key"]
            )[5:]  # strip DPAPI prefix
            # Use Windows DPAPI to decrypt the AES key
            import ctypes
            class DATA_BLOB(ctypes.Structure):
                _fields_ = [("cbData", ctypes.c_ulong), ("pbData", ctypes.POINTER(ctypes.c_char))]
            p = ctypes.create_string_buffer(enc_key, len(enc_key))
            blobin = DATA_BLOB(ctypes.sizeof(p), p)
            blobout = DATA_BLOB()
            retval = ctypes.windll.crypt32.CryptUnprotectData(
                ctypes.byref(blobin), None, None, None, None, 0, ctypes.byref(blobout)
            )
            if retval:
                key = ctypes.string_at(blobout.pbData, blobout.cbData)
                ctypes.windll.kernel32.LocalFree(blobout.pbData)
            else:
                return None
            # Decrypt cookie value with AES-256-GCM
            from Crypto.Cipher import AES
            nonce = encrypted_value[3:15]
            ciphertext = encrypted_value[15:-16]
            tag = encrypted_value[-16:]
            cipher = AES.new(key, AES.MODE_GCM, nonce=nonce)
            return cipher.decrypt_and_verify(ciphertext, tag).decode("utf-8")
        except Exception:
            return None

    def decrypt_dpapi(encrypted_value):
        """Fallback: plain DPAPI (older Chrome versions)."""
        try:
            import ctypes
            class DATA_BLOB(ctypes.Structure):
                _fields_ = [("cbData", ctypes.c_ulong), ("pbData", ctypes.POINTER(ctypes.c_char))]
            p = ctypes.create_string_buffer(encrypted_value, len(encrypted_value))
            blobin = DATA_BLOB(ctypes.sizeof(p), p)
            blobout = DATA_BLOB()
            retval = ctypes.windll.crypt32.CryptUnprotectData(
                ctypes.byref(blobin), None, None, None, None, 0, ctypes.byref(blobout)
            )
            if retval:
                val = ctypes.string_at(blobout.pbData, blobout.cbData)
                ctypes.windll.kernel32.LocalFree(blobout.pbData)
                return val.decode("utf-8", errors="replace")
        except Exception:
            pass
        return None

    localappdata = Path(os.environ.get("LOCALAPPDATA", ""))
    roamingappdata = Path(os.environ.get("APPDATA", ""))

    chromium_profiles = [
        # (browser_name, cookie_db_path, local_state_path)
        ("Chrome", localappdata / "Google" / "Chrome" / "User Data" / "Default" / "Network" / "Cookies",
                   localappdata / "Google" / "Chrome" / "User Data" / "Local State"),
        ("Chrome", localappdata / "Google" / "Chrome" / "User Data" / "Default" / "Cookies",
                   localappdata / "Google" / "Chrome" / "User Data" / "Local State"),
        ("Edge",   localappdata / "Microsoft" / "Edge" / "User Data" / "Default" / "Network" / "Cookies",
                   localappdata / "Microsoft" / "Edge" / "User Data" / "Local State"),
        ("Edge",   localappdata / "Microsoft" / "Edge" / "User Data" / "Default" / "Cookies",
                   localappdata / "Microsoft" / "Edge" / "User Data" / "Local State"),
    ]

    def _read_chromium(browser, cookie_db, local_state):
        tmp = Path(tempfile.mktemp(suffix=".db"))
        try:
            try:
                import ctypes, ctypes.wintypes
                GENERIC_READ=0x80000000; OPEN_EXISTING=3; FILE_ATTRIBUTE_NORMAL=0x80; FILE_SHARE=0x7
                k32=ctypes.windll.kernel32
                h=k32.CreateFileW(str(cookie_db),GENERIC_READ,FILE_SHARE,None,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,None)
                INVALID=-1&0xFFFFFFFFFFFFFFFF
                if h==INVALID: raise OSError("Cannot open")
                sh=ctypes.c_ulong(0); sz=k32.GetFileSize(h,ctypes.byref(sh))
                total=sz|(sh.value<<32); buf=ctypes.create_string_buffer(total); rd=ctypes.c_ulong(0)
                k32.ReadFile(h,buf,total,ctypes.byref(rd),None); k32.CloseHandle(h)
                tmp.write_bytes(buf.raw[:rd.value])
            except Exception:
                shutil.copy2(str(cookie_db), str(tmp))
            conn = sqlite3.connect(str(tmp))
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT name, encrypted_value, value FROM cookies "
                "WHERE host_key LIKE '%fclm-portal.amazon.com%' "
                "   OR host_key LIKE '%amazon.com%'"
            ).fetchall()
            conn.close()
        except Exception as e:
            log(f"       Could not read {browser} DB: {e}")
            return {}
        finally:
            try: tmp.unlink()
            except: pass
        if not rows:
            log(f"       No Amazon cookies found in {browser}.")
            return {}
        jar = {}
        for row in rows:
            name = row["name"]; raw = row["encrypted_value"]; val = row["value"]
            if raw:
                decrypted = decrypt_value(raw, local_state)
                if not decrypted:
                    decrypted = decrypt_dpapi(raw)
                if decrypted:
                    val = decrypted
            if val:
                jar[name] = val
        return jar

    def _try_chromium(profiles):
        for browser, cookie_db, local_state in profiles:
            if not cookie_db.exists(): continue
            log(f"       Trying {browser} cookies...")
            jar = _read_chromium(browser, cookie_db, local_state)
            if jar:
                log(f"       Loaded {len(jar)} Amazon cookies from {browser}.")
                return jar
        return {}

    def _try_firefox():
        ff_profiles_dir = roamingappdata / "Mozilla" / "Firefox" / "Profiles"
        if not ff_profiles_dir.exists():
            return {}
        ff_profile = None
        for p in sorted(ff_profiles_dir.iterdir()):
            if p.is_dir() and "default-release" in p.name:
                ff_profile = p; break
        if not ff_profile:
            for p in sorted(ff_profiles_dir.iterdir()):
                if p.is_dir() and "default" in p.name:
                    ff_profile = p; break
        if not ff_profile:
            return {}
        cookie_db = ff_profile / "cookies.sqlite"
        if not cookie_db.exists():
            return {}
        log("       Trying Firefox cookies...")
        tmp = Path(tempfile.mktemp(suffix=".db"))
        rows = []
        try:
            try:
                import ctypes, ctypes.wintypes
                GENERIC_READ=0x80000000; OPEN_EXISTING=3; FILE_ATTRIBUTE_NORMAL=0x80; FILE_SHARE=0x7
                k32=ctypes.windll.kernel32
                h=k32.CreateFileW(str(cookie_db),GENERIC_READ,FILE_SHARE,None,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,None)
                INVALID=-1&0xFFFFFFFFFFFFFFFF
                if h==INVALID: raise OSError("Cannot open")
                sh=ctypes.c_ulong(0); sz=k32.GetFileSize(h,ctypes.byref(sh))
                total=sz|(sh.value<<32); buf=ctypes.create_string_buffer(total); rd=ctypes.c_ulong(0)
                k32.ReadFile(h,buf,total,ctypes.byref(rd),None); k32.CloseHandle(h)
                tmp.write_bytes(buf.raw[:rd.value])
            except Exception:
                shutil.copy2(str(cookie_db), str(tmp))
            conn = sqlite3.connect(str(tmp))
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT name, value FROM moz_cookies "
                "WHERE host LIKE '%fclm-portal.amazon.com%' "
                "   OR host LIKE '%amazon.com%'"
            ).fetchall()
            conn.close()
        except Exception as e:
            log(f"       Could not read Firefox DB: {e}"); rows = []
        finally:
            try: tmp.unlink()
            except: pass
        if rows:
            jar = {row["name"]: row["value"] for row in rows if row["value"]}
            if jar:
                log(f"       Loaded {len(jar)} Amazon cookies from Firefox.")
                return jar
        log("       No Amazon cookies found in Firefox.")
        return {}

    # Determine try order based on user preference
    if prefer == "firefox":
        order = [_try_firefox,
                 lambda: _try_chromium(chromium_profiles[:2]),
                 lambda: _try_chromium(chromium_profiles[2:])]
    else:
        order = [lambda: _try_chromium(chromium_profiles[:2]),
                 _try_firefox,
                 lambda: _try_chromium(chromium_profiles[2:])]

    for fn in order:
        result = fn()
        if result:
            return result
    return {}

# ── UI ─────────────────────────────────────────────────────────────────────
class App:
    NAVY   = "#1F3864"
    DARK   = "#152642"
    PANEL  = "#243B55"
    PURPLE = "#7030A0"
    PURPLH = "#9B59B6"
    WHITE  = "#FFFFFF"
    GRAY   = "#AAB7D4"
    GREEN  = "#70AD47"
    GOLD   = "#FFD966"
    RED    = "#FF6B6B"

    def __init__(self, root):
        self.root = root
        root.title(f"PS Tracker — {WAREHOUSE_ID}")
        root.configure(bg=self.NAVY)
        root.resizable(False, False)
        root.geometry("540x700")

        # Extension cookie server (starts immediately; extension sends cookies in background)
        self._ext_cookies = {}
        self._ext_lock = threading.Lock()
        threading.Thread(target=self._run_ext_server, daemon=True).start()

        tk.Label(root, text="IB Problem Solve Tracker",
                 bg=self.NAVY, fg=self.WHITE,
                 font=("Calibri", 20, "bold")).pack(pady=(24, 2))
        tk.Label(root, text=f"{WAREHOUSE_ID}  ·  {REPORT_LABEL}  ·  Trailing 4 Weeks",
                 bg=self.NAVY, fg=self.GRAY,
                 font=("Calibri", 10)).pack(pady=(0, 4))
        tk.Label(root, text="v1.0  ·  Created by raiedwar",
                 bg=self.NAVY, fg=self.GRAY,
                 font=("Calibri", 8)).pack(pady=(0, 12))

        # Date row
        panel = tk.Frame(root, bg=self.PANEL, pady=12)
        panel.pack(fill="x", padx=30)
        tk.Label(panel, text="Week Start Date (Sunday):",
                 bg=self.PANEL, fg=self.WHITE,
                 font=("Calibri", 11)).pack(side="left", padx=(16, 8))
        self.date_var = tk.StringVar(value=last_sunday())
        tk.Entry(panel, textvariable=self.date_var, font=("Calibri", 13),
                 width=13, justify="center",
                 bg=self.DARK, fg=self.WHITE,
                 insertbackground=self.WHITE, relief="flat",
                 bd=4).pack(side="left")
        tk.Label(panel, text="YYYY/MM/DD",
                 bg=self.PANEL, fg=self.GRAY,
                 font=("Calibri", 9)).pack(side="left", padx=8)


        # Browser cookie source selector
        brow_frame = tk.Frame(root, bg=self.NAVY)
        brow_frame.pack(pady=(6, 0))
        tk.Label(brow_frame, text="Cookie source:",
                 bg=self.NAVY, fg=self.GRAY,
                 font=("Calibri", 9)).pack(side="left", padx=(0, 10))
        self.browser_var = tk.StringVar(value="chrome")
        for _lbl, _val in [("Chrome", "chrome"), ("Firefox", "firefox")]:
            tk.Radiobutton(
                brow_frame, text=_lbl,
                variable=self.browser_var, value=_val,
                bg=self.NAVY, fg=self.GRAY,
                selectcolor=self.DARK,
                activebackground=self.NAVY,
                activeforeground=self.WHITE,
                font=("Calibri", 9)
            ).pack(side="left", padx=(0, 14))

        # Session status bar
        sbar = tk.Frame(root, bg=self.NAVY)
        sbar.pack(fill="x", padx=30, pady=(6, 0))
        self.session_lbl = tk.Label(sbar, text="Checking session...",
                                    bg=self.NAVY, fg=self.GRAY,
                                    font=("Calibri", 9))
        self.session_lbl.pack(side="left")
        self.refresh_btn = tk.Button(sbar, text="\U0001f511  Refresh Session",
                                     font=("Calibri", 9),
                                     bg=self.PANEL, fg=self.GRAY,
                                     activebackground=self.DARK,
                                     activeforeground=self.WHITE,
                                     relief="flat", bd=0,
                                     padx=10, pady=4,
                                     cursor="hand2",
                                     command=self._refresh_session)
        self.refresh_btn.pack(side="right")

        self.btn = tk.Button(root, text="\u25b6   RUN",
                             font=("Calibri", 22, "bold"),
                             bg=self.PURPLE, fg=self.WHITE,
                             activebackground=self.PURPLH,
                             activeforeground=self.WHITE,
                             relief="flat", bd=0,
                             padx=50, pady=18,
                             cursor="hand2",
                             command=self.start)
        self.btn.pack(pady=(16, 6))

        self.test_btn = tk.Button(root, text="Test  (use existing CSVs)",
                                  font=("Calibri", 10),
                                  bg=self.PANEL, fg=self.GRAY,
                                  activebackground=self.DARK,
                                  activeforeground=self.WHITE,
                                  relief="flat", bd=0,
                                  padx=10, pady=6,
                                  cursor="hand2",
                                  command=self.start_test)
        self.test_btn.pack(pady=(0, 10))

        self.log = scrolledtext.ScrolledText(
            root, height=12, width=62,
            bg=self.DARK, fg="#D0D0D0",
            font=("Consolas", 9),
            state="disabled", bd=0,
            relief="flat")
        self.log.pack(padx=20, pady=(0, 16))

        self.log.tag_configure("green",  foreground=self.GREEN)
        self.log.tag_configure("gold",   foreground=self.GOLD)
        self.log.tag_configure("red",    foreground=self.RED)
        self.log.tag_configure("gray",   foreground=self.GRAY)
        self.log.tag_configure("white",  foreground=self.WHITE)

        self.write("Ready. Enter the week start date and click RUN.\n", "gray")
        root.after(1500, self._update_session_status)

    def _update_session_status(self):
        """Check session validity every 10 s and update the status bar."""
        import time
        # Extension cookies take a few seconds to arrive — check first
        with self._ext_lock:
            has_ext = bool(self._ext_cookies)
        if has_ext:
            self.session_lbl.config(
                text="\u25cf  Extension connected — session OK", fg=self.GREEN)
            self.root.after(10000, self._update_session_status)
            return

        # Check Midway cookie for real session tokens (user_name alone is not enough)
        mw = get_midway_cookies()
        session_keys = {k: v for k, v in mw.items() if k != "user_name"}
        if session_keys:
            self.session_lbl.config(
                text="\u25cf  Midway session OK", fg=self.GREEN)
        else:
            self.session_lbl.config(
                text="\u26a0  Session expired — click Refresh Session", fg=self.GOLD)
        self.root.after(10000, self._update_session_status)

    def _refresh_session(self):
        """Open a terminal running mwinit so the user can refresh their Amazon session."""
        self.write("Opening session refresh — sign in with your badge in the browser.", "gold")
        self.write("When the terminal says \'Done\', close it and click RUN.", "gray")
        try:
            msg = "echo Refreshing Amazon session... && mwinit && echo. && echo Done. Close this window and click RUN in PS Tracker."
            subprocess.Popen(f'start cmd /k "{msg}"', shell=True)
        except Exception as e:
            self.write(f"Could not open terminal: {e}", "red")
            self.write("Open Command Prompt and run:  mwinit", "gold")

    def _run_ext_server(self):
        """Tiny HTTP server on 127.0.0.1:9876 — Chrome extension POSTs Amazon cookies here."""
        import json
        from http.server import HTTPServer, BaseHTTPRequestHandler
        app = self

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                if self.path == "/ping":
                    self.send_response(200)
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(b"ok")
                else:
                    self.send_response(404); self.end_headers()

            def do_OPTIONS(self):
                self.send_response(204)
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                self.send_header("Access-Control-Allow-Headers", "Content-Type")
                self.end_headers()

            def do_POST(self):
                if self.path == "/cookies":
                    try:
                        length = int(self.headers.get("Content-Length", 0))
                        body = self.rfile.read(length)
                        with app._ext_lock:
                            for ck in json.loads(body):
                                app._ext_cookies[ck["name"]] = ck["value"]
                    except Exception:
                        pass
                    self.send_response(200)
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(b"ok")
                else:
                    self.send_response(404); self.end_headers()

            def log_message(self, *args): pass

        try:
            server = HTTPServer(("127.0.0.1", 9876), Handler)
            server.serve_forever()
        except OSError:
            pass  # port in use — extension path unavailable, other methods still work

    def write(self, msg, tag="white"):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def start(self):
        self.btn.configure(state="disabled", text="  Running...  ")
        self.test_btn.configure(state="disabled")
        self.log.configure(state="normal"); self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        threading.Thread(target=self._worker, daemon=True).start()

    def start_test(self):
        self.btn.configure(state="disabled")
        self.test_btn.configure(state="disabled", text="Running...")
        self.log.configure(state="normal"); self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        threading.Thread(target=self._test_worker, daemon=True).start()

    def _worker(self):
        try:    self._run()
        except Exception as e:
            self.write(f"\nUnexpected error: {e}", "red")
        finally:
            self.btn.configure(state="normal", text="▶   RUN")
            self.test_btn.configure(state="normal", text="Test  (use existing CSVs)")

    def _test_worker(self):
        try:
            self.write("Test mode — skipping download, using existing CSVs.", "gold")
            self.write("─" * 48, "gray")
            self.write("\nGenerating Excel report...", "gold")
            result = subprocess.run(
                [str(UVX), "run", "--with", "openpyxl", str(SCRIPT), str(BASE)],
                capture_output=True, text=True
            )
            if result.returncode == 0:
                for line in result.stdout.splitlines():
                    self.write(f"  {line}", "gray")
                self.write("\n  Done! Opening report...", "green")
                os.startfile(str(OUTPUT))
            else:
                self.write(result.stderr[:800], "red")
        except Exception as e:
            self.write(f"Error: {e}", "red")
        finally:
            self.btn.configure(state="normal", text="▶   RUN")
            self.test_btn.configure(state="normal", text="Test  (use existing CSVs)")

    def _run(self):
        try:
            dt = datetime.strptime(self.date_var.get().strip(), "%Y/%m/%d")
        except ValueError:
            self.write("Bad date format — use YYYY/MM/DD", "red"); return

        sundays = [dt - timedelta(weeks=i) for i in range(4, -1, -1)]
        self.write(f"Trailing 4 weeks ending {dt.strftime('%B %d, %Y')}", "white")
        self.write("─" * 48, "gray")

        import urllib3; urllib3.disable_warnings()
        session = requests.Session()
        session.verify = False

        self.write("\n[1/2]  Downloading CSVs", "gold")

        # Read cookie from file directly — avoids Tkinter/Tcl truncating on semicolons
        manual_cookie = COOKIE_FILE.read_text().strip() if COOKIE_FILE.exists() else ""
        auth = None

        # Tier 0: Chrome extension (server has been running since app start)
        with self._ext_lock:
            ext_cookies = dict(self._ext_cookies)
        if ext_cookies:
            for name, value in ext_cookies.items():
                session.cookies.set(name, value, domain="fclm-portal.amazon.com")
                session.cookies.set(name, value, domain=".amazon.com")
            self.write(f"       Got {len(ext_cookies)} cookies from Chrome extension.", "green")
            # Save fresh session so V2/other tools can use it without the extension
            _fkeys = ["session-id","session-id-time","session-token","ubid-main","program-cookie","JSESSIONID"]
            _fset = {k: ext_cookies[k] for k in _fkeys if k in ext_cookies}
            if "session-id" in _fset and "session-token" in _fset:
                COOKIE_FILE.write_text("; ".join(f"{k}={v}" for k,v in _fset.items()))

        # Tier 1: Midway cookie + Kerberos — same pattern as internal FCLM tools
        if not len(session.cookies) and not session.headers.get("Cookie"):
            mw = get_midway_cookies()
            if mw:
                session.cookies.update(mw)
                try:
                    from requests_kerberos import HTTPKerberosAuth, OPTIONAL as KRB_OPTIONAL
                    auth = HTTPKerberosAuth(mutual_authentication=KRB_OPTIONAL)
                    self.write(f"       Midway cookie ({len(mw)} values) + Kerberos auth.", "green")
                except ImportError:
                    self.write(f"       Midway cookie ({len(mw)} values) — Kerberos unavailable.", "gold")
            else:
                self.write("       Midway cookie not found or expired — run mwinit.", "gold")

        # Tier 2: browser cookie file (works when Chrome is fully closed)
        if not len(session.cookies) and not session.headers.get("Cookie"):
            self.write("       Trying browser cookie file...", "gray")
            try:
                browser_cookies = get_browser_cookies(log_fn=self.write, prefer=self.browser_var.get())
                for name, value in browser_cookies.items():
                    session.cookies.set(name, value, domain="fclm-portal.amazon.com")
                    session.cookies.set(name, value, domain=".amazon.com")
            except Exception as e:
                self.write(f"       Cookie read error: {e}", "gold")

        # Tier 2: manual saved cookie — send as raw Cookie header (bypasses domain matching)
        if not len(session.cookies) and manual_cookie:
            parsed = parse_cookie_string(manual_cookie)
            session.headers.update({"Cookie": manual_cookie})
            self.write(f"       Using saved session cookies ({len(parsed)} values).", "gold")

        # Tier 3: Windows SSPI / Kerberos (last resort)
        if not len(session.cookies) and not session.headers.get('Cookie'):
            try:
                from requests_negotiate_sspi import HttpNegotiateAuth
                auth = HttpNegotiateAuth()
                self.write("       Trying Windows auth (mwinit)...", "gold")
            except ImportError:
                pass

        if not len(session.cookies) and auth is None:
            self.write("       No session found.", "red")
            self.write("       Click  \U0001f511 Refresh Session  to sign in, then click RUN.", "gold")
            return

        cookie_hdr = session.headers.get("Cookie", "")
        if auth:
            auth_label = "SSPI/Kerberos"
        elif cookie_hdr:
            auth_label = f"{len(parse_cookie_string(cookie_hdr))} cookies (header)"
        else:
            auth_label = f"{len(session.cookies)} cookies (jar)"
        self.write(f"       Auth: {auth_label}", "gray")

        # Clear any leftover week CSVs before downloading — prevents stale files skewing the report
        for old_csv in CSV_FOLDER.glob(f"functionRollupReport-{WAREHOUSE_ID}-{PATH_NAME}-Week-*.csv"):
            try: old_csv.unlink()
            except: pass

        downloaded = []
        for wk in sundays:
            fname = csv_filename(wk)
            dest  = CSV_FOLDER / fname
            wk_lbl = wk.strftime("%b %d")
            self.write(f"       Wk {wk_lbl}  —  downloading...", "gray")
            try:
                resp = session.get(csv_url(wk), timeout=30, allow_redirects=True, auth=auth)
                data = resp.content
                is_html = data[:200].lstrip().startswith(b"<") or b"<html" in data[:500].lower()
                if resp.status_code == 200 and len(data) > 1000 and not is_html:
                    CSV_FOLDER.mkdir(parents=True, exist_ok=True)
                    dest.write_bytes(data)
                    downloaded.append(dest)
                    self.write(f"       Wk {wk_lbl}  —  saved ({len(data):,} bytes)", "green")
                elif is_html or len(data) <= 1000:
                    self.write(f"       Wk {wk_lbl}  —  got HTML, session not valid.", "red")
                    self.write("       Session expired — click  🔑 Refresh Session  then click RUN again.", "gold")
                    return
                else:
                    self.write(f"       Wk {wk_lbl}  —  HTTP {resp.status_code}.", "red")
                    return
            except requests.RequestException as e:
                self.write(f"       Wk {wk_lbl}  —  error: {e}", "red"); return

        self.write("\n[2/2]  Generating Excel report", "gold")
        result = subprocess.run(
            [str(UVX), "run", "--with", "openpyxl", str(SCRIPT), str(BASE)],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                self.write(f"       {line}", "gray")
            # Delete all CSVs before opening Excel — report is saved, no longer needed
            _deleted = []
            for _f in CSV_FOLDER.glob(f"functionRollupReport-{WAREHOUSE_ID}-{PATH_NAME}-Week-*.csv"):
                try: _f.unlink(); _deleted.append(_f)
                except: pass
            self.write(f"       Cleaned up {len(_deleted)} CSV files.", "gray")
            self.write("\n  Done! Opening report...", "green")
            os.startfile(str(OUTPUT))
        else:
            err = result.stderr
            if "Excel has the report open" in err or "File locked" in err:
                self.write("\n  Excel has the report open.", "red")
                self.write("  Close Excel completely, then click RUN again.", "gold")
            else:
                self.write("Script error:", "red")
                self.write(err[:1200], "red")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
