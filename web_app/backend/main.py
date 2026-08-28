import os
import sys
import asyncio
import io
import pandas as pd
from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Dict, Any, Optional

import threading
import builtins

# Tambahkan path root agar bisa import bot.py
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(ROOT_DIR)

import bot  # Import skrip bot.py asli

# ---------------------------------------------------------
# BOT WAITER (UNTUK MENGGANTIKAN INPUT())
# ---------------------------------------------------------
bot_resume_event = threading.Event()
is_bot_waiting = False
stop_bot_event = False

def web_input(prompt=""):
    global is_bot_waiting, stop_bot_event
    print(prompt)
    print("\n>>> 🟡 SISTEM MENUNGGU 🟡 <<<")
    print(">>> Silakan berinteraksi dengan Chrome, lalu klik tombol 'LANJUTKAN' di UI. <<<\n")
    is_bot_waiting = True
    bot_resume_event.clear()
    bot_resume_event.wait()
    is_bot_waiting = False
    
    if stop_bot_event:
        print("\n🛑 BOT DIHENTIKAN OLEH PENGGUNA.")
        raise KeyboardInterrupt("Pengguna menekan tombol Hentikan.")
        
    return ""

app = FastAPI(title="SIPGN Bot API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------
# WEBSOCKET & LOGGING
# ---------------------------------------------------------
connected_clients: List[WebSocket] = []

@app.websocket("/ws/logs")
async def websocket_endpoint(websocket: WebSocket):
    await websocket.accept()
    connected_clients.append(websocket)
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        connected_clients.remove(websocket)

class WsStdout(io.StringIO):
    def __init__(self, loop):
        super().__init__()
        self.loop = loop
    
    def write(self, s):
        super().write(s)
        # Hapus newline di akhir jika ada, karena front-end akan menampilkannya sebagai elemen terpisah
        msg = s.strip()
        if msg:
            # Panggil coroutine secara thread-safe
            asyncio.run_coroutine_threadsafe(self.broadcast(msg), self.loop)

    async def broadcast(self, message: str):
        for client in connected_clients:
            try:
                await client.send_text(message)
            except Exception:
                pass

# ---------------------------------------------------------
# DATA API
# ---------------------------------------------------------
@app.get("/api/data")
def get_excel_data(file: str = "kloter1.xlsx"):
    file_path = os.path.join(ROOT_DIR, file)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File tidak ditemukan")
    
    try:
        # Hanya baca sheet 'DISTRIBUSI'
        df = pd.read_excel(file_path, sheet_name="DISTRIBUSI", dtype=str)
        # Bersihkan NaN
        df = df.fillna("")
        
        # Format tanggal_wo ke dd/mm/yyyy
        if 'tanggal_wo' in df.columns:
            def format_date(x):
                if pd.isna(x) or str(x).strip() == "": return ""
                try:
                    if str(x).isdigit():
                        return pd.to_datetime(int(x), unit='D', origin='1899-12-30').strftime('%d/%m/%Y')
                    return pd.to_datetime(x).strftime('%d/%m/%Y')
                except:
                    return str(x)
            df['tanggal_wo'] = df['tanggal_wo'].apply(format_date)
            
        # Pastikan kolom _excel_row ada untuk melacak baris asli
        if "_excel_row" not in df.columns:
            # Baris di Excel mulai dari 2 (1 untuk header)
            df.insert(0, "_excel_row", df.index + 2)
            
        return {
            "headers": df.columns.tolist(),
            "rows": df.to_dict(orient="records")
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

class SaveDataRequest(BaseModel):
    rows: List[Dict[str, Any]]

@app.post("/api/data")
def save_excel_data(request: SaveDataRequest, file: str = "kloter1.xlsx"):
    file_path = os.path.join(ROOT_DIR, file)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
    try:
        df = pd.DataFrame(request.rows)
        # Jika '_excel_row' ada, kita hapus sebelum save karena ini cuma indeks virtual
        if "_excel_row" in df.columns:
            df = df.drop(columns=["_excel_row"])
            
        # Kembalikan format tanggal_wo dari dd/mm/yyyy menjadi integer (format Excel)
        if 'tanggal_wo' in df.columns:
            def parse_date_to_excel(x):
                if pd.isna(x) or str(x).strip() == "": return ""
                try:
                    dt = pd.to_datetime(str(x), format='%d/%m/%Y', errors='coerce')
                    if pd.notna(dt):
                        return (dt - pd.Timestamp('1899-12-30')).days
                except:
                    pass
                return x
            df['tanggal_wo'] = df['tanggal_wo'].apply(parse_date_to_excel)
            
        # Gunakan mode append (a) agar sheet lain (seperti MASTER, dll) tidak terhapus
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="DISTRIBUSI", index=False)
            
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# Simpan event loop utama agar bisa diakses thread lain
@app.on_event("startup")
def startup_event():
    app.state.loop = asyncio.get_running_loop()

# ---------------------------------------------------------
# BOT RUNNER API
# ---------------------------------------------------------
class RunRequest(BaseModel):
    mode: str
    excel_file: str
    row: Optional[int] = None

@app.post("/api/resume")
def resume_bot_endpoint():
    bot_resume_event.set()
    return {"status": "success"}

@app.post("/api/stop")
def stop_bot_endpoint():
    global stop_bot_event
    stop_bot_event = True
    bot_resume_event.set()  # Bangunkan jika sedang tidur
    return {"status": "success"}

@app.get("/api/status")
def status_endpoint():
    return {"waiting": is_bot_waiting}

import shutil

@app.post("/api/clear-cache")
def clear_cache_endpoint():
    profile_dir = os.path.join(ROOT_DIR, "chrome-profile", "Default")
    cache_dirs = [
        os.path.join(profile_dir, "Cache"),
        os.path.join(profile_dir, "Code Cache"),
        os.path.join(profile_dir, "GPUCache"),
        os.path.join(profile_dir, "Service Worker", "CacheStorage")
    ]
    
    cleared = []
    for d in cache_dirs:
        if os.path.exists(d):
            try:
                shutil.rmtree(d, ignore_errors=True)
                cleared.append(os.path.basename(d))
            except Exception as e:
                pass
                
    return {"status": "success", "cleared": cleared}

@app.post("/api/run")
def run_bot(request: RunRequest):
    global stop_bot_event
    stop_bot_event = False
    
    # Menyimpan stdout asli
    original_stdout = sys.stdout
    
    # Dapatkan loop utama dari state
    loop = app.state.loop

    if loop:
        ws_stdout = WsStdout(loop)
        sys.stdout = ws_stdout
    
    # Path absolut excel
    excel_path = os.path.join(ROOT_DIR, request.excel_file)
    
    print(f"\n==============================")
    print(f"Menjalankan mode: {request.mode.upper()}")
    if request.row:
        print(f"Target baris: {request.row}")
    print(f"==============================")
    
    try:
        original_input = builtins.input
        builtins.input = web_input
        bot.input = web_input

        # Panggil fungsi bot sesuai mode (Synchronous call yang memblokir backend)
        # Karena ini MVP lokal, kita jalankan langsung. Idealnya di background task/thread.
        if request.mode == "login":
            bot.login_mode()
        elif request.mode == "berangkat":
            bot.run_berangkat_mode(excel_path, request.row)
        elif request.mode == "ambil":
            bot.run_ambil_mode(excel_path, request.row)
        elif request.mode == "kembali":
            bot.run_kembali_mode(excel_path, request.row)
        elif request.mode == "terima":
            bot.run_terima_mode(excel_path, request.row)
        elif request.mode == "mulaicuci":
            bot.run_mulaicuci_mode(excel_path, request.row)
        elif request.mode == "selesaicuci":
            bot.run_selesaicuci_mode(excel_path, request.row)
        else:
            bot.run_mode(excel_path, request.row)
            
        print("✓ Eksekusi bot selesai.")
    except Exception as e:
        print(f"❌ ERROR: {e}")
    finally:
        # Kembalikan stdout dan input
        builtins.input = original_input
        bot.input = original_input
        sys.stdout = original_stdout
        
    return {"status": "success"}

# Trigger reload

# Trigger reload untuk task scheduler

# Trigger reload untuk pembatalan schtasks
