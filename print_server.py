"""
大阪テック 検査表印刷サーバー（印刷キュー版）
==========================================
統合管理システムの「印刷」は Supabase の印刷キュー（print_jobs テーブル）に登録されます。
このサーバーは数秒おきにキューを見て、届いた依頼を複合機／ラベルプリンターで印刷します。
→ どのPCから印刷しても、このサーバーを動かしている1台で印刷されます。

【使い方】
1. 印刷サーバーにするPCで、このファイルと同じフォルダでコマンドプロンプトを開いて:
     python print_server.py
2. 「印刷キューの監視を開始しました」と表示されたら準備完了
   （このウィンドウは開いたままにしておく。閉じると印刷されません）
【終了】Ctrl+C

【印刷サーバーPCに必要なもの】
- Python 3（requests / pywin32 は初回起動時に自動インストール）
- Microsoft Excel（検査表への書き込みとPDF化に使用）
- 複合機「RICOH IM C3500」のプリンタードライバー（PRINTER_NAME と同じ名前で登録）
- SumatraPDF（推奨）または Adobe Acrobat / Reader
- データシールを使う場合：Brother QL-800 のドライバーと b-PAC、ラベルレイアウト(.lbx)
"""

import json
import os
import re
import sys
import time
import socket
import shutil
import tempfile
import threading
import subprocess
import urllib.parse
from datetime import datetime, timezone, timedelta

VERSION = "queue-1.0"

# ========== 設定 ==========
# Supabase（統合管理システムと同じプロジェクト。公開キー）
SUPABASE_URL = "https://qmyvsuhuoikimwktfufb.supabase.co"
SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFteXZzdWh1b2lraW13a3RmdWZiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAyNTQ3NDcsImV4cCI6MjA5NTgzMDc0N30.6NE3RDRS91gahTH1cP2dgjd0sQ7ku_7e28Jk9vPc7go'

# キューを見に行く間隔（秒）と、稼働通知の間隔（秒）
POLL_SEC = 2
HEARTBEAT_SEC = 5
# これより古い未処理の依頼は印刷せずに「期限切れ」にする（時間）
EXPIRE_HOURS = 12

# 複合機のプリンター名（「デバイスとプリンター」で表示される名前）
PRINTER_NAME = "RICOH IM C3500 JPN RPCS"

HOME = os.environ.get("USERPROFILE", os.path.expanduser("~"))
# 検査表フォルダのローカル同期パス（Supabaseからダウンロードできなかったときの予備）
INSPECTION_FOLDER = os.path.join(HOME, "大阪テック株式会社", "OSAKA TEC - ドキュメント", "共有フォルダー", "検査表")

# ===== データシール（Brother QL-800 / b-PAC）設定 =====
LABEL_PRINTER_NAME = "Brother QL-800"
# P-touch Editorで作成したラベルレイアウト(.lbx)のフルパス
# レイアウト内に下記の名前でテキストオブジェクトを作っておく（差し込み対象）:
#   productno / productname / otnumber / quantity / delivery / client
LABEL_TEMPLATE_PATH = os.path.join(HOME, "osakatec", "label_datasheet.lbx")
# ==========================

SERVER_ID = socket.gethostname()
_stop = threading.Event()


def ensure_packages():
    for pkg, mod in [("requests", "requests"), ("pywin32", "win32com")]:
        try:
            __import__(mod)
        except ImportError:
            print(f"{pkg}をインストールします...")
            subprocess.run([sys.executable, "-m", "pip", "install", pkg], check=True)


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def log(msg):
    print(f"{datetime.now():%H:%M:%S} {msg}", flush=True)


# ---------------- Supabase REST ----------------
def sb_headers(extra=None):
    h = {"apikey": SUPABASE_KEY, "Authorization": "Bearer " + SUPABASE_KEY, "Content-Type": "application/json"}
    if extra:
        h.update(extra)
    return h


def sb_get(path):
    import requests
    r = requests.get(f"{SUPABASE_URL}/rest/v1/{path}", headers=sb_headers(), timeout=15)
    r.raise_for_status()
    return r.json()


def sb_patch(path, body, returning=False):
    import requests
    extra = {"Prefer": "return=representation"} if returning else {"Prefer": "return=minimal"}
    r = requests.patch(f"{SUPABASE_URL}/rest/v1/{path}", headers=sb_headers(extra), data=json.dumps(body), timeout=15)
    r.raise_for_status()
    return r.json() if returning else None


def heartbeat_loop():
    """数秒おきに「このサーバーは動いている」とSupabaseへ書く（画面はこれで印刷サーバーの起動を判定する）"""
    import requests
    body = lambda: json.dumps({"id": SERVER_ID, "last_seen": now_iso(), "printer": PRINTER_NAME, "version": VERSION})
    fails = 0
    while not _stop.is_set():
        try:
            r = requests.post(f"{SUPABASE_URL}/rest/v1/print_server_status",
                              headers=sb_headers({"Prefer": "resolution=merge-duplicates,return=minimal"}),
                              data=body(), timeout=10)
            r.raise_for_status()
            if fails:
                log("[稼働通知] 復旧しました")
            fails = 0
        except Exception as e:
            fails += 1
            if fails in (1, 10) or fails % 60 == 0:
                log(f"[稼働通知] 失敗（{fails}回目）: {e}")
        _stop.wait(HEARTBEAT_SEC)


def claim(job_id):
    """依頼を「印刷中」にして自分が担当する（すでに他で取られていたら None）"""
    rows = sb_patch(f"print_jobs?id=eq.{job_id}&status=eq.pending",
                    {"status": "printing", "server": SERVER_ID, "started_at": now_iso()}, returning=True)
    return rows[0] if rows else None


def finish(job_id, status, message):
    try:
        sb_patch(f"print_jobs?id=eq.{job_id}", {"status": status, "message": (message or "")[:500], "finished_at": now_iso()})
    except Exception as e:
        log(f"[結果の書き込み失敗] id={job_id}: {e}")


# ---------------- 検査表 ----------------
def find_inspection_excel(ot_number):
    """OT番号に対応するExcelファイルをローカル同期フォルダから検索（予備）"""
    if not ot_number or not os.path.isdir(INSPECTION_FOLDER):
        return None
    for ext in (".xlsx", ".xlsm", ".xls"):
        p = os.path.join(INSPECTION_FOLDER, ot_number + ext)
        if os.path.exists(p):
            return p
    return None


def download_inspection(file_url, file_name, tmp_dir):
    """Supabaseに登録した検査表をダウンロード"""
    import requests
    name = file_name or urllib.parse.unquote(file_url.rsplit("/", 1)[-1])
    name = re.sub(r'[\\/:*?"<>|]', "_", name) or "inspection.xlsx"
    dst = os.path.join(tmp_dir, name)
    r = requests.get(file_url, timeout=60)
    r.raise_for_status()
    with open(dst, "wb") as f:
        f.write(r.content)
    return dst


def fill_excel_to_pdf(xlsx_path, pdf_path, p):
    """検査表Excelに書き込んでPDF化（元ファイルは一時コピーなので保存しない）"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        try:
            excel.Visible = False
            excel.DisplayAlerts = False
        except Exception:
            pass
        wb = excel.Workbooks.Open(xlsx_path)
        try:
            ws = wb.Sheets(1)
            drawingno = (p.get("drawingno") or p.get("orderno") or "").strip()
            shanai_no = (p.get("shanai_no") or "").strip()
            sealno = (p.get("sealno") or "").strip()
            ws.Range("F5").Value = drawingno          # 客先番号（客先図番）
            if shanai_no:
                ws.Range("F6").Value = shanai_no      # 社内番号（OT番号）
            ws.Range("F7").Value = sealno             # 製造No.（シールNo）
            # 検査日：'year'=西暦だけ入れる ／ 'none'=何も入れない（月・日は入れない）
            if p.get("date_mode") == "year":
                ws.Range("F8").Value = datetime.now().year
            wb.ExportAsFixedFormat(0, pdf_path)       # 0 = xlTypePDF
        finally:
            wb.Close(False)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def do_inspection(p):
    ot = (p.get("ot_number") or "").strip()
    tmp_dir = tempfile.mkdtemp(prefix="ostec_ins_")
    try:
        src = None
        if p.get("file_url"):
            try:
                src = download_inspection(p["file_url"], p.get("file_name"), tmp_dir)
            except Exception as e:
                log(f"  ダウンロード失敗（ローカルフォルダを探します）: {e}")
        if not src:
            local = find_inspection_excel(ot)
            if local:
                src = os.path.join(tmp_dir, os.path.basename(local))
                shutil.copy2(local, src)
        if not src:
            return "skipped", f"{ot}: 検査表ファイルなし（スキップ）"

        if src.lower().endswith(".pdf"):
            pdf = src                                 # PDFで登録された検査表はそのまま印刷
        else:
            pdf = os.path.join(tmp_dir, "inspection.pdf")
            fill_excel_to_pdf(src, pdf, p)
        ok, msg = print_pdf(pdf)
        return ("done" if ok else "error"), f"{ot}: {msg}"
    finally:
        # 印刷アプリがPDFを読み終わるまで待ってから一時フォルダを消す
        threading.Thread(target=lambda: (time.sleep(60), shutil.rmtree(tmp_dir, ignore_errors=True)), daemon=True).start()


# ---------------- PDF印刷 ----------------
def _close_later(proc, delay=30):
    def _worker():
        time.sleep(delay)
        subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    threading.Thread(target=_worker, daemon=True).start()


def _popen_hidden(cmd):
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 7  # SW_SHOWMINNOACTIVE（最小化・フォーカスを奪わない）
    return subprocess.Popen(cmd, startupinfo=si, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def print_pdf(pdf_path):
    """PDFを複合機へ印刷（できる限り画面を前面に出さない）"""
    username = os.environ.get("USERNAME", "")
    pdf_apps = [
        (r"C:\Program Files\SumatraPDF\SumatraPDF.exe", ["-print-to", PRINTER_NAME, "-silent", pdf_path], True),
        (r"C:\Users\%s\AppData\Local\SumatraPDF\SumatraPDF.exe" % username, ["-print-to", PRINTER_NAME, "-silent", pdf_path], True),
        (r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe", ["/n", "/h", "/t", pdf_path, PRINTER_NAME], False),
        (r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe", ["/n", "/h", "/t", pdf_path, PRINTER_NAME], False),
        (r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe", ["/n", "/h", "/t", pdf_path, PRINTER_NAME], False),
    ]
    for app_path, args, silent in pdf_apps:
        if os.path.exists(app_path):
            name = os.path.basename(app_path)
            if silent:
                # SumatraPDFは印刷が終わるまで待つ（順番どおりに印刷するため）
                subprocess.run([app_path] + args, timeout=120)
            else:
                _close_later(_popen_hidden([app_path] + args), delay=30)
                time.sleep(5)
            return True, f"{name}で印刷しました"

    import ctypes
    ret = ctypes.windll.shell32.ShellExecuteW(None, "print", pdf_path, f'"{PRINTER_NAME}"', None, 2)
    if ret > 32:
        time.sleep(5)
        return True, "ShellExecuteで印刷しました"
    return False, f"印刷アプリが見つかりません (ShellExecute: {ret})"


# ---------------- データシール ----------------
def do_label(p):
    """Brother QL-800 へ b-PAC でデータシールを印刷する"""
    import pythoncom
    import win32com.client
    fields = {k: str(p.get(k, "")).strip() for k in ("productno", "productname", "otnumber", "quantity", "delivery", "client")}
    if not os.path.exists(LABEL_TEMPLATE_PATH):
        return "error", f"ラベルテンプレートが見つかりません: {LABEL_TEMPLATE_PATH}"
    pythoncom.CoInitialize()
    doc = None
    try:
        doc = win32com.client.Dispatch("bpac.Document")
        if not doc.Open(LABEL_TEMPLATE_PATH):
            return "error", "ラベルテンプレートを開けませんでした（.lbxのパスを確認）"
        try:
            doc.SetPrinter(LABEL_PRINTER_NAME, True)
        except Exception as e:
            log(f"  [警告] プリンター指定に失敗（既定のまま続行）: {e}")
        for key, value in fields.items():
            try:
                obj = doc.GetObject(key)
                if obj is not None:
                    obj.Text = value
            except Exception as e:
                log(f"  [警告] オブジェクト '{key}' への差し込み失敗: {e}")
        doc.StartPrint("OsakaTecLabel", 0)
        doc.PrintOut(1, 0)
        doc.EndPrint()
        doc.Close()
        doc = None
        return "done", f"{LABEL_PRINTER_NAME} でシールを発行しました（製番 {fields['productno']}）"
    finally:
        try:
            if doc is not None:
                doc.Close()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# ---------------- キュー処理 ----------------
def process(job):
    jid, kind, p = job["id"], job["kind"], job.get("payload") or {}
    who = job.get("requested_by") or "?"
    if kind == "inspection":
        log(f"[検査表] id={jid} 依頼={who} OT={p.get('ot_number')} 図番={p.get('drawingno')} シール={p.get('sealno')} 検査日={p.get('date_mode')}")
        return do_inspection(p)
    if kind == "label":
        log(f"[シール] id={jid} 依頼={who} 製番={p.get('productno')} OT={p.get('otnumber')}")
        return do_label(p)
    return "error", f"不明な依頼の種類: {kind}"


def recover_and_expire():
    # 前回このPCで「印刷中」のまま止まった依頼は、もう一度印刷待ちに戻す
    try:
        sb_patch(f"print_jobs?status=eq.printing&server=eq.{urllib.parse.quote(SERVER_ID)}", {"status": "pending"})
    except Exception as e:
        log(f"[復旧] 失敗: {e}")
    expire()


def expire():
    limit = (datetime.now(timezone.utc) - timedelta(hours=EXPIRE_HOURS)).isoformat()
    try:
        sb_patch(f"print_jobs?status=eq.pending&created_at=lt.{urllib.parse.quote(limit)}",
                 {"status": "error", "message": f"期限切れ（{EXPIRE_HOURS}時間以上印刷されなかったため取り消し）", "finished_at": now_iso()})
    except Exception as e:
        log(f"[期限切れ処理] 失敗: {e}")


def queue_loop():
    fails, last_expire = 0, time.time()
    while not _stop.is_set():
        try:
            jobs = sb_get("print_jobs?status=eq.pending&order=id.asc&limit=20&select=id,kind,payload,requested_by")
            if fails:
                log("[キュー] 接続が復旧しました")
            fails = 0
        except Exception as e:
            fails += 1
            if fails in (1, 10) or fails % 60 == 0:
                log(f"[キュー] 取得失敗（{fails}回目）: {e}")
            _stop.wait(POLL_SEC * 3)
            continue

        for job in jobs:
            if _stop.is_set():
                break
            try:
                if not claim(job["id"]):
                    continue
            except Exception as e:
                log(f"[キュー] 取得失敗 id={job['id']}: {e}")
                continue
            try:
                status, msg = process(job)
            except Exception as e:
                import traceback
                traceback.print_exc()
                status, msg = "error", str(e)
            log(f"  → {status}: {msg}")
            finish(job["id"], status, msg)

        if time.time() - last_expire > 600:
            expire()
            last_expire = time.time()
        if not jobs:
            _stop.wait(POLL_SEC)


def main():
    ensure_packages()
    print("=" * 56)
    print("  大阪テック 検査表印刷サーバー（印刷キュー版 " + VERSION + "）")
    print("=" * 56)
    print(f"  サーバー名     : {SERVER_ID}")
    print(f"  プリンター     : {PRINTER_NAME}")
    print(f"  ラベル         : {LABEL_PRINTER_NAME}（{'レイアウトあり' if os.path.exists(LABEL_TEMPLATE_PATH) else 'レイアウトなし'}）")
    print(f"  予備の検査表   : {INSPECTION_FOLDER}（{'あり' if os.path.isdir(INSPECTION_FOLDER) else 'なし'}）")
    print("  終了           : Ctrl+C")
    print("=" * 56)

    recover_and_expire()
    threading.Thread(target=heartbeat_loop, daemon=True).start()
    log("印刷キューの監視を開始しました（このウィンドウは開いたままにしてください）")
    try:
        queue_loop()
    except KeyboardInterrupt:
        pass
    finally:
        _stop.set()
        print("\n印刷サーバーを終了しました")


if __name__ == "__main__":
    main()
