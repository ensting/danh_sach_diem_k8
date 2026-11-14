"""Web app tra cứu điểm hàng loạt — giao diện web cho crawl_doc.py.

Chạy:  python app.py   rồi mở http://127.0.0.1:5000
"""
import os
import re
import time
import uuid
import threading
import traceback
from io import BytesIO
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template
from playwright.sync_api import sync_playwright

from crawler_core import (
    doc_danh_sach, tim_frame_co_form,
    tra_cuu_mot_sv, luu_excel, LUU_SAU_MOI,
)

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE, "outputs")
os.makedirs(OUT_DIR, exist_ok=True)

URL_MAC_DINH = "https://script.google.com/macros/s/AKfycbzh0NPH8xQiE-7Wq1_a7Y3ZIoTTAfCPx5mWbbGh1596Uiw1VSFUFPk1y4bH41XfoDIuyw/exec"
HEADLESS = os.environ.get("HEADLESS", "1") != "0"  # HEADLESS=0 để xem trình duyệt khi debug local

# Arg Chromium cho container RAM nhỏ (Render free 512MB)
CHROME_ARGS = ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]

app = Flask(__name__)
JOBS = {}


def log(job, msg):
    job["logs"].append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def chay_job(job, danh_sach):
    job["status"] = "running"
    job["t_bd"] = time.time()
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=job["headless"], args=CHROME_ARGS)
            page = browser.new_page()

            log(job, f"Đang mở trang: {job['url']}")
            page.goto(job["url"], timeout=60000)
            frame = tim_frame_co_form(page)
            if frame is None:
                log(job, "Không tìm thấy form (#mssv) trên trang!")
                job["status"] = "error"
                job["error"] = "Không tìm thấy form tra cứu trên trang"
                browser.close()
                return
            log(job, f"Đã tìm thấy form. Bắt đầu tra cứu {len(danh_sach)} sinh viên...")

            for idx, sv in enumerate(danh_sach, 1):
                if job["stop"]:
                    break
                t0 = time.time()
                try:
                    hang = tra_cuu_mot_sv(frame, sv)
                    log(job, f"[{idx}/{len(danh_sach)}] {sv['mssv']} {sv['ten']}: "
                             f"{hang[3]} / {hang[4]} / {hang[5]}  ({time.time() - t0:.1f}s) "
                             f"{hang[7] if hang[7] != 'OK' else ''}")
                except Exception as e:
                    job["loi"] += 1
                    loi = str(e).splitlines()[0][:80]
                    log(job, f"[{idx}/{len(danh_sach)}] {sv['mssv']} {sv['ten']}: LỖI - {loi}")
                    hang = [sv["mssv"], sv["ten"], sv["lop"], "", "", "", "",
                            "Lỗi: không có kết quả (sai mã tra cứu hoặc server chậm)",
                            datetime.now().strftime("%d/%m/%Y %H:%M:%S")]

                job["rows"].append(hang)
                job["done"] += 1

                if job["done"] % LUU_SAU_MOI == 0:
                    luu_excel(job["rows"], job["file_ra"])

            browser.close()

        if luu_excel(job["rows"], job["file_ra"]):
            log(job, f"Đã xuất Excel: {os.path.basename(job['file_ra'])}")
        job["status"] = "stopped" if job["stop"] else "done"
        tong = time.time() - job["t_bd"]
        log(job, f"Hoàn tất: {job['done']}/{job['total']} sinh viên, "
                 f"{job['loi']} lỗi, mất {tong / 60:.1f} phút")
    except Exception as e:
        job["status"] = "error"
        job["error"] = str(e).splitlines()[0][:200]
        log(job, f"LỖI CHƯƠNG TRÌNH: {job['error']}")
        traceback.print_exc()


def don_job_cu():
    """Xoá job quá 6 tiếng để không đầy bộ nhớ."""
    het = time.time() - 6 * 3600
    for jid in [j for j, v in JOBS.items() if v["created"] < het]:
        JOBS.pop(jid, None)


# =========================
# ROUTES
# =========================
@app.route("/")
def index():
    return render_template("index.html", url_mac_dinh=URL_MAC_DINH)


@app.route("/api/start", methods=["POST"])
def start():
    # Free tier RAM 512MB chỉ chạy được 1 Chromium — chặn job đồng thời
    for j in JOBS.values():
        if j["status"] in ("starting", "running"):
            return jsonify({"ok": False,
                            "error": "Đang có 1 lượt tra cứu chạy — hãy đợi xong hoặc dừng nó"}), 429

    url = (request.form.get("url") or "").strip()
    if not re.match(r"^https?://", url):
        return jsonify({"ok": False,
                        "error": "Link crawl không hợp lệ (phải bắt đầu bằng http:// hoặc https://)"}), 400

    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"ok": False,
                        "error": "Bắt buộc upload file mã tra cứu (.xlsx)"}), 400
    try:
        danh_sach = doc_danh_sach(BytesIO(f.read()))
    except Exception as e:
        return jsonify({"ok": False, "error": f"Không đọc được file Excel: {e}"}), 400

    if not danh_sach:
        return jsonify({"ok": False,
                        "error": "File không có dữ liệu — cần ít nhất cột MSSV + Mã tra cứu"}), 400

    don_job_cu()
    job_id = uuid.uuid4().hex[:8]
    job = {
        "id": job_id, "url": url, "status": "starting", "error": None,
        "total": len(danh_sach), "done": 0, "loi": 0,
        "logs": [], "rows": [], "stop": False,
        "headless": HEADLESS, "t_bd": None, "created": time.time(),
        "file_ra": os.path.join(OUT_DIR, f"bao_cao_diem_{job_id}.xlsx"),
    }
    JOBS[job_id] = job
    threading.Thread(target=chay_job, args=(job, danh_sach), daemon=True).start()
    return jsonify({"ok": True, "job_id": job_id, "total": len(danh_sach)})


@app.route("/api/status/<job_id>")
def status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"ok": False, "error": "Không tìm thấy job"}), 404
    log_from = int(request.args.get("log_from", 0))
    row_from = int(request.args.get("row_from", 0))
    return jsonify({
        "ok": True, "status": job["status"], "error": job["error"],
        "total": job["total"], "done": job["done"], "loi": job["loi"],
        "logs": job["logs"][log_from:], "log_next": len(job["logs"]),
        "rows": job["rows"][row_from:], "row_next": len(job["rows"]),
        "elapsed": round(time.time() - job["t_bd"], 1) if job["t_bd"] else 0,
        "co_file": os.path.exists(job["file_ra"]),
    })


@app.route("/api/stop/<job_id>", methods=["POST"])
def stop(job_id):
    job = JOBS.get(job_id)
    if job and job["status"] == "running":
        job["stop"] = True
        log(job, "Đang dừng sau khi xong sinh viên hiện tại...")
    return jsonify({"ok": True})


@app.route("/api/download/<job_id>")
def download(job_id):
    job = JOBS.get(job_id)
    if not job or not os.path.exists(job["file_ra"]):
        return "Chưa có file kết quả", 404
    return send_file(job["file_ra"], as_attachment=True, download_name="bao_cao_diem.xlsx")


if __name__ == "__main__":
    # Render set biến PORT (mặc định 10000); local dùng 5000
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)),
            threaded=True, debug=False)
