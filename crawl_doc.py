from playwright.sync_api import sync_playwright
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import unicodedata
import time
import re
import os

# =========================
# CẤU HÌNH
# =========================
URL = "https://script.google.com/macros/s/AKfycbzh0NPH8xQiE-7Wq1_a7Y3ZIoTTAfCPx5mWbbGh1596Uiw1VSFUFPk1y4bH41XfoDIuyw/exec"

THU_MUC = os.path.dirname(os.path.abspath(__file__))
FILE_VAO = os.path.join(THU_MUC, "MaTraCuu.xlsx")        # file danh sách MSSV + mã tra cứu
FILE_RA = os.path.join(THU_MUC, "bao_cao_diem.xlsx")      # file báo cáo xuất ra

HEADLESS = True        # True = chạy ẩn cho nhanh; đặt False nếu muốn xem trình duyệt
TIMEOUT_MOI_SV = 40    # giây, tối đa chờ kết quả cho 1 sinh viên
LUU_SAU_MOI = 10       # cứ xử lý 10 sinh viên thì lưu file 1 lần (tránh mất dữ liệu nếu bị ngắt)

TIEU_DE = [
    "MSSV", "Họ và tên", "Lớp",
    "Điểm tích luỹ (thang 10)", "Điểm tích luỹ (thang 4)",
    "Học lực", "Số TC tính điểm / chương trình",
    "Trạng thái", "Thời gian tra cứu",
]


# =========================
# HÀM PHỤ
# =========================
def bo_dau(s):
    """Bỏ dấu + chữ thường để so khớp nhãn ('luỹ' và 'lũy' đều thành 'luy')."""
    s = unicodedata.normalize("NFD", s or "").replace("\xa0", " ")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    return re.sub(r"\s+", " ", s).strip().lower()


def sang_so(x):
    try:
        return float(str(x).replace(",", "."))
    except Exception:
        return x


def doc_danh_sach(duong_dan):
    """Đọc file Excel đầu vào, trả về list dict. MSSV đọc dạng chữ để giữ số 0 đầu."""
    wb = load_workbook(duong_dan, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    tieu_de = [bo_dau(str(c)) if c else "" for c in rows[0]]

    def cot(*tu_khoa):
        for i, t in enumerate(tieu_de):
            if any(k in t for k in tu_khoa):
                return i
        raise ValueError(f"Không thấy cột {tu_khoa} trong file {duong_dan}")

    i_mssv = cot("ma so sinh vien", "mssv")
    i_ten = cot("ho va ten", "ho ten")
    i_lop = cot("lop")
    i_ma = cot("ma tra cuu")

    ds = []
    for r in rows[1:]:
        if not r or not r[i_mssv]:
            continue
        ds.append({
            "mssv": str(r[i_mssv]).strip(),
            "ten": str(r[i_ten] or "").strip(),
            "lop": str(r[i_lop] or "").strip(),
            "ma": str(r[i_ma] or "").strip(),
        })
    return ds


def doc_ket_qua_cu():
    """Nếu đã có file báo cáo (lần chạy trước bị ngắt), lấy lại các dòng đã xong để chạy tiếp."""
    if not os.path.exists(FILE_RA):
        return {}
    try:
        wb = load_workbook(FILE_RA, data_only=True)
        ws = wb.active
        cu = {}
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and r[0] and r[7] == "OK":      # chỉ giữ dòng thành công, dòng lỗi sẽ được thử lại
                cu[str(r[0])] = list(r)
        wb.close()
        return cu
    except Exception:
        return {}


def luu_excel(cac_hang):
    wb = Workbook()
    ws = wb.active
    ws.title = "Báo cáo điểm"
    ws.append(TIEU_DE)
    for h in cac_hang:
        ws.append(h)

    vien = Side(style="thin", color="999999")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=vien, right=vien, top=vien, bottom=vien)
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for i, col in enumerate(ws.columns, 1):
        rong = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(i)].width = min(max(rong + 4, 14), 40)

    # Cột MSSV là chữ để không mất số 0 đầu
    for row in ws.iter_rows(min_row=2, max_col=1):
        for c in row:
            c.number_format = "@"

    try:
        wb.save(FILE_RA)
        return True
    except PermissionError:
        print(f"\n!! Không ghi được '{FILE_RA}' vì file đang mở trong Excel. Hãy đóng file lại.")
        return False


def tim_frame_co_form(page, timeout=20):
    het_gio = time.time() + timeout
    while time.time() < het_gio:
        for frame in page.frames:
            try:
                if frame.locator("#mssv").count() > 0:
                    return frame
            except Exception:
                pass
        page.wait_for_timeout(200)
    return None


def tra_cuu_mot_sv(frame, sv):
    """Điền form, bấm nút, chờ kết quả CỦA ĐÚNG sinh viên này rồi đọc dữ liệu."""
    frame.locator("#mssv").fill(sv["mssv"])
    frame.locator("#ma").fill(sv["ma"])
    frame.locator("#btn").click()

    # Chờ đến khi #ketqua có chứa MSSV vừa nhập (tránh đọc nhầm kết quả của người trước)
    frame.wait_for_function(
        """mssv => {
            const el = document.querySelector('#ketqua');
            return el && el.innerText.includes(mssv);
        }""",
        arg=sv["mssv"],
        timeout=TIMEOUT_MOI_SV * 1000,
    )

    ketqua = frame.locator("#ketqua")
    cac_cap = ketqua.evaluate("""
        el => [...el.querySelectorAll('span')].map(s => {
            let b = s.parentElement.querySelector('b');
            if (!b && s.nextElementSibling && s.nextElementSibling.tagName === 'B')
                b = s.nextElementSibling;
            return { label: s.innerText, value: b ? b.innerText : null };
        })
    """)

    def tim(tu_khoa):
        tk = bo_dau(tu_khoa)
        for c in cac_cap:
            if c["value"] and tk in bo_dau(c["label"]):
                return c["value"].strip()
        return None

    ten = tim("tên") or sv["ten"]
    lop = tim("lớp") or sv["lop"]
    d10 = tim("điểm tích luỹ (thang 10)")
    d4 = tim("điểm tích luỹ (thang 4)")
    hoc_luc = tim("học lực")
    so_tc = tim("số tc tính điểm")

    thieu = [n for n, v in [("thang 10", d10), ("thang 4", d4),
                            ("học lực", hoc_luc), ("số TC", so_tc)] if v is None]
    trang_thai = "OK" if not thieu else "Thiếu: " + ", ".join(thieu)

    return [
        sv["mssv"], ten, lop,
        sang_so(d10) if d10 else "", sang_so(d4) if d4 else "",
        hoc_luc or "", so_tc or "",
        trang_thai, datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    ]


# =========================
# CHƯƠNG TRÌNH CHÍNH
# =========================
def main():
    if not os.path.exists(FILE_VAO):
        print(f"Không thấy file '{FILE_VAO}'. Hãy để MaTraCuu.xlsx cùng thư mục với file .py.")
        return

    danh_sach = doc_danh_sach(FILE_VAO)
    print(f"Đọc được {len(danh_sach)} sinh viên từ {os.path.basename(FILE_VAO)}")

    da_xong = doc_ket_qua_cu()
    if da_xong:
        print(f"Đã có {len(da_xong)} sinh viên thành công từ lần chạy trước, sẽ bỏ qua và chạy tiếp.")

    ket_qua = dict(da_xong)  # mssv -> hàng

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()

        print("Đang mở trang...")
        page.goto(URL)
        frame = tim_frame_co_form(page)
        if frame is None:
            print("Không tìm thấy form!")
            browser.close()
            return
        print("Đã tìm thấy form. Bắt đầu tra cứu...\n")

        t_bd = time.time()
        so_loi = 0
        xu_ly = 0

        for idx, sv in enumerate(danh_sach, 1):
            if sv["mssv"] in da_xong:
                continue

            t0 = time.time()
            try:
                hang = tra_cuu_mot_sv(frame, sv)
                print(f"[{idx}/{len(danh_sach)}] {sv['mssv']} {sv['ten']}: "
                      f"{hang[3]} / {hang[4]} / {hang[5]}  ({time.time() - t0:.1f}s)  {hang[7] if hang[7] != 'OK' else ''}")
            except Exception as e:
                so_loi += 1
                loi = str(e).splitlines()[0][:80]
                print(f"[{idx}/{len(danh_sach)}] {sv['mssv']} {sv['ten']}: LỖI - {loi}")
                hang = [sv["mssv"], sv["ten"], sv["lop"], "", "", "", "",
                        "Lỗi: không có kết quả (sai mã tra cứu hoặc server chậm)",
                        datetime.now().strftime("%d/%m/%Y %H:%M:%S")]

            ket_qua[sv["mssv"]] = hang
            xu_ly += 1

            # Lưu định kỳ, theo đúng thứ tự trong file đầu vào
            if xu_ly % LUU_SAU_MOI == 0:
                luu_excel([ket_qua[s["mssv"]] for s in danh_sach if s["mssv"] in ket_qua])

        browser.close()

    ok = luu_excel([ket_qua[s["mssv"]] for s in danh_sach if s["mssv"] in ket_qua])

    tong = time.time() - t_bd
    print("\n==============================")
    print(f"Hoàn tất: {len(ket_qua)}/{len(danh_sach)} sinh viên, {so_loi} lỗi, mất {tong / 60:.1f} phút")
    if ok:
        print(f"Đã xuất Excel: {FILE_RA}")
    print("==============================")


if __name__ == "__main__":
    main()