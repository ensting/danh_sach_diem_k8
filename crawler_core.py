"""Logic crawl tra cứu điểm — tách từ crawl_doc.py để app.py (web) tái sử dụng."""
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import unicodedata
import time
import re
import os

TIMEOUT_MOI_SV = 40    # giây, tối đa chờ kết quả cho 1 sinh viên
LUU_SAU_MOI = 10       # cứ xử lý 10 sinh viên thì lưu file 1 lần

TIEU_DE = [
    "MSSV", "Họ và tên", "Lớp",
    "Điểm tích luỹ (thang 10)", "Điểm tích luỹ (thang 4)",
    "Học lực", "Số TC tính điểm / chương trình",
    "Trạng thái", "Thời gian tra cứu",
]


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


def doc_danh_sach(nguon):
    """Đọc file Excel đầu vào (đường dẫn hoặc file-like), trả về list dict.

    Chỉ bắt buộc MSSV + Mã tra cứu — Họ tên/Lớp trang tra cứu trả về đủ.
    Nhận diện cột theo tiêu đề; file không có tiêu đề thì hiểu theo vị trí:
    cột đầu = MSSV, cột cuối = Mã tra cứu, các cột giữa = Tên/Lớp."""
    wb = load_workbook(nguon, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        return []

    tieu_de = [bo_dau(str(c)) if c else "" for c in rows[0]]

    def tim_cot(*tu_khoa, bo_qua=frozenset()):
        for i, t in enumerate(tieu_de):
            if i not in bo_qua and any(k in t for k in tu_khoa):
                return i
        return None

    i_mssv = tim_cot("ma so sinh vien", "mssv")
    bo_qua = {i_mssv} if i_mssv is not None else set()
    i_ma = tim_cot("ma tra cuu", "ma xac nhan", "mat khau", "ma tc", "code",
                   bo_qua=bo_qua)

    # File 2 cột: nhận ra được 1 cột thì cột còn lại là cột kia
    if len(tieu_de) == 2:
        if i_mssv is not None and i_ma is None:
            i_ma = 1 - i_mssv
        elif i_ma is not None and i_mssv is None:
            i_mssv = 1 - i_ma

    if (i_mssv is None) != (i_ma is None):
        thieu = "Mã tra cứu" if i_ma is None else "MSSV"
        raise ValueError(f"Không thấy cột '{thieu}' trong file đầu vào")

    ds = []
    if i_mssv is not None:
        # Có dòng tiêu đề — Tên/Lớp tuỳ chọn
        i_ten = tim_cot("ho va ten", "ho ten", bo_qua={i_mssv, i_ma})
        i_lop = tim_cot("lop", bo_qua={i_mssv, i_ma})

        def lay(r, i):
            return str(r[i]).strip() if i is not None and i < len(r) and r[i] is not None else ""

        for r in rows[1:]:
            if not r or not lay(r, i_mssv):
                continue
            ds.append({"mssv": lay(r, i_mssv), "ten": lay(r, i_ten),
                       "lop": lay(r, i_lop), "ma": lay(r, i_ma)})
        return ds

    # Không có tiêu đề: MSSV, (Tên), (Lớp), Mã — theo vị trí
    for r in rows:
        if not r:
            continue
        vals = [str(c).strip() if c is not None else "" for c in r]
        while vals and vals[-1] == "":
            vals.pop()
        if len(vals) < 2:
            continue
        if not ds and ("mssv" in bo_dau(vals[0]) or "ma so" in bo_dau(vals[0])
                       or "ma tra cuu" in bo_dau(vals[-1])):
            continue  # bỏ dòng tiêu đề lạ không nhận diện được
        ds.append({
            "mssv": vals[0],
            "ten": vals[1] if len(vals) >= 3 else "",
            "lop": vals[2] if len(vals) >= 4 else "",
            "ma": vals[-1],
        })
    return ds


def doc_danh_sach_tu_text(text):
    """Parse danh sách dán vào ô text. Mỗi dòng: MSSV, Họ tên, Lớp, Mã tra cứu
    (tối thiểu MSSV + Mã tra cứu; phân cách bằng Tab, dấu phẩy, chấm phẩy, | hoặc khoảng trắng)."""
    ds = []
    for dong in (text or "").splitlines():
        dong = dong.strip()
        if not dong:
            continue
        parts = [p.strip() for p in dong.split("\t")]
        if len(parts) == 1:
            parts = [p.strip() for p in re.split(r"[,;|]", dong)]
        if len(parts) == 1:
            parts = dong.split()
        if len(parts) < 2:
            continue
        if not ds and ("mssv" in bo_dau(parts[0]) or "ma so" in bo_dau(parts[0])):
            continue  # bỏ dòng tiêu đề nếu có
        ds.append({
            "mssv": parts[0],
            "ten": parts[1] if len(parts) >= 3 else "",
            "lop": parts[2] if len(parts) >= 4 else "",
            "ma": parts[-1],
        })
    return ds


def luu_excel(cac_hang, file_ra):
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
        wb.save(file_ra)
        return True
    except PermissionError:
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
    # Mỗi cặp nhãn/giá trị là 1 <b> trong div cha — nhãn = phần text còn lại của div.
    # Bắt được cả "Họ và tên: <b>X</b>" (.info) lẫn "<span>nhãn</span><b>giá trị</b>" (.tong).
    cac_cap = ketqua.evaluate("""
        el => [...el.querySelectorAll('b')].map(b => {
            const d = b.parentElement;
            if (!d || d === el) return null;
            const label = d.innerText.replace(b.innerText, '').replace(/[:：]\\s*$/, '').trim();
            return { label, value: b.innerText.trim() };
        }).filter(x => x && x.label)
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
