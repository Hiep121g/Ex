"""PyExcel - an Excel-like spreadsheet built with tkinter.

Run:  python excel_app.py [file.xlsx | file.pyxl | file.csv]
"""

import base64
import csv
import datetime as dt
import heapq
import io
import itertools
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import webbrowser
import tkinter as tk
import tkinter.font as tkfont
import unicodedata
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import charts
import dialogs
import formula_engine as fe
import images
import pivot as pv
import printing
import richclip
import session
import ttf
from formula_assist import FormulaAssist
from formula_engine import MAX_COLS, MAX_ROWS, cell_name, col_name, range_name
from function_info import CATEGORIES, FUNCTION_INFO
from pivot_pane import PivotPane
from print_preview import PrintPreview
from ribbon import (RIBBON_BG, TAB_BAR_BG, TAB_HOVER_BG, MenuIcons, Ribbon, RibbonButton, RibbonStyle, StyleTile,
                    Tooltip, color_popup, gallery_popup, mix)

APP_NAME = "PyExcel"

ACCENT = "#107c41"
GRID = "#e1e1e1"
HEADER_BG = "#f3f3f3"
HEADER_FG = "#555555"
HEADER_LINE = "#cfcfcf"
HEADER_SEL_BG = "#d2e9da"
HEADER_FULL_BG = "#a6d4b4"
HEADER_SEL_FG = "#0b5a2f"
SELECTION_TINT = "#a9cdb5"
SHEETBAR_BG = "#ececec"
FREEZE_LINE = "#8a8a8a"
BREAK_LINE = {False: "#8c8c8c", True: "#4f4f4f"}  # automatic / manual page breaks
OUTLINE_BG, OUTLINE_FG = "#fbfbfb", "#7a7a7a"
HYPERLINK_FG = "#0563c1"  # Excel's Hyperlink cell style: this colour, underlined
PROTECTED_MSG = ("Ô hoặc đối tượng bạn muốn thay đổi nằm trên trang tính đang được bảo vệ.\n"
                 "Để thay đổi, hãy bỏ bảo vệ: Xem lại → Bỏ bảo vệ Trang tính.")
FILTERED_ROW_FG = "#2b579a"  # row numbers inside a filtered range, as in Excel
NOTE_MARK = "#d13438"
NOTE_BG = "#ffffe1"
NOTE_LINE = "#9a9a7a"
REF_COLORS = ("#2f6fd6", "#d13438", "#8e44ad", "#c55a11", "#00838f", "#c2185b")

FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
NUMBER_FORMATS = [
    ("Chung", None), ("Số", "0.00"), ("Tiền tệ", "#,##0 ₫"), ("Kế toán", dialogs.accounting_format("₫", 0)),
    ("Ngày ngắn", "dd/mm/yyyy"), ("Ngày dài", dialogs.LONG_DATE), ("Thời gian", "hh:mm:ss"),
    ("Phần trăm", "0.00%"), ("Phân số", "# ?/?"), ("Khoa học", "0.00E+00"), ("Văn bản", "@"),
]
MORE_FORMATS = "Định dạng số khác..."
CURRENCIES = [("₫ Tiếng Việt", "₫"), ("$ Tiếng Anh (Mỹ)", "$"), ("€ Euro", "€"), ("¥ Tiếng Nhật", "¥"),
              ("£ Tiếng Anh (Anh)", "£")]
BORDER_MENU = [("Viền dưới", "bottom"), ("Viền trên", "top"), ("Viền trái", "left"), ("Viền phải", "right"), None,
               ("Không viền", "none"), ("Tất cả viền", "all"), ("Viền ngoài", "outside"),
               ("Viền ngoài dày", "thick_outside"), None, ("Viền dưới kép", "bottom_double"),
               ("Viền dưới dày", "thick_bottom"), ("Viền trên và dưới", "top_bottom"),
               ("Viền trên và dưới dày", "top_thick_bottom"), ("Viền trên và dưới kép", "top_double_bottom")]
DATA_BARS = [("Xanh dương", "#638ec6"), ("Xanh lá", "#63c384"), ("Đỏ", "#ff555a"), ("Cam", "#ffb628"),
             ("Xanh nhạt", "#008aef"), ("Tím", "#d6007b")]
COLOR_SCALES = [("Xanh lá - Vàng - Đỏ", ["#63be7b", "#ffeb84", "#f8696b"]),
                ("Đỏ - Vàng - Xanh lá", ["#f8696b", "#ffeb84", "#63be7b"]),
                ("Xanh lá - Trắng - Đỏ", ["#63be7b", "#fcfcff", "#f8696b"]),
                ("Đỏ - Trắng - Xanh lá", ["#f8696b", "#fcfcff", "#63be7b"]),
                ("Xanh dương - Trắng - Đỏ", ["#5a8ac6", "#fcfcff", "#f8696b"]),
                ("Đỏ - Trắng - Xanh dương", ["#f8696b", "#fcfcff", "#5a8ac6"]),
                ("Trắng - Đỏ", ["#fcfcff", "#f8696b"]), ("Đỏ - Trắng", ["#f8696b", "#fcfcff"]),
                ("Xanh lá - Trắng", ["#63be7b", "#fcfcff"]), ("Trắng - Xanh lá", ["#fcfcff", "#63be7b"])]
TABLE_STYLES = [  # (name, header fill, header text, band fill, line colour)
    ("Xanh dương", "#4472c4", "#ffffff", "#d9e1f2", "#8ea9db"), ("Cam", "#ed7d31", "#ffffff", "#fce4d6", "#f4b084"),
    ("Xám", "#a5a5a5", "#ffffff", "#ededed", "#c9c9c9"), ("Vàng", "#ffc000", "#000000", "#fff2cc", "#ffd966"),
    ("Xanh lam", "#5b9bd5", "#ffffff", "#ddebf7", "#9bc2e6"), ("Xanh lá", "#70ad47", "#ffffff", "#e2efda", "#a9d08e"),
    ("Tối", "#404040", "#ffffff", "#d9d9d9", "#a6a6a6"), ("Xanh đậm", "#1f4e78", "#ffffff", "#bdd7ee", "#2f75b5"),
]


def _all_borders(style, color):
    return {key: [style, color] for key in fe.BORDER_KEYS}


ACCENTS = [("Nhấn 1", "#4472c4"), ("Nhấn 2", "#ed7d31"), ("Nhấn 3", "#a5a5a5"), ("Nhấn 4", "#ffc000"),
           ("Nhấn 5", "#5b9bd5"), ("Nhấn 6", "#70ad47")]
CELL_STYLES = [
    ("Tốt, Xấu và Trung tính", [
        ("Chuẩn", {}), ("Xấu", {"bg": "#ffc7ce", "fg": "#9c0006"}), ("Tốt", {"bg": "#c6efce", "fg": "#006100"}),
        ("Trung tính", {"bg": "#ffeb9c", "fg": "#9c5700"})]),
    ("Dữ liệu và Mô hình", [
        ("Tính toán", {"bg": "#f2f2f2", "fg": "#fa7d00", "bold": True, **_all_borders("thin", "#7f7f7f")}),
        ("Ô kiểm tra", {"bg": "#a5a5a5", "fg": "#ffffff", "bold": True, **_all_borders("double", "#3f3f3f")}),
        ("Văn bản giải thích", {"fg": "#7f7f7f", "italic": True}),
        ("Nhập", {"bg": "#ffcc99", "fg": "#3f3f76", **_all_borders("thin", "#7f7f7f")}),
        ("Ô liên kết", {"fg": "#fa7d00", "bb": ["double", "#ff8001"]}),
        ("Ghi chú", {"bg": "#ffffcc", **_all_borders("thin", "#b2b2b2")}),
        ("Đầu ra", {"bg": "#f2f2f2", "fg": "#3f3f3f", "bold": True, **_all_borders("thin", "#3f3f3f")}),
        ("Văn bản cảnh báo", {"fg": "#ff0000"})]),
    ("Tiêu đề và Đầu đề", [
        ("Đầu đề 1", {"bold": True, "size": 15, "fg": "#44546a", "bb": ["thick", "#4472c4"]}),
        ("Đầu đề 2", {"bold": True, "size": 13, "fg": "#44546a", "bb": ["thick", "#a9c4eb"]}),
        ("Đầu đề 3", {"bold": True, "fg": "#44546a", "bb": ["medium", "#a9c4eb"]}),
        ("Đầu đề 4", {"bold": True, "fg": "#44546a"}),
        ("Tiêu đề", {"size": 18, "fg": "#44546a", "font": "Calibri Light"}),
        ("Tổng", {"bold": True, "bt": ["thin", "#4472c4"], "bb": ["double", "#4472c4"]})]),
    ("Kiểu ô theo chủ đề", [(f"{pct}% - {name}", {"bg": mix(color, "#ffffff", t), "fg": fg})
                            for pct, t, fg in ((20, 0.8, "#000000"), (40, 0.6, "#000000"), (60, 0.4, "#ffffff"))
                            for name, color in ACCENTS] + [(name, {"bg": color, "fg": "#ffffff"}) for name, color in ACCENTS]),
    ("Định dạng số", [("Dấu phẩy", {"num": "#,##0.00"}), ("Dấu phẩy [0]", {"num": "#,##0"}),
                      ("Tiền tệ", {"num": "#,##0 ₫"}), ("Phần trăm", {"num": "0%"})]),
]
STYLE_LOOKUP = {name: style for _, items in CELL_STYLES for name, style in items}
QUICK_STYLES = ["Chuẩn", "Trung tính", "Xấu", "Tốt", "Ô kiểm tra", "Ghi chú", "Ô liên kết", "Tính toán"]
STYLE_KEYS = ("font", "size", "bold", "italic", "underline", "strike", "fg", "bg") + fe.BORDER_KEYS

# picture shown beside a menu entry, by label (without a trailing "..."); see ribbon.ICONS
MENU_ICONS = {
    "Mới": "file_new", "Mở": "file_open", "Lưu": "save", "Lưu thành": "file_save_as", "Thoát": "file_exit",
    "Mở gần đây": "file_open", "Lưu tự động bản khôi phục": "save",
    "In": "print", "Xuất PDF": "pdf", "Thiết lập Trang": "page_setup",
    "Nhóm": "group", "Bỏ nhóm": "ungroup", "Tổng phụ": "subtotal", "Xóa dàn bài": "clear",
    "Trình Quản lý Tên": "name_manager", "Liên kết": "link", "Mở liên kết": "link", "Loại bỏ liên kết": "delete_x",
    "Thay ảnh": "picture", "Bảo vệ Trang tính": "protect", "Khóa ô": "lock_cell",
    "Cắt": "cut", "Sao chép": "copy", "Dán": "paste", "Tùy chọn dán": "paste", "Dán công thức": "paste_formulas",
    "Dán giá trị": "paste_values", "Dán định dạng": "paste_formats", "Dán chuyển vị": "paste_transpose",
    "Dán đặc biệt": "paste_special", "Gạch chân": "underline", "Gạch chân kép": "underline_double",
    "Hợp nhất & Căn giữa": "merge", "Hợp nhất ngang hàng": "merge_across", "Hợp nhất ô": "merge_cells",
    "Bỏ hợp nhất ô": "merge_unmerge", "Định dạng kế toán khác": "format_cells", "Định dạng ô": "format_cells",
    "Định dạng căn chỉnh ô": "format_cells", "Thêm viền": "format_cells",
    "Chèn": "insert_cells", "Chèn ô": "insert_cells", "Chèn hàng trang tính": "insert_row",
    "Chèn cột trang tính": "insert_col", "Chèn trang tính": "sheet", "Xóa": "delete_cells", "Xóa ô": "delete_cells",
    "Xóa hàng trang tính": "delete_row", "Xóa cột trang tính": "delete_col", "Xóa trang tính": "sheet_delete",
    "Chiều cao hàng": "row_size", "Tự động khớp chiều cao hàng": "row_autofit", "Độ rộng cột": "col_size",
    "Tự động khớp độ rộng cột": "col_autofit", "Ẩn & Bỏ ẩn": "hide_row", "Ẩn hàng": "hide_row", "Ẩn cột": "hide_col",
    "Bỏ ẩn hàng": "unhide_row", "Bỏ ẩn cột": "unhide_col", "Đổi tên trang tính": "sheet_rename",
    "Đổi tên": "sheet_rename", "Sao chép trang tính": "sheet_copy", "Màu tab": "sheet_tab_color",
    "Di chuyển sang trái": "sheet_left", "Di chuyển sang phải": "sheet_right",
    "Tổng": "autosum", "Trung bình": "stat_avg", "Đếm số": "stat_count", "Lớn nhất": "stat_max",
    "Nhỏ nhất": "stat_min", "Hàm khác": "function",
    "Xuống": "fill_down", "Sang phải": "fill_right", "Lên": "fill_up", "Sang trái": "fill_left", "Chuỗi": "fill_series",
    "Xóa tất cả": "clear", "Xóa định dạng": "clear_formats", "Xóa nội dung": "clear_contents",
    "Tìm": "find", "Thay thế": "replace", "Đi tới": "goto",
    "Sắp xếp": "sort_az", "Sắp xếp từ A đến Z": "sort_az", "Sắp xếp từ Z đến A": "sort_za",
    "Sắp xếp tùy chỉnh": "sort_custom", "Lọc": "filter", "Bật/tắt bộ lọc": "filter", "Xóa bộ lọc": "filter_clear",
    "Áp dụng lại bộ lọc": "filter_reapply", "Lọc theo giá trị của ô đã chọn": "filter_value",
    "Lọc theo màu của ô đã chọn": "filter_color", "Lọc theo màu phông của ô đã chọn": "filter_color",
    "Xoay ngược chiều kim đồng hồ": "angle_ccw", "Xoay theo chiều kim đồng hồ": "angle_cw",
    "Văn bản dọc": "angle_vertical", "Xoay văn bản lên": "angle_up", "Xoay văn bản xuống": "angle_down",
    "Quy tắc tô sáng ô": "cf_highlight", "Quy tắc trên/dưới": "cf_topbottom", "Thanh dữ liệu": "cf_databar",
    "Thang màu": "cf_scale", "Xóa quy tắc": "clear", "Xóa quy tắc khỏi ô đã chọn": "clear",
    "Xóa quy tắc khỏi toàn bộ trang tính": "clear", "Quản lý quy tắc": "cf_manage",
    "Lớn hơn": "cf_rule_gt", "Nhỏ hơn": "cf_rule_lt", "Nằm giữa": "cf_rule_between", "Bằng": "cf_rule_eq",
    "Văn bản có chứa": "cf_rule_text", "Giá trị trùng lặp": "cf_rule_dup", "10 mục đầu": "rank_top",
    "10% đầu": "rank_top_pct", "10 mục cuối": "rank_bottom", "10% cuối": "rank_bottom_pct",
    "Trên trung bình": "rank_above", "Dưới trung bình": "rank_below",
    "Cố định ngăn": "freeze", "Cố định hàng trên cùng": "freeze_row", "Cố định cột đầu tiên": "freeze_col",
    "Xác thực dữ liệu": "validation", "Xóa xác thực dữ liệu": "validation_clear",
    "Chèn/Sửa ghi chú": "note", "Xóa ghi chú": "note_delete",
    "Đổi kiểu biểu đồ": "chart_column", "Biểu đồ cột": "chart_column", "Biểu đồ thanh": "chart_bar",
    "Biểu đồ đường": "chart_line", "Biểu đồ tròn": "chart_pie", "Dữ liệu và tiêu đề biểu đồ": "chart_edit",
    "Xóa biểu đồ": "delete_x", "PivotTable": "pivot", "Làm mới tất cả": "refresh",
}

SHIFT, CONTROL, ALT = 0x0001, 0x0004, 0x20000
MODIFIER_KEYS = {"Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock", "Win_L", "Win_R"}
POINT_AFTER = ("=", "+", "-", "*", "/", "^", "(", ",", ";", ":", "<", ">", "&")
BIG_AREA = 20000  # selections larger than this are clipped to the used area

OPEN_TYPES = [("Bảng tính", "*.xlsx *.pyxl *.csv"), ("Excel", "*.xlsx"), ("PyExcel", "*.pyxl"),
              ("CSV", "*.csv"), ("Tất cả tệp", "*.*")]
SAVE_TYPES = [("Excel", "*.xlsx"), ("PyExcel", "*.pyxl"), ("CSV (chỉ trang hiện tại)", "*.csv")]
if not fe.HAS_XLSX:
    OPEN_TYPES = [t for t in OPEN_TYPES if t[0] != "Excel"]
    SAVE_TYPES = SAVE_TYPES[1:]

HELP_TEXT = """\
THANH RIBBON
  • Tab Trang chủ có đủ các nhóm như Excel: Bộ nhớ tạm, Phông chữ, Căn chỉnh, Số, Kiểu, Ô, Chỉnh sửa.
  • Nút có mũi tên ▾ mở thêm lựa chọn; mũi tên nhỏ ↘ ở góc nhóm mở hộp thoại Định dạng ô (Ctrl+1).
  • Ô "Cho tôi biết bạn muốn làm gì" trên thanh xanh: gõ tên lệnh (ví dụ "viền", "lọc") rồi Enter.
  • Nhấp đúp vào tên tab để thu gọn/mở rộng ribbon.
  • Xem → Cố định ngăn: giữ các hàng phía trên và cột bên trái ô đang chọn luôn hiển thị khi cuộn.
  • Dữ liệu → Xác thực dữ liệu: giới hạn giá trị được nhập (số, ngày, độ dài) hoặc tạo danh sách thả xuống
    (nguồn  Có,Không  hoặc  =$A$1:$A$10). Nhấp mũi tên cạnh ô hoặc Alt+↓ để chọn.
  • Xem lại → Ghi chú mới (Shift+F2): ô có ghi chú có tam giác đỏ ở góc, rê chuột vào để xem.
  • PivotTable: quét chọn bảng có hàng tiêu đề rồi bấm Chèn → PivotTable. Trong ngăn Các trường PivotTable,
    tích trường hoặc kéo trường vào Bộ lọc / Cột / Hàng / Giá trị; nhấp vào trường trong một vùng để sắp xếp,
    lọc mục, nhóm ngày hay chọn Tổng/Đếm/Trung bình... Sửa dữ liệu nguồn xong bấm Làm mới; nhấp đúp vào
    một số để xem các dòng dữ liệu tạo nên số đó.
  • Lọc dữ liệu: quét chọn bảng (kể cả hàng tiêu đề) rồi bấm Dữ liệu → Lọc (Ctrl+Shift+L). Nút ▾ ở tiêu đề
    có: sắp xếp (cả theo màu), lọc theo màu, Bộ lọc số / văn bản / ngày (lớn hơn, nằm giữa, có chứa, 10 mục
    đầu, trên trung bình, tháng này...), Bộ lọc tùy chỉnh (2 điều kiện Và/Hoặc, dùng * và ?) và danh sách
    giá trị có ô tìm kiếm. Thanh trạng thái cho biết số bản ghi tìm thấy.
  • Trang chủ → Tìm và Chọn → Đi tới Đặc biệt: chọn ô trống, hằng số, công thức, ô có ghi chú... (ví dụ chọn
    các ô trống rồi gõ giá trị và nhấn Ctrl+Enter để điền vào tất cả). Chuột phải vào tab để Ẩn / Bỏ ẩn trang tính.
  • Xem lại → Bảo vệ Trang tính: không cho sửa các ô bị khóa (mặc định mọi ô đều khóa). Trước khi bảo vệ, chọn
    các ô cần cho nhập rồi bấm Khóa ô (hoặc Định dạng ô → Bảo vệ, bỏ chọn Khóa); "Ẩn công thức" giấu công thức
    trên thanh công thức. Có thể đặt mật khẩu (kiểu Excel, chỉ chống sửa nhầm).
  • Dữ liệu → Văn bản thành Cột: tách một cột (ví dụ "Họ tên, SĐT, Địa chỉ") thành nhiều cột theo dấu phẩy,
    dấu chấm phẩy, dấu cách... hoặc theo độ rộng cố định; có thể giữ số 0 ở đầu số điện thoại.
  • Tên vùng: chọn vùng, gõ tên (ví dụ ĐơnGiá) vào ô Tên bên trái thanh công thức rồi Enter; sau đó dùng
    =SUM(ĐơnGiá). Công thức → Trình Quản lý Tên (Ctrl+F3) để sửa/xóa, Tạo từ Vùng chọn (Ctrl+Shift+F3) để đặt
    tên theo hàng tiêu đề. Chọn tên trong danh sách của ô Tên để nhảy tới vùng đó.
  • Chèn → Liên kết (Ctrl+K): ô dẫn tới trang web, tệp, một ô/tên trong sổ làm việc hoặc email. Nhấp vào ô để mở
    liên kết; nhấp và giữ để chọn ô. Hàm =HYPERLINK("https://...", "Tên") cũng tạo liên kết.
  • Chèn → Ảnh: ảnh PNG, GIF, JPEG đặt nổi trên trang tính; kéo để di chuyển, kéo góc để đổi cỡ (giữ tỉ lệ),
    chuột phải để thay ảnh / đặt lại kích thước / xóa. Ảnh được in và lưu cùng tệp. Cài Pillow (pip install
    pillow) để ảnh co giãn mịn hơn và xem được ảnh JPEG trên màn hình.
  • Dữ liệu → Nhóm (Alt+Shift+→): gom các hàng/cột đang chọn; nút −/+ ở lề trái (hoặc trên) thu gọn / mở
    rộng nhóm, các nút 1 2 3 ở góc hiện theo mức. Bỏ nhóm (Alt+Shift+←), Xóa dàn bài, Hiện/Ẩn chi tiết.
  • Dữ liệu → Tổng phụ: với bảng đã sắp xếp theo một cột (ví dụ Khu vực), thêm hàng "Bắc Tổng", "Nam Tổng"...
    dùng hàm SUBTOTAL (tổng, đếm, trung bình...) và hàng Tổng Cuối, kèm dàn bài để xem gọn theo mức.
  • Chèn → Biểu đồ (Cột, Thanh, Đường, Tròn): chọn vùng dữ liệu kèm tiêu đề rồi bấm. Kéo để di chuyển, kéo
    các nút ở viền để đổi kích thước, nhấp đúp để sửa dữ liệu/tiêu đề, Delete để xóa.

NHẬP DỮ LIỆU
  • Chọn ô rồi gõ; Enter/Tab để xác nhận, Esc để hủy. F2 hoặc nhấp đúp để sửa.
  • Công thức bắt đầu bằng "=":  =SUM(A1:A10)*2   =IF(B2>=5,"Đạt","Chưa đạt")   =Sheet2!A1+1
  • Khi đang gõ công thức, nhấp hoặc kéo chuột lên ô để chèn địa chỉ.
  • Gõ vài chữ đầu của tên hàm (=vl...) sẽ hiện danh sách hàm kèm mô tả: ↑/↓ để chọn, Tab hoặc nhấp đúp để
    chèn, Esc để đóng. Trong dấu ngoặc của hàm có chú thích các đối số, đối số đang nhập được in đậm.
  • F4 khi đang gõ công thức: đổi địa chỉ cạnh con trỏ lần lượt A1 → $A$1 → A$1 → $A1 → A1.
  • Gõ ' ở đầu để giữ nguyên dạng chữ ('007). Ngày nhập dạng 21/10/2026; phần trăm như 15%.

CHUỘT
  • Kéo để chọn vùng; Shift+nhấp mở rộng vùng; nhấp tiêu đề để chọn cả cột/hàng.
  • Kéo mép tiêu đề cột/hàng để đổi kích thước; nhấp đúp mép để tự khớp.
  • Kéo ô vuông nhỏ ở góc vùng chọn để sao chép hoặc điền chuỗi (1, 2 → 3, 4...; "Mục 1" → "Mục 2").
  • Chổi "Sao chép định dạng": nhấp một lần để dán định dạng một lần, nhấp đúp để dán nhiều lần.
  • Ctrl + lăn chuột để phóng to/thu nhỏ.

PHÍM TẮT
  Ctrl+C / X / V      Sao chép / Cắt / Dán         Ctrl+Z / Y        Hoàn tác / Làm lại
  Ctrl+B / I / U / 5  Đậm / Nghiêng / Gạch chân / Gạch ngang
  Ctrl+1              Định dạng ô                  Ctrl+Shift+L      Bật/tắt bộ lọc
  Ctrl+D / R          Điền xuống / sang phải       Ctrl+Enter        Nhập cho mọi ô đang chọn
  Ctrl+F / H / G      Tìm / Thay thế / Đi tới       Alt+=             Tự động tính tổng
  Ctrl++ / Ctrl+-     Chèn / Xóa ô                  Ctrl+9 / Ctrl+0   Ẩn hàng / cột
  Ctrl+;              Nhập ngày hôm nay            Ctrl+`            Hiện công thức
  Ctrl+Shift+&        Viền ngoài                   Ctrl+Shift+_      Bỏ viền
  Ctrl+mũi tên        Nhảy tới mép dữ liệu         Ctrl+Home / End   Về A1 / ô cuối
  Ctrl+N / O / S, F12 Mới / Mở / Lưu / Lưu thành   F9                Tính lại
  Ctrl+P              In / Xem trước khi in
  Ctrl+PgUp / PgDn    Trang tính trước / sau       Alt+Enter         Xuống dòng trong ô
  Ctrl+K              Chèn liên kết                Ctrl+F3           Trình Quản lý Tên
  Ctrl+nhấp chuột     Chọn thêm vùng rời (định dạng, xóa, Ctrl+Enter cho tất cả các vùng)

TỆP
  .xlsx mở được bằng Microsoft Excel (giữ công thức, định dạng, ô hợp nhất, định dạng có điều kiện,
  bộ lọc); .pyxl là định dạng riêng (JSON); .csv chỉ lưu giá trị của trang tính đang mở.
  Tệp → Mở gần đây: 10 tệp mở gần nhất. PyExcel tự lưu bản khôi phục mỗi 5 phút khi có thay đổi chưa lưu
  (Tệp → Lưu tự động bản khôi phục để đổi số phút); nếu máy tắt đột ngột, lần mở sau sẽ hỏi khôi phục.
  Mỗi trang tính rộng như Excel: 1.048.576 hàng × 16.384 cột (A → XFD).

TỆP CSV / VĂN BẢN
  • Mở tệp .csv tự nhận bảng mã như Excel: UTF-8 (có hoặc không có BOM), UTF-16 ("Văn bản Unicode"),
    còn lại là bảng mã ANSI của Windows - Windows-1258 khi tệp là tiếng Việt. Dấu thanh viết tách rời
    (kiểu "Unicode tổ hợp" / Windows-1258) được ghép lại để hiển thị đúng.
  • Tự nhận dấu phân cách (phẩy, chấm phẩy, Tab, |); tệp dùng dấu chấm phẩy với số kiểu 1.234,5 được đọc
    thành số. Thanh trạng thái cho biết bảng mã và dấu phân cách đã dùng.
  • Nếu chữ vẫn sai: Dữ liệu → Từ Văn bản/CSV để tự chọn bảng mã (kể cả TCVN3 / ABC và VNI của các
    phông .Vn..., VNI-...), dấu phân cách và kiểu số, có xem trước; dữ liệu vào một trang tính mới.

IN VÀ XUẤT PDF
  • Tệp → In... (Ctrl+P): xem trước từng trang; chọn in trang tính hiện hành, cả sổ làm việc hoặc vùng chọn,
    đổi hướng giấy, khổ giấy, lề, co giãn (vừa 1 trang / vừa chiều rộng) rồi bấm In hoặc Xuất PDF.
  • Nút In tạo tệp PDF rồi gửi tới máy in mặc định qua chương trình đọc PDF của máy; nếu máy chưa có
    chương trình in PDF trực tiếp, tệp PDF sẽ được mở để bạn bấm In trong đó.
  • Tab Bố cục Trang: Lề, Hướng, Khổ giấy, Vùng in (in riêng vùng chọn), Ngắt trang (chèn phía trên / bên
    trái ô đang chọn), In tiêu đề (hàng lặp lại ở đầu mỗi trang, ví dụ $1:$2), Co giãn, In đường lưới.
  • Thiết lập Trang có thêm Đầu trang/Chân trang với mã &P số trang, &N tổng số trang, &D ngày, &T giờ,
    &A tên trang tính, &F tên tệp (ví dụ "Trang &P / &N"). Đường nét đứt trên lưới cho biết chỗ ngắt trang.
  • PDF dùng phông chữ TrueType có trên máy (nhúng phần cần dùng) nên giữ đúng tiếng Việt; biểu đồ, viền,
    màu nền, thanh dữ liệu đều được in. Các thiết lập trang được lưu cùng tệp .xlsx / .pyxl.

SAO CHÉP SANG PHẦN MỀM KHÁC (Windows)
  Ctrl+C đưa vùng chọn lên bộ nhớ tạm như Excel: chữ, bảng (HTML, RTF) và hình vẽ (EMF, bitmap).
  Dán được vào Word, AutoCAD...; trong AutoCAD dùng Ctrl+V, hoặc lệnh PASTESPEC để chọn dạng dán.
"""

KEEP = object()


def normalize_text(text):
    """Lower-case text without Vietnamese diacritics, for the 'tell me' search."""
    text = unicodedata.normalize("NFD", text.lower().replace("đ", "d"))
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


class ChangeSet:
    """Cell edits collected so they can be applied, and undone, as one step."""

    def __init__(self, wb):
        self.wb = wb
        self.items = {}  # (sheet, r, c) -> [old_raw, new_raw, old_fmt, new_fmt]

    def set(self, sheet, r, c, raw=KEEP, fmt=KEEP):
        key = (sheet, r, c)
        item = self.items.get(key)
        if item is None:
            sh = self.wb.sheets[sheet]
            raw0, fmt0 = sh.cells.get((r, c)), sh.formats.get((r, c))
            item = self.items[key] = [raw0, raw0, fmt0, fmt0]
        if raw is not KEEP:
            item[1] = raw or None
        if fmt is not KEEP:
            item[3] = dict(fmt) if fmt else None

    def changes(self):
        return [(key, *item) for key, item in self.items.items() if item[0] != item[1] or item[2] != item[3]]


class SpreadsheetApp:
    def __init__(self, root, path=None):
        self.root = root
        self.scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        self.zoom = 1.0
        self.wb = fe.Workbook()
        self.sheet = "Sheet1"
        self.path = None
        self.dirty = False
        self.top = self.left = 0       # first visible row / column
        self.cur = self.ext = (0, 0)   # active cell and the moving corner of the selection
        self.areas = []                # more selected blocks (Ctrl+click, Go To Special), besides cur/ext
        self.views = {}                # sheet -> (top, left, cur, ext)
        self.undo_stack, self.redo_stack = [], []
        self.edit_serial = 0           # counts changes, for caches of derived data
        self._stats_key = self._stats_job = self._autosave_thread = None
        self.settings = session.Settings()
        self.recovery = session.Recovery(self.settings.folder)
        self._font_library, self._layout_cache = None, (None, None)
        self.breaks_shown = set()      # sheets showing page-break lines (after a print preview)
        self.clip = None
        self.editor = None
        self.edit_var = tk.StringVar(root)
        self.edit_cell = None
        self.edit_mode = None          # "enter" (typing replaced the cell) or "edit" (F2 / double-click)
        self.bar_editing = False
        self.point = None              # (start, end) of a reference inserted by clicking while typing a formula
        self.drag = None
        self.painter = None            # formats picked up by the Format Painter
        self.show_grid = self.show_headings = True
        self.show_formulas = False
        self.show_notes = False        # every note shown, not just the one under the pointer
        self.hover_note = None         # (sheet, row, col) of the note shown under the pointer
        self._list_button = None       # canvas box of the active cell's validation-list arrow
        self.chart_sel = None          # index of the selected chart on this sheet
        self.show_pivot_fields = True  # the PivotTable Fields pane appears in a pivot report
        self._pane_key = None          # (sheet, pivot id, layout) the pane shows
        self._chart_boxes = []         # [(index, canvas box)] of the charts drawn
        self.last_fill, self.last_fg = "#ffff00", "#ff0000"
        self.border_style, self.border_color, self.last_border = "thin", "#000000", "bottom"
        self.vis_rows, self.vis_cols = [], []
        self._scroll_rows = 0
        self._filter_buttons = []
        self._outline_buttons = []
        self._autoscroll_job = None
        self._last_mouse = (0, 0)
        self._cursor = ""
        self._v_total = self._h_total = 1
        self._ribbon_level = 0
        self._ribbon_widths = {}
        self._zoom_busy = False
        self.find_win = None
        self.fonts, self._widths, self._tints, self._images = {}, {}, {}, []
        self.commands = []  # (label, callable) for the "tell me" search
        self.command_icons = {}  # label -> icon of the ribbon button it came from

        families = set(tkfont.families(root))
        self.font_families = sorted(f for f in families if not f.startswith("@"))
        self.default_family = "Calibri" if "Calibri" in families else "Arial"
        self.ui_family = "Segoe UI" if "Segoe UI" in families else "Arial"
        self._init_metrics()
        self.menu_icons = MenuIcons(root, self.scale, os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons"))
        self.pictures = images.PictureCache(root)
        self._link_click = self._link_tip = self._link_tip_job = self._link_hover = None
        self._build_styles()
        self._build_ribbon()
        self._build_formula_bar()
        self._build_status_bar()
        self._build_sheet_bar()
        self._build_grid()
        self._build_context_menus()
        self._bind_keys()
        self.edit_var.trace_add("write", self._on_edit_text)
        self.fx_var.trace_add("write", self._on_fx_text)
        root.protocol("WM_DELETE_WINDOW", self.quit)
        if path:
            self.open_path(path)
        self.refresh_all()
        self.canvas.focus_set()
        self.recovery.beat(path=self.path)
        root.after(1200, self._offer_recovery)
        root.after(60_000, self._autosave_tick)

    @property
    def sh(self):
        return self.wb.sheets[self.sheet]

    # ------------------------------------------------------------------ metrics and fonts

    def _init_metrics(self):
        """Pixel sizes for the current DPI and zoom."""
        self.px = self.scale * self.zoom
        self.fonts.clear()
        self._widths.clear()
        size = -max(6, round(12 * self.px))
        self.header_font = tkfont.Font(root=self.root, family=self.ui_family, size=size)
        self.header_bold = tkfont.Font(root=self.root, family=self.ui_family, size=size, weight="bold")
        self._head_h0 = round(fe.DEFAULT_ROW_HEIGHT * self.px) if self.show_headings else 0
        self._head_w0 = self.header_font.measure("100000") + round(8 * self.px) if self.show_headings else 0
        self._outline_metrics()
        self.pad = max(1, round(3 * self.px))
        self.default_row_px = max(1, round(fe.DEFAULT_ROW_HEIGHT * self.px))

    def _outline_metrics(self):
        """Room for the outline symbols left of the row headings and above the column headings."""
        sh = self.wb.sheets.get(self.sheet)
        rows = max(sh.row_levels.values(), default=0) if sh else 0
        cols = max(sh.col_levels.values(), default=0) if sh else 0
        self.outline_step = round(16 * self.scale)
        self.outline_w = (rows + 1) * self.outline_step + 2 if rows else 0
        self.outline_h = (cols + 1) * self.outline_step + 2 if cols else 0
        self.header_w = max(self._head_w0, self.outline_step if cols else 0) + self.outline_w
        self.header_h = max(self._head_h0, self.outline_step if rows else 0) + self.outline_h

    def font_for(self, fmt):
        fmt = fmt or {}
        underline = fmt.get("underline")
        key = (fmt.get("font") or self.default_family, fmt.get("size") or 11, bool(fmt.get("bold")),
               bool(fmt.get("italic")), underline is True or underline == "single", bool(fmt.get("strike")))
        font = self.fonts.get(key)
        if font is None:
            font = self.fonts[key] = tkfont.Font(
                root=self.root, family=key[0], size=-max(1, round(key[1] * 4 / 3 * self.px)),
                weight="bold" if key[2] else "normal", slant="italic" if key[3] else "roman",
                underline=key[4], overstrike=key[5])
        return font

    def text_width(self, font, text):
        key = (font.name, text)
        width = self._widths.get(key)
        if width is None:
            if len(self._widths) > 20000:
                self._widths.clear()
            width = self._widths[key] = font.measure(text)
        return width

    def clip_text(self, text, font, width):
        if width <= 0:
            return ""
        if self.text_width(font, text) <= width:
            return text
        lo, hi = 0, len(text)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if font.measure(text[:mid]) <= width:
                lo = mid
            else:
                hi = mid - 1
        return text[:lo]

    def wrap_lines(self, text, font, width):
        """Word-wrap text to a pixel width (long words are broken)."""
        lines = []
        for para in text.split("\n"):
            line = ""
            for word in para.split(" "):
                candidate = f"{line} {word}" if line else word
                if self.text_width(font, candidate) <= width:
                    line = candidate
                    continue
                if line:
                    lines.append(line)
                while len(word) > 1 and self.text_width(font, word) > width:
                    cut = max(1, len(self.clip_text(word, font, width)))
                    lines.append(word[:cut])
                    word = word[cut:]
                line = word
            lines.append(line)
        return lines

    def fit_number(self, value, numfmt, font, width):
        """Number text that fits the column, rounding General numbers like Excel, else '###'."""
        if numfmt and numfmt != "@":
            text = fe.format_number(value, numfmt)
            if self.text_width(font, text) <= width:
                return text
        else:
            for digits in (None, 9, 8, 7, 6, 5, 4, 3, 2, 1):
                text = fe.format_general(value, digits)
                if self.text_width(font, text) <= width:
                    return text
        return "#" * max(1, int(width // max(1, self.text_width(font, "#"))))

    def tint(self, color):
        """`color` blended with the selection highlight."""
        tinted = self._tints.get(color)
        if tinted is None:
            r, g, b = (v // 257 for v in self.root.winfo_rgb(color))
            tinted = self._tints[color] = mix("#%02x%02x%02x" % (r, g, b), SELECTION_TINT, 0.35)
        return tinted

    # ------------------------------------------------------------------ UI construction

    def _build_styles(self):
        style = ttk.Style(self.root)
        style.configure("Ribbon.TCheckbutton", background=RIBBON_BG)

    def _menu_image(self, label, icon=None):
        """image/compound options showing the entry's picture (by MENU_ICONS unless `icon` is given)."""
        image = self.menu_icons.get(icon or MENU_ICONS.get(label.rstrip(". ")))
        return {"image": image, "compound": "left"} if image is not None else {}

    def _make_menu(self, items, parent=None):
        """tk.Menu from [(label, command[, accelerator[, icon]]) | (label, [sub items]) | (label, None) | None].

        Entries get the picture MENU_ICONS gives their label, or the icon named in the item."""
        menu = tk.Menu(parent or self.root, tearoff=0)
        for item in items:
            if item is None:
                menu.add_separator()
                continue
            label, target = item[0], item[1]
            image = self._menu_image(label, item[3] if len(item) > 3 else None)
            if isinstance(target, list):
                menu.add_cascade(label=label, menu=self._make_menu(target, menu), **image)
            elif isinstance(target, tuple):  # ("palette", on_pick, none_label): colour submenu
                menu.add_cascade(label=label, menu=self._palette_menu(menu, target[1], target[2]), **image)
            elif target is None:
                menu.add_command(label=label, state="disabled")
            else:
                menu.add_command(label=label, command=target, accelerator=item[2] if len(item) > 2 else "", **image)
                self.commands.append((label.rstrip(". "), target))
        return menu

    def _swatch(self, color):
        size = round(14 * self.scale)
        img = tk.PhotoImage(master=self.root, width=size, height=size)
        img.put("#7f7f7f", to=(0, 0, size, size))
        img.put(color, to=(1, 1, size - 1, size - 1))
        self._images.append(img)
        return img

    def _palette_menu(self, parent, on_pick, none_label=None, more=True):
        menu = tk.Menu(parent, tearoff=0)
        if none_label:
            menu.add_command(label=none_label, command=lambda: on_pick(None))
            menu.add_separator()
        for name, color in (("Đen", "#000000"), ("Trắng", "#ffffff"), ("Đỏ", "#ff0000"), ("Cam", "#ffc000"),
                            ("Vàng", "#ffff00"), ("Xanh lá", "#00b050"), ("Xanh dương", "#0070c0"),
                            ("Tím", "#7030a0"), ("Xám", "#808080")):
            menu.add_command(label=name, image=self._swatch(color), compound="left",
                             command=lambda c=color: on_pick(c))
        if more:
            menu.add_separator()
            menu.add_command(label="Màu khác...", command=lambda: self._ask_color(on_pick))
        return menu

    def _ask_color(self, on_pick, initial=None):
        _, color = colorchooser.askcolor(color=initial, parent=self.root, title="Chọn màu")
        if color:
            on_pick(color)

    def _btn(self, group, parent=None, large=False, side="left", label=None, **kw):
        """Ribbon button that is also registered for the 'tell me' search."""
        name = label or kw.get("text", "").replace("\n", " ") or (kw.get("tooltip") or "").split(" (")[0]
        if kw.get("command") and name:
            self.commands.append((name.split("\n")[0], kw["command"]))
            self.command_icons.setdefault(name.split("\n")[0], kw.get("icon"))
        if large:
            return group.large(**kw)
        return group.small(parent, side=side, **kw)

    def _build_ribbon(self):
        ui = self.ui = RibbonStyle(self.root, self.scale)
        rb = self.ribbon = Ribbon(self.root, ui)
        rb.pack(side="top", fill="x")
        for icon, command, tip in (("save", self.save_file, "Lưu (Ctrl+S)"), ("undo", self.undo, "Hoàn tác (Ctrl+Z)"),
                                   ("redo", self.redo, "Làm lại (Ctrl+Y)")):
            RibbonButton(rb.qat, ui, icon=icon, command=command, tooltip=tip, color="#ffffff", bg=TAB_BAR_BG,
                         hover_bg=TAB_HOVER_BG).pack(side="left")
        self.file_menu = self._make_menu([
            ("Mới", self.new_file, "Ctrl+N"), ("Mở...", self.open_file, "Ctrl+O"), ("Mở gần đây", []), None,
            ("Lưu", self.save_file, "Ctrl+S"), ("Lưu thành...", self.save_as, "F12"), None,
            ("In...", self.print_preview, "Ctrl+P"), ("Xuất PDF...", self.export_pdf),
            ("Thiết lập Trang...", self.page_setup_dialog), None,
            ("Lưu tự động bản khôi phục...", self.autosave_options), None,
            ("Thoát", self.quit)])
        self.recent_menu = tk.Menu(self.file_menu, tearoff=0, postcommand=self._fill_recent_menu)
        self.file_menu.entryconfigure(self.file_menu.index("Mở gần đây"), menu=self.recent_menu)
        rb.add_tab("Tệp", command=lambda label: self.file_menu.tk_popup(
            label.winfo_rootx(), label.winfo_rooty() + label.winfo_height()))
        self._build_home(rb.add_tab("Trang chủ"))
        self._build_insert_tab(rb.add_tab("Chèn"))
        self._build_page_layout_tab(rb.add_tab("Bố cục Trang"))
        self._build_formulas_tab(rb.add_tab("Công thức"))
        self._build_data_tab(rb.add_tab("Dữ liệu"))
        self._build_review_tab(rb.add_tab("Xem lại"))
        self._build_view_tab(rb.add_tab("Xem"))
        self._build_help_tab(rb.add_tab("Trợ giúp"))
        rb.select("Trang chủ")
        self._build_tell_me(rb.extra)
        rb.body.bind("<Configure>", self._fit_ribbon, add="+")

    def _build_tell_me(self, parent):
        hint = "Cho tôi biết bạn muốn làm gì..."
        self.tell_var = tk.StringVar(value=hint)
        entry = tk.Entry(parent, textvariable=self.tell_var, bg=TAB_BAR_BG, fg="#cfe6d8", relief="flat",
                         insertbackground="#ffffff", font=self.ui.font, width=34, highlightthickness=0)
        entry.pack(side="left", pady=self.ui.px(3), ipady=self.ui.px(2))

        def focus_in(e):
            if self.tell_var.get() == hint:
                self.tell_var.set("")
                entry.config(fg="#ffffff", bg=TAB_HOVER_BG)

        def focus_out(e):
            if not self.tell_var.get().strip():
                self.tell_var.set(hint)
            entry.config(fg="#cfe6d8", bg=TAB_BAR_BG)

        entry.bind("<FocusIn>", focus_in)
        entry.bind("<FocusOut>", focus_out)
        entry.bind("<Return>", lambda e: self._tell_me(entry))
        entry.bind("<Escape>", lambda e: self.canvas.focus_set())
        self.tell_entry = entry

    def _tell_me(self, entry):
        query = normalize_text(self.tell_var.get().strip())
        if not query:
            return
        seen, matches = set(), []
        for label, command in self.commands:
            if query in normalize_text(label) and label not in seen:
                seen.add(label)
                matches.append((label, command))
        if not matches:
            self.root.bell()
            return
        menu = tk.Menu(self.root, tearoff=0)
        for label, command in matches[:15]:
            menu.add_command(label=label, command=lambda c=command: (self.canvas.focus_set(), c()),
                             **self._menu_image(label, self.command_icons.get(label)))
        menu.tk_popup(entry.winfo_rootx(), entry.winfo_rooty() + entry.winfo_height())
        menu.grab_release()

    def _build_home(self, page):
        rb = self.ribbon
        # Clipboard
        g = rb.group(page, "Bộ nhớ tạm")
        paste_menu = self._make_menu([
            ("Dán", self.paste, "Ctrl+V"), ("Dán công thức", lambda: self.paste(mode="formulas")),
            ("Dán giá trị", lambda: self.paste(mode="values")), ("Dán định dạng", lambda: self.paste(mode="formats")),
            ("Dán chuyển vị", lambda: self.paste(transpose=True)), None, ("Dán đặc biệt...", self.paste_special)])
        self._btn(g, large=True, icon="paste", text="Dán", command=self.paste, menu=paste_menu, split=True,
                  tooltip="Dán (Ctrl+V)")
        col = g.column()
        self._btn(g, col, side="top", icon="cut", text="Cắt", command=self.cut, tooltip="Cắt (Ctrl+X)")
        self._btn(g, col, side="top", icon="copy", text="Sao chép", command=self.copy, tooltip="Sao chép (Ctrl+C)")
        self.btn_painter = self._btn(g, col, side="top", icon="painter", text="Sao định dạng",
                                     command=self.start_painter,
                                     tooltip="Sao chép định dạng\nNhấp đúp để dán định dạng nhiều lần")
        self.btn_painter.bind("<Double-Button-1>", lambda e: self.start_painter(sticky=True))
        # Font
        g = rb.group(page, "Phông chữ", lambda: self.format_cells_dialog("font"), "Định dạng phông chữ (Ctrl+1)")
        row0, row1, _ = g.rows()
        self.font_var, self.size_var = tk.StringVar(self.root), tk.StringVar(self.root)
        self.font_box = ttk.Combobox(row0, textvariable=self.font_var, values=self.font_families, width=15,
                                     font=self.ui.font)
        self.font_box.pack(side="left", padx=(0, 3), pady=1)
        self.size_box = ttk.Combobox(row0, textvariable=self.size_var, values=FONT_SIZES, width=4, font=self.ui.font)
        self.size_box.pack(side="left", padx=(0, 3), pady=1)
        for box, command in ((self.font_box, self._apply_font_box), (self.size_box, self._apply_size_box)):
            box.bind("<<ComboboxSelected>>", lambda e, c=command: c())
            box.bind("<Return>", lambda e, c=command: c())
            box.bind("<MouseWheel>", lambda e: "break")
        Tooltip(self.font_box, "Phông chữ", self.ui.small_font)
        Tooltip(self.size_box, "Cỡ phông", self.ui.small_font)
        self._btn(g, row0, icon="grow_font", command=lambda: self.step_font_size(1), tooltip="Tăng cỡ phông")
        self._btn(g, row0, icon="shrink_font", command=lambda: self.step_font_size(-1), tooltip="Giảm cỡ phông")
        self.btn_bold = self._btn(g, row1, icon="bold", label="Đậm", command=lambda: self.toggle_style("bold"),
                                  tooltip="Đậm (Ctrl+B)")
        self.btn_italic = self._btn(g, row1, icon="italic", label="Nghiêng",
                                    command=lambda: self.toggle_style("italic"), tooltip="Nghiêng (Ctrl+I)")
        underline_menu = self._make_menu([("Gạch chân", lambda: self.set_underline("single")),
                                          ("Gạch chân kép", lambda: self.set_underline("double"))])
        self.btn_underline = self._btn(g, row1, icon="underline", label="Gạch chân",
                                       command=lambda: self.toggle_style("underline"), menu=underline_menu,
                                       split=True, tooltip="Gạch chân (Ctrl+U)")
        g.gap(row1)
        self.btn_border = self._btn(g, row1, icon=f"border_{self.last_border}", label="Viền",
                                    command=lambda: self.apply_border(self.last_border),
                                    menu=self._border_menu(), split=True, tooltip="Viền")
        g.gap(row1)
        self.btn_fill = self._btn(g, row1, icon="fill", color=self.last_fill, label="Màu tô",
                                  command=lambda: self.apply_format({"bg": self.last_fill}),
                                  menu=lambda b: self._color_dropdown(b, "bg"), split=True, tooltip="Màu tô")
        self.btn_fg = self._btn(g, row1, icon="font_color", color=self.last_fg, label="Màu phông",
                                command=lambda: self.apply_format({"fg": self.last_fg}),
                                menu=lambda b: self._color_dropdown(b, "fg"), split=True, tooltip="Màu phông")
        # Alignment
        g = rb.group(page, "Căn chỉnh", lambda: self.format_cells_dialog("align"), "Định dạng căn chỉnh ô")
        row0, row1, _ = g.rows()
        self.btn_vtop = self._btn(g, row0, icon="valign_top", label="Căn trên",
                                  command=lambda: self.set_valign("top"), tooltip="Căn trên")
        self.btn_vmid = self._btn(g, row0, icon="valign_center", label="Căn giữa theo chiều dọc",
                                  command=lambda: self.set_valign("center"), tooltip="Căn giữa")
        self.btn_vbot = self._btn(g, row0, icon="valign_bottom", label="Căn dưới",
                                  command=lambda: self.set_valign(None), tooltip="Căn dưới")
        g.gap(row0)
        self._btn(g, row0, icon="orientation", menu=self._orientation_menu(), tooltip="Hướng văn bản")
        g.gap(row0)
        self.btn_wrap = self._btn(g, row0, icon="wrap", text="Xuống dòng tự động", command=self.toggle_wrap,
                                  tooltip="Xuống dòng tự động\nHiện toàn bộ văn bản trên nhiều dòng")
        self.btn_left = self._btn(g, row1, icon="align_left", label="Căn trái",
                                  command=lambda: self.set_align("left"), tooltip="Căn trái")
        self.btn_center = self._btn(g, row1, icon="align_center", label="Căn giữa",
                                    command=lambda: self.set_align("center"), tooltip="Căn giữa")
        self.btn_right = self._btn(g, row1, icon="align_right", label="Căn phải",
                                   command=lambda: self.set_align("right"), tooltip="Căn phải")
        g.gap(row1)
        self._btn(g, row1, icon="indent_dec", label="Giảm thụt lề", command=lambda: self.change_indent(-1),
                  tooltip="Giảm thụt lề")
        self._btn(g, row1, icon="indent_inc", label="Tăng thụt lề", command=lambda: self.change_indent(1),
                  tooltip="Tăng thụt lề")
        g.gap(row1)
        merge_menu = self._make_menu([("Hợp nhất & Căn giữa", lambda: self.merge(center=True)),
                                      ("Hợp nhất ngang hàng", lambda: self.merge(across=True)),
                                      ("Hợp nhất ô", self.merge), ("Bỏ hợp nhất ô", self.unmerge)])
        self.btn_merge = self._btn(g, row1, icon="merge", text="Hợp nhất & Căn giữa", command=self.merge_center,
                                   menu=merge_menu, split=True, tooltip="Hợp nhất & Căn giữa")
        # Number
        g = rb.group(page, "Số", lambda: self.format_cells_dialog("number"), "Định dạng số")
        row0, row1, _ = g.rows()
        self.numfmt_box = ttk.Combobox(row0, state="readonly", width=14, font=self.ui.font,
                                       values=[label for label, _ in NUMBER_FORMATS] + [MORE_FORMATS])
        self.numfmt_box.pack(side="left", pady=1)
        self.numfmt_box.bind("<<ComboboxSelected>>", self._on_numfmt)
        self.numfmt_box.bind("<MouseWheel>", lambda e: "break")
        accounting_menu = self._make_menu(
            [(label, lambda s=sym: self.apply_format({"num": dialogs.accounting_format(s, 0 if s in "₫¥" else 2)}))
             for label, sym in CURRENCIES] + [None, ("Định dạng kế toán khác...", lambda: self.format_cells_dialog("number"))])
        self._btn(g, row1, icon="accounting", label="Định dạng số kế toán", menu=accounting_menu, split=True,
                  command=lambda: self.apply_format({"num": dialogs.accounting_format("₫", 0)}),
                  tooltip="Định dạng số kế toán")
        self._btn(g, row1, icon="percent", label="Kiểu phần trăm", command=lambda: self.apply_format({"num": "0%"}),
                  tooltip="Kiểu phần trăm")
        self._btn(g, row1, icon="comma", label="Kiểu dấu phẩy", command=lambda: self.apply_format({"num": "#,##0.00"}),
                  tooltip="Kiểu dấu phẩy")
        g.gap(row1)
        self._btn(g, row1, icon="dec_inc", label="Tăng số thập phân", command=lambda: self.change_decimals(1),
                  tooltip="Tăng số thập phân")
        self._btn(g, row1, icon="dec_dec", label="Giảm số thập phân", command=lambda: self.change_decimals(-1),
                  tooltip="Giảm số thập phân")
        # Styles
        g = rb.group(page, "Kiểu")
        self._btn(g, large=True, icon="cond_format", text="Định dạng có\nđiều kiện", menu=self._cond_menu(),
                  tooltip="Định dạng có điều kiện")
        self._btn(g, large=True, icon="format_table", text="Định dạng\nnhư Bảng", menu=self._table_gallery,
                  tooltip="Định dạng như Bảng")
        self._build_style_gallery(g)
        # Cells
        g = rb.group(page, "Ô")
        insert_menu = self._make_menu([("Chèn ô...", self.insert_cells_dialog),
                                       ("Chèn hàng trang tính", self.insert_rows),
                                       ("Chèn cột trang tính", self.insert_cols), None,
                                       ("Chèn trang tính", self.add_sheet)])
        delete_menu = self._make_menu([("Xóa ô...", self.delete_cells_dialog),
                                       ("Xóa hàng trang tính", self.delete_rows),
                                       ("Xóa cột trang tính", self.delete_cols), None,
                                       ("Xóa trang tính", lambda: self.delete_sheet(self.sheet))])
        tab_colors = ("palette", self.set_tab_color, "Không màu")
        format_menu = self._make_menu([
            ("KÍCH THƯỚC Ô", None), ("Chiều cao hàng...", self.ask_row_height),
            ("Tự động khớp chiều cao hàng", self.autofit_selected_rows), ("Độ rộng cột...", self.ask_column_width),
            ("Tự động khớp độ rộng cột", self.autofit_selected_columns), None,
            ("HIỂN THỊ", None), ("Ẩn & Bỏ ẩn", [("Ẩn hàng", self.hide_rows, "Ctrl+9"), ("Ẩn cột", self.hide_cols, "Ctrl+0"),
                                                ("Ẩn trang tính", self.hide_sheet, "", "sheet_hide"), None,
                                                ("Bỏ ẩn hàng", self.unhide_rows), ("Bỏ ẩn cột", self.unhide_cols),
                                                ("Bỏ ẩn trang tính...", self.unhide_sheet, "", "sheet_unhide")]),
            None, ("SẮP XẾP TRANG TÍNH", None), ("Đổi tên trang tính", lambda: self.rename_sheet(self.sheet)),
            ("Sao chép trang tính", self.copy_sheet), ("Màu tab", tab_colors), None,
            ("BẢO VỆ", None), ("Bảo vệ Trang tính...", self.toggle_protection), ("Khóa ô", self.toggle_lock_cells),
            ("Định dạng ô...", lambda: self.format_cells_dialog(), "Ctrl+1")])
        self._btn(g, large=True, icon="insert_cells", text="Chèn", command=self.insert_smart, menu=insert_menu,
                  split=True, tooltip="Chèn ô, hàng, cột hoặc trang tính")
        self._btn(g, large=True, icon="delete_cells", text="Xóa", command=self.delete_smart, menu=delete_menu,
                  split=True, tooltip="Xóa ô, hàng, cột hoặc trang tính")
        self._btn(g, large=True, icon="format_cells", text="Định dạng", menu=format_menu,
                  tooltip="Kích thước hàng/cột, ẩn, trang tính...")
        # Editing
        g = rb.group(page, "Chỉnh sửa")
        col = g.column()
        autosum_menu = self._make_menu([("Tổng", lambda: self.autosum("SUM")), ("Trung bình", lambda: self.autosum("AVERAGE")),
                                        ("Đếm số", lambda: self.autosum("COUNT")), ("Lớn nhất", lambda: self.autosum("MAX")),
                                        ("Nhỏ nhất", lambda: self.autosum("MIN")), None,
                                        ("Hàm khác...", self.insert_function_dialog)])
        self.btn_autosum = self._btn(g, col, side="top", icon="autosum", text="Tổng tự động",
                                     command=lambda: self.autosum("SUM"), menu=autosum_menu, split=True,
                                     tooltip="Tổng tự động (Alt+=)")
        self.btn_fill_menu = self._btn(g, col, side="top", icon="fill_down", text="Điền", menu=self._make_menu([
            ("Xuống", self.fill_down, "Ctrl+D"), ("Sang phải", self.fill_right, "Ctrl+R"), ("Lên", self.fill_up),
            ("Sang trái", self.fill_left), None, ("Chuỗi...", self.fill_series_dialog)]), tooltip="Điền")
        self.btn_clear_menu = self._btn(g, col, side="top", icon="clear", text="Xóa", menu=self._make_menu([
            ("Xóa tất cả", self.clear_all), ("Xóa định dạng", self.clear_formats),
            ("Xóa nội dung", self.clear_contents, "Delete")]), tooltip="Xóa")
        self._btn(g, large=True, icon="sort_filter", text="Sắp xếp\nvà Lọc", menu=self._sort_menu(),
                  tooltip="Sắp xếp và Lọc")
        self._btn(g, large=True, icon="find", text="Tìm và\nChọn", menu=self._make_menu([
            ("Tìm...", self.open_find, "Ctrl+F"), ("Thay thế...", lambda: self.open_find(True), "Ctrl+H"),
            ("Đi tới...", self.focus_name_box, "Ctrl+G"), ("Đi tới Đặc biệt...", self.go_to_special, "", "goto")]),
                  tooltip="Tìm và Chọn")
        # labels dropped, level by level, when the window is too narrow for the whole tab
        level2 = [self.btn_wrap, self.btn_merge, self.btn_autosum, self.btn_fill_menu, self.btn_clear_menu]
        level3 = [w for w in self.btn_painter.master.winfo_children() if isinstance(w, RibbonButton)]
        level3 += [w for group in page.winfo_children() if isinstance(group, tk.Frame) and hasattr(group, "body")
                   for w in group.body.winfo_children() if isinstance(w, RibbonButton) and w.large and w.text
                   and w.icon != "paste"]
        self._collapsible = {2: [(b, b.text) for b in level2], 3: [(b, b.text) for b in level3]}

    def _build_style_gallery(self, g):
        frame = self.style_gallery = tk.Frame(g.body, bg=RIBBON_BG)
        frame.pack(side="left", fill="y", padx=(4, 0))
        tiles = tk.Frame(frame, bg="#ffffff", highlightthickness=1, highlightbackground="#c6c6c6")
        tiles.pack(side="left", pady=(self.ui.px(4), 0))
        w, h = self.ui.px(82), (self.ui.content_h - self.ui.px(14)) // 2
        for i, name in enumerate(QUICK_STYLES):
            tile = StyleTile(tiles, self.ui, name, self._style_preview(STYLE_LOOKUP[name]), w, h,
                             lambda n=name: self.apply_cell_style(n))
            tile.grid(row=i // 4, column=i % 4, padx=1, pady=1)
            self.commands.append((f"Kiểu ô {name}", lambda n=name: self.apply_cell_style(n)))
        RibbonButton(frame, self.ui, menu=self._styles_popup, tooltip="Tất cả kiểu ô",
                     height=self.ui.content_h - self.ui.px(10)).pack(side="left", padx=1, pady=(self.ui.px(4), 0))
        self.btn_styles_compact = RibbonButton(g.body, self.ui, large=True, icon="format_cells", text="Kiểu\nô",
                                               menu=self._styles_popup, tooltip="Kiểu ô")

    def _build_insert_tab(self, page):
        rb = self.ribbon
        g = rb.group(page, "Bảng")
        self._btn(g, large=True, icon="pivot", text="PivotTable\n", label="PivotTable", command=self.insert_pivot,
                  tooltip="Chèn PivotTable\nTóm tắt dữ liệu: tổng, đếm, trung bình... theo hàng và cột")
        self._btn(g, large=True, icon="format_table", text="Bảng", menu=self._table_gallery, tooltip="Tạo bảng có định dạng")
        g = rb.group(page, "Ô và trang tính")
        col = g.column()
        self._btn(g, col, side="top", icon="insert_row", text="Chèn hàng", command=self.insert_rows)
        self._btn(g, col, side="top", icon="insert_col", text="Chèn cột", command=self.insert_cols)
        self._btn(g, col, side="top", icon="insert_cells", text="Chèn ô...", command=self.insert_cells_dialog)
        col = g.column()
        self._btn(g, col, side="top", icon="sheet", text="Chèn trang tính", command=self.add_sheet)
        self._btn(g, col, side="top", icon="copy", text="Sao chép trang tính", command=self.copy_sheet)
        g = rb.group(page, "Biểu đồ")
        for kind, text in (("column", "Cột"), ("bar", "Thanh"), ("line", "Đường"), ("pie", "Tròn")):
            self._btn(g, large=True, icon="chart_" + kind, text=f"{text}\n", label=f"Biểu đồ {text.lower()}",
                      command=lambda k=kind: self.insert_chart(k),
                      tooltip=f"Chèn biểu đồ {text.lower()}\nVẽ từ vùng dữ liệu đang chọn")
        g = rb.group(page, "Minh họa")
        self._btn(g, large=True, icon="picture", text="Ảnh\n", command=self.insert_picture,
                  tooltip="Chèn ảnh\nẢnh PNG, GIF, JPEG từ máy tính, đặt nổi trên các ô")
        g = rb.group(page, "Liên kết")
        self._btn(g, large=True, icon="link", text="Liên\nkết", command=self.insert_hyperlink,
                  tooltip="Chèn liên kết (Ctrl+K)\nTới trang web, tệp, một ô trong sổ làm việc hoặc email")
        g = rb.group(page, "Hàm")
        self._btn(g, large=True, icon="function", text="Chèn\nhàm", command=self.insert_function_dialog,
                  tooltip="Chèn hàm (Shift+F3)")
        g = rb.group(page, "Ngày giờ")
        col = g.column()
        self._btn(g, col, side="top", icon="calendar", text="Ngày hôm nay", command=lambda: self.insert_now(False),
                  tooltip="Nhập ngày hôm nay (Ctrl+;)")
        self._btn(g, col, side="top", icon="clock", text="Giờ hiện tại", command=lambda: self.insert_now(True))

    def _build_formulas_tab(self, page):
        rb = self.ribbon
        g = rb.group(page, "Thư viện hàm")
        self._btn(g, large=True, icon="function", text="Chèn\nhàm", command=self.insert_function_dialog)
        self._btn(g, large=True, icon="autosum", text="Tổng tự\nđộng", command=lambda: self.autosum("SUM"),
                  menu=self._make_menu([("Tổng", lambda: self.autosum("SUM")), ("Trung bình", lambda: self.autosum("AVERAGE")),
                                        ("Đếm số", lambda: self.autosum("COUNT")), ("Lớn nhất", lambda: self.autosum("MAX")),
                                        ("Nhỏ nhất", lambda: self.autosum("MIN"))]), split=True)
        labels = {"Toán học & Lượng giác": "Toán học &\nLượng giác", "Thống kê": "Thống\nkê", "Logic": "Logic\n",
                  "Văn bản": "Văn\nbản", "Ngày & Giờ": "Ngày &\nGiờ", "Tra cứu & Tham chiếu": "Tra cứu &\nTham chiếu",
                  "Thông tin": "Thông\ntin"}
        for cat in CATEGORIES:
            names = sorted(n for n, info in FUNCTION_INFO.items() if info[0] == cat)
            menu = tk.Menu(self.root, tearoff=0)
            for n in names:
                menu.add_command(label=n, command=lambda n=n: self.insert_function(n))
            self._btn(g, large=True, icon="book", text=labels[cat], menu=menu, tooltip=cat)
        g = rb.group(page, "Tên đã Xác định")
        self._btn(g, large=True, icon="name_manager", text="Trình Quản\nlý Tên", command=self.name_manager,
                  tooltip="Trình Quản lý Tên (Ctrl+F3)\nTạo, sửa, xóa các tên dùng trong công thức")
        col = g.column()
        self._btn(g, col, side="top", icon="name_define", text="Xác định Tên", command=self.define_name_dialog,
                  tooltip="Đặt tên cho vùng đang chọn")
        use_menu = tk.Menu(self.root, tearoff=0)
        use_menu.configure(postcommand=lambda: self._names_menu(use_menu))
        self._btn(g, col, side="top", icon="name_use", text="Dùng trong Công thức", menu=use_menu)
        self._btn(g, col, side="top", icon="name_create", text="Tạo từ Vùng chọn",
                  command=self.create_names_from_selection, tooltip="Tạo tên từ tiêu đề (Ctrl+Shift+F3)")
        g = rb.group(page, "Kiểm tra công thức")
        self.btn_show_formulas = self._btn(g, large=True, icon="show_formulas", text="Hiện\ncông thức",
                                           command=self.toggle_show_formulas, tooltip="Hiện công thức (Ctrl+`)")
        g = rb.group(page, "Tính toán")
        self._btn(g, large=True, icon="calc", text="Tính\nngay", command=self.recalculate, tooltip="Tính lại (F9)")

    def _build_data_tab(self, page):
        rb = self.ribbon
        g = rb.group(page, "Sắp xếp và Lọc")
        col = g.column()
        self._btn(g, col, side="top", icon="sort_az", label="Sắp xếp A đến Z", command=lambda: self.sort_selection(False),
                  tooltip="Sắp xếp từ A đến Z")
        self._btn(g, col, side="top", icon="sort_za", label="Sắp xếp Z đến A", command=lambda: self.sort_selection(True),
                  tooltip="Sắp xếp từ Z đến A")
        self._btn(g, large=True, icon="sort_filter", text="Sắp xếp", command=self.custom_sort, tooltip="Sắp xếp tùy chỉnh")
        self.btn_filter = self._btn(g, large=True, icon="filter", text="Lọc", command=self.toggle_filter,
                                    tooltip="Lọc (Ctrl+Shift+L)")
        col = g.column()
        self._btn(g, col, side="top", icon="filter_clear", text="Xóa bộ lọc", command=self.clear_filter)
        self._btn(g, col, side="top", icon="calc", text="Áp dụng lại", command=self.reapply_filter)
        g = rb.group(page, "PivotTable")
        self._btn(g, large=True, icon="refresh", text="Làm mới\ntất cả", command=self.refresh_all_pivots,
                  tooltip="Làm mới tất cả PivotTable theo dữ liệu nguồn hiện tại")
        self.btn_pivot_fields = self._btn(g, large=True, icon="pivot_fields", text="Danh sách\ntrường",
                                          command=self.toggle_pivot_fields,
                                          tooltip="Hiện/ẩn ngăn Các trường PivotTable")
        g = rb.group(page, "Dàn bài")
        self._btn(g, large=True, icon="group", text="Nhóm\n", command=lambda: self.group_selection(1),
                  tooltip="Nhóm (Alt+Shift+→)\nGom các hàng hoặc cột để thu gọn / mở rộng bằng nút −/+")
        self._btn(g, large=True, icon="ungroup", text="Bỏ\nnhóm", command=lambda: self.group_selection(-1),
                  menu=self._make_menu([("Bỏ nhóm...", lambda: self.group_selection(-1), "Alt+Shift+←", "ungroup"),
                                        ("Xóa dàn bài", self.clear_outline, "", "clear")]), split=True,
                  tooltip="Bỏ nhóm (Alt+Shift+←)")
        self._btn(g, large=True, icon="subtotal", text="Tổng\nphụ", command=self.subtotal_dialog,
                  tooltip="Tổng phụ\nTính tổng (đếm, trung bình...) cho từng nhóm giá trị giống nhau")
        col = g.column()
        self._btn(g, col, side="top", icon="show_detail", text="Hiện chi tiết", command=lambda: self.outline_detail(True))
        self._btn(g, col, side="top", icon="hide_detail", text="Ẩn chi tiết", command=lambda: self.outline_detail(False))
        g = rb.group(page, "Lấy dữ liệu")
        self._btn(g, large=True, icon="import_text", text="Từ Văn\nbản/CSV", command=self.import_text,
                  tooltip="Từ Văn bản/CSV\nNhập tệp .csv/.txt, tự chọn bảng mã (UTF-8, Windows-1258, TCVN3, VNI...)\n"
                          "và dấu phân cách, có xem trước")
        g = rb.group(page, "Công cụ dữ liệu")
        self._btn(g, large=True, icon="text_columns", text="Văn bản\nthành Cột", command=self.text_to_columns,
                  tooltip="Văn bản thành Cột\nTách nội dung một cột thành nhiều cột theo dấu phẩy, dấu cách...")
        self._btn(g, large=True, icon="remove_dup", text="Loại bỏ\ntrùng lặp", command=self.remove_duplicates)
        self._btn(g, large=True, icon="validation", text="Xác thực\ndữ liệu", command=self.data_validation_dialog,
                  menu=self._make_menu([("Xác thực dữ liệu...", self.data_validation_dialog),
                                        ("Xóa xác thực dữ liệu", self.clear_validation)]), split=True,
                  tooltip="Xác thực dữ liệu\nGiới hạn những gì được nhập vào ô, ví dụ danh sách thả xuống")

    def _build_review_tab(self, page):
        rb = self.ribbon
        g = rb.group(page, "Ghi chú")
        self._btn(g, large=True, icon="note", text="Ghi chú\nmới", command=self.edit_note,
                  tooltip="Thêm hoặc sửa ghi chú của ô (Shift+F2)")
        col = g.column()
        self._btn(g, col, side="top", icon="note_delete", text="Xóa ghi chú", command=self.delete_notes)
        self._btn(g, col, side="top", icon="note_prev", text="Ghi chú trước", command=lambda: self.goto_note(-1))
        self._btn(g, col, side="top", icon="note_next", text="Ghi chú tiếp", command=lambda: self.goto_note(1))
        self.btn_show_notes = self._btn(g, large=True, icon="note", text="Hiện tất cả\nghi chú",
                                        command=self.toggle_show_notes, tooltip="Hiện/ẩn mọi ghi chú trên trang tính")
        g = rb.group(page, "Bảo vệ")
        self.btn_protect = self._btn(g, large=True, icon="protect", text="Bảo vệ\nTrang tính", label="Bảo vệ Trang tính",
                                     command=self.toggle_protection,
                                     tooltip="Bảo vệ Trang tính\nKhông cho sửa các ô bị khóa (có thể đặt mật khẩu).\n"
                                             "Bỏ khóa ô trước bằng Định dạng ô → Bảo vệ.")
        self._btn(g, large=True, icon="lock_cell", text="Khóa\nô", command=self.toggle_lock_cells,
                  tooltip="Khóa ô\nBật/tắt khóa các ô đang chọn (có tác dụng khi trang tính được bảo vệ)")

    def _build_page_layout_tab(self, page):
        rb = self.ribbon

        def cm(inches):
            return f"{inches * 2.54:.2f}".replace(".", ",")
        margins = [(name, lambda m=m: self.update_page_setup({"margins": m}),
                    f"Trên {cm(m[0])}  Dưới {cm(m[1])}  Trái {cm(m[2])}  Phải {cm(m[3])} cm", "margins")
                   for name, m in (("Thường", printing.MARGINS["normal"]), ("Rộng", printing.MARGINS["wide"]),
                                   ("Hẹp", printing.MARGINS["narrow"]))]
        g = rb.group(page, "Thiết lập Trang", launcher=self.page_setup_dialog, launcher_tip="Thiết lập Trang")
        self._btn(g, large=True, icon="margins", text="Lề\n", menu=self._make_menu(
            margins + [None, ("Lề tùy chỉnh...", lambda: self.page_setup_dialog("margins"), "", "page_setup")]),
            tooltip="Lề\nChọn khoảng cách từ mép giấy tới nội dung")
        self._btn(g, large=True, icon="page_orient", text="Hướng\n", menu=self._make_menu([
            ("Dọc", lambda: self.update_page_setup({"orientation": "portrait"}), "", "portrait"),
            ("Ngang", lambda: self.update_page_setup({"orientation": "landscape"}), "", "landscape")]),
            tooltip="Hướng\nTrang dọc hay trang ngang")
        self._btn(g, large=True, icon="paper_size", text="Khổ\ngiấy", menu=self._make_menu(
            [(label, lambda k=key: self.update_page_setup({"paper": k}), "", "paper_size")
             for key, label in dialogs.PAPER_LABELS.items()] +
            [None, ("Khổ giấy khác...", lambda: self.page_setup_dialog("page"), "", "page_setup")]),
            tooltip="Khổ giấy\nChọn cỡ giấy in")
        self._btn(g, large=True, icon="print_area", text="Vùng\nin", menu=self._make_menu([
            ("Đặt vùng in", self.set_print_area, "", "print_area"),
            ("Xóa vùng in", self.clear_print_area, "", "clear")]),
            tooltip="Vùng in\nChỉ in vùng đang chọn của trang tính")
        self._btn(g, large=True, icon="page_break", text="Ngắt\ntrang", menu=self._make_menu([
            ("Chèn ngắt trang", self.insert_page_break, "", "page_break"),
            ("Loại bỏ ngắt trang", self.remove_page_break, "", "delete_x"),
            ("Đặt lại tất cả ngắt trang", self.reset_page_breaks, "", "refresh")]),
            tooltip="Ngắt trang\nBắt đầu trang mới phía trên / bên trái ô đang chọn")
        self._btn(g, large=True, icon="print_titles", text="In tiêu\nđề", command=lambda: self.page_setup_dialog("sheet"),
                  tooltip="In tiêu đề\nChọn các hàng lặp lại ở đầu mỗi trang in")
        g = rb.group(page, "Co giãn cho vừa", launcher=lambda: self.page_setup_dialog("page"),
                     launcher_tip="Tùy chọn co giãn")
        scales = [(f"{n}%", lambda n=n: self.update_page_setup({"fit": None, "scale": n}))
                  for n in (50, 75, 85, 100, 125, 150)]
        self._btn(g, large=True, icon="fit_page", text="Co\ngiãn", menu=self._make_menu([
            ("Không co giãn", lambda: self.update_page_setup({"fit": None, "scale": 100}), "", "zoom100"),
            ("Vừa tất cả các cột trên một trang", lambda: self.update_page_setup({"fit": "width"}), "", "fit_page"),
            ("Vừa trang tính trên một trang", lambda: self.update_page_setup({"fit": "page"}), "", "fit_page"),
            None, ("Tỷ lệ", scales), ("Tùy chọn co giãn...", lambda: self.page_setup_dialog("page"), "", "page_setup")]),
            tooltip="Co giãn\nThu nhỏ bản in để vừa số trang mong muốn")
        g = rb.group(page, "Tùy chọn Trang tính")
        col = g.column()
        self.print_grid_var = tk.BooleanVar(self.root, value=False)
        ttk.Checkbutton(col, text="In đường lưới", variable=self.print_grid_var, command=self.toggle_print_gridlines,
                        style="Ribbon.TCheckbutton").pack(side="top", anchor="w", pady=1)
        self.commands.append(("In đường lưới", lambda: (self.print_grid_var.set(not self.print_grid_var.get()),
                                                       self.toggle_print_gridlines())))
        g = rb.group(page, "In")
        self._btn(g, large=True, icon="print_preview", text="Xem trước\nkhi in", command=self.print_preview,
                  tooltip="Xem trước khi in và In (Ctrl+P)")
        self._btn(g, large=True, icon="print", text="In\n", command=self.print_now,
                  tooltip="In nhanh\nGửi trang tính tới máy in mặc định")
        self._btn(g, large=True, icon="pdf", text="Xuất\nPDF", command=self.export_pdf,
                  tooltip="Xuất PDF\nLưu trang tính thành tệp PDF")

    def _build_view_tab(self, page):
        rb = self.ribbon
        g = rb.group(page, "Hiển thị")
        col = g.column()
        self.grid_var = tk.BooleanVar(self.root, value=True)
        self.headings_var = tk.BooleanVar(self.root, value=True)
        self.fxbar_var = tk.BooleanVar(self.root, value=True)
        for text, var, command in (("Đường lưới", self.grid_var, self.toggle_grid),
                                   ("Tiêu đề hàng/cột", self.headings_var, self.toggle_headings),
                                   ("Thanh công thức", self.fxbar_var, self.toggle_formula_bar)):
            ttk.Checkbutton(col, text=text, variable=var, command=command, style="Ribbon.TCheckbutton").pack(
                side="top", anchor="w", pady=1)
            self.commands.append((text, lambda v=var, c=command: (v.set(not v.get()), c())))
        g = rb.group(page, "Thu phóng")
        self._btn(g, large=True, icon="zoom", text="Thu\nphóng", command=self.zoom_dialog)
        self._btn(g, large=True, icon="zoom100", text="100%\n", command=lambda: self.set_zoom(1.0))
        self._btn(g, large=True, icon="zoom_sel", text="Phóng theo\nvùng chọn", command=self.zoom_to_selection)
        g = rb.group(page, "Cửa sổ")
        freeze_menu = self._make_menu([
            ("Cố định ngăn", lambda: self.freeze_panes(None if self.sh.freeze else "cell")),
            ("Cố định hàng trên cùng", lambda: self.freeze_panes("row")),
            ("Cố định cột đầu tiên", lambda: self.freeze_panes("col"))])
        # like Excel, the first item turns into "Unfreeze Panes" while panes are frozen
        freeze_menu.configure(postcommand=lambda: freeze_menu.entryconfigure(
            0, label="Bỏ cố định ngăn" if self.sh.freeze else "Cố định ngăn"))
        self.commands.append(("Bỏ cố định ngăn", lambda: self.freeze_panes(None)))
        self._btn(g, large=True, icon="freeze", text="Cố định\nngăn", menu=freeze_menu,
                  tooltip="Cố định ngăn\nGiữ hàng và cột tiêu đề luôn hiển thị khi cuộn")

    def _build_help_tab(self, page):
        g = self.ribbon.group(page, "Trợ giúp")
        self._btn(g, large=True, icon="help", text="Hướng dẫn\n", command=self.show_help, tooltip="Hướng dẫn (F1)")
        self._btn(g, large=True, icon="info", text="Giới\nthiệu", command=self.show_about)

    def _border_menu(self):
        menu = tk.Menu(self.root, tearoff=0)
        for item in BORDER_MENU:
            if item is None:
                menu.add_separator()
                continue
            label, kind = item
            menu.add_command(label=label, command=lambda k=kind: self.apply_border(k),
                             **self._menu_image(label, f"border_{kind}"))
            self.commands.append((label, lambda k=kind: self.apply_border(k)))
        menu.add_separator()
        menu.add_cascade(label="Màu đường kẻ", menu=self._palette_menu(
            menu, lambda c: setattr(self, "border_color", c or "#000000"), "Tự động"),
            **self._menu_image("", "line_color"))
        styles = tk.Menu(menu, tearoff=0)
        self.border_style_var = tk.StringVar(self.root, value="thin")
        for key, name in dialogs.BORDER_NAMES:
            styles.add_radiobutton(label=name, value=key, variable=self.border_style_var,
                                   command=lambda: setattr(self, "border_style", self.border_style_var.get()),
                                   **self._menu_image("", f"line_{key}"))
        menu.add_cascade(label="Kiểu đường kẻ", menu=styles, **self._menu_image("", "line_medium"))
        menu.add_separator()
        menu.add_command(label="Thêm viền...", command=lambda: self.format_cells_dialog("border"),
                         **self._menu_image("Thêm viền"))
        return menu

    def _orientation_menu(self):
        items = [("Xoay ngược chiều kim đồng hồ", 45), ("Xoay theo chiều kim đồng hồ", -45), ("Văn bản dọc", 255),
                 ("Xoay văn bản lên", 90), ("Xoay văn bản xuống", -90)]
        return self._make_menu([(label, lambda a=angle: self.set_rotation(a)) for label, angle in items] +
                               [None, ("Định dạng căn chỉnh ô...", lambda: self.format_cells_dialog("align"))])

    def _cond_menu(self):
        highlight = [("Lớn hơn...", lambda: self.cf_rule("cell", ">")), ("Nhỏ hơn...", lambda: self.cf_rule("cell", "<")),
                     ("Nằm giữa...", lambda: self.cf_rule("cell", "between")), ("Bằng...", lambda: self.cf_rule("cell", "=")),
                     ("Văn bản có chứa...", lambda: self.cf_rule("text")), ("Giá trị trùng lặp...", lambda: self.cf_rule("dup"))]
        top = [("10 mục đầu...", lambda: self.cf_rule("top")), ("10% đầu...", lambda: self.cf_rule("top", percent=True)),
               ("10 mục cuối...", lambda: self.cf_rule("bottom")),
               ("10% cuối...", lambda: self.cf_rule("bottom", percent=True)), None,
               ("Trên trung bình...", lambda: self.cf_rule("above")), ("Dưới trung bình...", lambda: self.cf_rule("below"))]
        bars = [(f"Thanh dữ liệu {name.lower()}", lambda c=color: self.cf_add({"type": "databar", "color": c}), "",
                 f"cf_databar@{color}") for name, color in DATA_BARS]
        scales = [(f"Thang màu {name}", lambda cs=colors: self.cf_add({"type": "scale", "colors": list(cs)}), "",
                   f"cf_scale@{','.join(colors)}") for name, colors in COLOR_SCALES]
        return self._make_menu([
            ("Quy tắc tô sáng ô", highlight), ("Quy tắc trên/dưới", top), ("Thanh dữ liệu", bars),
            ("Thang màu", scales), None,
            ("Xóa quy tắc", [("Xóa quy tắc khỏi ô đã chọn", self.cf_clear_selection),
                             ("Xóa quy tắc khỏi toàn bộ trang tính", self.cf_clear_sheet)]),
            ("Quản lý quy tắc...", self.cf_manage)])

    def _sort_menu(self):
        return self._make_menu([("Sắp xếp từ A đến Z", lambda: self.sort_selection(False)),
                                ("Sắp xếp từ Z đến A", lambda: self.sort_selection(True)),
                                ("Sắp xếp tùy chỉnh...", self.custom_sort), None,
                                ("Lọc", self.toggle_filter, "Ctrl+Shift+L"), ("Xóa bộ lọc", self.clear_filter),
                                ("Áp dụng lại bộ lọc", self.reapply_filter)])

    def _fit_ribbon(self, event=None):
        """Collapse parts of the Home tab when the window is too narrow (like Excel)."""
        page = self.ribbon.pages.get("Trang chủ")
        if page is None or self.ribbon.current != "Trang chủ":
            return
        avail = self.ribbon.body.winfo_width() - self.ui.px(8)
        while self._ribbon_level > 0 and self._ribbon_widths.get(self._ribbon_level - 1, 10 ** 6) <= avail:
            self._set_ribbon_level(self._ribbon_level - 1)
        while self._ribbon_level < 4:
            page.update_idletasks()
            self._ribbon_widths[self._ribbon_level] = page.winfo_reqwidth()
            if self._ribbon_widths[self._ribbon_level] <= avail:
                break
            self._set_ribbon_level(self._ribbon_level + 1)

    def _set_ribbon_level(self, level):
        self._ribbon_level = level
        if level >= 1:
            self.style_gallery.pack_forget()
            self.btn_styles_compact.pack(side="left", fill="y", padx=1)
        else:
            self.btn_styles_compact.pack_forget()
            self.style_gallery.pack(side="left", fill="y", padx=(4, 0))
        for min_level, buttons in self._collapsible.items():
            for btn, text in buttons:
                btn.set_text("" if level >= min_level else text)
        self.font_box.config(width=9 if level >= 4 else 15)
        self.numfmt_box.config(width=9 if level >= 4 else 14)

    def _build_formula_bar(self):
        bar = self.formula_bar = tk.Frame(self.root, bg="#ffffff", highlightthickness=0)
        bar.pack(side="top", fill="x")
        inner = tk.Frame(bar, bg="#ffffff")
        inner.pack(fill="x", padx=6, pady=4)
        self.name_box = ttk.Combobox(inner, width=12, justify="left", font=self.ui.font,
                                     postcommand=lambda: self.name_box.configure(
                                         values=[d["name"] for d in self.wb.sorted_names()]))
        self.name_box.pack(side="left")
        self.name_box.bind("<Return>", self._on_name_box)
        self.name_box.bind("<<ComboboxSelected>>", self._on_name_box)
        self.name_box.bind("<Escape>", lambda e: self._leave_name_box())
        self.name_box.bind("<FocusIn>", lambda e: self.name_box.select_range(0, "end"))
        Tooltip(self.name_box, "Hộp tên: gõ địa chỉ (B5, A1:C9, Sheet2!A1) hoặc tên vùng rồi Enter.\n"
                               "Gõ một tên mới để đặt tên cho vùng đang chọn.", self.ui.small_font)
        tk.Label(inner, text="⋮", bg="#ffffff", fg="#a0a0a0", font=self.ui.font).pack(side="left", padx=(6, 2))
        self.btn_x = tk.Label(inner, text="✕", bg="#ffffff", fg="#c0c0c0", font=self.ui.font, padx=5, cursor="hand2")
        self.btn_ok = tk.Label(inner, text="✓", bg="#ffffff", fg="#c0c0c0", font=self.ui.font, padx=5, cursor="hand2")
        fx = tk.Label(inner, text="fx", bg="#ffffff", fg="#555555", padx=6, cursor="hand2",
                      font=("Cambria", self.ui.font.actual("size"), "italic"))
        for widget in (self.btn_x, self.btn_ok, fx):
            widget.pack(side="left")
        self.btn_x.bind("<ButtonRelease-1>", lambda e: self.cancel_edit())
        self.btn_ok.bind("<ButtonRelease-1>", lambda e: self.commit_edit())
        fx.bind("<ButtonRelease-1>", lambda e: self.insert_function_dialog())
        Tooltip(self.btn_x, "Hủy", self.ui.small_font)
        Tooltip(self.btn_ok, "Nhập", self.ui.small_font)
        Tooltip(fx, "Chèn hàm", self.ui.small_font)
        self.fx_var = tk.StringVar(self.root)
        self.fx = ttk.Entry(inner, textvariable=self.fx_var, font=self.ui.font)
        self.fx.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.fx.bind("<FocusIn>", self._on_fx_focus)
        self.fx.bind("<Return>", lambda e: self._editor_commit(e, (1, 0)))
        self.fx.bind("<KP_Enter>", lambda e: self._editor_commit(e, (1, 0)))
        self.fx.bind("<Tab>", lambda e: self._editor_commit(e, (0, 1)))
        self.fx.bind("<Control-Return>", lambda e: self._key_action(lambda: self.commit_edit(fill=True)))
        self.fx.bind("<Escape>", lambda e: self._key_action(self.cancel_edit))
        self.fx.bind("<Alt-Return>", lambda e: self._key_action(lambda: self.fx.insert("insert", "\n")))
        self.fx.bind("<Key>", self._reset_point, add="+")
        self._formula_assist(self.fx)
        tk.Frame(self.root, bg=HEADER_LINE, height=1).pack(side="top", fill="x")

    def _build_status_bar(self):
        bar = tk.Frame(self.root, bg="#f3f3f3")
        bar.pack(side="bottom", fill="x")
        self.mode_label = tk.Label(bar, text="Sẵn sàng", bg="#f3f3f3", fg="#444444", font=self.ui.small_font)
        self.mode_label.pack(side="left", padx=8, pady=2)
        self.zoom_label = tk.Label(bar, text="100%", bg="#f3f3f3", fg="#444444", font=self.ui.small_font,
                                   width=5, cursor="hand2")
        self.zoom_label.pack(side="right", padx=(2, 8))
        self.zoom_label.bind("<ButtonRelease-1>", lambda e: self.zoom_dialog())
        plus = tk.Label(bar, text="+", bg="#f3f3f3", fg="#444444", font=self.ui.font, cursor="hand2")
        plus.pack(side="right")
        self.zoom_var = tk.DoubleVar(self.root, value=100)
        scale = ttk.Scale(bar, from_=25, to=300, variable=self.zoom_var, length=self.ui.px(110),
                          command=lambda v: None if self._zoom_busy else self.set_zoom(float(v) / 100))
        scale.pack(side="right", padx=2)
        minus = tk.Label(bar, text="−", bg="#f3f3f3", fg="#444444", font=self.ui.font, cursor="hand2")
        minus.pack(side="right")
        plus.bind("<ButtonRelease-1>", lambda e: self.set_zoom(self.zoom + 0.1))
        minus.bind("<ButtonRelease-1>", lambda e: self.set_zoom(self.zoom - 0.1))
        self.stats_label = tk.Label(bar, text="", bg="#f3f3f3", fg="#444444", font=self.ui.small_font)
        self.stats_label.pack(side="right", padx=16)

    def _build_sheet_bar(self):
        bar = tk.Frame(self.root, bg=SHEETBAR_BG)
        bar.pack(side="bottom", fill="x")
        self.tab_frame = tk.Frame(bar, bg=SHEETBAR_BG)
        self.tab_frame.pack(side="left", padx=(self.ui.px(24), 0))
        plus = RibbonButton(bar, self.ui, icon="sheet", command=self.add_sheet, tooltip="Trang tính mới (Shift+F11)",
                            bg=SHEETBAR_BG)
        plus.pack(side="left", padx=4)
        self.hbar = ttk.Scrollbar(bar, orient="horizontal", command=self.on_hscroll)
        self.hbar.pack(side="left", fill="x", expand=True, padx=(16, 0))

    def _build_grid(self):
        frame = ttk.Frame(self.root)
        frame.pack(side="top", fill="both", expand=True)
        self.canvas = tk.Canvas(frame, bg="white", highlightthickness=0, bd=0, takefocus=1)
        self.vbar = ttk.Scrollbar(frame, orient="vertical", command=self.on_vscroll)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.pivot_pane = PivotPane(frame, self.ui, self.menu_icons, self._pivot_pane_change, self._pivot_pane_action)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        cv = self.canvas
        cv.bind("<Configure>", lambda e: self.refresh_view())
        cv.bind("<Button-1>", self.on_press)
        cv.bind("<B1-Motion>", self.on_drag)
        cv.bind("<ButtonRelease-1>", self.on_release)
        cv.bind("<Double-Button-1>", self.on_double)
        cv.bind("<Button-3>", self.on_context)
        cv.bind("<Motion>", self.on_motion)
        cv.bind("<Leave>", lambda e: (self._set_hover_note(None), self._hide_link_tip()))
        self.root.bind_all("<MouseWheel>", self.on_wheel)
        self.root.bind_all("<Shift-MouseWheel>", lambda e: self.on_wheel(e, horizontal=True))

    def _build_context_menus(self):
        paste_options = [("Dán", self.paste), ("Dán giá trị", lambda: self.paste(mode="values")),
                         ("Dán công thức", lambda: self.paste(mode="formulas")),
                         ("Dán định dạng", lambda: self.paste(mode="formats")),
                         ("Dán chuyển vị", lambda: self.paste(transpose=True))]
        self.cell_menu = self._make_menu([
            ("Cắt", self.cut, "Ctrl+X"), ("Sao chép", self.copy, "Ctrl+C"), ("Tùy chọn dán", paste_options),
            ("Dán đặc biệt...", self.paste_special), None,
            ("Chèn...", self.insert_cells_dialog, "Ctrl++"), ("Xóa...", self.delete_cells_dialog, "Ctrl+-"),
            ("Xóa nội dung", self.clear_contents, "Delete"), None,
            ("Chèn/Sửa ghi chú...", self.edit_note, "Shift+F2"), ("Xóa ghi chú", self.delete_notes), None,
            ("Liên kết...", self.insert_hyperlink, "Ctrl+K"), ("Mở liên kết", self.open_hyperlink),
            ("Loại bỏ liên kết", self.remove_hyperlinks), None,
            ("Lọc", [("Lọc theo giá trị của ô đã chọn", self.filter_by_value),
                     ("Lọc theo màu của ô đã chọn", lambda: self.filter_by_color("bg")),
                     ("Lọc theo màu phông của ô đã chọn", lambda: self.filter_by_color("fg")), None,
                     ("Bật/tắt bộ lọc", self.toggle_filter)]),
            ("Sắp xếp", [("Sắp xếp từ A đến Z", lambda: self.sort_selection(False)),
                         ("Sắp xếp từ Z đến A", lambda: self.sort_selection(True)),
                         ("Sắp xếp tùy chỉnh...", self.custom_sort)]), None,
            ("Định dạng ô...", lambda: self.format_cells_dialog(), "Ctrl+1")])
        self.chart_menu = self._make_menu([
            ("Đổi kiểu biểu đồ", [(f"Biểu đồ {label.lower()}", lambda k=kind: self.set_chart_type(k))
                                  for kind, label in dialogs.CHART_KINDS]),
            ("Dữ liệu và tiêu đề biểu đồ...", self.edit_chart), None, ("Xóa biểu đồ", self.delete_chart)])
        self.picture_menu = self._make_menu([
            ("Đặt lại kích thước ảnh", self.reset_picture_size, "", "picture"), ("Thay ảnh...", self.change_picture),
            None, ("Xóa ảnh", self.delete_chart, "Delete", "delete_x")])
        common = [("Cắt", self.cut), ("Sao chép", self.copy), ("Dán", self.paste), None]
        self.row_menu = self._make_menu(common + [
            ("Chèn", self.insert_rows, "", "insert_row"), ("Xóa", self.delete_rows, "", "delete_row"),
            ("Xóa nội dung", self.clear_contents), None,
            ("Định dạng ô...", lambda: self.format_cells_dialog()), ("Chiều cao hàng...", self.ask_row_height),
            ("Ẩn", self.hide_rows, "", "hide_row"), ("Bỏ ẩn", self.unhide_rows, "", "unhide_row")])
        self.col_menu = self._make_menu(common + [
            ("Chèn", self.insert_cols, "", "insert_col"), ("Xóa", self.delete_cols, "", "delete_col"),
            ("Xóa nội dung", self.clear_contents), None,
            ("Định dạng ô...", lambda: self.format_cells_dialog()), ("Độ rộng cột...", self.ask_column_width),
            ("Ẩn", self.hide_cols, "", "hide_col"), ("Bỏ ẩn", self.unhide_cols, "", "unhide_col")])

    def _bind_keys(self):
        cv = self.canvas
        cv.bind("<Key>", self.on_key)
        canvas_keys = {
            "c": self.copy, "x": self.cut, "v": self.paste, "z": self.undo, "y": self.redo, "a": self.select_all,
            "b": lambda: self.toggle_style("bold"), "i": lambda: self.toggle_style("italic"),
            "u": lambda: self.toggle_style("underline"), "d": self.fill_down, "r": self.fill_right,
        }
        for key, func in canvas_keys.items():
            for k in (key, key.upper()):
                cv.bind(f"<Control-{k}>", lambda e, f=func: self._key_action(f))
        extra = {"<Alt-equal>": self.autosum, "<Control-L>": self.toggle_filter,
                 "<Control-Key-5>": lambda: self.toggle_style("strike"), "<Control-Key-9>": self.hide_rows,
                 "<Control-Key-0>": self.hide_cols, "<Control-grave>": self.toggle_show_formulas,
                 "<Control-semicolon>": lambda: self.insert_now(False), "<Control-plus>": self.insert_cells_dialog,
                 "<Control-KP_Add>": self.insert_cells_dialog, "<Control-minus>": self.delete_cells_dialog,
                 "<Control-KP_Subtract>": self.delete_cells_dialog,
                 "<Control-ampersand>": lambda: self.apply_border("outside"),
                 "<Control-underscore>": lambda: self.apply_border("none"), "<F9>": self.recalculate,
                 "<Shift-F3>": self.insert_function_dialog, "<Shift-F2>": self.edit_note,
                 "<Alt-Down>": self.open_validation_list,
                 "<Alt-Shift-Right>": lambda: self.group_selection(1), "<Alt-Shift-Left>": lambda: self.group_selection(-1)}
        for seq, func in extra.items():
            cv.bind(seq, lambda e, f=func: self._key_action(f))
        global_keys = {"n": self.new_file, "o": self.open_file, "s": self.save_file, "f": self.open_find,
                       "h": lambda: self.open_find(True), "g": self.focus_name_box, "p": self.print_preview,
                       "k": self.insert_hyperlink}
        for key, func in global_keys.items():
            for k in (key, key.upper()):
                self.root.bind_all(f"<Control-{k}>", lambda e, f=func: self._key_action(f))
        self.root.bind_all("<Control-Key-1>", lambda e: self._key_action(self.format_cells_dialog))
        self.root.bind_all("<F12>", lambda e: self._key_action(self.save_as))
        self.root.bind_all("<Control-F2>", lambda e: self._key_action(self.print_preview))
        self.root.bind_all("<Control-F3>", lambda e: self._key_action(self.name_manager))
        self.root.bind_all("<Control-Shift-F3>", lambda e: self._key_action(self.create_names_from_selection))
        self.root.bind_all("<F1>", lambda e: self._key_action(self.show_help))
        self.root.bind_all("<Shift-F11>", lambda e: self._key_action(self.add_sheet))

    @staticmethod
    def _key_action(func):
        func()
        return "break"

    # ------------------------------------------------------------------ geometry

    def col_w(self, c, sh=None):
        sh = sh or self.sh
        if c in sh.hidden_cols:
            return 0
        return round(sh.col_widths.get(c, fe.DEFAULT_COL_WIDTH) * self.px)

    def row_h(self, r, sh=None):
        sh = sh or self.sh
        if sh.row_hidden(r):
            return 0
        return round(sh.row_heights.get(r, fe.DEFAULT_ROW_HEIGHT) * self.px)

    def frozen(self):
        """(r0, r1, c0, c1): rows r0..r1-1 and columns c0..c1-1 are frozen (empty spans when not)."""
        return self.sh.freeze or (0, 0, 0, 0)

    def pane_origin(self):
        """Canvas (x, y) where the scrolling part of the grid starts, right of / below the frozen panes."""
        r0, r1, c0, c1 = self.frozen()
        return (self.header_w + sum(self.col_w(k) for k in range(c0, c1)),
                self.header_h + sum(self.row_h(k) for k in range(r0, r1)))

    def _clamp_view(self):
        """Keep the scrolled area out of the frozen rows and columns."""
        _, r1, _, c1 = self.frozen()
        self.top, self.left = max(self.top, r1), max(self.left, c1)

    def col_x(self, c):
        """Canvas x of the left edge of column c. With frozen columns, edges of columns hidden behind
        them are clamped, so a range's box covers just its visible part."""
        c0, c1 = self.frozen()[2:]
        if c1 > c0:
            if c <= c1:
                return self.header_w + sum(self.col_w(k) for k in range(c0, c)) - (4 if c < c0 else 0)
            x0 = self.pane_origin()[0]
            return x0 if c <= self.left else x0 + sum(self.col_w(k) for k in range(self.left, c))
        x = self.header_w
        if c >= self.left:
            for k in range(self.left, c):
                x += self.col_w(k)
        else:
            for k in range(c, self.left):
                x -= self.col_w(k)
        return x

    def row_y(self, r):
        """Canvas y of the top of row r (rows far off screen are estimated; with frozen rows, edges
        hidden behind them are clamped as in col_x)."""
        r0, r1 = self.frozen()[:2]
        y = self.header_h
        if r1 > r0:
            if r <= r1:
                return y + sum(self.row_h(k) for k in range(r0, r)) - (4 if r < r0 else 0)
            y = self.pane_origin()[1]
            if r <= self.top:
                return y
        far = self.canvas.winfo_height() + 4000
        if r >= self.top:
            for k in range(self.top, r):
                y += self.row_h(k)
                if y > far:
                    return y + (r - k - 1) * self.default_row_px
            return y
        for k in range(self.top - 1, r - 1, -1):
            y -= self.row_h(k)
            if y < -far:
                return y - (k - r) * self.default_row_px
        return y

    def _rect_bbox(self, r1, c1, r2, c2):
        return self.col_x(c1), self.row_y(r1), self.col_x(c2 + 1), self.row_y(r2 + 1)

    def _clip_box(self, box, W, H):
        x0, y0, x1, y1 = box
        if x1 <= self.header_w or y1 <= self.header_h or x0 >= W or y0 >= H or x1 <= x0 or y1 <= y0:
            return None
        return max(x0, self.header_w - 4), max(y0, self.header_h - 4), min(x1, W + 4), min(y1, H + 4)

    def _visible_rows(self):
        """Rows a page of scrolling moves (the scrolled part of the window, less one)."""
        return max(1, self._scroll_rows - 1)

    def row_at(self, y):
        r0, r1 = self.frozen()[:2]
        if y < self.header_h:
            return r0 if r1 > r0 else max(0, self.top - 1)
        edge = self.header_h
        for r in range(r0, r1):
            edge += self.row_h(r)
            if y < edge:
                return r
        r = self.top
        while r < MAX_ROWS - 1:
            edge += self.row_h(r)
            if y < edge:
                return r
            r += 1
        return MAX_ROWS - 1

    def col_at(self, x):
        c0, c1 = self.frozen()[2:]
        if x < self.header_w:
            return c0 if c1 > c0 else max(0, self.left - 1)
        edge = self.header_w
        for c in range(c0, c1):
            edge += self.col_w(c)
            if x < edge:
                return c
        c = self.left
        while c < MAX_COLS - 1:
            edge += self.col_w(c)
            if x < edge:
                return c
            c += 1
        return MAX_COLS - 1

    def hit(self, x, y):
        """('corner' | 'colhead' | 'rowhead' | 'cell', row, col) under a canvas point."""
        if y < self.header_h and x < self.header_w:
            return "corner", None, None
        if y < self.header_h:
            return "colhead", None, self.col_at(x)
        if x < self.header_w:
            return "rowhead", self.row_at(y), None
        return "cell", self.row_at(y), self.col_at(x)

    def _border_hit(self, x, y):
        """('col', c) / ('row', r) when the pointer is on a header border used for resizing."""
        tolerance = 3 * self.scale
        if self.outline_h <= y < self.header_h and x >= self.header_w:
            for c, _, x1 in self.vis_cols:
                if abs(x - x1) <= tolerance:
                    return "col", c
        if self.outline_w <= x < self.header_w and y >= self.header_h:
            for r, _, y1 in self.vis_rows:
                if abs(y - y1) <= tolerance:
                    return "row", r
        return None

    def _on_fill_handle(self, x, y):
        _, _, x1, y1 = self._rect_bbox(*self.sel_rect())
        tolerance = 5 * self.scale
        return x1 > self.header_w and y1 > self.header_h and abs(x - x1) <= tolerance and abs(y - y1) <= tolerance

    def _cell_box(self, r, c):
        m = self.sh.merge_at(r, c)
        return m if m else (r, c, r, c)

    def _expand(self, r1, c1, r2, c2):
        """Grow a rectangle until it covers every merged block it touches."""
        merges = self.sh.merges
        grew = bool(merges)
        while grew:
            grew = False
            for m in merges:
                if m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1 and \
                        (m[0] < r1 or m[1] < c1 or m[2] > r2 or m[3] > c2):
                    r1, c1, r2, c2 = min(r1, m[0]), min(c1, m[1]), max(r2, m[2]), max(c2, m[3])
                    grew = True
        return r1, c1, r2, c2

    def sel_rect(self):
        (ra, ca), (rb, cb) = self.cur, self.ext
        return self._expand(min(ra, rb), min(ca, cb), max(ra, rb), max(ca, cb))

    def _sel_state(self):
        return self.sheet, self.cur, self.ext

    @staticmethod
    def _clamp(r, c):
        return max(0, min(MAX_ROWS - 1, r)), max(0, min(MAX_COLS - 1, c))

    @staticmethod
    def _is_huge(r1, c1, r2, c2):
        """Whole rows/columns or very large selections, which commands clip to the data."""
        full = (r1 == 0 and r2 == MAX_ROWS - 1) or (c1 == 0 and c2 == MAX_COLS - 1)
        return full or (r2 - r1 + 1) * (c2 - c1 + 1) > BIG_AREA

    def ensure_visible(self, r, c):
        H, W = self.canvas.winfo_height(), self.canvas.winfo_width()
        self._clamp_view()
        x0, y0 = self.pane_origin()
        _, r1, _, c1 = self.frozen()
        if r < r1:  # frozen rows are always on screen
            pass
        elif r < self.top:
            self.top = r
        elif self.row_y(r + 1) > H:
            avail, acc, t = H - y0, 0, r
            while t >= r1:
                acc += self.row_h(t)
                if acc > avail:
                    break
                t -= 1
            self.top = min(r, t + 1)
        if c < c1:
            pass
        elif c < self.left:
            self.left = c
        elif self.col_x(c + 1) > W:
            avail, acc, t = W - x0, 0, c
            while t >= c1:
                acc += self.col_w(t)
                if acc > avail:
                    break
                t -= 1
            self.left = min(c, t + 1)

    # ------------------------------------------------------------------ drawing

    def redraw(self):
        cv = self.canvas
        cv.delete("all")
        W, H = cv.winfo_width(), cv.winfo_height()
        if W < 20 or H < 20:
            return
        sh = self.sh
        cells, formats = sh.cells, sh.formats
        self._outline_metrics()
        hw, hh, pad = self.header_w, self.header_h, self.pad
        self._clamp_view()
        fr0, fr1, fc0, fc1 = self.frozen()

        def strip(first, frozen, limit, size, origin, end):
            """[(index, start, end)] of the frozen lines, then of the scrolled ones from `first`."""
            out, pos = [], origin
            for i in itertools.chain(frozen, range(first, limit)):
                if pos >= end:
                    break
                n = size(i)
                if n:
                    out.append((i, pos, pos + n))
                pos += n
            return out

        cols = strip(self.left, range(fc0, fc1), MAX_COLS, self.col_w, hw, W)
        rows = strip(self.top, range(fr0, fr1), MAX_ROWS, self.row_h, hh, H)
        self.vis_cols, self.vis_rows = cols, rows
        self._scroll_rows = sum(1 for r, _, _ in rows if r >= self.top)
        self._filter_buttons = []
        self._outline_buttons = []
        r1, c1, r2, c2 = self.sel_rect()
        if not cols or not rows:
            self._draw_headers(cv, W, H, (r1, c1, r2, c2))
            self._draw_outline(cv, W, H, {}, {})
            return
        col_pos = {c: (x0, x1) for c, x0, x1 in cols}
        row_pos = {r: (y0, y1) for r, y0, y1 in rows}
        right, bottom = cols[-1][2], rows[-1][2]
        vr1, vr2, vc1, vc2 = rows[0][0], rows[-1][0], cols[0][0], cols[-1][0]
        active = self._cell_box(*self.cur)
        multi = (r1, c1, r2, c2) != active
        areas = [a for a in self.areas if a[0] <= vr2 and a[2] >= vr1 and a[1] <= vc2 and a[3] >= vc1]

        cf = {}
        if sh.cond_formats:
            for r in row_pos:
                for c in col_pos:
                    style = self.wb.cond_format(self.sheet, r, c)
                    if style:
                        cf[(r, c)] = style

        def fill_of(r, c):
            style = cf.get((r, c))
            fmt = formats.get((r, c))
            bg = (style or {}).get("bg") or (fmt.get("bg") if fmt else None) or "#ffffff"
            inside = active[0] <= r <= active[2] and active[1] <= c <= active[3]
            if multi and r1 <= r <= r2 and c1 <= c <= c2 and not inside:
                return self.tint(bg)
            if areas and not inside and any(a[0] <= r <= a[2] and a[1] <= c <= a[3] for a in areas):
                return self.tint(bg)
            return bg

        def visible(d):
            if len(d) < len(rows) * len(cols):
                return [k for k in d if k[0] in row_pos and k[1] in col_pos]
            return [(r, c) for r in row_pos for c in col_pos if (r, c) in d]

        merges = [m for m in sh.merges if m[0] <= vr2 and m[2] >= vr1 and m[1] <= vc2 and m[3] >= vc1]
        covered = set()
        for m in merges:
            for r in range(max(m[0], vr1), min(m[2], vr2) + 1):
                for c in range(max(m[1], vc1), min(m[3], vc2) + 1):
                    if (r, c) != (m[0], m[1]):
                        covered.add((r, c))

        # backgrounds: selection tint, filled cells, the active cell left clear, data bars
        if multi:
            box = self._clip_box(self._rect_bbox(r1, c1, r2, c2), W, H)
            if box:
                cv.create_rectangle(*box, fill=self.tint("#ffffff"), outline="")
        for area in areas:
            box = self._clip_box(self._rect_bbox(*area), W, H)
            if box:
                cv.create_rectangle(*box, fill=self.tint("#ffffff"), outline="")
        filled = {k for k in visible(formats) if formats[k].get("bg")} | {k for k, s in cf.items() if s.get("bg")}
        for r, c in filled:
            cv.create_rectangle(col_pos[c][0], row_pos[r][0], col_pos[c][1], row_pos[r][1], fill=fill_of(r, c),
                                outline="")
        if multi:
            box = self._clip_box(self._rect_bbox(*active), W, H)
            if box:
                cv.create_rectangle(*box, fill=fill_of(*self.cur), outline="")
        for (r, c), style in cf.items():
            bar = style.get("databar")
            if bar and (r, c) not in covered:
                (x0, x1), (y0, y1) = col_pos[c], row_pos[r]
                width = round((x1 - x0 - 5) * bar[0])
                if width > 0:
                    cv.create_rectangle(x0 + 2, y0 + 2, x0 + 2 + width, y1 - 3, fill=mix(bar[1], "#ffffff", 0.45),
                                        outline=bar[1])
        if self.show_grid:
            for c, x0, x1 in cols:
                cv.create_line(x1 - 1, hh, x1 - 1, bottom, fill=GRID)
            for r, y0, y1 in rows:
                cv.create_line(hw, y1 - 1, right, y1 - 1, fill=GRID)
        for m in merges:  # merged blocks hide their inner grid lines
            box = self._clip_box(self._rect_bbox(*m), W, H)
            if box:
                cv.create_rectangle(box[0], box[1], box[2] - 1, box[3] - 1, fill=fill_of(m[0], m[1]), outline="")

        # cell borders
        for r, c in visible(formats):
            fmt = formats[(r, c)]
            if not any(k in fmt for k in fe.BORDER_KEYS):
                continue
            m = sh.merge_at(r, c)
            (x0, x1), (y0, y1) = col_pos[c], row_pos[r]
            for key in fe.BORDER_KEYS:
                spec = fmt.get(key)
                if not spec:
                    continue
                if m and ((key == "bt" and r != m[0]) or (key == "bb" and r != m[2]) or
                          (key == "bl" and c != m[1]) or (key == "br" and c != m[3])):
                    continue
                self._draw_border(cv, key, spec, x0, y0, x1, y1)

        # filter dropdown buttons (sized first so header text can leave room for them)
        f = sh.filter
        btn_size = 0
        filter_cols = ()
        if f and f["range"][0] in row_pos:
            fy0, fy1 = row_pos[f["range"][0]]
            btn_size = max(8, min(fy1 - fy0 - 4, round(16 * self.px)))
            filter_cols = range(f["range"][1], f["range"][3] + 1)

        def draw_text(r, c, box, style, spill):
            fmt = formats.get((r, c)) or {}
            raw = cells.get((r, c))
            if self.show_formulas and fe.is_formula(raw):
                value, numfmt = raw, None
            else:
                value, numfmt = self.wb.cell_display(self.sheet, r, c)
            if value is None or value == "":
                return
            eff = fmt
            if style and any(k in style for k in ("fg", "bold", "italic")):
                eff = {**fmt, **{k: style[k] for k in ("fg", "bold", "italic") if k in style}}
            font = self.font_for(eff)
            x0, y0, x1, y1 = box
            if f and r == f["range"][0] and c in filter_cols:
                x1 -= btn_size + 4
            indent = (fmt.get("indent") or 0) * round(9 * self.px)
            avail = x1 - x0 - 2 * pad - indent
            is_num = False
            if isinstance(value, fe.CellError):
                text, align = value.code, "center"
            elif isinstance(value, bool):
                text, align = ("TRUE" if value else "FALSE"), "center"
            elif fe.is_number(value):
                text, align, is_num = self.fit_number(value, numfmt, font, avail), "right", True
            else:
                text, align = str(value), "left"
            if fe._COMBINING.search(text):  # "Unicode tổ hợp" (separate tone marks): Tk draws the marks apart
                text = unicodedata.normalize("NFC", text)
            align = fmt.get("align") or align
            valign = fmt.get("valign") or "bottom"
            fg = eff.get("fg") or "#000000"
            ls = font.metrics("linespace")
            rotate = fmt.get("rotate") or 0
            if rotate and rotate != 255:
                text = text.replace("\n", " ")
                tw, a = self.text_width(font, text), math.radians(abs(rotate))
                bw, bh = tw * math.cos(a) + ls * math.sin(a), tw * math.sin(a) + ls * math.cos(a)
                cx = {"left": x0 + pad + indent + bw / 2, "center": (x0 + x1) / 2, "right": x1 - pad - bw / 2}[align]
                cy = {"top": y0 + 2 + bh / 2, "center": (y0 + y1) / 2, "bottom": y1 - 2 - bh / 2}[valign]
                cv.create_text(cx, cy, text=text, angle=rotate, font=font, fill=fg)
                return
            if rotate == 255:
                lines = list(text.replace("\n", ""))[:max(1, (y1 - y0) // ls)]
                align = "center" if align == "left" and not fmt.get("align") else align
            elif fmt.get("wrap") and not is_num:
                lines = self.wrap_lines(text, font, max(4, avail))[:max(1, (y1 - y0 - 2) // ls)]
            else:
                text = text.replace("\r", " ").replace("\n", " ")
                if not is_num:
                    if spill and align == "left" and isinstance(value, str):
                        needed, end, nc = self.text_width(font, text), x1, c + 1
                        while (needed > end - x0 - 2 * pad - indent and nc in col_pos and (r, nc) not in cells
                               and (r, nc) not in covered and not sh.merge_at(r, nc)
                               and not (f and r == f["range"][0] and nc in filter_cols)):
                            if self.show_grid:
                                boundary = col_pos[nc][0] - 1
                                cv.create_line(boundary, y0, boundary, y1 - 1, fill=fill_of(r, nc - 1))
                            end, nc = col_pos[nc][1], nc + 1
                        avail = end - x0 - 2 * pad - indent
                    text = self.clip_text(text, font, avail)
                lines = [text]
            block = len(lines) * ls
            ty = {"top": y0 + 1 + block / 2, "center": (y0 + y1) / 2, "bottom": y1 - 2 - block / 2}[valign]
            if align == "center":
                tx, anchor = (x0 + x1) / 2, "center"
            elif align == "right":
                tx, anchor = x1 - pad - 1 - indent, "e"
            else:
                tx, anchor = x0 + pad + indent, "w"
            item = cv.create_text(tx, ty, text="\n".join(lines), anchor=anchor, justify=align, font=font, fill=fg)
            if fmt.get("underline") == "double":
                bx0, _, bx1, by1 = cv.bbox(item)
                base = by1 - font.metrics("descent")
                for dy in (1, 3):
                    cv.create_line(bx0, base + dy, bx1, base + dy, fill=fg)

        targets = [k for k in visible(cells) if k not in covered]
        seen = set(targets)
        targets += [(m[0], m[1]) for m in merges if (m[0], m[1]) in cells and (m[0], m[1]) not in seen]
        for r, c in targets:
            m = sh.merge_at(r, c)
            if m:
                box = self._rect_bbox(*m)
                if box[2] <= box[0] or box[3] <= box[1]:  # scrolled behind the frozen panes
                    continue
                style = cf.get((r, c)) if (r, c) in cf else self.wb.cond_format(self.sheet, r, c)
            else:
                box = (col_pos[c][0], row_pos[r][0], col_pos[c][1], row_pos[r][1])
                style = cf.get((r, c))
            draw_text(r, c, box, style, spill=not m)

        mark = max(4, round(5 * self.px))
        for r, c in visible(sh.notes):
            m = sh.merge_at(r, c)
            if (r, c) in covered or (m and (r, c) != (m[0], m[1])):
                continue
            if m:
                _, y0, x1, _ = self._rect_bbox(*m)
            else:
                x1, y0 = col_pos[c][1], row_pos[r][0]
            cv.create_polygon(x1 - 1 - mark, y0, x1 - 1, y0, x1 - 1, y0 + mark, fill=NOTE_MARK, outline="")

        if btn_size:
            fy0, fy1 = row_pos[f["range"][0]]
            for c in filter_cols:
                if c not in col_pos:
                    continue
                bx1, by1 = col_pos[c][1] - 3, fy1 - 3
                bx0, by0 = bx1 - btn_size, by1 - btn_size
                cv.create_rectangle(bx0, by0, bx1, by1, fill="#ffffff", outline="#a6a6a6")
                s = btn_size
                if c in f["criteria"]:
                    cv.create_polygon(bx0 + s * .15, by0 + s * .22, bx0 + s * .6, by0 + s * .22, bx0 + s * .44, by0 + s * .48,
                                      bx0 + s * .44, by0 + s * .8, bx0 + s * .32, by0 + s * .86, bx0 + s * .32, by0 + s * .48,
                                      fill="#2b579a", outline="")
                    cv.create_polygon(bx0 + s * .64, by0 + s * .5, bx0 + s * .88, by0 + s * .5, bx0 + s * .76, by0 + s * .68,
                                      fill="#444444", outline="")
                else:
                    mx, my, d = (bx0 + bx1) / 2, (by0 + by1) / 2, s * .22
                    cv.create_polygon(mx - d, my - d / 2, mx + d, my - d / 2, mx, my + d * .8, fill="#444444", outline="")
                self._filter_buttons.append((bx0, by0, bx1, by1, c))

        line = max(2, round(2 * self.px))
        if self.is_editing() and self.edit_cell[0] == self.sheet:
            text = self.edit_var.get() if self.editor is not None else self.fx_var.get()
            for i, (ref_sheet, a1, b1, a2, b2) in enumerate(fe.formula_references(text)):
                if ref_sheet is not None and self.wb.resolve_sheet(ref_sheet) != self.sheet:
                    continue
                if not (0 <= a1 and 0 <= b1 and a2 < MAX_ROWS and b2 < MAX_COLS):
                    continue
                box = self._clip_box(self._rect_bbox(a1, b1, a2, b2), W, H)
                if box:
                    cv.create_rectangle(box[0], box[1], box[2] - 1, box[3] - 1,
                                        outline=REF_COLORS[i % len(REF_COLORS)], width=line)
        if self.clip and self.clip["sheet"] == self.sheet:
            box = self._clip_box(self._rect_bbox(*self.clip["rect"]), W, H)
            if box:
                cv.create_rectangle(box[0] + 1, box[1] + 1, box[2] - 2, box[3] - 2, outline=ACCENT, dash=(5, 3),
                                    width=line)
        for area in areas:
            box = self._clip_box(self._rect_bbox(*area), W, H)
            if box:
                cv.create_rectangle(box[0], box[1], box[2] - 1, box[3] - 1, outline=ACCENT)
        sel_box = self._rect_bbox(r1, c1, r2, c2)
        box = self._clip_box(sel_box, W, H)
        if box:
            cv.create_rectangle(box[0], box[1], box[2] - 1, box[3] - 1, outline=ACCENT, width=line)
            hx, hy = sel_box[2] - 1, sel_box[3] - 1
            corner_shown = not (fr1 <= r2 < self.top or fc1 <= c2 < self.left)
            if hw < hx < W and hh < hy < H and corner_shown and not self.is_editing():
                s = max(3, round(3 * self.px))
                cv.create_rectangle(hx - s, hy - s, hx + s, hy + s, fill=ACCENT, outline="white")
        if self.drag and self.drag.get("mode") == "fill" and self.drag.get("target"):
            box = self._clip_box(self._rect_bbox(*self.drag["target"]), W, H)
            if box:
                cv.create_rectangle(box[0], box[1], box[2] - 1, box[3] - 1, outline="#555555", dash=(4, 3))
        self._draw_charts(cv, W, H)
        if sh.page_setup or self.sheet in self.breaks_shown:
            self._draw_page_breaks(cv, col_pos, row_pos, right, bottom)
        px, py = self.pane_origin()
        if fr1 > fr0:
            cv.create_line(hw, py - 1, right, py - 1, fill=FREEZE_LINE)
        if fc1 > fc0:
            cv.create_line(px - 1, hh, px - 1, bottom, fill=FREEZE_LINE)
        self._draw_headers(cv, W, H, (r1, c1, r2, c2), sel_box)
        self._draw_outline(cv, W, H, row_pos, col_pos)
        self._draw_cell_extras(cv, W, H, col_pos, row_pos)
        self._place_editor()

    def _sheet_x(self, c):
        """Canvas x of column c's left edge, also when scrolled out of view (col_x clamps those)."""
        c0, c1 = self.frozen()[2:]
        if c0 <= c < c1:
            return self.col_x(c)
        x = self.pane_origin()[0]
        if c >= self.left:
            return x + sum(self.col_w(k) for k in range(self.left, c))
        return x - sum(self.col_w(k) for k in range(c, self.left))

    def _sheet_y(self, r):
        """Canvas y of row r's top edge, like _sheet_x (rows far off screen are estimated)."""
        r0, r1 = self.frozen()[:2]
        if r0 <= r < r1:
            return self.row_y(r)
        y, far = self.pane_origin()[1], self.canvas.winfo_height() + 4000
        step, rows = (1, range(self.top, r)) if r >= self.top else (-1, range(self.top - 1, r - 1, -1))
        for n, k in enumerate(rows):
            y += step * self.row_h(k)
            if abs(y) > far:
                return y + step * (len(rows) - n - 1) * self.default_row_px
        return y

    def _chart_box(self, chart):
        r, c, dx, dy = chart["anchor"]
        x, y = self._sheet_x(c) + dx * self.px, self._sheet_y(r) + dy * self.px
        return x, y, x + chart["size"][0] * self.px, y + chart["size"][1] * self.px

    def _draw_charts(self, cv, W, H):
        self._chart_boxes = []
        if self.chart_sel is not None and self.chart_sel >= len(self.sh.charts):
            self.chart_sel = None
        for i, chart in enumerate(self.sh.charts):
            box = self._chart_box(chart)
            if box[2] < self.header_w or box[0] > W or box[3] < self.header_h or box[1] > H:
                continue
            if chart["type"] == fe.PICTURE:
                self._draw_picture(cv, chart, box)
            else:
                labels, series = self.wb.chart_data(self.sheet, chart)
                charts.draw_chart(cv, box, chart["type"], chart.get("title"), labels, series,
                                  self.font_for({"size": 14}), self.font_for({"size": 9}), self.px)
            self._chart_boxes.append((i, box))
            if i == self.chart_sel:
                x0, y0, x1, y1 = box
                cv.create_rectangle(x0, y0, x1, y1, outline="#7a7a7a", width=max(1, round(self.px)))
                s = max(3, round(4 * self.px))
                for hx, hy in self._chart_handles(box).values():
                    cv.create_rectangle(hx - s, hy - s, hx + s, hy + s, fill="#ffffff", outline="#5f5f5f")

    def _draw_picture(self, cv, picture, box):
        x0, y0, x1, y1 = box
        photo = self.pictures.photo(picture["id"], picture["image"], picture.get("format", "png"), x1 - x0, y1 - y0)
        if photo is not None:
            # without Pillow pictures scale in whole steps: centre what Tk made on the box
            cv.create_image((x0 + x1) / 2, (y0 + y1) / 2, image=photo, anchor="center")
        else:
            cv.create_rectangle(x0, y0, x1, y1, fill="#f2f2f2", outline="#a6a6a6")
            hint = ("Ảnh JPEG: cài Pillow để xem trên màn hình\n(pip install pillow)"
                    if picture.get("format") == "jpeg" else "Không hiển thị được ảnh")
            cv.create_text((x0 + x1) / 2, (y0 + y1) / 2, width=max(40, x1 - x0 - 10), fill="#595959",
                           font=self.ui.small_font, justify="center", text=hint)

    @staticmethod
    def _chart_handles(box):
        """Resize handles of a chart box: {edges moved ("l", "rb", ...): (x, y)}."""
        x0, y0, x1, y1 = box
        xm, ym = (x0 + x1) / 2, (y0 + y1) / 2
        return {"lt": (x0, y0), "t": (xm, y0), "rt": (x1, y0), "r": (x1, ym), "rb": (x1, y1), "b": (xm, y1),
                "lb": (x0, y1), "l": (x0, ym)}

    def _chart_at(self, x, y):
        """(index, box, handle or None) of the topmost chart under a canvas point, or None."""
        tolerance = max(4, 6 * self.px)
        for i, box in reversed(self._chart_boxes):
            if i == self.chart_sel:
                for edges, (hx, hy) in self._chart_handles(box).items():
                    if abs(x - hx) <= tolerance and abs(y - hy) <= tolerance:
                        return i, box, edges
            if self._inside(box, x, y):
                return i, box, None
        return None

    def _place_chart(self, index, box):
        """Re-anchor chart `index` so that it covers canvas box (x0, y0, x1, y1)."""
        x0, y0, x1, y1 = box
        r, c = self.row_at(y0), self.col_at(x0)
        self.sh.charts[index] = dict(self.sh.charts[index], anchor=(
            r, c, max(0, round((x0 - self._sheet_x(c)) / self.px)), max(0, round((y0 - self._sheet_y(r)) / self.px))),
            size=(round((x1 - x0) / self.px), round((y1 - y0) / self.px)))

    def _drag_chart(self, x, y):
        d = self.drag
        dx, dy = x - d["start"][0], y - d["start"][1]
        x0, y0, x1, y1 = d["box"]
        least = 60 * self.px
        if d["edges"] is None:
            x0, x1 = x0 + dx, x1 + dx
            y0, y1 = y0 + dy, y1 + dy
            push = max(0, self.header_w - x0), max(0, self.header_h - y0)  # keep it off the headers
            x0, x1, y0, y1 = x0 + push[0], x1 + push[0], y0 + push[1], y1 + push[1]
        else:
            edges = d["edges"]
            if "l" in edges:
                x0 = max(self.header_w, min(x0 + dx, x1 - least))
            if "r" in edges:
                x1 = max(x1 + dx, x0 + least)
            if "t" in edges:
                y0 = max(self.header_h, min(y0 + dy, y1 - least))
            if "b" in edges:
                y1 = max(y1 + dy, y0 + least)
            if len(edges) == 2 and self.sh.charts[d["index"]]["type"] == fe.PICTURE:  # corners keep the shape
                bx0, by0, bx1, by1 = d["box"]
                height = (x1 - x0) * (by1 - by0) / max(1, bx1 - bx0)
                y0, y1 = (y1 - height, y1) if "t" in edges else (y0, y0 + height)
        self._place_chart(d["index"], (x0, y0, x1, y1))
        self.redraw()

    def _draw_cell_extras(self, cv, W, H, col_pos, row_pos):
        """The validation-list arrow and input message of the active cell, and note callouts."""
        self._list_button = None
        r, c = self.cur
        box = self._rect_bbox(*self._cell_box(r, c))
        shown = self._clip_box(box, W, H) is not None and box[2] > box[0] and box[3] > box[1]
        rule = self.wb.validation_at(self.sheet, r, c)
        if rule and shown and rule.get("type") == "list" and rule.get("dropdown", True) and not self.is_editing():
            size = min(box[3] - box[1], round(18 * self.px))
            bx0, by1 = box[2] + 1, box[3]
            self._list_button = (bx0, by1 - size, bx0 + size, by1)
            cv.create_rectangle(*self._list_button, fill="#f0f0f0", outline="#a6a6a6")
            mx, my, d = bx0 + size / 2, by1 - size / 2, size * 0.2
            cv.create_polygon(mx - d, my - d / 2, mx + d, my - d / 2, mx, my + d * 0.8, fill="#444444", outline="")
        if rule and shown and (rule.get("prompt") or rule.get("prompt_title")):
            self._callout(cv, W, H, box[0] + 8, box[3] + 6, rule.get("prompt_title"), rule.get("prompt"), None)
        notes = self.sh.notes
        if self.show_notes:
            keys = [k for k in notes if k[0] in row_pos and k[1] in col_pos]
        elif self.hover_note and self.hover_note[0] == self.sheet and self.hover_note[1:] in notes:
            keys = [self.hover_note[1:]]
        else:
            keys = []
        for nr, nc in sorted(keys):
            nbox = self._rect_bbox(*self._cell_box(nr, nc))
            if nbox[2] <= nbox[0] or nbox[3] <= nbox[1]:
                continue
            self._callout(cv, W, H, nbox[2] + 14, nbox[1] - 4, None, notes[(nr, nc)], (nbox[2] - 1, nbox[1]))

    def _callout(self, cv, W, H, x, y, title, text, pointer):
        """A note-style box at (x, y), moved left of `pointer` (x, y) or up when it would leave the canvas."""
        font, bold, pad = self.ui.small_font, self.header_bold, round(6 * self.scale)
        width = round(190 * self.scale)
        lines = self.wrap_lines(text or "", font, width - 2 * pad) if text else []
        ls = font.metrics("linespace")
        height = 2 * pad + len(lines) * ls + (bold.metrics("linespace") + 2 if title else 0)
        if x + width > W and pointer:
            x = pointer[0] - 14 - width
        x = max(self.header_w + 2, min(x, W - width - 2))
        y = max(self.header_h + 2, min(y, H - height - 2))
        if pointer:
            cv.create_line(pointer[0], pointer[1], x if x > pointer[0] else x + width, y + 6, fill=NOTE_LINE)
        cv.create_rectangle(x + 2, y + 2, x + width + 2, y + height + 2, fill="#d8d8d8", outline="")
        cv.create_rectangle(x, y, x + width, y + height, fill=NOTE_BG, outline=NOTE_LINE)
        ty = y + pad
        if title:
            cv.create_text(x + pad, ty, text=title, anchor="nw", font=bold, fill="#000000")
            ty += bold.metrics("linespace") + 2
        if lines:
            cv.create_text(x + pad, ty, text="\n".join(lines), anchor="nw", font=font, fill="#000000")

    def _draw_border(self, cv, key, spec, x0, y0, x1, y1):
        style, color = spec[0], spec[1] if len(spec) > 1 else "#000000"
        width = max(1, round({"medium": 2, "thick": 3}.get(style, 1) * min(self.px, 2)))
        dash = {"dashed": "-", "dotted": "."}.get(style, "")
        points = {"bt": ((x0 - 1, y0 - 1, x1, y0 - 1), (0, 1)), "bb": ((x0 - 1, y1 - 1, x1, y1 - 1), (0, -1)),
                  "bl": ((x0 - 1, y0 - 1, x0 - 1, y1), (1, 0)), "br": ((x1 - 1, y0 - 1, x1 - 1, y1), (-1, 0))}
        (a, b, c, d), (dx, dy) = points[key]
        if style == "double":
            gap = max(2, round(2 * self.px))
            cv.create_line(a, b, c, d, fill=color)
            cv.create_line(a + dx * gap, b + dy * gap, c + dx * gap, d + dy * gap, fill=color)
        else:
            cv.create_line(a, b, c, d, fill=color, width=width, dash=dash)

    def _draw_headers(self, cv, W, H, sel, sel_box=None):
        hw, hh = self.header_w, self.header_h
        ow, oh = self.outline_w, self.outline_h  # outline symbols take the room before the headings
        if not self.show_headings:
            return
        r1, c1, r2, c2 = sel
        cv.create_rectangle(0, oh, W, hh, fill=HEADER_BG, outline="")
        cv.create_rectangle(ow, 0, hw, H, fill=HEADER_BG, outline="")
        full_cols = r1 == 0 and r2 == MAX_ROWS - 1
        full_rows = c1 == 0 and c2 == MAX_COLS - 1
        for c, x0, x1 in self.vis_cols:
            sel_col = c1 <= c <= c2
            if sel_col:
                cv.create_rectangle(x0, oh, x1, hh, fill=HEADER_FULL_BG if full_cols else HEADER_SEL_BG, outline="")
            cv.create_line(x1 - 1, oh, x1 - 1, hh, fill=HEADER_LINE)
            cv.create_text((x0 + x1) // 2, (oh + hh) // 2, text=col_name(c), fill=HEADER_SEL_FG if sel_col else HEADER_FG,
                           font=self.header_bold if sel_col else self.header_font)
        f = self.sh.filter
        fr1, fr2 = (f["range"][0], f["range"][2]) if f and f["criteria"] else (0, -1)
        for r, y0, y1 in self.vis_rows:
            sel_row = r1 <= r <= r2
            if sel_row:
                cv.create_rectangle(ow, y0, hw, y1, fill=HEADER_FULL_BG if full_rows else HEADER_SEL_BG, outline="")
            cv.create_line(ow, y1 - 1, hw, y1 - 1, fill=HEADER_LINE)
            if y1 - y0 >= self.header_font.metrics("linespace") * 0.6:
                fg = HEADER_SEL_FG if sel_row else FILTERED_ROW_FG if fr1 < r <= fr2 else HEADER_FG
                cv.create_text((ow + hw) // 2, (y0 + y1) // 2, text=str(r + 1), fill=fg,
                               font=self.header_bold if sel_row else self.header_font)
        cv.create_line(ow, hh - 1, W, hh - 1, fill=HEADER_LINE)
        cv.create_line(hw - 1, oh, hw - 1, H, fill=HEADER_LINE)
        if sel_box:
            sx0, sy0, sx1, sy1 = sel_box
            if sx1 > hw and sx0 < W:
                cv.create_line(max(sx0, hw), hh - 1, min(sx1, W), hh - 1, fill=ACCENT, width=2)
            if sy1 > hh and sy0 < H:
                cv.create_line(hw - 1, max(sy0, hh), hw - 1, min(sy1, H), fill=ACCENT, width=2)
        cv.create_rectangle(ow, oh, hw - 1, hh - 1, fill=HEADER_BG, outline="")
        t = round((hh - oh) * 0.45)
        cv.create_polygon(hw - 5, hh - 5 - t, hw - 5, hh - 5, hw - 5 - t, hh - 5, outline="",
                          fill=ACCENT if full_cols and full_rows else "#c4c4c4")

    def _draw_outline(self, cv, W, H, row_pos, col_pos):
        """Excel's outline symbols beside the headings: level buttons 1, 2, ..., a bracket along each group
        and a +/- button at its summary row or column. Button boxes go to self._outline_buttons."""
        self._outline_buttons = []
        ow, oh, hw, hh, step = self.outline_w, self.outline_h, self.header_w, self.header_h, self.outline_step
        if not (ow or oh):
            return
        sh = self.sh
        cv.create_rectangle(0, 0, W, oh, fill=OUTLINE_BG, outline="")
        cv.create_rectangle(0, 0, ow, H, fill=OUTLINE_BG, outline="")
        size = max(7, round(11 * self.scale))
        font = (self.ui_family, -max(7, round(10 * self.scale)))

        def button(x, y, text, action, side):
            x0, y0, x1, y1 = x - side / 2, y - side / 2, x + side / 2, y + side / 2
            cv.create_rectangle(x0, y0, x1, y1, fill="#ffffff", outline=OUTLINE_FG)
            if text in "+-":
                arm = side * 0.28
                cv.create_line(x - arm, y, x + arm + 1, y, fill="#404040")
                if text == "+":
                    cv.create_line(x, y - arm, x, y + arm + 1, fill="#404040")
            else:
                cv.create_text(x, y, text=text, font=font, fill="#404040")
            self._outline_buttons.append((x0, y0, x1, y1, action))

        for axis, levels, pos, vis, room in ((0, sh.row_levels, row_pos, self.vis_rows, ow),
                                             (1, sh.col_levels, col_pos, self.vis_cols, oh)):
            if not room or not levels:
                continue

            def pt(across, along, axis=axis):
                return (across, along) if axis == 0 else (along, across)
            hidden = sh.hidden_rows if axis == 0 else sh.hidden_cols
            start, end = (hh, H) if axis == 0 else (hw, W)
            band = (oh + hh) / 2 if axis == 0 else (ow + hw) / 2
            for n in range(1, max(levels.values()) + 2):
                button(*pt((n - 1) * step + step / 2 + 1, band), str(n), ("level", axis, n), size + 2)
            for level, a, b in fe.outline_groups(levels):
                across = (level - 1) * step + step / 2 + 1
                collapsed = a in hidden and b in hidden
                summary = pos.get(b + 1)
                if not collapsed and vis:
                    seen = [(p0, p1) for i, p0, p1 in vis if a <= i <= b]
                    span = ((seen[0][0], seen[-1][1]) if seen else
                            (start, end) if a < vis[0][0] and b > vis[-1][0] else None)
                    if span:
                        lo = span[0] + 3 if a in pos else start
                        hi = (summary[0] + summary[1]) / 2 - size / 2 if summary else span[1]
                        if hi > lo:
                            cv.create_line(*pt(across, lo), *pt(across, hi), fill=OUTLINE_FG)
                        if a in pos:
                            cv.create_line(*pt(across, lo), *pt(across + step * 0.4, lo), fill=OUTLINE_FG)
                if summary:
                    button(*pt(across, (summary[0] + summary[1]) / 2), "+" if collapsed else "-",
                           ("toggle", axis, level, a, b, not collapsed), size)

    def update_scrollbars(self):
        rows, cols = self.wb.used_extent(self.sheet)
        vis_r = self._visible_rows()
        total_r = min(MAX_ROWS, max(rows, self.top + vis_r, self.cur[0] + 1) + vis_r)
        self._v_total = max(1, total_r)
        self.vbar.set(self.top / self._v_total, min(1.0, (self.top + vis_r) / self._v_total))
        vis_c = max(1, len(self.vis_cols) - 1)
        total_c = min(MAX_COLS, max(cols, self.left + vis_c, self.cur[1] + 1) + vis_c)
        self._h_total = max(1, total_c)
        self.hbar.set(self.left / self._h_total, min(1.0, (self.left + vis_c) / self._h_total))

    # ------------------------------------------------------------------ refresh

    def refresh_view(self):
        self.redraw()
        self.update_scrollbars()

    def refresh(self):
        self.refresh_view()
        self.update_formula_bar()
        self.update_name_box()
        self.update_toolbar()
        self.update_status()
        self._sync_pivot_pane()

    def refresh_all(self):
        self.build_tabs()
        self.refresh()
        self.update_title()
        mode = self.mode_label.cget("text")
        if mode == "Sẵn sàng" or mode.startswith("Bộ lọc:"):
            self.set_mode("Sẵn sàng")

    def _focus_is(self, widget):
        # compare Tcl names: focus_get() raises KeyError while a combobox list is open
        return str(self.root.tk.call("focus")) == str(widget)

    def update_formula_bar(self):
        if not self.is_editing():
            raw = self.wb.get_raw(self.sheet, *self.cur)
            if self.sh.protection and (self.sh.formats.get(self.cur) or {}).get("hide_formula"):
                raw = ""  # Format Cells > Protection > Hidden: the formula is not shown while protected
            self.fx_var.set(raw)
        editing = self.is_editing()
        self.btn_x.config(fg="#c42b1c" if editing else "#c0c0c0")
        self.btn_ok.config(fg="#107c41" if editing else "#c0c0c0")

    def update_name_box(self):
        if self._focus_is(self.name_box):
            return
        r1, c1, r2, c2 = self.sel_rect()
        if self.drag and self.drag.get("mode") in ("select", "colsel", "rowsel") and (r1, c1, r2, c2) != self._cell_box(*self.cur):
            text = f"{r2 - r1 + 1}H x {c2 - c1 + 1}C"
        else:
            text = cell_name(*self.cur)
            if self.wb.names:  # a block that has a name shows the name, as in Excel
                text = next((d["name"] for d in self.wb.sorted_names()
                             if self.wb.name_target(d["name"]) == (self.sheet, (r1, c1, r2, c2))), text)
        self.name_box.delete(0, "end")
        self.name_box.insert(0, text)

    def update_toolbar(self):
        fmt = self.sh.formats.get(self.cur) or {}
        if not self._focus_is(self.font_box):
            self.font_var.set(fmt.get("font") or self.default_family)
        if not self._focus_is(self.size_box):
            size = fmt.get("size") or 11
            self.size_var.set(str(int(size)) if float(size).is_integer() else str(size))
        self.btn_bold.set_checked(fmt.get("bold"))
        self.btn_italic.set_checked(fmt.get("italic"))
        self.btn_underline.set_checked(fmt.get("underline"))
        valign, align = fmt.get("valign"), fmt.get("align")
        self.btn_vtop.set_checked(valign == "top")
        self.btn_vmid.set_checked(valign == "center")
        self.btn_vbot.set_checked(valign in (None, "bottom"))
        self.btn_left.set_checked(align == "left")
        self.btn_center.set_checked(align == "center")
        self.btn_right.set_checked(align == "right")
        self.btn_wrap.set_checked(fmt.get("wrap"))
        self.btn_merge.set_checked(self.sh.merge_at(*self.cur) is not None)
        self.btn_painter.set_checked(self.painter is not None)
        self.btn_show_formulas.set_checked(self.show_formulas)
        self.btn_filter.set_checked(self.sh.filter is not None)
        self.btn_protect.set_checked(self.sh.protection is not None)
        self.print_grid_var.set(bool((self.sh.page_setup or {}).get("gridlines")))
        num = dialogs.canonical_format(fmt.get("num"))
        self.numfmt_box.set(next((label for label, f in NUMBER_FORMATS if dialogs.canonical_format(f) == num),
                                 "Tùy chỉnh"))

    def update_status(self):
        """Average / Count / Sum of the selection in the status bar. Large selections are added up a piece
        at a time between events, so moving around never waits for them; an unchanged selection is not
        added up again."""
        rects = [self.sel_rect()] + self.areas
        if len(rects) == 1 and rects[0] == self._cell_box(*self.cur):
            self._stats_key = None
            if self._stats_job:
                self.root.after_cancel(self._stats_job)
                self._stats_job = None
            self.stats_label.config(text="")
            return
        rows, cols = self.wb.used_extent(self.sheet)
        rects = [(r1, c1, min(r2, rows - 1), min(c2, cols - 1)) for r1, c1, r2, c2 in rects]
        key = (id(self.wb), self.sheet, tuple(rects), self.edit_serial)
        if key == self._stats_key:
            return
        self._stats_key = key
        if self._stats_job:
            self.root.after_cancel(self._stats_job)
        work = self._stats_work(rects)
        if sum(max(0, r2 - r1 + 1) * max(0, c2 - c1 + 1) for r1, c1, r2, c2 in rects) > 5000:
            self.stats_label.config(text="")
        self._stats_job = None
        self._stats_step(work)

    def _stats_work(self, rects):
        """Generator adding up the selected cells; yields now and then so the work can pause, and
        finally yields (count, numbers, total)."""
        cells, sh, get = self.sh.cells, self.sh, self.wb.get_value
        sheet, seen = self.sheet, set()
        count = nums = 0
        total = 0.0
        for r1, c1, r2, c2 in rects:
            if r2 < r1 or c2 < c1:
                continue
            if (r2 - r1 + 1) * (c2 - c1 + 1) > len(cells):
                keys = sorted(k for k in cells if r1 <= k[0] <= r2 and c1 <= k[1] <= c2)
            else:
                keys = ((r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) in cells)
            for i, (r, c) in enumerate(keys):
                if i % 4000 == 3999:
                    yield None
                if (r, c) in seen or sh.row_hidden(r):
                    continue
                if len(rects) > 1:
                    seen.add((r, c))
                v = get(sheet, r, c)
                if v is None or v == "":
                    continue
                count += 1
                if fe.is_number(v):
                    nums += 1
                    total += v
        yield count, nums, total

    def _stats_step(self, work):
        """Run the adding up for a moment, then let events through and carry on."""
        end = time.perf_counter() + 0.04
        result = None
        while result is None and time.perf_counter() < end:
            result = next(work)
        if result is None:
            self._stats_job = self.root.after(1, lambda: self._stats_step(work))
            return
        self._stats_job = None
        count, nums, total = result
        parts = []
        if nums:
            parts.append(f"Trung bình: {fe.format_general(total / nums)}")
        if count:
            parts.append(f"Đếm: {count}")
        if nums:
            parts.append(f"Tổng: {fe.format_general(total)}")
        self.stats_label.config(text="      ".join(parts))

    def update_title(self):
        name = os.path.basename(self.path) if self.path else "Sổ làm việc mới"
        self.root.title(f"{'*' if self.dirty else ''}{name} - {APP_NAME}")

    def set_mode(self, text):
        if text == "Sẵn sàng":
            text = self._ready_text()
        self.mode_label.config(text=text)

    def _ready_text(self):
        """'Ready', or how many records a filter shows (Excel's "5 of 20 records found")."""
        f = self.sh.filter if self.sheet in self.wb.sheets else None
        if f and f["criteria"]:
            total = f["range"][2] - f["range"][0]
            return f"Bộ lọc: tìm thấy {total - len(f['hidden'])} trên {total} bản ghi"
        return "Sẵn sàng"

    def mark_dirty(self):
        if not self.dirty:
            self.dirty = True
            self.update_title()

    def build_tabs(self):
        for widget in self.tab_frame.winfo_children():
            widget.destroy()
        for name, sheet in self.wb.sheets.items():
            if sheet.hidden:
                continue
            active = name == self.sheet
            strip = sheet.tab_color or (ACCENT if active else SHEETBAR_BG)
            holder = tk.Frame(self.tab_frame, bg=strip)
            holder.pack(side="left", padx=(0, 1))
            label = tk.Label(holder, text=name, padx=14, pady=3, cursor="hand2", font=self.ui.font,
                             bg="white" if active else "#e0e0e0", fg=ACCENT if active else "#333333")
            if active:
                label.config(font=(self.ui.family, -round(12 * self.scale), "bold"))
            label.pack(pady=(0, 3 if active or sheet.tab_color else 0))
            label.bind("<Button-1>", lambda e, n=name: self.switch_sheet(n))
            label.bind("<Double-Button-1>", lambda e, n=name: self.rename_sheet(n))
            label.bind("<Button-3>", lambda e, n=name: self._tab_menu(e, n))

    # ------------------------------------------------------------------ selection and navigation

    def _anchor(self, pos):
        m = self.sh.merge_at(*pos)
        return (m[0], m[1]) if m else pos

    def _step(self, pos, dr, dc):
        """Cell reached by one arrow-key step, jumping over merged blocks and hidden rows/columns."""
        r, c = pos
        m = self.sh.merge_at(r, c)
        if m:
            r = m[2] if dr > 0 else m[0] if dr < 0 else r
            c = m[3] if dc > 0 else m[1] if dc < 0 else c
        r, c = r + dr, c + dc
        while dr and 0 <= r < MAX_ROWS and self.sh.row_hidden(r):
            r += dr
        while dc and 0 <= c < MAX_COLS and c in self.sh.hidden_cols:
            c += dc
        if not (0 <= r < MAX_ROWS and 0 <= c < MAX_COLS):
            return pos
        return r, c

    def select_to(self, target, extend=False):
        self.areas = []
        if extend:
            self.ext = target
        else:
            self.cur = self.ext = self._anchor(target)
        self.ensure_visible(*target)
        self.refresh()

    def move_active(self, dr, dc):
        self.select_to(self._step(self.cur, dr, dc))

    def select_all(self):
        self.commit_edit()
        self.areas = []
        self.cur, self.ext = (0, 0), (MAX_ROWS - 1, MAX_COLS - 1)
        self.refresh()

    def _jump(self, pos, dr, dc):
        """Ctrl+arrow: to the edge of the current block of data, or the next filled cell."""
        cells = self.sh.cells

        def inside(r, c):
            return 0 <= r < MAX_ROWS and 0 <= c < MAX_COLS

        r, c = pos
        nr, nc = r + dr, c + dc
        if not inside(nr, nc):
            return pos
        if (r, c) in cells and (nr, nc) in cells:
            while inside(nr + dr, nc + dc) and (nr + dr, nc + dc) in cells:
                nr, nc = nr + dr, nc + dc
            return nr, nc
        while inside(nr, nc) and (nr, nc) not in cells:
            nr, nc = nr + dr, nc + dc
        return (nr, nc) if inside(nr, nc) else (nr - dr, nc - dc)

    def on_key(self, e):
        ks = e.keysym
        if self.chart_sel is not None and ks not in MODIFIER_KEYS:
            if ks in ("Delete", "BackSpace"):
                self.delete_chart()
                return "break"
            self.chart_sel = None  # any other key goes back to the cells
            self.redraw()
            if ks == "Escape":
                return "break"
        shift, ctrl = bool(e.state & SHIFT), bool(e.state & CONTROL)
        arrows = {"Up": (-1, 0), "Down": (1, 0), "Left": (0, -1), "Right": (0, 1)}
        base = self.ext if shift else self.cur
        if ks in arrows:
            dr, dc = arrows[ks]
            target = self._jump(base, dr, dc) if ctrl else self._step(base, dr, dc)
            self.select_to(target, extend=shift)
        elif ks in ("Return", "KP_Enter"):
            self.move_active(-1 if shift else 1, 0)
        elif ks == "Tab":
            self.move_active(0, -1 if shift else 1)
        elif ks in ("Prior", "Next") and ctrl:
            self.step_sheet(-1 if ks == "Prior" else 1)
        elif ks in ("Prior", "Next"):
            n = self._visible_rows() * (-1 if ks == "Prior" else 1)
            self.top = max(0, min(MAX_ROWS - 1, self.top + n))
            self.select_to(self._clamp(base[0] + n, base[1]), extend=shift)
        elif ks == "Home":
            self.select_to((0, 0) if ctrl else (base[0], 0), extend=shift)
        elif ks == "End" and ctrl:
            rows, cols = self.wb.used_extent(self.sheet)
            self.select_to((max(rows - 1, 0), max(cols - 1, 0)), extend=shift)
        elif ks == "F2":
            self.start_edit(mode="edit")
        elif ks == "Delete":
            self.clear_contents()
        elif ks == "BackSpace":
            self.start_edit(initial="", mode="enter")
        elif ks == "Escape":
            if self.painter:
                self.stop_painter()
            elif self.clip:
                self.clip = None
                self.set_mode("Sẵn sàng")
                self.redraw()
        elif ks == "space" and ctrl:
            self.cur, self.ext = (0, self.cur[1]), (MAX_ROWS - 1, self.ext[1])
            self.refresh()
        elif ks == "space" and shift:
            self.cur, self.ext = (self.cur[0], 0), (self.ext[0], MAX_COLS - 1)
            self.refresh()
        elif e.char and e.char.isprintable() and (not ctrl or e.state & ALT):
            self.start_edit(initial=e.char, mode="enter")
        else:
            return None
        return "break"

    # ------------------------------------------------------------------ mouse

    def on_press(self, e):
        if self._list_button and self._inside(self._list_button, e.x, e.y):
            self.open_validation_list()
            return "break"
        hit = self._chart_at(e.x, e.y)
        prot = self.sh.protection
        if hit is not None and prot and "objects" not in prot.get("allow", ()):
            hit = None  # objects on a protected sheet can't be selected
        if hit is not None and not self._point_mode_ok():
            self.commit_edit()
            self.canvas.focus_set()
            index, box, edges = hit
            self.chart_sel = index
            self.drag = {"mode": "chart", "index": index, "box": box, "edges": edges, "start": (e.x, e.y),
                         "before": self.wb.snapshot(), "sel": self._sel_state()}
            self.redraw()
            return "break"
        if self.chart_sel is not None:
            self.chart_sel = None
            self.redraw()
        for bx0, by0, bx1, by1, action in self._outline_buttons:
            if bx0 - 2 <= e.x <= bx1 + 2 and by0 - 2 <= e.y <= by1 + 2:
                self.commit_edit()
                self.outline_action(action)
                return "break"
        if e.x < self.outline_w or e.y < self.outline_h:
            return "break"
        for bx0, by0, bx1, by1, col in self._filter_buttons:
            if bx0 <= e.x <= bx1 and by0 <= e.y <= by1:
                self.commit_edit()
                self.open_filter_popup(col, self.canvas.winfo_rootx() + int(bx0), self.canvas.winfo_rooty() + int(by1) + 2)
                return "break"
        kind, r, c = self.hit(e.x, e.y)
        if kind == "cell" and self._point_mode_ok():
            self._point_start(r, c)
            return "break"
        self.commit_edit()
        self.canvas.focus_set()
        border = self._border_hit(e.x, e.y)
        if border is not None and prot and "format_rows_cols" not in prot.get("allow", ()):
            border = None
        if border is not None:
            axis, index = border
            size = self.col_w(index) if axis == "col" else self.row_h(index)
            self.drag = {"mode": "resize_" + axis, "index": index, "start": e.x if axis == "col" else e.y,
                         "size": size, "before": self.wb.snapshot(), "sel": self._sel_state()}
            return None
        if kind == "cell" and self._on_fill_handle(e.x, e.y) and not self.painter:
            self.drag = {"mode": "fill", "target": None}
            return None
        shift = bool(e.state & SHIFT)
        if kind == "cell" and e.state & CONTROL and not shift and not self.painter:
            self.areas.append(self.sel_rect())  # Ctrl+click: keep what is selected and start another block
        elif not shift:
            self.areas = []
        if kind == "corner":
            self.cur, self.ext = (0, 0), (MAX_ROWS - 1, MAX_COLS - 1)
            self.drag = None
        elif kind == "colhead":
            self.cur = (0, self.cur[1] if shift else c)
            self.ext = (MAX_ROWS - 1, c)
            self.drag = {"mode": "colsel"}
        elif kind == "rowhead":
            self.cur = (self.cur[0] if shift else r, 0)
            self.ext = (r, MAX_COLS - 1)
            self.drag = {"mode": "rowsel"}
        else:
            if shift:
                self.ext = (r, c)
            else:
                self.cur = self.ext = self._anchor((r, c))
                if not self.painter and not e.state & CONTROL and self.wb.hyperlink_at(self.sheet, *self.cur):
                    self._link_click = (self.sheet, self.cur, e.time)  # opened on release unless held
            self.drag = {"mode": "select"}
        self.refresh()
        return None

    def on_drag(self, e):
        if not self.drag:
            return
        self._last_mouse = (e.x, e.y)
        self._drag_to(e.x, e.y)
        W, H = self.canvas.winfo_width(), self.canvas.winfo_height()
        outside = e.x < self.header_w or e.y < self.header_h or e.x > W or e.y > H
        if outside and self._autoscroll_job is None and not self.drag["mode"].startswith(("resize", "chart")):
            self._autoscroll_job = self.root.after(50, self._autoscroll)

    def _drag_to(self, x, y):
        d = self.drag
        mode = d["mode"]
        if mode == "chart":
            self._drag_chart(x, y)
            return
        if mode.startswith("resize"):
            pos = x if mode == "resize_col" else y
            size = max(0, round((d["size"] + pos - d["start"]) / self.px))
            if mode == "resize_col":
                self.sh.col_widths[d["index"]] = max(fe.MIN_COL_WIDTH, size)
            else:
                self.sh.row_heights[d["index"]] = max(4, size)
                self.sh.custom_rows.add(d["index"])
            self.refresh_view()
            return
        r, c = self.row_at(y), self.col_at(x)
        if mode == "select":
            self.ext = (r, c)
        elif mode == "colsel":
            self.ext = (MAX_ROWS - 1, c)
        elif mode == "rowsel":
            self.ext = (r, MAX_COLS - 1)
        elif mode == "point":
            self._point_set(*d["anchor"], r, c)
        elif mode == "fill":
            d["target"] = self._fill_target(r, c)
        self.redraw()
        self.update_status()
        self.update_name_box()

    def _autoscroll(self):
        self._autoscroll_job = None
        if not self.drag or self.drag["mode"].startswith("resize"):
            return
        x, y = self._last_mouse
        W, H = self.canvas.winfo_width(), self.canvas.winfo_height()
        mode, moved = self.drag["mode"], False
        _, r1, _, c1 = self.frozen()
        if y > H and self.top < MAX_ROWS - 1:
            self.top, moved = self.top + 1, True
        elif y < self.header_h and self.top > r1 and mode != "colsel":
            self.top, moved = self.top - 1, True
        if x > W and self.left < MAX_COLS - 1:
            self.left, moved = self.left + 1, True
        elif x < self.header_w and self.left > c1 and mode != "rowsel":
            self.left, moved = self.left - 1, True
        if moved:
            self._drag_to(x, y)
            self.update_scrollbars()
            self._autoscroll_job = self.root.after(50, self._autoscroll)

    def on_release(self, e):
        d, self.drag = self.drag, None
        click, self._link_click = self._link_click, None
        self._hide_link_tip()
        if self._autoscroll_job is not None:
            self.root.after_cancel(self._autoscroll_job)
            self._autoscroll_job = None
        if not d:
            return
        if d["mode"] == "fill" and d.get("target"):
            self.do_fill(d["target"])
        elif d["mode"].startswith("resize"):
            if d["mode"] == "resize_col":
                self.autofit_wrapped_rows(d["index"])
            self._push_snapshot(d["before"], d["sel"])
        elif d["mode"] == "chart":
            self._push_snapshot(d["before"], d["sel"])
            self.redraw()
            return
        elif d["mode"] in ("select", "colsel", "rowsel") and self.painter:
            self._apply_painter()
        if (click and d["mode"] == "select" and click[:2] == (self.sheet, self.cur) and self.ext == self.cur
                and 0 <= e.time - click[2] < 600):  # a click, not a hold, on a linked cell follows the link
            link = self.wb.hyperlink_at(self.sheet, *self.cur)
            if link:
                self.refresh()
                self.follow_link(link)
                return
        self.refresh()

    def on_double(self, e):
        if self._chart_at(e.x, e.y) is not None:
            self.edit_chart()
            return "break"
        border = self._border_hit(e.x, e.y)
        if border is not None:
            if border[0] == "col":
                self.autofit_columns(border[1], border[1])
            else:
                self.autofit_rows_manual([border[1]])
            return "break"
        kind, r, c = self.hit(e.x, e.y)
        if kind == "cell" and not self.is_editing() and not self.painter:
            p = pv.pivot_at(self.wb, self.sheet, r, c)
            if p is not None:
                self.pivot_details(p, r, c)
                return "break"
            self.cur = self.ext = self._anchor((r, c))
            self.start_edit(mode="edit")
        return "break"

    @staticmethod
    def _inside(box, x, y):
        return box[0] <= x <= box[2] and box[1] <= y <= box[3]

    def _set_hover_note(self, key):
        if key != self.hover_note:
            self.hover_note = key
            self.redraw()

    def on_motion(self, e):
        if self.drag:
            return
        kind, r, c = self.hit(e.x, e.y)
        anchor = self._anchor((r, c)) if kind == "cell" else None
        chart = self._chart_at(e.x, e.y)
        self._set_hover_note((self.sheet, *anchor) if anchor in self.sh.notes and not chart else None)
        border = self._border_hit(e.x, e.y)
        if chart is not None:
            cursor = {None: "fleur", "l": "sb_h_double_arrow", "r": "sb_h_double_arrow", "t": "sb_v_double_arrow",
                      "b": "sb_v_double_arrow", "lt": "top_left_corner", "rb": "bottom_right_corner",
                      "rt": "top_right_corner", "lb": "bottom_left_corner"}[chart[2]]
        elif border is not None:
            cursor = "sb_h_double_arrow" if border[0] == "col" else "sb_v_double_arrow"
        elif any(b[0] <= e.x <= b[2] and b[1] <= e.y <= b[3] for b in self._filter_buttons) or (
                self._list_button and self._inside(self._list_button, e.x, e.y)):
            cursor = "hand2"
        elif not self.is_editing() and self._on_fill_handle(e.x, e.y):
            cursor = "crosshair"
        elif self.painter:
            cursor = "pencil"
        else:
            cursor = {"colhead": "sb_down_arrow", "rowhead": "sb_right_arrow", "cell": "plus"}.get(self.hit(e.x, e.y)[0], "")
        link = self.wb.hyperlink_at(self.sheet, *anchor) if cursor == "plus" and not self.is_editing() else None
        if link:
            cursor = "hand2"
        hover = (self.sheet, anchor) if link else None
        if hover != self._link_hover:
            self._link_hover = hover
            self._hide_link_tip()
            if link:
                self._show_link_tip(link, e.x_root, e.y_root)
        if cursor != self._cursor:
            self._cursor = cursor
            self.canvas.config(cursor=cursor)

    def on_wheel(self, e, horizontal=False):
        widget = self.root.winfo_containing(e.x_root, e.y_root)
        if widget is not self.canvas and widget is not self.editor:
            return None
        notches = round(e.delta / 120) or (1 if e.delta > 0 else -1)
        if e.state & CONTROL:
            self.set_zoom(self.zoom + 0.1 * notches)
            return "break"
        if horizontal:
            self.left = max(0, min(MAX_COLS - 1, self.left - notches))
        else:
            self.top = max(0, min(MAX_ROWS - 1, self.top - 3 * notches))
        self.refresh_view()
        return "break"

    def on_vscroll(self, *args):
        if args[0] == "moveto":
            self.top = int(float(args[1]) * self._v_total)
        elif args[0] == "scroll":
            self.top += int(args[1]) * (self._visible_rows() if args[2] == "pages" else 1)
        self.top = max(0, min(self.top, MAX_ROWS - 1))
        self.refresh_view()

    def on_hscroll(self, *args):
        if args[0] == "moveto":
            self.left = int(float(args[1]) * self._h_total)
        elif args[0] == "scroll":
            self.left += int(args[1]) * (max(1, len(self.vis_cols) - 1) if args[2] == "pages" else 1)
        self.left = max(0, min(self.left, MAX_COLS - 1))
        self.refresh_view()

    def on_context(self, e):
        self.commit_edit()
        self.canvas.focus_set()
        hit = self._chart_at(e.x, e.y)
        if hit is not None:
            self.chart_sel = hit[0]
            self.redraw()
            menu = self.picture_menu if self.sh.charts[hit[0]]["type"] == fe.PICTURE else self.chart_menu
            try:
                menu.tk_popup(e.x_root, e.y_root)
            finally:
                menu.grab_release()
            return
        self.chart_sel = None
        kind, r, c = self.hit(e.x, e.y)
        r1, c1, r2, c2 = self.sel_rect()
        menu = self.cell_menu
        if kind == "cell" and not (r1 <= r <= r2 and c1 <= c <= c2):
            self.cur = self.ext = self._anchor((r, c))
        elif kind == "colhead":
            if not (c1 <= c <= c2 and r1 == 0 and r2 == MAX_ROWS - 1):
                self.cur, self.ext = (0, c), (MAX_ROWS - 1, c)
            menu = self.col_menu
        elif kind == "rowhead":
            if not (r1 <= r <= r2 and c1 == 0 and c2 == MAX_COLS - 1):
                self.cur, self.ext = (r, 0), (r, MAX_COLS - 1)
            menu = self.row_menu
        self.refresh()
        try:
            menu.tk_popup(e.x_root, e.y_root)
        finally:
            menu.grab_release()

    # ------------------------------------------------------------------ editing

    def is_editing(self):
        return self.editor is not None or self.bar_editing

    def _pivot_guard(self, r1, c1, r2, c2):
        """True (after telling the user) when the cells overlap a pivot report, which is not edited by hand."""
        for p in self.sh.pivots:
            a1, b1, a2, b2 = p["rect"] or (0, 0, -1, -1)
            if a1 <= r2 and r1 <= a2 and b1 <= c2 and c1 <= b2:
                messagebox.showinfo(APP_NAME, "Không thể sửa ô của báo cáo PivotTable. Hãy sửa dữ liệu nguồn rồi "
                                    "bấm Làm mới, hoặc dùng ngăn Các trường PivotTable.", parent=self.root)
                return True
        return False

    def start_edit(self, initial=None, mode="edit"):
        if self.is_editing():
            return
        if self._pivot_guard(*self._cell_box(*self._anchor(self.cur))):
            return
        if self._locked(cells=[self._anchor(self.cur)]):
            return
        if self.painter:
            self.stop_painter()
        r, c = self.cur = self._anchor(self.cur)
        self.ensure_visible(r, c)
        fmt = self.sh.formats.get((r, c)) or {}
        text = self.wb.get_raw(self.sheet, r, c) if initial is None else initial
        self.edit_cell, self.edit_mode = (self.sheet, r, c), mode
        self._edit_font = self.font_for(fmt)
        e = tk.Entry(self.canvas, textvariable=self.edit_var, font=self._edit_font, relief="flat", bd=0,
                     highlightthickness=2, highlightbackground=ACCENT, highlightcolor=ACCENT,
                     bg=fmt.get("bg") or "white", fg=fmt.get("fg") or "black",
                     insertwidth=max(1, round(self.scale)))
        self.editor = e
        self.edit_var.set(text)
        e.icursor("end")
        e.bind("<Return>", lambda ev: self._editor_commit(ev, (1, 0)))
        e.bind("<KP_Enter>", lambda ev: self._editor_commit(ev, (1, 0)))
        e.bind("<Tab>", lambda ev: self._editor_commit(ev, (0, 1)))
        e.bind("<Control-Return>", lambda ev: self._key_action(lambda: self.commit_edit(fill=True)))
        e.bind("<Escape>", lambda ev: self._key_action(self.cancel_edit))
        e.bind("<Control-semicolon>", lambda ev: self._key_action(lambda: self.insert_now(False)))
        e.bind("<Alt-Return>", lambda ev: self._key_action(lambda: e.insert("insert", "\n")))
        for key, d in (("Up", (-1, 0)), ("Down", (1, 0)), ("Left", (0, -1)), ("Right", (0, 1))):
            e.bind(f"<{key}>", lambda ev, d=d: self._editor_arrow(d))
        e.bind("<Key>", self._reset_point, add="+")
        self._formula_assist(e)
        self.set_mode("Nhập" if mode == "enter" else "Sửa")
        self.update_formula_bar()
        self.redraw()
        e.focus_set()

    def _place_editor(self):
        if self.editor is None:
            return
        sheet, r, c = self.edit_cell
        x0, y0, x1, y1 = self._rect_bbox(*self._cell_box(r, c))
        W, H = self.canvas.winfo_width(), self.canvas.winfo_height()
        if (sheet != self.sheet or y0 < self.header_h - 1 or y0 >= H or x1 <= self.header_w or x0 >= W
                or x1 <= x0 or y1 <= y0):  # off screen, or scrolled behind frozen panes
            self.editor.place_forget()
            return
        needed = self.text_width(self._edit_font, self.edit_var.get()) + 4 * self.pad
        width = max(x1 - x0, min(needed, W - x0))
        height = max(y1 - y0, self._edit_font.metrics("linespace") + 4)
        self.editor.place(x=x0 - 1, y=y0 - 1, width=width + 1, height=height + 1)

    def _on_edit_text(self, *_):
        if self.editor is not None:
            self.fx_var.set(self.edit_var.get())
            self.redraw()

    def _on_fx_text(self, *_):
        if self.bar_editing:
            self.redraw()

    def _on_fx_focus(self, e):
        if self.editor is not None:  # move an in-cell edit into the formula bar
            text = self.edit_var.get()
            self.editor.destroy()
            self.editor = None
            self.bar_editing = True
            self.fx_var.set(text)
        elif not self.bar_editing:
            self.bar_editing = True
            self.cur = self._anchor(self.cur)
            self.edit_cell, self.edit_mode = (self.sheet, *self.cur), "edit"
        self.set_mode("Sửa")
        self.update_formula_bar()

    def toggle_protection(self):
        """Review > Protect Sheet / Unprotect Sheet."""
        self.commit_edit()
        sh = self.sh
        if sh.protection:
            if sh.protection.get("hash"):
                pw = simpledialog.askstring("Bỏ bảo vệ Trang tính", "Mật khẩu:", show="*", parent=self.root)
                if pw is None:
                    return
                if fe.password_hash(pw) != sh.protection["hash"]:
                    messagebox.showerror(APP_NAME, "Mật khẩu không đúng.", parent=self.root)
                    return
            self._structural(lambda: setattr(sh, "protection", None))
            self.set_mode("Đã bỏ bảo vệ trang tính")
            return
        result = dialogs.ProtectSheetDialog(self.root).show()
        if result is not None:
            password, allow = result
            self._structural(lambda: setattr(sh, "protection", {"hash": fe.password_hash(password), "allow": allow}))
            self.set_mode("Trang tính đang được bảo vệ")

    def toggle_lock_cells(self):
        """Format > Lock Cell: lock or unlock the selected cells (what protection then keeps from changes)."""
        self.commit_edit()
        if self.sh.protection:
            messagebox.showinfo(APP_NAME, PROTECTED_MSG, parent=self.root)
            return
        targets = self._format_targets()
        unlock = any(fe.cell_locked(self.sh.formats.get(k)) for k in targets)
        cs = ChangeSet(self.wb)
        for r, c in targets:
            fmt = dict(self.sh.formats.get((r, c)) or {})
            if unlock:
                fmt["unlocked"] = True
            else:
                fmt.pop("unlocked", None)
            cs.set(self.sheet, r, c, fmt=fmt)
        self.commit(cs)
        self.set_mode("Đã bỏ khóa ô" if unlock else "Đã khóa ô")

    def _formula_assist(self, entry):
        """Function (and defined name) list while typing, argument tip inside a call, F4 for $ references."""
        functions = fe.function_names()

        def names():
            return sorted(functions + [d["name"] for d in self.wb.names.values()], key=str.upper)

        def info(name):
            item = FUNCTION_INFO.get(name)  # function names are upper case; defined names keep their case
            if item:
                return item[1:]
            d = self.wb.names.get(name.upper())
            if d:
                return d["name"], f"Tên: {d['ref'][1:]}" + (f" — {d['comment']}" if d.get("comment") else ""), False
            return None

        def changed():
            self.point = None
            self.set_mode("Nhập" if self.edit_mode == "enter" else "Sửa")
        return FormulaAssist(entry, self.ui.font, self.ui.small_font, names, info, changed)

    def _editor_commit(self, e, move):
        if e.state & SHIFT:
            move = (-move[0], -move[1])
        self.commit_edit(move=move)
        return "break"

    def _editor_arrow(self, move):
        if self.edit_mode == "enter" and not self.edit_var.get().startswith("="):
            self.commit_edit(move=move)
            return "break"
        self.point = None
        return None

    def _reset_point(self, e):
        if e.keysym not in MODIFIER_KEYS and self.point is not None:
            self.point = None
            self.set_mode("Nhập" if self.edit_mode == "enter" else "Sửa")

    def commit_edit(self, move=None, fill=False):
        """Store the text being edited; Ctrl+Enter (fill=True) writes it into every selected cell."""
        if not self.is_editing():
            return
        text = self.edit_var.get() if self.editor is not None else self.fx_var.get()
        sheet, r, c = self.edit_cell
        self._end_edit()
        if sheet in self.wb.sheets:
            raw = fe.normalize_formula(text, self.wb.names)
            if not self._entry_allowed(sheet, r, c, raw, text):
                self.refresh()
                return
            cs = ChangeSet(self.wb)
            targets = self._format_targets() if fill and sheet == self.sheet else [(r, c)]
            for rr, cc in targets:
                cs.set(sheet, rr, cc, raw=fe.shift_formula(raw, rr - r, cc - c))
                if "\n" in raw and not fe.is_formula(raw):  # Alt+Enter line breaks turn on Wrap Text, as in Excel
                    cs.set(sheet, rr, cc, fmt=dict(self.wb.get_format(sheet, rr, cc) or {}, wrap=True))
            self.commit(cs)
        if move:
            self.move_active(*move)
        else:
            self.refresh()

    def _entry_allowed(self, sheet, r, c, raw, text):
        """Check an entry against the cell's data validation, asking the user like Excel does.
        "Retry" (or "No" on a warning) reopens the editor with the text typed."""
        rule = self.wb.validate(sheet, r, c, raw)
        if rule is None:
            return True
        title = rule.get("error_title") or APP_NAME
        msg = rule.get("error") or ("Giá trị này không khớp với các giới hạn xác thực dữ liệu đã đặt cho ô này."
                                    if rule.get("type") != "list" else "Hãy chọn một giá trị trong danh sách.")
        style = rule.get("style", "stop")
        if style == "warning":
            answer = messagebox.askyesnocancel(title, msg + "\n\nTiếp tục?", icon="warning", parent=self.root)
        elif style == "info":
            answer = messagebox.askokcancel(title, msg, icon="info", parent=self.root) or None
        else:
            answer = False if messagebox.askretrycancel(title, msg, icon="error", parent=self.root) else None
        if answer:
            return True
        if answer is False:
            self.root.after_idle(lambda: self._reopen_edit(sheet, r, c, text))
        return False

    def _reopen_edit(self, sheet, r, c, text):
        if sheet not in self.wb.sheets or self.is_editing():
            return
        if sheet != self.sheet:
            self.switch_sheet(sheet)
        self.cur = self.ext = (r, c)
        self.start_edit(initial=text, mode="edit")

    def cancel_edit(self):
        if not self.is_editing():
            return
        self._end_edit()
        self.refresh()

    def _end_edit(self):
        if self.editor is not None:
            self.editor.destroy()
            self.editor = None
        self.bar_editing = False
        self.edit_cell = self.edit_mode = self.point = None
        if self.drag and self.drag.get("mode") == "point":
            self.drag = None
        self.set_mode("Sẵn sàng")
        self.canvas.focus_set()

    # pointing at cells while typing a formula
    def _edit_widget(self):
        return self.editor if self.editor is not None else self.fx

    def _point_mode_ok(self):
        if not self.is_editing() or self.edit_cell[0] != self.sheet:
            return False
        w = self._edit_widget()
        text = w.get()
        if not text.startswith("="):
            return False
        pos = w.index("insert")
        if self.point and pos == self.point[1]:
            return True
        return text[:pos].rstrip()[-1:] in POINT_AFTER

    def _point_start(self, r, c):
        pos = self._edit_widget().index("insert")
        if not (self.point and pos == self.point[1]):
            self.point = (pos, pos)
        self.drag = {"mode": "point", "anchor": (r, c)}
        self._point_set(r, c, r, c)

    def _point_set(self, ra, ca, rb, cb):
        w = self._edit_widget()
        start, end = self.point
        r1, c1, r2, c2 = self._expand(min(ra, rb), min(ca, cb), max(ra, rb), max(ca, cb))
        ref = range_name(r1, c1, r2, c2)
        w.delete(start, end)
        w.insert(start, ref)
        self.point = (start, start + len(ref))
        w.icursor(self.point[1])
        w.focus_set()
        self.set_mode("Trỏ")

    # ------------------------------------------------------------------ changes and undo

    def commit(self, cs):
        changes = cs.changes()
        if not changes:
            return False
        for (sheet, r, c), raw0, raw1, fmt0, fmt1 in changes:  # protected sheets: locked cells keep their content
            prot = self.wb.sheets[sheet].protection
            if prot and ((raw0 != raw1 and (fe.cell_locked(fmt0) or fe.cell_locked(fmt1)))
                         or (fmt0 != fmt1 and "format_cells" not in prot.get("allow", ()))):
                messagebox.showinfo(APP_NAME, PROTECTED_MSG, parent=self.root)
                self.refresh()
                return False
        self._apply_changes(changes, forward=True)
        self._push_undo(("cells", changes, self._sel_state()))
        return True

    def _apply_changes(self, changes, forward):
        rows = {}
        if forward:
            self.wb.write_cells((sheet, r, c, raw1, fmt1) for (sheet, r, c), raw0, raw1, fmt0, fmt1 in changes)
        else:
            self.wb.write_cells((sheet, r, c, raw0, fmt0) for (sheet, r, c), raw0, raw1, fmt0, fmt1 in changes)
        for (sheet, r, c), *_ in changes:
            rows.setdefault(sheet, set()).add(r)
        for sheet, changed in rows.items():
            self.autofit_rows(changed, sheet)

    def _push_undo(self, action):
        self.undo_stack.append(action)
        del self.undo_stack[:-200]
        self.redo_stack.clear()
        self.edit_serial += 1
        self.mark_dirty()

    def _push_snapshot(self, before, sel):
        after = self.wb.snapshot()
        if after != before:
            self._share_snapshots(before, after)
            self._push_undo(("snap", before, after, sel))
            self._trim_undo()

    def _share_snapshots(self, before, after):
        """Let snapshot parts with the same content be one object: before/after of this change, and this
        change's before with the previous change's after (big sheets would otherwise be copied each time)."""
        prev = next((a for a in reversed(self.undo_stack) if a[0] == "snap"), None)
        previous = dict(prev[2][0]) if prev else {}
        old = dict(before[0])
        for name, state in after[0]:
            b, p = old.get(name), previous.get(name)
            if b is None:
                continue
            for key in fe.Sheet._FLAT:
                if p is not None and b[key] is not p[key] and b[key] == p[key]:
                    b[key] = p[key]
                if state[key] is not b[key] and state[key] == b[key]:
                    state[key] = b[key]

    UNDO_BUDGET = 6_000_000  # cells and formats kept in undo copies of whole sheets

    def _trim_undo(self):
        """Forget the oldest whole-sheet undo copies when they hold more than UNDO_BUDGET cells."""
        seen, size = set(), 0
        for i in range(len(self.undo_stack) - 1, -1, -1):
            action = self.undo_stack[i]
            if action[0] != "snap":
                continue
            for _, state in action[1][0] + action[2][0]:
                for key in ("cells", "formats"):
                    obj = state[key]
                    if id(obj) not in seen:
                        seen.add(id(obj))
                        size += len(obj)
            if size > self.UNDO_BUDGET and i < len(self.undo_stack) - 1:
                del self.undo_stack[:i]
                return

    def _structural(self, func):
        """Run a change to sheets, rows, columns, merges or rules with whole-workbook undo."""
        self.commit_edit()
        before, sel = self.wb.snapshot(), self._sel_state()
        try:
            func()
        except ValueError as err:
            messagebox.showerror(APP_NAME, str(err), parent=self.root)
            self.wb.restore(before)
            return False
        if self._protection_broken(before):
            self.wb.restore(before)
            self.refresh_all()
            return False
        self.wb.invalidate()
        self._push_snapshot(before, sel)
        self.refresh_all()
        return True

    def _protection_broken(self, before):
        """After a change: True (having told the user) when it altered a protected sheet in a way its
        protection does not allow; sheets protected or unprotected by the change itself are not checked."""
        old = dict(before[0])
        for name, sh in self.wb.sheets.items():
            state = old.get(name)
            if sh.protection and state and state.get("protection"):
                if fe.protection_problem(state, sh.copy_state(), sh.protection.get("allow", ())):
                    messagebox.showinfo(APP_NAME, PROTECTED_MSG, parent=self.root)
                    return True
        return False

    def _locked(self, sheet=None, cells=None, need=None):
        """True (after telling the user) when the sheet is protected and the cells are locked, or `need`
        (a protection option) is not allowed."""
        sh = self.wb.sheets[sheet or self.sheet]
        prot = sh.protection
        if not prot:
            return False
        if (need and need not in prot.get("allow", ())) or any(fe.cell_locked(sh.formats.get(k)) for k in cells or ()):
            messagebox.showinfo(APP_NAME, PROTECTED_MSG, parent=self.root)
            return True
        return False

    def undo(self):
        self._undo_redo(self.undo_stack, self.redo_stack, forward=False)

    def redo(self):
        self._undo_redo(self.redo_stack, self.undo_stack, forward=True)

    def _undo_redo(self, source, target, forward):
        self.commit_edit()
        if not source:
            self.root.bell()
            return
        action = source.pop()
        self.edit_serial += 1
        self.areas = []
        if action[0] == "cells":
            self._apply_changes(action[1], forward)
        else:
            self.wb.restore(action[2] if forward else action[1])
        target.append(action)
        sheet, cur, ext = action[-1]
        if sheet in self.wb.sheets:
            self._show_sheet(sheet)
            self.cur, self.ext = cur, ext
        elif self.sheet not in self.wb.sheets:
            self._show_sheet(next(iter(self.wb.sheets)))
        self.ensure_visible(*self.cur)
        self.mark_dirty()
        self.refresh_all()

    # ------------------------------------------------------------------ row heights

    def autofit_rows(self, rows, sheet=None):
        """Recompute automatic row heights from fonts, wrapped and rotated text."""
        sheet = sheet or self.sheet
        sh = self.wb.sheets[sheet]
        rows = {r for r in rows if r not in sh.custom_rows}
        if not rows:
            return
        by_row = {}
        for (r, c), fmt in sh.formats.items():
            if r in rows and any(k in fmt for k in ("size", "font", "wrap", "rotate")):
                by_row.setdefault(r, []).append((c, fmt))
        for r in rows:
            need = fe.DEFAULT_ROW_HEIGHT
            for c, fmt in by_row.get(r, ()):
                m = sh.merge_at(r, c)
                if m and m[0] != m[2]:
                    continue  # like Excel, merged blocks spanning rows don't size them
                need = max(need, self._needed_height(sheet, sh, r, c, fmt))
            if need > fe.DEFAULT_ROW_HEIGHT:
                sh.row_heights[r] = need
            else:
                sh.row_heights.pop(r, None)

    def _needed_height(self, sheet, sh, r, c, fmt):
        font = self.font_for(fmt)
        ls = font.metrics("linespace")
        need = ls + 2 * self.px
        rotate, wrap = fmt.get("rotate"), fmt.get("wrap")
        if rotate or wrap:
            text = self.wb.display_text(sheet, r, c)
            if text:
                if rotate == 255:
                    need = len(text) * ls + 4 * self.px
                elif rotate:
                    a = math.radians(abs(rotate))
                    need = font.measure(text) * math.sin(a) + ls * math.cos(a) + 6 * self.px
                else:
                    m = sh.merge_at(r, c)
                    span = range(m[1], m[3] + 1) if m else [c]
                    width = sum(self.col_w(k, sh) for k in span) - 2 * self.pad - \
                        (fmt.get("indent") or 0) * round(9 * self.px)
                    need = len(self.wrap_lines(text, font, max(4, width))) * ls + 2 * self.px
        return math.ceil(need / self.px)

    def autofit_wrapped_rows(self, col):
        rows = {r for (r, c), fmt in self.sh.formats.items() if c == col and fmt.get("wrap")}
        self.autofit_rows(rows)

    def autofit_rows_manual(self, rows):
        before = self.wb.snapshot()
        self.sh.custom_rows.difference_update(rows)
        for r in rows:
            self.sh.row_heights.pop(r, None)
        self.autofit_rows(rows)
        self._push_snapshot(before, self._sel_state())
        self.refresh()

    def autofit_selected_rows(self):
        self.commit_edit()
        r1, _, r2, _ = self._copy_rect()
        self.autofit_rows_manual(list(range(r1, r2 + 1)))

    def ask_row_height(self):
        self.commit_edit()
        r1, _, r2, _ = self.sel_rect()
        if self._is_huge(*self.sel_rect()):
            r2 = min(r2, max(self.wb.used_extent(self.sheet)[0], self.top + self._visible_rows()) - 1)
        current = self.sh.row_heights.get(r1, fe.DEFAULT_ROW_HEIGHT) * 0.75
        height = simpledialog.askfloat("Chiều cao hàng", "Chiều cao hàng (điểm):", parent=self.root,
                                       initialvalue=round(current, 2), minvalue=0, maxvalue=409)
        if height is None:
            return

        def change():
            for r in range(r1, r2 + 1):
                self.sh.row_heights[r] = max(1, round(height * 4 / 3))
                self.sh.custom_rows.add(r)
                if height == 0:
                    self.sh.hidden_rows.add(r)
        self._structural(change)

    # ------------------------------------------------------------------ clipboard

    def _copy_rect(self):
        """The selection, shrunk to the used area when it is huge."""
        r1, c1, r2, c2 = self.sel_rect()
        if self._is_huge(r1, c1, r2, c2):
            rows, cols = self.wb.used_extent(self.sheet)
            r2, c2 = min(r2, rows - 1), min(c2, cols - 1)
        return r1, c1, r2, c2

    def _format_targets(self):
        """Cells a formatting command applies to (huge selections clipped to the used/visible area)."""
        out, seen = [], set()
        for r1, c1, r2, c2 in [self.sel_rect()] + self.areas:
            if self._is_huge(r1, c1, r2, c2):
                rows, cols = self.wb.used_extent(self.sheet)
                r2 = min(r2, max(rows, self.top + self._visible_rows()) - 1)
                c2 = min(c2, max(cols, self.left + len(self.vis_cols)) - 1)
            for key in ((r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)):
                if key not in seen:
                    seen.add(key)
                    out.append(key)
        return out

    @staticmethod
    def _tsv_field(text):
        if any(ch in text for ch in '\t\n"'):
            return '"' + text.replace('"', '""') + '"'
        return text

    def copy(self, cut=False):
        self.commit_edit()
        if self.areas:
            messagebox.showinfo(APP_NAME, "Không thể sao chép hoặc cắt nhiều vùng chọn rời nhau cùng lúc.",
                                parent=self.root)
            return
        r1, c1, r2, c2 = self._copy_rect()
        if r2 < r1 or c2 < c1:
            return
        sh, cols = self.sh, range(c1, c2 + 1)
        block = [[(sh.cells.get((r, c)), sh.formats.get((r, c))) for c in cols] for r in range(r1, r2 + 1)]
        values, lines = [], []
        for r in range(r1, r2 + 1):  # each value once, for both the paste-values copy and the text
            shown = [self.wb.cell_display(self.sheet, r, c) if (r, c) in sh.cells else (None, None) for c in cols]
            values.append([fe.value_to_raw(v) for v, _ in shown])
            lines.append("\t".join(self._tsv_field(fe.format_value(*vf)) for vf in shown))
        text = "\n".join(lines)
        try:  # on Windows, also as HTML/RTF tables and a picture, like Excel (AutoCAD needs those)
            rich = (r2 - r1 + 1) * (c2 - c1 + 1) <= BIG_AREA and richclip.copy_to_clipboard(
                self.root.winfo_id(), text, self._clip_table(r1, c1, r2, c2))
        except Exception:  # noqa: BLE001 - the rich formats are a bonus; plain text must always work
            rich = False
        if not rich:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        self.clip = {"sheet": self.sheet, "rect": (r1, c1, r2, c2), "cells": block, "values": values,
                     "cut": cut, "text": text}
        self.set_mode("Chọn ô đích rồi nhấn Ctrl+V (Esc để hủy)")
        self.redraw()

    def cut(self):
        self.copy(cut=True)

    def _clip_table(self, r1, c1, r2, c2):
        """The copied block as richclip describes a table: visible rows and columns only, with each
        cell's text as displayed, its alignment and its format (conditional formatting included)."""
        sh = self.sh
        rows = [r for r in range(r1, r2 + 1) if not sh.row_hidden(r)]
        cols = [c for c in range(c1, c2 + 1) if c not in sh.hidden_cols]
        cells = []
        for r in rows:
            row = []
            for c in cols:
                value = self.wb.get_value(self.sheet, r, c)
                fmt = dict(sh.formats.get((r, c)) or {})
                style = self.wb.cond_format(self.sheet, r, c) or {}
                fmt.update({k: style[k] for k in ("bg", "fg", "bold", "italic") if k in style})
                default = ("right" if fe.is_number(value) else
                           "center" if isinstance(value, (bool, fe.CellError)) else "left")
                row.append({"text": self.wb.display_text(self.sheet, r, c), "align": fmt.get("align") or default,
                            "fmt": fmt})
            cells.append(row)
        row_at, col_at = {r: i for i, r in enumerate(rows)}, {c: j for j, c in enumerate(cols)}
        merges = []
        for m in sh.merges:
            if r1 <= m[0] and m[2] <= r2 and c1 <= m[1] and m[3] <= c2:
                mr = [row_at[r] for r in range(m[0], m[2] + 1) if r in row_at]
                mc = [col_at[c] for c in range(m[1], m[3] + 1) if c in col_at]
                if mr and mc and (len(mr) > 1 or len(mc) > 1):
                    merges.append((mr[0], mc[0], mr[-1], mc[-1]))
        return {"cols": [sh.col_widths.get(c, fe.DEFAULT_COL_WIDTH) for c in cols],
                "rows": [sh.row_heights.get(r, fe.DEFAULT_ROW_HEIGHT) for r in rows], "cells": cells, "merges": merges}

    def _paste_targets(self, h, w):
        """(row, col, i, j) for pasting an h x w block; a selection that is a multiple of
        the block is tiled like in Excel. Also selects the pasted area."""
        r1, c1, r2, c2 = self.sel_rect()
        sel_h, sel_w = r2 - r1 + 1, c2 - c1 + 1
        tile = (sel_h, sel_w) != (1, 1) and sel_h % h == 0 and sel_w % w == 0 and sel_h * sel_w <= BIG_AREA
        reps_h, reps_w = (sel_h // h, sel_w // w) if tile else (1, 1)
        targets = []
        for a in range(reps_h):
            for b in range(reps_w):
                for i in range(h):
                    for j in range(w):
                        r, c = r1 + a * h + i, c1 + b * w + j
                        if r < MAX_ROWS and c < MAX_COLS:
                            targets.append((r, c, i, j))
        self.cur = (r1, c1)
        self.ext = self._clamp(r1 + reps_h * h - 1, c1 + reps_w * w - 1)
        return targets

    def paste(self, mode="all", transpose=False, skip_blanks=False):
        """Paste the clipboard: mode is "all", "formulas", "values" or "formats"."""
        self.commit_edit()
        if self._pivot_guard(*self.sel_rect()):
            return
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            text = ""
        text = text.replace("\r\n", "\n").rstrip("\n")
        cs = ChangeSet(self.wb)
        clip = self.clip
        if clip and text == clip["text"].rstrip("\n"):
            block, values = clip["cells"], clip["values"]
            sr, sc = clip["rect"][0], clip["rect"][1]
            if transpose:
                block = [list(row) for row in zip(*block)]
                values = [list(row) for row in zip(*values)]
            moving = clip["cut"] and mode == "all" and not transpose
            if moving and clip["sheet"] in self.wb.sheets:
                for i, row in enumerate(block):
                    for j in range(len(row)):
                        cs.set(clip["sheet"], sr + i, sc + j, raw=None, fmt=None)
            for r, c, i, j in self._paste_targets(len(block), len(block[0])):
                raw, fmt = block[i][j]
                if skip_blanks and raw is None:
                    continue
                src_r, src_c = (sr + j, sc + i) if transpose else (sr + i, sc + j)
                if mode in ("all", "formulas"):
                    cs.set(self.sheet, r, c, raw=raw if moving or not raw else fe.shift_formula(raw, r - src_r, c - src_c))
                elif mode == "values":
                    cs.set(self.sheet, r, c, raw=values[i][j])
                if mode in ("all", "formats"):
                    cs.set(self.sheet, r, c, fmt=fmt)
            if moving:
                self.clip = None
                self.set_mode("Sẵn sàng")
        elif text and mode != "formats":
            rows = list(csv.reader(io.StringIO(text), delimiter="\t")) or [[""]]
            width = max(len(row) for row in rows)
            rows = [row + [""] * (width - len(row)) for row in rows]
            if transpose:
                rows = [list(row) for row in zip(*rows)]
            for r, c, i, j in self._paste_targets(len(rows), len(rows[0])):
                if not (skip_blanks and rows[i][j] == ""):
                    cs.set(self.sheet, r, c, raw=rows[i][j])
        else:
            return
        self.commit(cs)
        self.ensure_visible(*self.cur)
        self.refresh()

    def paste_special(self):
        result = dialogs.PasteSpecialDialog(self.root).show()
        if result:
            self.paste(**result)

    # ------------------------------------------------------------------ format painter

    def start_painter(self, sticky=False):
        self.commit_edit()
        if self.painter and not sticky:
            self.stop_painter()
            return
        r1, c1, r2, c2 = self._copy_rect()
        block = [[self.sh.formats.get((r, c)) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)] or [[None]]
        self.painter = {"block": block, "sticky": sticky}
        self.set_mode("Chọn ô để dán định dạng" + (" (Esc để thoát)" if sticky else ""))
        self.update_toolbar()

    def stop_painter(self):
        self.painter = None
        self.set_mode("Sẵn sàng")
        self.update_toolbar()

    def _apply_painter(self):
        block = self.painter["block"]
        h, w = len(block), len(block[0])
        r1, c1, r2, c2 = self.sel_rect()
        if (r1, c1, r2, c2) == self._cell_box(*self.cur):
            r2, c2 = min(MAX_ROWS - 1, r1 + h - 1), min(MAX_COLS - 1, c1 + w - 1)
        elif self._is_huge(r1, c1, r2, c2):
            rows, cols = self.wb.used_extent(self.sheet)
            r2 = min(r2, max(rows, self.top + self._visible_rows()) - 1)
            c2 = min(c2, max(cols, self.left + len(self.vis_cols)) - 1)
        cs = ChangeSet(self.wb)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cs.set(self.sheet, r, c, fmt=block[(r - r1) % h][(c - c1) % w])
        self.commit(cs)
        self.ext = (r2, c2) if self.cur == (r1, c1) else self.ext
        if not self.painter["sticky"]:
            self.stop_painter()

    # ------------------------------------------------------------------ cell formatting

    def apply_format(self, update):
        """Merge `update` into the format of every selected cell (None/False removes a key). Cells that had
        the same format get one shared new format; very large selections are written in one go."""
        self.commit_edit()
        formats = self.sh.formats
        made = {}

        def merged(old):
            entry = made.get(id(old))
            if entry is None:
                fmt = dict(old or {})
                for key, value in update.items():
                    if value in (None, False, ""):
                        fmt.pop(key, None)
                    else:
                        fmt[key] = value
                entry = made[id(old)] = (old, fmt or None)  # (old kept alive: its id stays its own)
            return entry[1]
        targets = self._format_targets()
        if len(targets) <= 50000:
            cs = ChangeSet(self.wb)
            for r, c in targets:
                cs.set(self.sheet, r, c, fmt=merged(formats.get((r, c))))
            self.commit(cs)
        else:
            def change():
                with fe.bulk():
                    new = {key: merged(formats.get(key)) for key in targets}
                    kept = {k: v for k, v in formats.items() if k not in new}
                    self.sh.formats = {**kept, **{k: v for k, v in new.items() if v}}
                if any(k in update for k in ("size", "wrap", "rotate", "font")):
                    self.autofit_rows({r for r, _ in targets})
            self._structural(change)
        self.refresh()
        self.canvas.focus_set()

    def _active_fmt(self):
        return self.sh.formats.get(self.cur) or {}

    def toggle_style(self, key):
        self.apply_format({key: not self._active_fmt().get(key)})

    def set_underline(self, kind):
        current = self._active_fmt().get("underline")
        current = "single" if current is True else current
        self.apply_format({"underline": None if current == kind else kind})

    def set_align(self, align):
        self.apply_format({"align": None if self._active_fmt().get("align") == align else align})

    def set_valign(self, valign):
        self.apply_format({"valign": valign})

    def toggle_wrap(self):
        self.apply_format({"wrap": not self._active_fmt().get("wrap")})

    def set_rotation(self, angle):
        self.apply_format({"rotate": None if self._active_fmt().get("rotate") == angle else angle})

    def change_indent(self, delta):
        fmt = self._active_fmt()
        indent = max(0, min(15, (fmt.get("indent") or 0) + delta))
        update = {"indent": indent or None}
        if indent and fmt.get("align") not in ("left", "right"):
            update["align"] = "left"
        self.apply_format(update)

    def _apply_font_box(self):
        family = self.font_var.get().strip()
        if family:
            self.apply_format({"font": None if family == self.default_family else family})

    def _apply_size_box(self):
        try:
            size = float(self.size_var.get().replace(",", "."))
        except ValueError:
            self.root.bell()
            return
        size = max(1.0, min(409.0, size))
        size = int(size) if size.is_integer() else size
        self.apply_format({"size": None if size == 11 else size})

    def step_font_size(self, direction):
        size = self._active_fmt().get("size") or 11
        if direction > 0:
            new = next((s for s in FONT_SIZES if s > size), size + 10)
        else:
            new = next((s for s in reversed(FONT_SIZES) if s < size), max(1, size - 1))
        self.apply_format({"size": None if new == 11 else new})

    def _color_dropdown(self, button, key):
        color_popup(button, self.ui, lambda c: self._pick_color(key, c),
                    "Không tô màu" if key == "bg" else "Tự động",
                    lambda: self._ask_color(lambda c: self._pick_color(key, c)))

    def _pick_color(self, key, color):
        if color:
            if key == "bg":
                self.last_fill = color
                self.btn_fill.set_color(color)
            else:
                self.last_fg = color
                self.btn_fg.set_color(color)
        self.apply_format({key: color})

    def _on_numfmt(self, e):
        label = self.numfmt_box.get()
        if label == MORE_FORMATS:
            self.format_cells_dialog("number")
        else:
            formats = dict(NUMBER_FORMATS)
            if label in formats:
                self.apply_format({"num": formats[label]})
        self.canvas.focus_set()

    def change_decimals(self, delta):
        value, numfmt = self.wb.cell_display(self.sheet, *self.cur)
        self.apply_format({"num": fe.adjust_decimals(numfmt, value, delta)})

    def apply_border(self, kind):
        self.commit_edit()
        self.last_border = kind
        if self.btn_border.icon != f"border_{kind}":  # like Excel, the button repeats the last border used
            self.btn_border.icon = f"border_{kind}"
            self.btn_border.draw()
        targets = self._format_targets()
        if not targets:
            return
        r1, c1 = targets[0]
        r2, c2 = targets[-1]
        line = [self.border_style, self.border_color]
        thin, thick, double = ["thin", self.border_color], ["thick", self.border_color], ["double", self.border_color]
        cs = ChangeSet(self.wb)
        for r, c in targets:
            fmt = dict(self.sh.formats.get((r, c)) or {})
            top, bottom, left, right = r == r1, r == r2, c == c1, c == c2
            sets = {}
            if kind == "none":
                for key in fe.BORDER_KEYS:
                    fmt.pop(key, None)
            elif kind == "all":
                sets = dict.fromkeys(fe.BORDER_KEYS, line)
            elif kind in ("outside", "thick_outside"):
                v = thick if kind == "thick_outside" else line
                sets = {k: v for k, on in (("bt", top), ("bb", bottom), ("bl", left), ("br", right)) if on}
            elif kind in ("bottom", "top", "left", "right"):
                key = {"bottom": "bb", "top": "bt", "left": "bl", "right": "br"}[kind]
                if {"bottom": bottom, "top": top, "left": left, "right": right}[kind]:
                    sets = {key: line}
            elif kind == "bottom_double" and bottom:
                sets = {"bb": double}
            elif kind == "thick_bottom" and bottom:
                sets = {"bb": thick}
            elif kind in ("top_bottom", "top_thick_bottom", "top_double_bottom"):
                if top:
                    sets["bt"] = line if kind == "top_bottom" else thin
                if bottom:
                    sets["bb"] = {"top_bottom": line, "top_thick_bottom": thick, "top_double_bottom": double}[kind]
            fmt.update(sets)
            cs.set(self.sheet, r, c, fmt=fmt)
        self.commit(cs)
        self.refresh()

    def _apply_border_spec(self, borders):
        """Borders from the Format Cells dialog: outline edges and inside lines of the selection."""
        targets = self._format_targets()
        if not targets:
            return
        r1, c1 = targets[0]
        r2, c2 = targets[-1]
        cs = ChangeSet(self.wb)
        for r, c in targets:
            fmt = dict(self.sh.formats.get((r, c)) or {})
            spec = {"bt": borders["top"] if r == r1 else borders["inside_h"],
                    "bb": borders["bottom"] if r == r2 else borders["inside_h"],
                    "bl": borders["left"] if c == c1 else borders["inside_v"],
                    "br": borders["right"] if c == c2 else borders["inside_v"]}
            for key, value in spec.items():
                if value:
                    fmt[key] = value
                else:
                    fmt.pop(key, None)
            cs.set(self.sheet, r, c, fmt=fmt)
        self.commit(cs)

    def format_cells_dialog(self, tab="number"):
        self.commit_edit()
        r1, c1, r2, c2 = self.sel_rect()
        value = self.wb.get_value(self.sheet, *self.cur)
        merged = self.sh.merge_at(*self.cur) is not None
        dlg = dialogs.FormatCellsDialog(self.root, self.ui, self._active_fmt(), value if value is not None else "",
                                        self.font_families, self.default_family, tab, merged,
                                        multi_rows=r2 > r1, multi_cols=c2 > c1)
        changes = dlg.show()
        if not changes:
            return
        borders = changes.pop("borders", None)
        merge = changes.pop("merge", None)
        if changes:
            self.apply_format(changes)
        if borders:
            self._apply_border_spec(borders)
        if merge is True:
            self.merge()
        elif merge is False:
            self.unmerge()
        self.refresh()

    def apply_cell_style(self, name):
        style = STYLE_LOOKUP[name]
        self.commit_edit()
        cs = ChangeSet(self.wb)
        for r, c in self._format_targets():
            fmt = dict(self.sh.formats.get((r, c)) or {})
            if not style:
                fmt = {}
            elif set(style) == {"num"}:
                fmt["num"] = style["num"]
            else:
                fmt = {k: v for k, v in fmt.items() if k not in STYLE_KEYS}
                fmt.update(style)
            cs.set(self.sheet, r, c, fmt=fmt)
        self.commit(cs)
        self.refresh()

    def _style_preview(self, style):
        borders = [style.get(k) for k in fe.BORDER_KEYS]
        preview = {k: style.get(k) for k in ("bg", "fg", "bold", "italic")}
        if all(borders):
            preview["border"] = borders[0][1]
            preview["double"] = borders[0][0] == "double"
        elif style.get("bb"):
            preview["underline"] = style["bb"][1]
        if style.get("num"):
            preview["fg"] = "#444444"
        return preview

    def _styles_popup(self, button):
        sections = [(title, [(name, self._style_preview(style), name) for name, style in items])
                    for title, items in CELL_STYLES]
        gallery_popup(button, self.ui, sections, self.apply_cell_style)

    def _table_gallery(self, button):
        def preview(style):
            _, head, head_fg, band, line = style

            def draw(cv, w, h):
                rows, cols = 5, 4
                for i in range(rows):
                    fill = head if i == 0 else (band if i % 2 else "#ffffff")
                    y0 = 3 + i * (h - 6) / rows
                    cv.create_rectangle(4, y0, w - 4, y0 + (h - 6) / rows, fill=fill, outline=line)
            return draw

        items = [(style[0], preview(style), style) for style in TABLE_STYLES]
        gallery_popup(button, self.ui, [("Kiểu bảng", items)], self.format_as_table, tile_w=self.ui.px(62))

    def format_as_table(self, style):
        self.commit_edit()
        rect = self.sel_rect()
        if rect == self._cell_box(*self.cur):
            rect = self.current_region(*self.cur)
        else:
            rect = self._copy_rect()
        result = dialogs.TableDialog(self.root, range_name(*rect), True).show()
        if not result:
            return
        (r1, c1, r2, c2), header = result
        _, head, head_fg, band, line = style

        def apply():
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    fmt = {k: v for k, v in (self.sh.formats.get((r, c)) or {}).items() if k not in ("bg", "fg", "bold")}
                    if header and r == r1:
                        fmt.update(bg=head, fg=head_fg, bold=True)
                    elif (r - r1 - (1 if header else 0)) % 2 == 0:
                        fmt["bg"] = band
                    fmt["bt"] = fmt["bb"] = ["thin", line]
                    self.wb.set_format(self.sheet, r, c, fmt)
            if header:
                self.sh.filter = {"range": (r1, c1, r2, c2), "criteria": {}, "hidden": set()}
        self._structural(apply)

    # merging
    def merge(self, center=False, across=False):
        self.commit_edit()
        r1, c1, r2, c2 = self.sel_rect()
        if (r1, c1) == (r2, c2) or (across and c1 == c2):
            return
        if self._is_huge(r1, c1, r2, c2):
            messagebox.showinfo(APP_NAME, "Vùng chọn quá lớn để hợp nhất.", parent=self.root)
            return
        regions = [(r, c1, r, c2) for r in range(r1, r2 + 1)] if across else [(r1, c1, r2, c2)]
        lost = [k for k in self.sh.cells for reg in regions
                if reg[0] <= k[0] <= reg[2] and reg[1] <= k[1] <= reg[3] and k != (reg[0], reg[1])]
        if lost and not messagebox.askokcancel(
                APP_NAME, "Hợp nhất ô chỉ giữ lại giá trị ở ô trên cùng bên trái và bỏ các giá trị khác.",
                parent=self.root):
            return

        def apply():
            sh = self.sh
            for key in lost:
                self.wb.set_raw(self.sheet, *key, None)
            keep = [m for m in sh.merges if not (m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1)]
            sh.set_merges(keep + regions)
            if center:
                for reg in regions:
                    fmt = dict(sh.formats.get((reg[0], reg[1])) or {}, align="center")
                    self.wb.set_format(self.sheet, reg[0], reg[1], fmt)
        self._structural(apply)
        self.cur, self.ext = (r1, c1), (r2, c2)
        self.refresh()

    def unmerge(self):
        r1, c1, r2, c2 = self.sel_rect()
        self._structural(lambda: self.sh.set_merges(
            [m for m in self.sh.merges if not (m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1)]))

    def merge_center(self):
        r1, c1, r2, c2 = self.sel_rect()
        if any(m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1 for m in self.sh.merges):
            self.unmerge()
        else:
            self.merge(center=True)

    # conditional formatting
    def cf_rule(self, kind, op=None, percent=False):
        self.commit_edit()
        r1, c1, r2, c2 = self._copy_rect()
        nums = [v for r, c in list(self.sh.cells) if r1 <= r <= r2 and c1 <= c <= c2
                for v in [self.wb.get_value(self.sheet, r, c)] if fe.is_number(v)]
        default = fe.format_general(sum(nums) / len(nums)) if nums else ""
        prompts = {">": "Định dạng các ô LỚN HƠN:", "<": "Định dạng các ô NHỎ HƠN:", "between": "Định dạng các ô nằm GIỮA:",
                   "=": "Định dạng các ô BẰNG:"}
        titles = {">": "Lớn hơn", "<": "Nhỏ hơn", "between": "Nằm giữa", "=": "Bằng"}
        fields, choice = [], None
        if kind == "cell":
            title, prompt = titles[op], prompts[op]
            fields = [{"default": default}] + ([{"label": "và", "default": default}] if op == "between" else [])
        elif kind == "text":
            title, prompt, fields = "Văn bản có chứa", "Định dạng các ô có chứa văn bản:", [{"default": ""}]
        elif kind == "dup":
            title, prompt, choice = "Giá trị trùng lặp", "Định dạng các ô chứa:", ["Trùng lặp", "Duy nhất"]
        elif kind in ("top", "bottom"):
            title = f"10{'%' if percent else ' mục'} {'đầu' if kind == 'top' else 'cuối'}"
            prompt = f"Định dạng các ô xếp ở {'ĐẦU' if kind == 'top' else 'CUỐI'}:"
            fields = [{"default": "10", "spin": True, "after": "%" if percent else "mục"}]
        else:
            title = "Trên trung bình" if kind == "above" else "Dưới trung bình"
            prompt = f"Định dạng các ô {'TRÊN' if kind == 'above' else 'DƯỚI'} MỨC TRUNG BÌNH của vùng chọn"
        result = dialogs.RuleDialog(self.root, title, prompt, fields, choice).show()
        if not result:
            return
        rule = {"type": kind, "style": result["style"]}
        values = result["values"]
        if kind == "cell":
            rule.update(op=op, v1=values[0])
            if op == "between":
                rule["v2"] = values[1]
        elif kind == "text":
            rule["text"] = values[0]
        elif kind == "dup":
            rule["type"] = "dup" if result["choice"] == "Trùng lặp" else "unique"
        elif kind in ("top", "bottom"):
            try:
                rule.update(n=max(1, int(float(values[0]))), percent=percent)
            except ValueError:
                return
        self.cf_add(rule)

    def cf_add(self, rule):
        rule = dict(rule, range=self.sel_rect())
        self._structural(lambda: self.wb.set_cond_formats(self.sheet, self.sh.cond_formats + [rule]))

    def cf_clear_selection(self):
        r1, c1, r2, c2 = self.sel_rect()
        keep = [rule for rule in self.sh.cond_formats
                if not (rule["range"][0] <= r2 and rule["range"][2] >= r1 and rule["range"][1] <= c2 and rule["range"][3] >= c1)]
        self._structural(lambda: self.wb.set_cond_formats(self.sheet, keep))

    def cf_clear_sheet(self):
        self._structural(lambda: self.wb.set_cond_formats(self.sheet, []))

    def describe_rule(self, rule):
        kind, where = rule["type"], range_name(*rule["range"])
        ops = {">": "lớn hơn", "<": "nhỏ hơn", ">=": "≥", "<=": "≤", "=": "bằng", "<>": "khác",
               "between": "nằm giữa", "notbetween": "không nằm giữa"}
        if kind == "cell":
            text = f"Giá trị ô {ops.get(rule['op'], rule['op'])} {rule['v1']}"
            if "between" in rule["op"]:
                text += f" và {rule.get('v2')}"
        else:
            text = {"text": f"Văn bản chứa “{rule.get('text')}”", "dup": "Giá trị trùng lặp", "unique": "Giá trị duy nhất",
                    "top": f"{rule.get('n')}{'%' if rule.get('percent') else ''} đầu",
                    "bottom": f"{rule.get('n')}{'%' if rule.get('percent') else ''} cuối",
                    "above": "Trên trung bình", "below": "Dưới trung bình", "databar": "Thanh dữ liệu",
                    "scale": "Thang màu"}.get(kind, kind)
        return f"{text}    —  áp dụng cho {where}"

    def cf_manage(self):
        rules = dialogs.ManageRulesDialog(self.root, self.sh.cond_formats, self.describe_rule).show()
        if rules is not None:
            self._structural(lambda: self.wb.set_cond_formats(self.sheet, rules))

    # ------------------------------------------------------------------ cell commands

    def clear_contents(self):
        self.commit_edit()
        rects = [self.sel_rect()] + self.areas
        if any(self._pivot_guard(*rect) for rect in rects):
            return
        cs = ChangeSet(self.wb)
        for r1, c1, r2, c2 in rects:
            for r, c in [(r, c) for r, c in self.sh.cells if r1 <= r <= r2 and c1 <= c <= c2]:
                cs.set(self.sheet, r, c, raw=None)
        self.commit(cs)
        self.refresh()

    def clear_formats(self, contents=False):
        self.commit_edit()
        r1, c1, r2, c2 = self.sel_rect()

        def clear():
            sh = self.sh
            for key in list(sh.formats) + (list(sh.cells) if contents else []):
                if r1 <= key[0] <= r2 and c1 <= key[1] <= c2:
                    self.wb.set_format(self.sheet, *key, None)
                    if contents:
                        self.wb.set_raw(self.sheet, *key, None)
            sh.set_merges([m for m in sh.merges if not (m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1)])
            if contents:  # Clear All takes the links too
                sh.links = {k: v for k, v in sh.links.items() if not (r1 <= k[0] <= r2 and c1 <= k[1] <= c2)}
            self.autofit_rows({r for r in sh.row_heights if r1 <= r <= r2})
        self._structural(clear)

    def clear_all(self):
        self.clear_formats(contents=True)

    def fill_down(self):
        self._fill_from_edge(vertical=True)

    def fill_right(self):
        self._fill_from_edge(vertical=False)

    def fill_up(self):
        self._fill_from_edge(vertical=True, reverse=True)

    def fill_left(self):
        self._fill_from_edge(vertical=False, reverse=True)

    def _fill_from_edge(self, vertical, reverse=False):
        """Ctrl+D / Ctrl+R (and up/left): copy the first row (column) of the selection over the rest."""
        self.commit_edit()
        r1, c1, r2, c2 = self._copy_rect()
        if r2 < r1 or c2 < c1:
            return
        sh = self.sh
        cs = ChangeSet(self.wb)
        lo, hi = (r1, r2) if vertical else (c1, c2)
        if reverse:
            if lo == hi:
                return
            src, targets = hi, range(lo, hi)
        elif lo == hi:
            if lo == 0:
                return
            src, targets = lo - 1, [lo]
        else:
            src, targets = lo, range(lo + 1, hi + 1)
        for k in (range(c1, c2 + 1) if vertical else range(r1, r2 + 1)):
            pos = (src, k) if vertical else (k, src)
            raw, fmt = sh.cells.get(pos), sh.formats.get(pos)
            for t in targets:
                r, c = (t, k) if vertical else (k, t)
                cs.set(self.sheet, r, c, raw=fe.shift_formula(raw, r - pos[0], c - pos[1]) if raw else None, fmt=fmt)
        self.commit(cs)
        self.refresh()

    def fill_series_dialog(self):
        self.commit_edit()
        r1, c1, r2, c2 = self._copy_rect()
        result = dialogs.SeriesDialog(self.root, "col" if r2 - r1 >= c2 - c1 else "row").show()
        if not result:
            return
        vertical, step, stop = result["axis"] == "col", result["step"], result["stop"]
        lines = range(c1, c2 + 1) if vertical else range(r1, r2 + 1)
        length = (r2 - r1 + 1) if vertical else (c2 - c1 + 1)
        cs = ChangeSet(self.wb)
        for k in lines:
            start = (r1, k) if vertical else (k, c1)
            raw = self.sh.cells.get(start)
            if raw is None or fe.is_formula(raw):
                continue
            value, implied = fe.parse_literal(raw)
            if not fe.is_number(value):
                continue
            n = length if length > 1 or stop is None else MAX_ROWS
            for i in range(1, n):
                v = value * step ** i if result["kind"] == "growth" else value + step * i
                if stop is not None and ((step >= 0 and v > stop) or (step < 0 and v < stop)):
                    break
                r, c = (start[0] + i, k) if vertical else (k, start[1] + i)
                if r >= MAX_ROWS or c >= MAX_COLS:
                    break
                is_date = result["kind"] == "date" or (implied and fe.is_date_format(implied))
                cs.set(self.sheet, r, c, raw=fe.format_date(v, "dd/mm/yyyy") if is_date else fe.number_text(v))
        self.commit(cs)
        self.refresh()

    def _fill_target(self, r, c):
        """Area the fill handle covers when dragged to (r, c): the selection grown along one axis."""
        r1, c1, r2, c2 = self.sel_rect()
        down, up, right, left = r - r2, r1 - r, c - c2, c1 - c
        best = max(down, up, right, left)
        if best <= 0:
            return None
        if best == down:
            return r1, c1, r, c2
        if best == up:
            return r, c1, r2, c2
        if best == right:
            return r1, c1, r2, c
        return r1, c, r2, c2

    def do_fill(self, target):
        r1, c1, r2, c2 = self.sel_rect()
        t1, tc1, t2, tc2 = target
        sh = self.sh
        vertical = (tc1, tc2) == (c1, c2)
        if vertical:
            lines = [([(r, c) for r in range(r1, r2 + 1)], [(r, c) for r in range(t1, t2 + 1)]) for c in range(c1, c2 + 1)]
        else:
            lines = [([(r, c) for c in range(c1, c2 + 1)], [(r, c) for c in range(tc1, tc2 + 1)]) for r in range(r1, r2 + 1)]
        cs = ChangeSet(self.wb)
        for src, dst in lines:
            raws = [sh.cells.get(p) for p in src]
            series = fe.fill_series(raws)
            origin, n = src[0], len(src)
            for p in dst:
                k = p[0] - origin[0] if vertical else p[1] - origin[1]
                if 0 <= k < n:
                    continue
                i = k % n
                raw = series(k) if series else raws[i] and fe.shift_formula(raws[i], p[0] - src[i][0], p[1] - src[i][1])
                cs.set(self.sheet, p[0], p[1], raw=raw, fmt=sh.formats.get(src[i]))
        self.commit(cs)
        self.cur, self.ext = (t1, tc1), (t2, tc2)
        self.refresh()

    def autosum(self, func="SUM"):
        self.commit_edit()
        r1, c1, r2, c2 = self.sel_rect()

        def value(r, c):
            return self.wb.get_value(self.sheet, r, c)
        if (r1, c1, r2, c2) == self._cell_box(*self.cur):
            r, c = self.cur
            top, left = r, c
            while top > 0 and fe.is_number(value(top - 1, c)):
                top -= 1
            while left > 0 and fe.is_number(value(r, left - 1)):
                left -= 1
            if top < r:
                text = f"={func}({cell_name(top, c)}:{cell_name(r - 1, c)})"
            elif left < c:
                text = f"={func}({cell_name(r, left)}:{cell_name(r, c - 1)})"
            else:
                text = f"={func}()"
            self.start_edit(initial=text, mode="edit")
            if text.endswith("()"):
                self.editor.icursor(len(text) - 1)
            return
        r1, c1, r2, c2 = self._copy_rect()
        if r2 + 1 >= MAX_ROWS or r2 < r1:
            return
        cs = ChangeSet(self.wb)
        for c in range(c1, c2 + 1):
            cs.set(self.sheet, r2 + 1, c, raw=f"={func}({cell_name(r1, c)}:{cell_name(r2, c)})")
        self.commit(cs)
        self.refresh()

    def insert_function_dialog(self):
        name = dialogs.FunctionDialog(self.root).show()
        if name:
            self.insert_function(name)

    def insert_function(self, name):
        if self.is_editing():
            w = self._edit_widget()
            w.insert("insert", f"{name}(" if w.get().startswith("=") else f"={name}(")
            w.focus_set()
        else:
            self.start_edit(initial=f"={name}(", mode="edit")

    def insert_now(self, with_time):
        now = dt.datetime.now()
        text = now.strftime("%H:%M") if with_time else now.strftime("%d/%m/%Y")
        if self.is_editing():
            self._edit_widget().insert("insert", text)
            return
        cs = ChangeSet(self.wb)
        if with_time:
            cs.set(self.sheet, *self.cur, raw=fe.number_text(round((now.hour * 60 + now.minute) / 1440, 10)),
                   fmt=dict(self._active_fmt(), num="hh:mm"))
        else:
            cs.set(self.sheet, *self.cur, raw=text)
        self.commit(cs)
        self.refresh()

    def toggle_show_formulas(self):
        self.show_formulas = not self.show_formulas
        self.refresh()

    def recalculate(self):
        self.wb.invalidate()
        self.refresh()

    # outline: Data > Group / Ungroup / Subtotal
    def _outline_target(self, verb):
        """(axis, first, last) to group: whole rows or columns selected decide, otherwise ask like Excel."""
        r1, c1, r2, c2 = self.sel_rect()
        if r1 == 0 and r2 == MAX_ROWS - 1 and not (c1 == 0 and c2 == MAX_COLS - 1):
            return 1, c1, c2
        if c1 == 0 and c2 == MAX_COLS - 1:
            return (0, r1, r2) if r2 - r1 < 100000 else None
        axis = dialogs.ChoiceDialog(self.root, verb, [("Hàng", "rows"), ("Cột", "cols")], "rows").show()
        if axis is None:
            return None
        return (0, r1, r2) if axis == "rows" else (1, c1, c2)

    def group_selection(self, delta=1):
        """Data > Group (Alt+Shift+Right) / Ungroup (Alt+Shift+Left)."""
        self.commit_edit()
        target = self._outline_target("Nhóm" if delta > 0 else "Bỏ nhóm")
        if target is None:
            return
        done = {}

        def change():
            done["changed"] = self.wb.group(self.sheet, *target, delta)
        self._structural(change)
        if not done.get("changed"):
            self.root.bell()

    def clear_outline(self):
        self.commit_edit()
        self._structural(lambda: (self.wb.clear_outline(self.sheet, 0), self.wb.clear_outline(self.sheet, 1)))

    def outline_action(self, action):
        """A click on an outline symbol: ("level", axis, n) or ("toggle", axis, level, first, last, collapse)."""
        axis = action[1]
        if action[0] == "level":
            self._structural(lambda: self.wb.show_level(self.sheet, axis, action[2]))
        else:
            _, _, level, a, b, collapse = action
            self._structural(lambda: self.wb.set_collapsed(self.sheet, axis, level, a, b, collapse))
        hidden = self.sh.row_hidden(self.cur[0]) or self.cur[1] in self.sh.hidden_cols
        if hidden:  # the active cell went into a collapsed group: stand on the group's summary instead
            levels = self.sh.row_levels if axis == 0 else self.sh.col_levels
            index = self.cur[axis]
            end = max((b for _, a, b in fe.outline_groups(levels) if a <= index <= b), default=index)
            self.cur = self.ext = (end + 1, self.cur[1]) if axis == 0 else (self.cur[0], end + 1)
            self.refresh()

    def outline_detail(self, show):
        """Data > Show Detail / Hide Detail: expand or collapse the innermost group at the active cell."""
        self.commit_edit()
        r1, _, r2, _ = self.sel_rect()
        axes = (1,) if r1 == 0 and r2 == MAX_ROWS - 1 else (0, 1)
        for axis in axes:
            index = self.cur[axis]
            sh = self.sh
            levels, hidden = (sh.row_levels, sh.hidden_rows) if axis == 0 else (sh.col_levels, sh.hidden_cols)
            groups = [(lv, a, b) for lv, a, b in fe.outline_groups(levels)
                      if a <= index <= b + 1 and (a in hidden and b in hidden) == show]
            if groups:
                lv, a, b = max(groups)
                self.outline_action(("toggle", axis, lv, a, b, not show))
                return
        self.root.bell()

    def subtotal_dialog(self):
        """Data > Subtotal: totals after each group of equal values, with an outline to fold the table."""
        self.commit_edit()
        rect = self.sel_rect()
        if rect == self._cell_box(*self.cur):
            rect = self.current_region(*self.cur)
        if rect[2] <= rect[0]:
            messagebox.showinfo(APP_NAME, "Hãy chọn một ô trong bảng dữ liệu có hàng tiêu đề.", parent=self.root)
            return
        if self._pivot_guard(*rect):
            return
        columns = [(c, self.wb.display_text(self.sheet, rect[0], c)) for c in range(rect[1], rect[3] + 1)]
        result = dialogs.SubtotalDialog(self.root, columns).show()
        if not result:
            return
        out = {}

        def change():
            if result.get("remove"):
                out["rect"] = self.wb.remove_subtotals(self.sheet, rect)
            else:
                out["rect"] = self.wb.add_subtotals(self.sheet, rect, result["by"], result["func"],
                                                    result["columns"], result["replace"])
        if self._structural(change):
            r1, c1, r2, c2 = out["rect"]
            self.cur, self.ext = (r1, c1), (r2, c2)
            self.refresh()

    def import_text(self):
        """Data > From Text/CSV: read a text file with the encoding and delimiter the user picks, into a new sheet."""
        self.commit_edit()
        path = filedialog.askopenfilename(parent=self.root, title="Nhập từ Văn bản/CSV", filetypes=[
            ("Văn bản / CSV", "*.csv *.txt *.tsv *.prn"), ("Tất cả các tệp", "*.*")])
        if not path:
            return
        try:
            with open(path, "rb") as f:
                data = f.read()
        except OSError as err:
            messagebox.showerror(APP_NAME, f"Không mở được tệp:\n{path}\n\n{err}", parent=self.root)
            return
        name = os.path.splitext(os.path.basename(path))[0]
        options = dialogs.TextImportDialog(self.root, data, os.path.basename(path)).show()
        if options is None:
            return
        self._busy(f"Đang nhập {os.path.basename(path)}...")
        rows = fe.read_text_table(data, **options)[0]
        created = []

        def add():
            title, i = name[:31] or "Nhập", 2
            while self.wb.resolve_sheet(title) or self.wb.check_sheet_name(title):
                title = f"{name[:26]} ({i})"
                i += 1
                if i > 999:
                    title = None
                    break
            sheet = self.wb.add_sheet(title)
            self.wb.sheets[sheet].cells = fe.table_cells(rows)
            created.append(sheet)
        if self._structural(add):
            self.switch_sheet(created[0])
            self.autofit_columns(0, min(fe.MAX_COLS - 1, max((len(r) for r in rows[:200]), default=1) - 1))
        self._not_busy()
        self.set_mode(f"Đã nhập {len(rows):,} hàng")

    def text_to_columns(self):
        """Data > Text to Columns: split the text of one column into the columns to its right."""
        self.commit_edit()
        r1, c1, r2, c2 = self.sel_rect()
        if c1 != c2:
            messagebox.showinfo(APP_NAME, "Hãy chọn các ô trong một cột để tách.", parent=self.root)
            return
        if self._is_huge(r1, c1, r2, c2):
            r2 = min(r2, self.wb.used_extent(self.sheet)[0] - 1)
        rows = [r for r in range(r1, r2 + 1) if (r, c1) in self.sh.cells and not fe.is_formula(self.sh.cells[(r, c1)])]
        if not rows:
            self.root.bell()
            return
        texts = {r: self.wb.display_text(self.sheet, r, c1) for r in rows}
        opts = dialogs.TextToColumnsDialog(self.root, [texts[r] for r in rows]).show()
        if not opts:
            return
        parts = {r: [p.strip() for p in fe.split_text(texts[r], opts["delimiters"], opts["merge"], opts["breaks"])]
                 for r in rows}
        width = max(len(p) for p in parts.values())
        if c1 + width > MAX_COLS:
            self.root.bell()
            return
        if any((r, c) in self.sh.cells for r in rows for c in range(c1 + 1, c1 + width)) and not messagebox.askyesno(
                APP_NAME, "Ở các cột bên phải đã có dữ liệu. Bạn có muốn thay thế không?", parent=self.root):
            return
        if self._pivot_guard(r1, c1, r2, c1 + width - 1):
            return
        cs = ChangeSet(self.wb)
        for r, values in parts.items():
            for k in range(width):
                value = values[k] if k < len(values) else ""
                if opts["as_text"] and value and fe.parse_literal(value)[0] != value:
                    value = "'" + value  # kept as typed: phone numbers keep their leading 0
                cs.set(self.sheet, r, c1 + k, raw=value or None)
        self.commit(cs)
        self.cur, self.ext = (r1, c1), (r2, c1 + width - 1)
        self.refresh()

    def go_to_special(self):
        """Select the cells of one kind (blanks, constants, formulas...) in the selection or the whole sheet."""
        self.commit_edit()
        found = dialogs.GoToSpecialDialog(self.root).show()
        if not found:
            return
        kind, types = found
        rect = self.sel_rect()
        single = rect == self._cell_box(*self.cur) and not self.areas
        if kind == "region":
            r1, c1, r2, c2 = self.current_region(*self.cur)
            self.areas, self.cur, self.ext = [], (r1, c1), (r2, c2)
            self.refresh()
            return
        scope = (0, 0, MAX_ROWS - 1, MAX_COLS - 1) if single or kind == "last" else rect
        cells = self.wb.special_cells(self.sheet, scope, kind, types)
        if not cells:
            messagebox.showinfo(APP_NAME, "Không tìm thấy ô nào.", parent=self.root)
            return
        rects = fe.cells_to_rects(cells)[:20000]
        first = rects[0]
        self.cur, self.ext, self.areas = (first[0], first[1]), (first[2], first[3]), rects[1:]
        self.ensure_visible(*self.cur)
        self.refresh()
        self.set_mode(f"Đã chọn {len(cells)} ô ({len(rects)} vùng)")

    # sorting and filtering
    def current_region(self, r, c):
        """Bounding box of the block of filled cells around (r, c), like Excel's Ctrl+*.

        Rows are grown first, a row at a time (cheap: one row is a few lookups), then the columns are
        checked once along the whole height; repeated until nothing grows, so tall tables stay fast."""
        cells = self.sh.cells
        r1 = r2 = r
        c1 = c2 = c
        while True:
            cols = range(max(0, c1 - 1), min(MAX_COLS, c2 + 2))
            while r1 > 0 and any((r1 - 1, k) in cells for k in cols):
                r1 -= 1
            while r2 < MAX_ROWS - 1 and any((r2 + 1, k) in cells for k in cols):
                r2 += 1
            rows = range(max(0, r1 - 1), min(MAX_ROWS, r2 + 2))
            grew = False
            while c1 > 0 and any((k, c1 - 1) in cells for k in rows):
                c1, grew = c1 - 1, True
            while c2 < MAX_COLS - 1 and any((k, c2 + 1) in cells for k in rows):
                c2, grew = c2 + 1, True
            if not grew:
                return r1, c1, r2, c2

    def _has_header(self, r1, c1, r2, c2):
        """Guess, like Excel, whether the first row of a block holds column titles: all text, and
        formatted unlike the next row, or above numbers (numbers typed as text count), or the
        first row of a filter."""
        if r2 <= r1:
            return False
        first = [self.wb.get_value(self.sheet, r1, c) for c in range(c1, c2 + 1)]
        if not all(isinstance(v, str) and v for v in first):
            return False
        f = self.sh.filter
        if f and f["range"][0] == r1 and f["range"][1] <= c1 and c2 <= f["range"][3]:
            return True

        def look(r, c):
            fmt = self.sh.formats.get((r, c)) or {}
            return bool(fmt.get("bold")), fmt.get("bg"), bool(fmt.get("bb"))
        if any(look(r1, c) != look(r1 + 1, c) for c in range(c1, c2 + 1)):
            return True
        second = [self.wb.get_value(self.sheet, r1 + 1, c) for c in range(c1, c2 + 1)]
        return any(v not in (None, "") and fe.natural_key(v)[0] != 1 for v in second)

    def _sort_target(self):
        r1, c1, r2, c2 = self.sel_rect()
        if (r1, c1, r2, c2) == self._cell_box(*self.cur):
            if self.cur not in self.sh.cells:
                return None
            rect = self.current_region(*self.cur)
            return rect, self._has_header(*rect)
        return self._copy_rect(), False

    def sort_selection(self, descending=False):
        """Sort the selected rows by the active cell's column (a single cell sorts its data block)."""
        self.commit_edit()
        target = self._sort_target()
        if target:
            rect, header = target
            self._sort_rows(rect, [(min(max(self.cur[1], rect[1]), rect[3]), descending)], header)

    def custom_sort(self):
        self.commit_edit()
        target = self._sort_target()
        if not target:
            self.root.bell()
            return
        (r1, c1, r2, c2), header = target
        columns = [(c, self.wb.display_text(self.sheet, r1, c)) for c in range(c1, c2 + 1)]
        result = dialogs.SortDialog(self.root, columns, header).show()
        if result:
            self._sort_rows((r1, c1, r2, c2), result["keys"], result["header"])

    def _sort_rows(self, rect, keys, header, first=None):
        """Sort the rows of rect by keys [(col, descending)], or with first = (col, "bg" | "fg", colour)
        put the rows whose cell in col shows that colour on top (Excel's Sort by Color)."""
        r1, c1, r2, c2 = rect
        start = r1 + 1 if header else r1
        if r2 <= start or c2 < c1:
            return
        sh, get = self.sh, self.wb.get_value
        big = (r2 - start + 1) * (c2 - c1 + 1) > 20000
        cells, formats = sh.cells, sh.formats
        items = [(r, [get(self.sheet, r, k) for k, _ in keys],
                  None if big else [(cells.get((r, c)), formats.get((r, c))) for c in range(c1, c2 + 1)])
                 for r in range(start, r2 + 1)]
        for i in reversed(range(len(keys))):  # stable passes, most significant key last
            filled = [it for it in items if it[1][i] is not None and it[1][i] != ""]
            blanks = [it for it in items if it[1][i] is None or it[1][i] == ""]
            filled.sort(key=lambda it: fe.natural_key(it[1][i]), reverse=keys[i][1])  # 2 before 10
            items = filled + blanks
        if first is not None:
            col, what, color = first
            hits = {it[0] for it in items if self.wb.cell_color(self.sheet, it[0], col, what) == color}
            items = [it for it in items if it[0] in hits] + [it for it in items if it[0] not in hits]
        if not big:
            cs = ChangeSet(self.wb)
            for new_r, (old_r, _, cells) in zip(range(start, r2 + 1), items):
                for j, (raw, fmt) in enumerate(cells):
                    cs.set(self.sheet, new_r, c1 + j, raw=fe.shift_formula(raw, new_r - old_r, 0) if raw else None,
                           fmt=fmt)
            self.commit(cs)
            if sh.filter:
                self.wb.apply_filter(self.sheet)
            self.refresh()
            return

        def move():  # big tables: renumber the rows in one pass (undo keeps a copy of the sheet, not each cell)
            new_row = {old: new for new, (old, _, _) in zip(range(start, r2 + 1), items)}
            with fe.bulk():
                for name in ("cells", "formats"):
                    d = getattr(sh, name)
                    setattr(sh, name, {((new_row[r], c) if start <= r <= r2 and c1 <= c <= c2 else (r, c)): v
                                       for (r, c), v in d.items()})
                cells, old_row = sh.cells, {n: o for o, n in new_row.items()}
                formulas = [(k, v) for k, v in cells.items() if v[:1] == "=" and start <= k[0] <= r2 and c1 <= k[1] <= c2]
                for (r, c), raw in formulas:  # a formula's relative references move with its row
                    cells[(r, c)] = fe.shift_formula(raw, r - old_row[r], 0)
            self.wb._ast.clear()
            if sh.filter:
                self.wb.apply_filter(self.sheet)
        self._structural(move)

    def _filter_rect(self):
        r1, c1, r2, c2 = self.sel_rect()
        if (r1, c1, r2, c2) == self._cell_box(*self.cur):
            return self.current_region(*self.cur) if self.cur in self.sh.cells else None
        return self._copy_rect()

    def toggle_filter(self):
        self.commit_edit()
        if self.sh.filter:
            self._structural(lambda: setattr(self.sh, "filter", None))
            return
        rect = self._filter_rect()
        if rect is None:
            messagebox.showinfo(APP_NAME, "Hãy chọn một ô trong vùng dữ liệu để bật bộ lọc.", parent=self.root)
            return
        self._structural(lambda: setattr(self.sh, "filter", {"range": rect, "criteria": {}, "hidden": set()}))

    def clear_filter(self):
        if self.sh.filter and self.sh.filter["criteria"]:
            def clear():
                self.sh.filter["criteria"] = {}
                self.wb.apply_filter(self.sheet)
            self._structural(clear)

    def reapply_filter(self):
        if self.sh.filter:
            self._extend_filter_range()
            self._structural(lambda: self.wb.apply_filter(self.sheet))

    def _extend_filter_range(self):
        f = self.sh.filter
        r1, c1, r2, c2 = f["range"]
        while r2 + 1 < MAX_ROWS and any((r2 + 1, c) in self.sh.cells for c in range(c1, c2 + 1)):
            r2 += 1
        f["range"] = (r1, c1, r2, c2)

    def open_filter_popup(self, col, x, y):
        f = self.sh.filter
        if not f:
            return
        self._extend_filter_range()
        title = self.wb.display_text(self.sheet, f["range"][0], col) or f"Cột {col_name(col)}"
        crit = f["criteria"].get(col)
        if isinstance(crit, dict):  # tick the values a rule lets through
            r1, _, r2, _ = f["range"]
            stats = {}
            selected = {self.wb.display_text(self.sheet, r, col) for r in range(r1 + 1, r2 + 1)
                        if self.wb.filter_match(self.sheet, r, col, crit, stats)}
        else:
            selected = crit
        kind = self.wb.filter_kind(self.sheet, col)
        colors = {what: self.wb.filter_colors(self.sheet, col, what) for what in ("bg", "fg")}
        dialogs.FilterPopup(
            self.root, self.ui, title, self.wb.filter_values(self.sheet, col), selected, x, y, kind=kind,
            colors=colors, active=crit, icons=self.menu_icons,
            on_sort=lambda desc: self._sort_rows(f["range"], [(col, desc)], True),
            on_sort_color=lambda what, color: self._sort_rows(f["range"], [], True, first=(col, what, color)),
            on_clear=lambda: self._set_filter_criteria(col, None),
            on_apply=lambda allowed: self._set_filter_criteria(col, allowed),
            on_criterion=lambda new: self._set_filter_criteria(col, new),
            on_dialog=lambda name, preset: self._filter_dialog(col, title, name, preset, kind))

    def _filter_dialog(self, col, title, name, preset, kind):
        """The Top 10 ("top") or Custom AutoFilter ("custom") dialog for a filter column."""
        if name == "top":
            result = dialogs.Top10Dialog(self.root, preset).show()
        else:
            result = dialogs.CustomFilterDialog(self.root, title, self.wb.filter_values(self.sheet, col), preset,
                                                kind).show()
        if result:
            self._set_filter_criteria(col, result)

    def _set_filter_criteria(self, col, allowed):
        def change():
            criteria = self.sh.filter["criteria"]
            if allowed is None:
                criteria.pop(col, None)
            else:  # a set of texts from the checklist, or a rule (see formula_engine.FILTER_OPS)
                criteria[col] = dict(allowed) if isinstance(allowed, dict) else set(allowed)
            self.wb.apply_filter(self.sheet)
        self._structural(change)

    def filter_by_value(self):
        self._filter_by_cell(lambda r, c: {self.wb.display_text(self.sheet, r, c)})

    def filter_by_color(self, what="bg"):
        self._filter_by_cell(lambda r, c: {"type": "color", "what": what,
                                           "color": self.wb.cell_color(self.sheet, r, c, what)})

    def _filter_by_cell(self, criterion):
        """Filter the active cell's column by criterion(row, col), turning the filter on if needed."""
        self.commit_edit()
        r, c = self.cur
        crit = criterion(r, c)

        def change():
            if not self.sh.filter:
                rect = self.current_region(r, c)
                if rect[0] == rect[2]:
                    raise ValueError("Cần một vùng dữ liệu có tiêu đề để lọc.")
                self.sh.filter = {"range": rect, "criteria": {}, "hidden": set()}
            self.sh.filter["criteria"][c] = crit
            self.wb.apply_filter(self.sheet)
        self._structural(change)

    def remove_duplicates(self):
        self.commit_edit()
        target = self._sort_target()
        if not target:
            self.root.bell()
            return
        (r1, c1, r2, c2), header = target
        columns = [(c, self.wb.display_text(self.sheet, r1, c)) for c in range(c1, c2 + 1)]
        result = dialogs.RemoveDuplicatesDialog(self.root, columns, header).show()
        if not result:
            return
        start = r1 + 1 if result["header"] else r1
        seen, keep = set(), []
        for r in range(start, r2 + 1):
            key = tuple(self.wb.display_text(self.sheet, r, c).lower() for c in result["cols"])
            if key not in seen:
                seen.add(key)
                keep.append(r)
        removed = (r2 - start + 1) - len(keep)
        if removed:
            sh = self.sh
            cs = ChangeSet(self.wb)
            for i, src in enumerate(keep):
                dst = start + i
                for c in range(c1, c2 + 1):
                    raw = sh.cells.get((src, c))
                    cs.set(self.sheet, dst, c, raw=fe.shift_formula(raw, dst - src, 0) if raw else None,
                           fmt=sh.formats.get((src, c)))
            for dst in range(start + len(keep), r2 + 1):
                for c in range(c1, c2 + 1):
                    cs.set(self.sheet, dst, c, raw=None, fmt=None)
            self.commit(cs)
            self.refresh()
        messagebox.showinfo(APP_NAME, f"Đã tìm thấy và xóa {removed} giá trị trùng lặp; còn lại {len(keep)} giá trị duy nhất.",
                            parent=self.root)

    # ------------------------------------------------------------------ rows, columns, cells, sheets

    # ------------------------------------------------------------------ charts

    def insert_chart(self, kind):
        """Chart of the selected cells (or the data around the active cell), placed right of the data."""
        self.commit_edit()
        rect = self.sel_rect()
        if rect == self._cell_box(*self.cur):
            rect = self.current_region(*self.cur) if self.cur in self.sh.cells else None
        elif self._is_huge(*rect):
            rows, cols = self.wb.used_extent(self.sheet)
            rect = (rect[0], rect[1], min(rect[2], rows - 1), min(rect[3], cols - 1))
        if rect is None or rect[2] < rect[0] or rect[3] < rect[1] or not self.wb.chart_data(
                self.sheet, {"range": rect})[1]:
            messagebox.showinfo(APP_NAME, "Hãy chọn vùng dữ liệu (có thể kèm hàng/cột tiêu đề) để vẽ biểu đồ.",
                                parent=self.root)
            return
        chart = {"type": kind, "range": rect, "title": "", "anchor": (rect[0], min(rect[3] + 2, MAX_COLS - 1), 0, 0),
                 "size": fe.CHART_SIZE}
        if self._structural(lambda: self.sh.charts.append(chart)):
            self.chart_sel = len(self.sh.charts) - 1
            r, c = chart["anchor"][:2]
            self.ensure_visible(r, c)
            self.refresh_view()

    def _update_chart(self, **changes):
        index = self.chart_sel

        def change():
            self.sh.charts[index] = dict(self.sh.charts[index], **changes)
        self._structural(change)
        self.chart_sel = index
        self.redraw()

    def edit_chart(self):
        if self.chart_sel is not None and self.sh.charts[self.chart_sel]["type"] != fe.PICTURE:
            result = dialogs.ChartDialog(self.root, self.sh.charts[self.chart_sel]).show()
            if result:
                self._update_chart(**result)

    def set_chart_type(self, kind):
        if self.chart_sel is not None and self.sh.charts[self.chart_sel]["type"] != fe.PICTURE:
            self._update_chart(type=kind)

    def delete_chart(self):
        if self.chart_sel is not None:
            index, self.chart_sel = self.chart_sel, None
            self._structural(lambda: self.sh.charts.pop(index))

    # ------------------------------------------------------------------ pictures

    def _ask_picture(self):
        path = filedialog.askopenfilename(parent=self.root, title="Chèn ảnh", filetypes=[
            ("Ảnh", "*.png *.gif *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"), ("Tất cả các tệp", "*.*")])
        if not path:
            return None
        try:
            return images.load_picture(path) + (os.path.splitext(os.path.basename(path))[0],)
        except (OSError, ValueError) as err:
            messagebox.showerror(APP_NAME, f"Không chèn được ảnh:\n{path}\n\n{err}", parent=self.root)
            return None

    @staticmethod
    def _fit_picture(size, limit=(640, 480)):
        w, h = size
        k = min(1.0, limit[0] / max(1, w), limit[1] / max(1, h))
        return max(8, round(w * k)), max(8, round(h * k))

    def insert_picture(self):
        """Insert > Pictures: a picture floating over the cells at the active cell."""
        self.commit_edit()
        found = self._ask_picture()
        if not found:
            return
        data, fmt, size, name = found
        r, c = self.cur
        picture = {"type": fe.PICTURE, "image": data, "format": fmt, "id": fe.new_picture_id(), "name": name,
                   "anchor": (r, c, 0, 0), "size": self._fit_picture(size), "native": list(size)}
        if self._structural(lambda: self.sh.charts.append(picture)):
            self.chart_sel = len(self.sh.charts) - 1
            self.redraw()

    def change_picture(self):
        if self.chart_sel is None:
            return
        found = self._ask_picture()
        if found:
            data, fmt, size, name = found
            old = self.sh.charts[self.chart_sel]
            width = old["size"][0]
            self._update_chart(image=data, format=fmt, id=fe.new_picture_id(), name=name, native=list(size),
                               size=(width, max(8, round(width * size[1] / max(1, size[0])))))

    def reset_picture_size(self):
        if self.chart_sel is None:
            return
        picture = self.sh.charts[self.chart_sel]
        native = picture.get("native") or images.image_info(base64.b64decode(picture["image"]))[1:]
        self._update_chart(size=tuple(native))

    # ------------------------------------------------------------------ hyperlinks

    def insert_hyperlink(self):
        """Insert > Link (Ctrl+K): link the active cell to a web page, a file, a place or an e-mail address."""
        self.commit_edit()
        r, c = self.cur = self.ext = self._anchor(self.cur)
        if self._pivot_guard(r, c, r, c):
            return
        sheet, link = self.sheet, self.sh.links.get((r, c))
        text = self.wb.display_text(sheet, r, c)
        result = dialogs.HyperlinkDialog(self.root, text, link, list(self.wb.sheets),
                                         [d["name"] for d in self.wb.sorted_names()], cell_name(r, c),
                                         os.path.dirname(self.path) if self.path else None).show()
        if not result:
            return
        if result.get("remove"):
            self.remove_hyperlinks()
            return

        def change():
            self.wb.sheets[sheet].links[(r, c)] = {"target": result["target"], "tip": result["tip"]}
            if result["text"] != text or not self.wb.get_raw(sheet, r, c):
                shown = result["text"]
                self.wb.set_raw(sheet, r, c, "'" + shown if shown.startswith(("=", "'")) else shown)
            fmt = dict(self.wb.get_format(sheet, r, c) or {}, fg=HYPERLINK_FG, underline=True)
            self.wb.set_format(sheet, r, c, fmt)
        self._structural(change)

    def remove_hyperlinks(self):
        """Remove Hyperlinks: the links in the selection and their blue underlined look."""
        self.commit_edit()
        r1, c1, r2, c2 = self.sel_rect()
        keys = [k for k in self.sh.links if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]
        if not keys:
            self.root.bell()
            return
        sheet = self.sheet

        def change():
            for key in keys:
                del self.wb.sheets[sheet].links[key]
                fmt = dict(self.wb.get_format(sheet, *key) or {})
                if fmt.get("fg") == HYPERLINK_FG:
                    fmt.pop("fg")
                    fmt.pop("underline", None)
                self.wb.set_format(sheet, *key, fmt or None)
        self._structural(change)

    def open_hyperlink(self):
        link = self.wb.hyperlink_at(self.sheet, *self._anchor(self.cur))
        if link:
            self.follow_link(link)
        else:
            self.root.bell()

    def follow_link(self, link):
        """Open a link: '#Sheet!A1' or '#Name' goes there; web and mail addresses open in their programs;
        anything else is a file (relative to the workbook's folder)."""
        target = link["target"].strip()
        if target.startswith("#"):
            place = target[1:]
            if place.upper() in self.wb.names:
                if not self.go_to_name(place):
                    self.root.bell()
                return
            sheet_part, _, addr = place.rpartition("!")
            sheet = self.wb.resolve_sheet(sheet_part.strip("'").replace("''", "'")) if sheet_part else self.sheet
            rect = fe.parse_range(addr.replace("$", "").strip())
            if sheet is None or rect is None:
                messagebox.showerror(APP_NAME, f"Không tìm thấy vị trí {place}.", parent=self.root)
                return
            self.commit_edit()
            self.switch_sheet(sheet)
            self.cur, self.ext = rect[:2], rect[2:]
            self.ensure_visible(*self.cur)
            self.refresh()
            return
        if re.match(r"[a-z][a-z0-9+.-]+:", target, re.I) and not re.match(r"[a-z]:[\\/]", target, re.I):
            webbrowser.open(target)  # http:, https:, mailto:, ...
            return
        folder = os.path.dirname(self.path) if self.path else os.getcwd()
        path = target if os.path.isabs(target) else os.path.join(folder, target)
        try:
            self._open_document(path)
        except OSError as err:
            messagebox.showerror(APP_NAME, f"Không mở được:\n{path}\n\n{err}", parent=self.root)

    def _show_link_tip(self, link, x, y):
        """The ScreenTip of a linked cell under the pointer (after a short pause), like Excel's."""
        self._hide_link_tip()
        text = (link.get("tip") or link["target"]) + "\nNhấp một lần để mở. Nhấp và giữ để chọn ô này."

        def show():
            self._link_tip_job = None
            tip = self._link_tip = tk.Toplevel(self.root)
            tip.overrideredirect(True)
            tip.attributes("-topmost", True)
            tk.Label(tip, text=text, bg="#ffffff", fg="#333333", bd=1, relief="solid", padx=7, pady=4, justify="left",
                     font=self.ui.small_font).pack()
            tip.geometry(f"+{x + 14}+{y + 18}")
        self._link_tip_job = self.root.after(600, show)

    def _hide_link_tip(self):
        if self._link_tip_job:
            self.root.after_cancel(self._link_tip_job)
            self._link_tip_job = None
        if self._link_tip is not None:
            self._link_tip.destroy()
            self._link_tip = None

    # ------------------------------------------------------------------ PivotTables

    def insert_pivot(self):
        """Insert > PivotTable: a pivot of the selected block (or the data around the active cell)."""
        self.commit_edit()
        rect = self.sel_rect()
        if rect == self._cell_box(*self.cur):
            rect = self.current_region(*self.cur) if self.cur in self.sh.cells else (0, 0, 0, 0)
        elif self._is_huge(*rect):
            rows, cols = self.wb.used_extent(self.sheet)
            rect = (rect[0], rect[1], max(rect[0], min(rect[2], rows - 1)), max(rect[1], min(rect[3], cols - 1)))
        result = dialogs.PivotDialog(self.root, self.wb, self.sheet, rect).show()
        if not result:
            return
        names = {p["name"] for sh in self.wb.sheets.values() for p in sh.pivots}
        name = next(f"PivotTable{n}" for n in itertools.count(1) if f"PivotTable{n}" not in names)
        src_sheet, src_rect = result["source"]
        p = pv.new_pivot(name, src_sheet, src_rect, (2, 0))
        try:
            pv.Source(self.wb, p)
        except pv.PivotError as err:
            messagebox.showerror(APP_NAME, str(err), parent=self.root)
            return
        dest = result["dest"]
        if dest:
            p["anchor"] = dest[1:]
            if pv.overlaps(self.wb, dest[0], p, pv.compute(self.wb, p)) and not messagebox.askyesno(
                    APP_NAME, "Nơi đặt PivotTable đã có dữ liệu. Ghi đè lên dữ liệu đó?", parent=self.root):
                return
        placed = []

        def create():
            sheet = dest[0] if dest else self.wb.add_sheet(index=list(self.wb.sheets).index(self.sheet))
            self.wb.sheets[sheet].pivots.append(p)
            pv.place(self.wb, sheet, p, pv.compute(self.wb, p))
            placed.append(sheet)
        if self._structural(create):
            self.show_pivot_fields = True
            if placed[0] != self.sheet:
                self.switch_sheet(placed[0])
            self.select_to(p["anchor"])

    def _fit_report_columns(self, sheet, rect):
        """Size a pivot report's columns to its contents (Excel autofits them on every update)."""
        sh = self.wb.sheets[sheet]
        for c, width in self._content_widths(sheet, rect[1], rect[3], rows=(rect[0], rect[2])).items():
            sh.col_widths[c] = max(width, fe.DEFAULT_COL_WIDTH * 3 // 4)

    def _current_pivot(self):
        return pv.pivot_at(self.wb, self.sheet, *self.cur)

    def update_pivot(self, p, **changes):
        """Apply changes to pivot p on this sheet, recompute and rewrite its report (one undo step)."""
        sheet = self.sheet
        trial = dict(p, **changes)
        try:
            report = pv.compute(self.wb, trial)
            pv.report_rect(trial, report)
        except pv.PivotError as err:
            messagebox.showerror(APP_NAME, str(err), parent=self.root)
            self._pane_key = None
            self.refresh()
            return False
        if pv.overlaps(self.wb, sheet, trial, report) and not messagebox.askyesno(
                APP_NAME, "Báo cáo PivotTable mới sẽ ghi đè lên ô đang có dữ liệu. Tiếp tục?", parent=self.root):
            self._pane_key = None
            self.refresh()
            return False

        def change():
            p.update({key: trial[key] for key in changes})
            pv.place(self.wb, sheet, p, report)
            self._fit_report_columns(sheet, p["rect"])
        return self._structural(change)

    def refresh_pivot(self, p):
        self.update_pivot(p)

    def refresh_all_pivots(self):
        """Data > Refresh All: recompute every pivot from its source as it is now."""
        self.commit_edit()
        pivots = [(name, p) for name, sh in self.wb.sheets.items() for p in sh.pivots]
        if not pivots:
            messagebox.showinfo(APP_NAME, "Sổ làm việc chưa có PivotTable nào.", parent=self.root)
            return
        errors = []

        def change():
            for sheet, p in pivots:
                try:
                    pv.refresh(self.wb, sheet, p)
                    self._fit_report_columns(sheet, p["rect"])
                except pv.PivotError as err:
                    errors.append(f"{p['name']}: {err}")
        self._structural(change)
        self._pane_key = None
        self.refresh()
        if errors:
            messagebox.showwarning(APP_NAME, "\n".join(errors), parent=self.root)

    def delete_pivot(self, p):
        if not messagebox.askyesno(APP_NAME, f"Xóa {p['name']} và báo cáo của nó?", parent=self.root):
            return

        def change():
            a1, b1, a2, b2 = p["rect"] or (0, 0, -1, -1)
            for key in [k for k in set(self.sh.cells) | set(self.sh.formats) if a1 <= k[0] <= a2 and b1 <= k[1] <= b2]:
                self.sh.cells.pop(key, None)
                self.sh.formats.pop(key, None)
            self.sh.pivots.remove(p)
        self._structural(change)

    def pivot_details(self, p, r, c):
        """Double-click on a value: the source records behind it, on a new sheet (Excel's Show Details)."""
        try:
            source = pv.Source(self.wb, p)
            report = pv.compute(self.wb, dict(p), source)
        except pv.PivotError as err:
            messagebox.showerror(APP_NAME, str(err), parent=self.root)
            return
        key = report.sources.get((r - p["anchor"][0], c - p["anchor"][1]))
        if key is None:
            return
        names, records = pv.drill_down(self.wb, p, *key, source=source)
        made = []

        def create():
            sheet = self.wb.add_sheet(index=list(self.wb.sheets).index(self.sheet))
            sh = self.wb.sheets[sheet]
            for j, name in enumerate(names):
                sh.cells[(0, j)] = fe.value_to_raw(name) or name
                sh.formats[(0, j)] = {"bold": True, "bg": "#ddebf7", "bb": ["thin", "#9bc2e6"]}
            for i, rec in enumerate(records, 1):
                for j, (value, _) in enumerate(rec):
                    raw = fe.number_text(value) if fe.is_number(value) else fe.value_to_raw(value)
                    if raw:
                        sh.cells[(i, j)] = raw
                    fmt = source.formats.get(names[j]) if fe.is_number(value) else None
                    if fmt:
                        sh.formats[(i, j)] = {"num": fmt}
            sh.freeze = (0, 1, 0, 0)
            for col, width in self._content_widths(sheet, 0, len(names) - 1).items():
                sh.col_widths[col] = max(width, fe.DEFAULT_COL_WIDTH)
            made.append(sheet)
        if self._structural(create):
            self.switch_sheet(made[0])
            self.select_to((0, 0))

    def toggle_pivot_fields(self):
        self.show_pivot_fields = not self.show_pivot_fields
        self._pane_key = None
        self.refresh()

    def _sync_pivot_pane(self):
        """Show the PivotTable Fields pane while the active cell is in a pivot report, as Excel does."""
        self.btn_pivot_fields.set_checked(self.show_pivot_fields)
        p = self._current_pivot() if self.show_pivot_fields else None
        if p is None:
            if self.pivot_pane.winfo_ismapped() or self._pane_key is not None:
                self.pivot_pane.grid_remove()
                self._pane_key = None
            return
        key = (self.sheet, id(p), repr([p[area] for area in pv.AREAS]))
        if key != self._pane_key:
            try:
                source = pv.Source(self.wb, p)
                fields, kinds = source.names, source.kinds
            except pv.PivotError:
                fields, kinds = [], {}
            self.pivot_pane.show(p, fields, kinds)
            self._pane_key = key
        if not self.pivot_pane.winfo_ismapped():
            self.pivot_pane.grid(row=0, column=2, sticky="ns")

    def _pivot_pane_change(self, areas):
        p = self._current_pivot()
        if p is not None:
            self.update_pivot(p, **areas)

    def _pivot_pane_action(self, action, *args):
        p = self._current_pivot()
        if action == "close":
            self.toggle_pivot_fields()
            return
        if p is None:
            return
        if action == "refresh":
            self.refresh_pivot(p)
        elif action == "delete":
            self.delete_pivot(p)
        elif action == "options":
            result = dialogs.PivotOptionsDialog(self.root, p).show()
            if result:
                self.update_pivot(p, **result)
        elif action == "value":
            values = [dict(v) for v in p["values"]]
            result = dialogs.ValueFieldDialog(self.root, values[args[0]], pv.FUNCS, pv.SHOW_AS).show()
            if result:
                values[args[0]].update(result)
                self.update_pivot(p, values=values)
        elif action == "field":
            area, index = args
            field = p[area][index]
            try:
                source = pv.Source(self.wb, p)
            except pv.PivotError as err:
                messagebox.showerror(APP_NAME, str(err), parent=self.root)
                return
            result = dialogs.PivotFieldDialog(
                self.root, field, pv.field_items(source, p, field), p["items"].get(field),
                p["sort"].get(field) == "desc", p["group"].get(field), pv.GROUPS,
                source.kinds.get(field) == "date").show()
            if not result:
                return
            sort, group, items = dict(p["sort"]), dict(p["group"]), dict(p["items"])
            sort.pop(field, None)
            if result["desc"]:
                sort[field] = "desc"
            items.pop(field, None)
            if result["group"] != group.get(field):
                group.pop(field, None)  # new grouping, new items: all of them shown again
                if result["group"]:
                    group[field] = result["group"]
            elif result["items"] is not None:
                items[field] = result["items"]
            self.update_pivot(p, sort=sort, group=group, items=items)

    # ------------------------------------------------------------------ data validation and notes

    def _validation_target(self):
        r1, c1, r2, c2 = self.sel_rect()
        if self._is_huge(r1, c1, r2, c2) and not (r1 == 0 and r2 == MAX_ROWS - 1):
            rows, cols = self.wb.used_extent(self.sheet)
            r2, c2 = max(r1, min(r2, rows - 1)), max(c1, min(c2, cols - 1))
        return r1, c1, r2, c2

    def _replace_validation(self, rect, rule):
        """Rules after `rect` loses whatever validation it had and, unless rule is None, gets `rule`."""
        rules = [dict(old, range=part) for old in self.sh.validations for part in fe.rect_minus(old["range"], rect)]
        if rule is not None:
            rules.append(dict(rule, range=rect))
        return rules

    def data_validation_dialog(self):
        self.commit_edit()
        rect = self._validation_target()
        result = dialogs.ValidationDialog(self.root, self.wb.validation_at(self.sheet, *self.cur)).show()
        if result is None:
            return
        rule = None if result.get("clear") or result.get("type") == "any" else result
        rules = self._replace_validation(rect, rule)
        if rules != self.sh.validations:
            self._structural(lambda: setattr(self.sh, "validations", rules))

    def clear_validation(self):
        self.commit_edit()
        rules = self._replace_validation(self._validation_target(), None)
        if rules != self.sh.validations:
            self._structural(lambda: setattr(self.sh, "validations", rules))

    def open_validation_list(self):
        self.commit_edit()
        r, c = self.cur
        rule = self.wb.validation_at(self.sheet, r, c)
        if not rule or rule.get("type") != "list":
            return
        items = self.wb.validation_options(self.sheet, r, c, rule)
        if not items:
            self.root.bell()
            return
        self.ensure_visible(r, c)
        self.refresh_view()
        x0, _, x1, y1 = self._rect_bbox(*self._cell_box(r, c))
        cv = self.canvas
        width = max(x1 - x0 + (self._list_button[2] - self._list_button[0] if self._list_button else 0),
                    max(self.ui.font.measure(t) for t in items) + 24)

        def pick(text):
            cs = ChangeSet(self.wb)
            cs.set(self.sheet, r, c, raw="'" + text if text.startswith("=") else text)
            self.commit(cs)
            self.refresh()
            self.canvas.focus_set()

        dialogs.ListPopup(self.root, self.ui, items, self.wb.display_text(self.sheet, r, c),
                          cv.winfo_rootx() + int(x0), cv.winfo_rooty() + int(y1) + 1, width, pick)

    def edit_note(self):
        self.commit_edit()
        r, c = self.cur = self._anchor(self.cur)
        old = self.sh.notes.get((r, c), "")
        text = dialogs.NoteDialog(self.root, cell_name(r, c), old).show()
        if text is None or text == old:
            return

        def change():
            if text:
                self.sh.notes[(r, c)] = text
            else:
                self.sh.notes.pop((r, c), None)
        self._structural(change)

    def delete_notes(self):
        self.commit_edit()
        r1, c1, r2, c2 = self.sel_rect()
        doomed = [k for k in self.sh.notes if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]
        if not doomed:
            self.root.bell()
            return

        def change():
            for k in doomed:
                del self.sh.notes[k]
        self._structural(change)

    def goto_note(self, step):
        """Select the next (step 1) or previous (step -1) cell with a note, wrapping around."""
        self.commit_edit()
        keys = sorted(self.sh.notes)
        if not keys:
            self.root.bell()
            return
        after = [k for k in keys if (k > self.cur if step > 0 else k < self.cur)]
        target = (after[0] if step > 0 else after[-1]) if after else (keys[0] if step > 0 else keys[-1])
        self.select_to(target)

    def toggle_show_notes(self):
        self.show_notes = not self.show_notes
        self.btn_show_notes.set_checked(self.show_notes)
        self.redraw()

    def insert_rows(self):
        r1, _, r2, _ = self.sel_rect()
        count = 1 if r2 - r1 + 1 >= MAX_ROWS else r2 - r1 + 1
        self.clip = None
        self._structural(lambda: self.wb.insert_rows(self.sheet, r1, count))

    def delete_rows(self):
        r1, _, r2, _ = self.sel_rect()
        self.clip = None
        self._structural(lambda: self.wb.delete_rows(self.sheet, r1, r2 - r1 + 1))

    def insert_cols(self):
        _, c1, _, c2 = self.sel_rect()
        count = 1 if c2 - c1 + 1 >= MAX_COLS else c2 - c1 + 1
        self.clip = None
        self._structural(lambda: self.wb.insert_cols(self.sheet, c1, count))

    def delete_cols(self):
        _, c1, _, c2 = self.sel_rect()
        self.clip = None
        self._structural(lambda: self.wb.delete_cols(self.sheet, c1, c2 - c1 + 1))

    def _shift_cells(self, insert, shift):
        r1, c1, r2, c2 = self.sel_rect()
        if shift == "rows":
            return self.insert_rows() if insert else self.delete_rows()
        if shift == "cols":
            return self.insert_cols() if insert else self.delete_cols()
        if self._is_huge(r1, c1, r2, c2):
            r1, c1, r2, c2 = self._copy_rect()
        self.clip = None
        if insert:
            self._structural(lambda: self.wb.insert_cells(self.sheet, r1, c1, r2, c2, shift))
        else:
            self._structural(lambda: self.wb.delete_cells(self.sheet, r1, c1, r2, c2, shift))

    def insert_cells_dialog(self):
        self.commit_edit()
        choice = dialogs.cells_dialog(self.root, "insert")
        if choice:
            self._shift_cells(True, choice)

    def delete_cells_dialog(self):
        self.commit_edit()
        choice = dialogs.cells_dialog(self.root, "delete")
        if choice:
            self._shift_cells(False, choice)

    def insert_smart(self):
        r1, c1, r2, c2 = self.sel_rect()
        if c1 == 0 and c2 == MAX_COLS - 1:
            self.insert_rows()
        elif r1 == 0 and r2 == MAX_ROWS - 1:
            self.insert_cols()
        else:
            self._shift_cells(True, "down")

    def delete_smart(self):
        r1, c1, r2, c2 = self.sel_rect()
        if c1 == 0 and c2 == MAX_COLS - 1:
            self.delete_rows()
        elif r1 == 0 and r2 == MAX_ROWS - 1:
            self.delete_cols()
        else:
            self._shift_cells(False, "up")

    def hide_rows(self):
        r1, _, r2, _ = self.sel_rect()
        self._structural(lambda: self.sh.hidden_rows.update(range(r1, r2 + 1)))

    def hide_cols(self):
        _, c1, _, c2 = self.sel_rect()
        self._structural(lambda: self.sh.hidden_cols.update(range(c1, c2 + 1)))

    def unhide_rows(self):
        r1, _, r2, _ = self.sel_rect()
        self._structural(lambda: self.sh.hidden_rows.difference_update(range(r1, r2 + 1)))

    def unhide_cols(self):
        _, c1, _, c2 = self.sel_rect()
        self._structural(lambda: self.sh.hidden_cols.difference_update(range(c1, c2 + 1)))

    def _content_widths(self, sheet, c1, c2, rows=None):
        """{column: pixels at 96 DPI} its widest text needs (only rows r1..r2 when rows = (r1, r2))."""
        sh = self.wb.sheets[sheet]
        cells, formats = sh.cells, sh.formats
        by_col = {}
        for key in cells:
            if c1 <= key[1] <= c2 and (rows is None or rows[0] <= key[0] <= rows[1]):
                by_col.setdefault(key[1], []).append(key)
        widest = {}
        for c, keys in by_col.items():
            if len(keys) > 400:  # long columns: measure the longest texts and every cell in a larger font
                keys = set(heapq.nlargest(150, keys, key=lambda k: len(cells[k]))) | {
                    k for k in keys if k in formats and ((formats[k].get("size") or 11) > 11 or formats[k].get("indent"))}
            for r, c in keys:
                if sh.merge_at(r, c):
                    continue
                fmt = formats.get((r, c))
                text = self.wb.display_text(sheet, r, c)
                indent = ((fmt or {}).get("indent") or 0) * round(9 * self.px)
                widest[c] = max(widest.get(c, 0), self.text_width(self.font_for(fmt), text) + indent)
        return {c: max(fe.MIN_COL_WIDTH, round((w + 3 * self.pad) / self.px)) for c, w in widest.items()}

    def autofit_columns(self, c1, c2):
        sh = self.sh
        before = self.wb.snapshot()
        widest = self._content_widths(self.sheet, c1, c2)
        for c in range(c1, c2 + 1):
            if c in widest:
                sh.col_widths[c] = widest[c]
            else:
                sh.col_widths.pop(c, None)
        self._push_snapshot(before, self._sel_state())
        self.refresh_view()

    def autofit_selected_columns(self):
        self.commit_edit()
        _, c1, _, c2 = self.sel_rect()
        c2 = min(c2, max(c1, self.wb.used_extent(self.sheet)[1] - 1))
        self.autofit_columns(c1, c2)

    def ask_column_width(self):
        self.commit_edit()
        _, c1, _, c2 = self.sel_rect()
        c2 = min(c2, c1 + MAX_COLS)
        widths = self.sh.col_widths
        current = (widths.get(c1, fe.DEFAULT_COL_WIDTH) - 5) / 7
        chars = simpledialog.askfloat("Độ rộng cột", "Độ rộng cột (số ký tự):", parent=self.root,
                                      initialvalue=round(current, 2), minvalue=0, maxvalue=255)
        if chars is None:
            return

        def change():
            for c in range(c1, c2 + 1):
                if chars == 0:
                    self.sh.hidden_cols.add(c)
                else:
                    self.sh.col_widths[c] = max(fe.MIN_COL_WIDTH, round(chars * 7 + 5))
        self._structural(change)

    def _show_sheet(self, name):
        if self.sheet in self.wb.sheets:
            self.views[self.sheet] = (self.top, self.left, self.cur, self.ext)
        self.sheet = name
        self.chart_sel = None
        self.areas = []
        self.top, self.left, self.cur, self.ext = self.views.get(name, (0, 0, (0, 0), (0, 0)))

    def switch_sheet(self, name):
        if name == self.sheet:
            return
        self.commit_edit()
        self._show_sheet(name)
        self.refresh_all()

    def add_sheet(self):
        index = list(self.wb.sheets).index(self.sheet) + 1
        created = []
        if self._structural(lambda: created.append(self.wb.add_sheet(index=index))):
            self.switch_sheet(created[0])

    def copy_sheet(self):
        index = list(self.wb.sheets).index(self.sheet) + 1
        created = []
        if self._structural(lambda: created.append(self.wb.copy_sheet(self.sheet, index))):
            self.switch_sheet(created[0])

    def set_tab_color(self, color):
        self._structural(lambda: setattr(self.sh, "tab_color", color))

    def rename_sheet(self, name):
        self.commit_edit()
        new = simpledialog.askstring("Đổi tên trang tính", "Tên mới:", initialvalue=name, parent=self.root)
        if new is None or new.strip() == name:
            return
        new = new.strip()

        def rename():
            self.wb.rename_sheet(name, new)
            if self.sheet == name:
                self.sheet = new
            if name in self.views:
                self.views[new] = self.views.pop(name)
            if self.clip and self.clip["sheet"] == name:
                self.clip["sheet"] = new
        self._structural(rename)

    def delete_sheet(self, name):
        self.commit_edit()
        if len(self.wb.sheets) == 1:
            messagebox.showinfo(APP_NAME, "Sổ làm việc phải có ít nhất một trang tính.", parent=self.root)
            return
        if not messagebox.askyesno(APP_NAME, f"Xóa trang tính “{name}”?\n(Có thể hoàn tác bằng Ctrl+Z)", parent=self.root):
            return
        names = list(self.wb.sheets)
        i = names.index(name)
        fallback = names[i + 1] if i + 1 < len(names) else names[i - 1]

        def delete():
            self.wb.delete_sheet(name)
            self.views.pop(name, None)
            if self.sheet == name:
                self.sheet = fallback
                self.top, self.left, self.cur, self.ext = self.views.get(fallback, (0, 0, (0, 0), (0, 0)))
        self._structural(delete)

    def move_sheet(self, name, index):
        self._structural(lambda: self.wb.move_sheet(name, index))

    def hide_sheet(self, name=None):
        """Hide a sheet (its tab); the next visible sheet becomes active."""
        name = name or self.sheet
        visible = self.wb.visible_sheets()
        if len(visible) <= 1:
            messagebox.showinfo(APP_NAME, "Sổ làm việc phải có ít nhất một trang tính hiển thị.", parent=self.root)
            return
        i = visible.index(name)
        fallback = visible[i + 1] if i + 1 < len(visible) else visible[i - 1]
        if self._structural(lambda: self.wb.hide_sheet(name)) and self.sheet == name:
            self.switch_sheet(fallback)

    def unhide_sheet(self):
        hidden = [n for n, sh in self.wb.sheets.items() if sh.hidden]
        if not hidden:
            self.root.bell()
            return
        name = dialogs.ChoiceDialog(self.root, "Bỏ ẩn trang tính", [(n, n) for n in hidden], hidden[0]).show()
        if name and self._structural(lambda: self.wb.hide_sheet(name, False)):
            self.switch_sheet(name)

    def step_sheet(self, step):
        """Ctrl+PgDn / Ctrl+PgUp: the next / previous visible sheet."""
        visible = self.wb.visible_sheets()
        i = visible.index(self.sheet) if self.sheet in visible else 0
        if 0 <= i + step < len(visible):
            self.switch_sheet(visible[i + step])

    def _tab_menu(self, e, name):
        self.switch_sheet(name)
        index = list(self.wb.sheets).index(name)
        last = len(self.wb.sheets) - 1
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Chèn trang tính", command=self.add_sheet, **self._menu_image("Chèn trang tính"))
        menu.add_command(label="Xóa", command=lambda: self.delete_sheet(name), **self._menu_image("", "sheet_delete"))
        menu.add_command(label="Đổi tên...", command=lambda: self.rename_sheet(name), **self._menu_image("Đổi tên"))
        menu.add_command(label="Sao chép trang tính", command=self.copy_sheet, **self._menu_image("Sao chép trang tính"))
        menu.add_cascade(label="Màu tab", menu=self._palette_menu(menu, self.set_tab_color, "Không màu"),
                         **self._menu_image("Màu tab"))
        menu.add_command(label="Ẩn", command=lambda: self.hide_sheet(name), **self._menu_image("", "sheet_hide"),
                         state="normal" if len(self.wb.visible_sheets()) > 1 else "disabled")
        menu.add_command(label="Bỏ ẩn...", command=self.unhide_sheet, **self._menu_image("", "sheet_unhide"),
                         state="normal" if len(self.wb.visible_sheets()) < len(self.wb.sheets) else "disabled")
        menu.add_separator()
        menu.add_command(label="Di chuyển sang trái", command=lambda: self.move_sheet(name, index - 1),
                         state="normal" if index > 0 else "disabled", **self._menu_image("Di chuyển sang trái"))
        menu.add_command(label="Di chuyển sang phải", command=lambda: self.move_sheet(name, index + 1),
                         state="normal" if index < last else "disabled", **self._menu_image("Di chuyển sang phải"))
        try:
            menu.tk_popup(e.x_root, e.y_root)
        finally:
            menu.grab_release()

    # ------------------------------------------------------------------ view

    def set_zoom(self, zoom):
        zoom = max(0.25, min(3.0, round(zoom, 2)))
        if abs(zoom - self.zoom) < 1e-6:
            return
        self.commit_edit()
        self.zoom = zoom
        self._init_metrics()
        self._zoom_busy = True
        self.zoom_var.set(zoom * 100)
        self._zoom_busy = False
        self.zoom_label.config(text=f"{round(zoom * 100)}%")
        self.refresh()

    def zoom_dialog(self):
        value = simpledialog.askinteger("Thu phóng", "Tỷ lệ thu phóng (25 - 300%):", parent=self.root,
                                        initialvalue=round(self.zoom * 100), minvalue=25, maxvalue=300)
        if value:
            self.set_zoom(value / 100)

    def zoom_to_selection(self):
        r1, c1, r2, c2 = self._copy_rect()
        if r2 < r1 or c2 < c1:
            return
        width = sum(self.sh.col_widths.get(c, fe.DEFAULT_COL_WIDTH) for c in range(c1, c2 + 1)) * self.scale
        height = sum(self.sh.row_heights.get(r, fe.DEFAULT_ROW_HEIGHT) for r in range(r1, r2 + 1)) * self.scale
        avail_w = self.canvas.winfo_width() - self.header_w - 10
        avail_h = self.canvas.winfo_height() - self.header_h - 10
        self.top, self.left = r1, c1
        self.set_zoom(min(avail_w / max(1, width), avail_h / max(1, height)))
        self.top, self.left = r1, c1
        self.refresh()

    def freeze_panes(self, kind):
        """Freeze the rows above and columns left of the active cell ("cell"), the top visible row
        ("row") or the first visible column ("col"); None unfreezes."""
        self.commit_edit()
        self._clamp_view()
        r, c = self.cur
        if kind == "cell":
            freeze = (self.top, max(r, self.top), self.left, max(c, self.left))
        elif kind == "row":
            freeze = (self.top, self.top + 1, 0, 0)
        elif kind == "col":
            freeze = (0, 0, self.left, self.left + 1)
        else:
            freeze = None
        if freeze and freeze[1] == freeze[0] and freeze[3] == freeze[2]:
            messagebox.showinfo(APP_NAME, "Hãy chọn ô nằm dưới các hàng và bên phải các cột muốn cố định.",
                                parent=self.root)
            return
        if freeze == self.sh.freeze:
            return
        if self._structural(lambda: setattr(self.sh, "freeze", freeze)) and freeze:
            self.top, self.left = max(self.top, freeze[1]), max(self.left, freeze[3])
            self.refresh_view()

    def toggle_grid(self):
        self.show_grid = self.grid_var.get()
        self.refresh_view()

    def toggle_headings(self):
        self.show_headings = self.headings_var.get()
        self._init_metrics()
        self.refresh()

    def toggle_formula_bar(self):
        if self.fxbar_var.get():
            self.formula_bar.pack(side="top", fill="x", after=self.ribbon)
        else:
            self.formula_bar.pack_forget()

    # ------------------------------------------------------------------ name box, find, help

    def focus_name_box(self):
        self.commit_edit()
        self.name_box.focus_set()
        self.name_box.select_range(0, "end")

    def _leave_name_box(self):
        self.canvas.focus_set()
        self.update_name_box()

    def _on_name_box(self, e):
        text = self.name_box.get().strip()
        if text.upper() in self.wb.names:
            if not self.go_to_name(text):
                self.root.bell()
            return "break"
        if "!" not in text and fe.parse_address(text) is None and self.wb.check_name(text) is None:
            rect = self.sel_rect()  # a new name for the selection, as typing a name in Excel's Name Box does
            if self._structural(lambda: self.wb.define_name(text, "=" + fe.absolute_ref(self.sheet, rect))):
                self.set_mode(f"Đã đặt tên {text} cho {range_name(*rect)}")
            self.canvas.focus_set()
            self.update_name_box()
            return "break"
        sheet = self.sheet
        if "!" in text:
            sheet_part, text = text.rsplit("!", 1)
            sheet = self.wb.resolve_sheet(sheet_part.strip().strip("'"))
        target = fe.parse_address(text)
        if sheet is None or target is None:
            self.root.bell()
            return "break"
        self.switch_sheet(sheet)
        if len(target) == 2:
            self.cur = self.ext = self._anchor(target)
        else:
            self.cur, self.ext = target[:2], target[2:]
        self.ensure_visible(*self.cur)
        self.canvas.focus_set()
        self.refresh()
        return "break"

    def open_find(self, replace=False):
        self.commit_edit()
        if self.find_win is not None and self.find_win.winfo_exists():
            self.find_win.deiconify()
            self.find_win.lift()
            self.find_entry.focus_set()
            return
        win = self.find_win = tk.Toplevel(self.root)
        win.title("Tìm và thay thế")
        win.transient(self.root)
        win.resizable(False, False)
        frame = ttk.Frame(win, padding=10)
        frame.pack(fill="both", expand=True)
        self.find_var, self.repl_var = tk.StringVar(win), tk.StringVar(win)
        self.case_var = tk.BooleanVar(win)
        ttk.Label(frame, text="Tìm:").grid(row=0, column=0, sticky="w")
        self.find_entry = ttk.Entry(frame, textvariable=self.find_var, width=32)
        self.find_entry.grid(row=0, column=1, pady=2)
        ttk.Label(frame, text="Thay bằng:").grid(row=1, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(frame, textvariable=self.repl_var, width=32).grid(row=1, column=1, pady=2)
        ttk.Checkbutton(frame, text="Phân biệt chữ hoa/thường", variable=self.case_var).grid(row=2, column=1, sticky="w")
        self.find_status = ttk.Label(frame, text="", foreground="#666666")
        self.find_status.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
        buttons = ttk.Frame(frame)
        buttons.grid(row=0, column=2, rowspan=4, sticky="n", padx=(10, 0))
        for text, command in (("Tìm tiếp", self.find_next), ("Thay thế", self.replace_one),
                              ("Thay tất cả", self.replace_all), ("Đóng", win.destroy)):
            ttk.Button(buttons, text=text, command=command, width=12).pack(fill="x", pady=1)
        self.find_entry.bind("<Return>", lambda e: self.find_next())
        win.bind("<Escape>", lambda e: win.destroy())
        self.find_entry.focus_set()

    def _matches(self, r, c, needle):
        raw = self.wb.get_raw(self.sheet, r, c)
        shown = self.wb.display_text(self.sheet, r, c)
        if self.case_var.get():
            return needle in raw or needle in shown
        needle = needle.lower()
        return needle in raw.lower() or needle in shown.lower()

    def find_next(self):
        needle = self.find_var.get()
        if not needle:
            return False
        keys = sorted(self.sh.cells)
        ordered = [k for k in keys if k > self.cur] + [k for k in keys if k <= self.cur]
        for r, c in ordered:
            if self._matches(r, c, needle):
                self.cur = self.ext = self._anchor((r, c))
                self.ensure_visible(r, c)
                self.refresh()
                self.find_status.config(text=f"Tìm thấy tại {cell_name(r, c)}.")
                return True
        self.find_status.config(text="Không tìm thấy.")
        self.root.bell()
        return False

    def _replaced(self, raw):
        flags = 0 if self.case_var.get() else re.IGNORECASE
        replacement = self.repl_var.get()
        return re.sub(re.escape(self.find_var.get()), lambda m: replacement, raw, flags=flags)

    def replace_one(self):
        if not self.find_var.get():
            return
        raw = self.wb.get_raw(self.sheet, *self.cur)
        new = self._replaced(raw)
        if new != raw:
            cs = ChangeSet(self.wb)
            cs.set(self.sheet, *self.cur, raw=new)
            self.commit(cs)
            self.refresh()
        self.find_next()

    def replace_all(self):
        if not self.find_var.get():
            return
        cs = ChangeSet(self.wb)
        count = 0
        for (r, c), raw in list(self.sh.cells.items()):
            new = self._replaced(raw)
            if new != raw:
                cs.set(self.sheet, r, c, raw=new)
                count += 1
        self.commit(cs)
        self.refresh()
        self.find_status.config(text=f"Đã thay thế {count} ô.")

    def show_help(self):
        win = tk.Toplevel(self.root)
        win.title("Hướng dẫn PyExcel")
        win.geometry(f"{round(820 * self.scale)}x{round(620 * self.scale)}")
        text = tk.Text(win, wrap="word", font=(self.ui_family, 10), padx=14, pady=10, relief="flat", bg="white")
        bar = ttk.Scrollbar(win, command=text.yview)
        text.config(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        names = fe.function_names()
        text.insert("1.0", HELP_TEXT + f"\nHÀM HỖ TRỢ ({len(names)})\n  " + ", ".join(names) + "\n")
        text.config(state="disabled")

    def show_about(self):
        messagebox.showinfo(APP_NAME, "PyExcel — bảng tính kiểu Excel viết bằng Python + tkinter.\n\n"
                            f"{MAX_ROWS:,} hàng × {MAX_COLS} cột mỗi trang tính, {len(fe.function_names())} hàm,\n"
                            "mở/lưu .xlsx, .pyxl, .csv.", parent=self.root)

    # ------------------------------------------------------------------ defined names

    def go_to_name(self, name):
        """Select the block a name refers to; False when it is not a block of cells."""
        target = self.wb.name_target(name)
        if target is None:
            return False
        sheet, (r1, c1, r2, c2) = target
        self.commit_edit()
        self.switch_sheet(sheet)
        self.cur, self.ext = (r1, c1), (r2, c2)
        self.ensure_visible(r1, c1)
        self.canvas.focus_set()
        self.refresh()
        return True

    def _edit_name(self, parent, name=None):
        """New / Edit Name dialog; returns the name kept (for the Name Manager to select)."""
        if name:
            d = self.wb.names[name.upper()]
            values = dict(name=d["name"], ref=d["ref"], comment=d.get("comment", ""), old=d["name"])
        else:
            r1, c1, r2, c2 = self.sel_rect()
            label = self.wb.display_text(self.sheet, r1, c1) if (r1, c1) != (r2, c2) else ""
            guess = fe.name_from_label(label) if label else ""
            values = dict(name=guess if guess and self.wb.check_name(guess) is None else "",
                          ref="=" + fe.absolute_ref(self.sheet, (r1, c1, r2, c2)))
        result = dialogs.DefineNameDialog(parent, self.wb, **values).show()
        if not result:
            return name
        new, ref, comment = result
        self._structural(lambda: self.wb.define_name(new, ref, comment, old=values.get("old")))
        return new

    def define_name_dialog(self):
        self.commit_edit()
        self._edit_name(self.root)

    def name_manager(self):
        """Formulas > Name Manager (Ctrl+F3)."""
        self.commit_edit()

        def delete(name):
            self._structural(lambda: self.wb.delete_name(name))
        dialogs.NameManagerDialog(self.root, self.wb, lambda parent: self._edit_name(parent),
                                  lambda parent, name: self._edit_name(parent, name), delete, self.go_to_name).show()

    def create_names_from_selection(self):
        """Formulas > Create from Selection (Ctrl+Shift+F3): names from the labels along the block's edges."""
        self.commit_edit()
        r1, c1, r2, c2 = self.sel_rect()
        if (r1, c1) == (r2, c2) or r2 - r1 > 10000 or c2 - c1 > 1000:
            messagebox.showinfo(APP_NAME, "Hãy chọn vùng gồm tiêu đề và dữ liệu (ví dụ cả bảng có hàng tiêu đề).",
                                parent=self.root)
            return
        text = lambda r, c: self.wb.display_text(self.sheet, r, c)  # noqa: E731
        numeric = lambda r, c: fe.is_number(self.wb.get_value(self.sheet, r, c))  # noqa: E731
        top = r2 > r1 and all(text(r1, c) and not numeric(r1, c) for c in range(c1 + 1, c2 + 1) or [c1])
        left = c2 > c1 and all(text(r, c1) and not numeric(r, c1) for r in range(r1 + 1, r2 + 1))
        picked = dialogs.CreateNamesDialog(self.root, top or not left, left).show()
        if not picked:
            return
        rows = (r1 + ("top" in picked), r2 - ("bottom" in picked))
        cols = (c1 + ("left" in picked), c2 - ("right" in picked))
        items = []
        for edge in ("top", "bottom"):
            if edge in picked:
                label_row = r1 if edge == "top" else r2
                items += [(text(label_row, c), (rows[0], c, rows[1], c)) for c in range(cols[0], cols[1] + 1)]
        for edge in ("left", "right"):
            if edge in picked:
                label_col = c1 if edge == "left" else c2
                items += [(text(r, label_col), (r, cols[0], r, cols[1])) for r in range(rows[0], rows[1] + 1)]
        items = [(fe.name_from_label(label), rect) for label, rect in items
                 if label and rect[0] <= rect[2] and rect[1] <= rect[3]]
        items = [(name, rect) for name, rect in items if name]
        if not items:
            self.root.bell()
            return

        def change():
            for name, rect in items:
                old = name if name.upper() in self.wb.names else None
                if self.wb.check_name(name, ignore=old) is None:
                    self.wb.define_name(name, "=" + fe.absolute_ref(self.sheet, rect), old=old)
        if self._structural(change):
            self.set_mode(f"Đã tạo {len(items)} tên: " + ", ".join(n for n, _ in items[:6]) + ("..." if len(items) > 6 else ""))

    def _names_menu(self, menu):
        """Use in Formula: the names, inserted where a formula is being typed (or starting one)."""
        menu.delete(0, "end")
        if not self.wb.names:
            menu.add_command(label="(Chưa có tên nào)", state="disabled")
        for d in self.wb.sorted_names():
            menu.add_command(label=d["name"], command=lambda n=d["name"]: self.insert_name(n))
        menu.add_separator()
        menu.add_command(label="Trình Quản lý Tên...", command=self.name_manager, accelerator="Ctrl+F3")

    def insert_name(self, name):
        if not self.is_editing():
            self.start_edit(initial="=")
        w = self._edit_widget()
        w.insert("insert", name)
        w.focus_set()

    # ------------------------------------------------------------------ printing

    @property
    def font_library(self):
        """The TrueType fonts used to lay out and write printed pages (indexed on first use)."""
        if self._font_library is None:
            self.root.config(cursor="watch")  # scanning the fonts folder can take a moment the first time
            self.root.update_idletasks()
            try:
                self._font_library = ttf.FontLibrary()
                self._font_library.index()
            finally:
                self.root.config(cursor="")
        return self._font_library

    def file_label(self):
        return os.path.basename(self.path) if self.path else ""

    def print_jobs(self, scope):
        """[(sheet, setup)] to print: the active sheet, every sheet, or the selection of the active sheet."""
        if scope == "book":
            return [(name, printing.setup_of(sh)) for name, sh in self.wb.sheets.items() if not sh.hidden]
        setup = printing.setup_of(self.sh)
        if scope == "selection":
            setup = dict(setup, area=self.sel_rect())
        return [(self.sheet, setup)]

    def page_layout(self):
        """printing.Layout of the active sheet, kept until the next change (for the page-break lines)."""
        key = (id(self.wb), self.sheet, self.edit_serial)
        if self._layout_cache[0] != key:
            self._layout_cache = (key, printing.Layout(self.wb, self.sheet, printing.setup_of(self.sh)))
        return self._layout_cache[1]

    def update_page_setup(self, changes, sheet=None):
        """Change printing options of a sheet (with undo); only values that differ from the defaults are kept."""
        sh = self.wb.sheets[sheet or self.sheet]
        setup = {**printing.setup_of(sh), **changes}
        kept = {k: v for k, v in setup.items() if v != printing.DEFAULT_SETUP.get(k)} or None
        if kept != sh.page_setup:
            self._structural(lambda: setattr(sh, "page_setup", kept))

    def page_setup_dialog(self, tab="page", sheet=None, parent=None):
        self.commit_edit()
        sheet = sheet or self.sheet
        result = dialogs.PageSetupDialog(parent or self.root, printing.setup_of(self.wb.sheets[sheet]), tab).show()
        if result:
            self.update_page_setup(result, sheet)

    def print_preview(self, scope="sheet"):
        """File > Print (Ctrl+P): the preview with print and PDF buttons."""
        self.commit_edit()
        self.breaks_shown.add(self.sheet)
        PrintPreview(self, scope)
        self.redraw()

    def _write_pdf(self, scope, path):
        selection = (self.sheet, self.sel_rect()) if scope == "selection" else None
        sheets = self.wb.visible_sheets() if scope == "book" else [self.sheet]
        return printing.export_pdf(self.wb, sheets, path, self.font_library, self.file_label(), selection)

    def export_pdf(self, scope="sheet", parent=None):
        """File > Export PDF: write the pages to a PDF file and offer to open it."""
        self.commit_edit()
        parent = parent or self.root
        base = os.path.splitext(self.file_label())[0] or "Bang tinh"
        path = filedialog.asksaveasfilename(parent=parent, title="Xuất PDF", initialfile=base,
                                            defaultextension=".pdf", filetypes=[("Tệp PDF", "*.pdf")])
        if not path:
            return False
        try:
            pages = self._write_pdf(scope, path)
        except Exception as err:  # e.g. the PDF is open in a viewer that locks it
            messagebox.showerror(APP_NAME, f"Không xuất được PDF:\n{path}\n\n{err}", parent=parent)
            return False
        self.set_mode(f"Đã xuất {pages} trang ra {os.path.basename(path)}")
        if messagebox.askyesno(APP_NAME, f"Đã xuất {pages} trang ra:\n{path}\n\nMở tệp PDF ngay?", parent=parent):
            self._open_document(path)
        return True

    @staticmethod
    def _open_document(path, verb=None):
        """Open (or with verb "print", print) a file with the program the system uses for it."""
        if sys.platform == "win32":
            os.startfile(path, verb) if verb else os.startfile(path)
        elif verb == "print" and shutil.which("lp"):
            subprocess.Popen(["lp", path])
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])

    def print_now(self, scope="sheet", parent=None):
        """Print: the pages go to a PDF in the temp folder, which the system's PDF program sends to the
        default printer; when it cannot print directly the PDF is opened to print from there."""
        self.commit_edit()
        parent = parent or self.root
        base = os.path.splitext(self.file_label())[0] or "Bang tinh"
        path = os.path.join(tempfile.gettempdir(), f"PyExcel - {base} - {dt.datetime.now():%Y%m%d-%H%M%S}.pdf")
        try:
            pages = self._write_pdf(scope, path)
        except Exception as err:
            messagebox.showerror(APP_NAME, f"Không tạo được trang in:\n{err}", parent=parent)
            return False
        try:
            self._open_document(path, "print")
            self.set_mode(f"Đã gửi {pages} trang tới máy in")
        except OSError:
            try:
                self._open_document(path)
            except OSError as err:
                messagebox.showerror(APP_NAME, f"Không mở được trình xem PDF để in:\n{err}", parent=parent)
                return False
            messagebox.showinfo(APP_NAME, "Máy chưa có chương trình in PDF trực tiếp nên PyExcel đã mở tệp PDF;\n"
                                          "hãy bấm In (Ctrl+P) trong cửa sổ vừa mở.", parent=parent)
        return True

    def set_print_area(self):
        self.update_page_setup({"area": self.sel_rect()})

    def clear_print_area(self):
        self.update_page_setup({"area": None})

    def insert_page_break(self):
        """Excel's Insert Page Break: above the active row and left of the active column (not at row/column 1)."""
        r, c = self.cur
        setup = printing.setup_of(self.sh)
        changes = {}
        if r > 0:
            changes["breaks"] = sorted(set(setup["breaks"]) | {r})
        if c > 0:
            changes["col_breaks"] = sorted(set(setup["col_breaks"]) | {c})
        self.breaks_shown.add(self.sheet)
        self.update_page_setup(changes)

    def remove_page_break(self):
        r, c = self.cur
        setup = printing.setup_of(self.sh)
        self.update_page_setup({"breaks": [b for b in setup["breaks"] if b != r],
                                "col_breaks": [b for b in setup["col_breaks"] if b != c]})

    def reset_page_breaks(self):
        self.update_page_setup({"breaks": [], "col_breaks": []})

    def toggle_print_gridlines(self):
        self.update_page_setup({"gridlines": self.print_grid_var.get()})

    def _draw_page_breaks(self, cv, col_pos, row_pos, right, bottom):
        """Dashed lines where pages end, as Excel shows them once the sheet has been previewed or set up;
        manual breaks are drawn darker."""
        layout = self.page_layout()
        if not layout.pages:
            return
        rows, cols = layout.breaks()
        rows.add(max(r for page, _ in layout.pages for r in page) + 1)
        cols.add(max(c for _, page in layout.pages for c in page) + 1)
        manual_rows, manual_cols = set(layout.setup["breaks"]), set(layout.setup["col_breaks"])
        hw, hh = self.header_w, self.header_h
        for r in rows:
            y = row_pos[r][0] if r in row_pos else row_pos[r - 1][1] if r - 1 in row_pos else None
            if y is not None:
                manual = r in manual_rows
                cv.create_line(hw, y - 1, right, y - 1, fill=BREAK_LINE[manual], dash=(6, 3) if manual else (3, 3),
                               width=2 if manual else 1)
        for c in cols:
            x = col_pos[c][0] if c in col_pos else col_pos[c - 1][1] if c - 1 in col_pos else None
            if x is not None:
                manual = c in manual_cols
                cv.create_line(x - 1, hh, x - 1, bottom, fill=BREAK_LINE[manual], dash=(6, 3) if manual else (3, 3),
                               width=2 if manual else 1)

    # ------------------------------------------------------------------ files

    def confirm_discard(self):
        """Offer to save unsaved changes; False means the user cancelled."""
        self.commit_edit()
        if not self.dirty:
            return True
        answer = messagebox.askyesnocancel(APP_NAME, "Bạn có muốn lưu các thay đổi không?", parent=self.root)
        if answer is None:
            return False
        return self.save_file() if answer else True

    def new_file(self):
        if self.confirm_discard():
            self._load(fe.Workbook(), None)

    def open_file(self):
        if not self.confirm_discard():
            return
        path = filedialog.askopenfilename(parent=self.root, title="Mở bảng tính", filetypes=OPEN_TYPES)
        if path:
            self.open_path(path)

    def _busy(self, text):
        """Show what is going on (and a busy pointer) before a long job that holds up the window."""
        self.set_mode(text)
        self.root.config(cursor="watch")
        self.canvas.config(cursor="watch")
        self.root.update_idletasks()

    def _not_busy(self):
        self.root.config(cursor="")
        self._cursor = None  # the next pointer move sets the canvas cursor again
        self.canvas.config(cursor="")

    def open_path(self, path):
        self._busy(f"Đang mở {os.path.basename(path)}...")
        try:
            wb = fe.load_workbook(path)
        except Exception as err:  # any I/O or format problem is reported, not fatal
            self._not_busy()
            self.set_mode("Sẵn sàng")
            messagebox.showerror(APP_NAME, f"Không mở được tệp:\n{path}\n\n{err}", parent=self.root)
            return
        self._load(wb, path)
        self._not_busy()
        info = wb.import_info
        if info:  # tell how a text file was read (Data > From Text/CSV chooses otherwise)
            self.set_mode(f"Đã đọc {os.path.basename(path)}: {fe.ENCODING_NAMES.get(info['encoding'], info['encoding'])}"
                          f", {fe.DELIMITER_NAMES.get(info['delimiter'], info['delimiter'])}")
        if info and info.get("truncated"):
            messagebox.showwarning(APP_NAME, f"Tệp có nhiều hơn {MAX_ROWS:,} hàng hoặc {MAX_COLS:,} cột; phần vượt "
                                   "quá giới hạn của một trang tính không được mở (Excel cũng giới hạn như vậy).",
                                   parent=self.root)

    def _load(self, wb, path):
        self._end_edit()
        self.wb = wb
        self.sheet = wb.active if wb.active in wb.visible_sheets() else wb.visible_sheets()[0]
        self.views.clear()
        self.top = self.left = 0
        self.cur = self.ext = (0, 0)
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.clip = self.painter = None
        self.areas = []
        self.breaks_shown.clear()
        self.path = path
        if path:
            self.settings.add_recent(path)
        self.recovery.discard()
        self.dirty = False
        self.refresh_all()

    def save_file(self):
        if not self.path:
            return self.save_as()
        return self._write(self.path)

    def save_as(self):
        self.commit_edit()
        initial = os.path.splitext(os.path.basename(self.path))[0] if self.path else "Bang tinh"
        path = filedialog.asksaveasfilename(parent=self.root, title="Lưu bảng tính", initialfile=initial,
                                            defaultextension=".xlsx" if fe.HAS_XLSX else ".pyxl",
                                            filetypes=SAVE_TYPES)
        return self._write(path) if path else False

    def _write(self, path):
        self.commit_edit()
        if path.lower().endswith(".csv") and not messagebox.askokcancel(
                APP_NAME, "CSV chỉ lưu giá trị của trang tính đang mở (không có công thức, định dạng "
                          "và các trang tính khác). Tiếp tục?", parent=self.root):
            return False
        self._busy(f"Đang lưu {os.path.basename(path)}...")
        try:
            fe.save_workbook(self.wb, path, active=self.sheet)
        except Exception as err:  # e.g. the file is open in Excel
            self._not_busy()
            self.set_mode("Sẵn sàng")
            messagebox.showerror(APP_NAME, f"Không lưu được tệp:\n{path}\n\n{err}", parent=self.root)
            return False
        self._not_busy()
        self.path = path
        self.dirty = False
        self.settings.add_recent(path)
        self.recovery.discard()
        self.recovery.beat(path=path)
        self.update_title()
        self.set_mode(f"Đã lưu {os.path.basename(path)}")
        return True

    def quit(self):
        if self.confirm_discard():
            self.recovery.close()  # nothing left to recover: changes were saved or dropped on purpose
            self.root.destroy()

    # ------------------------------------------------------------------ recent files and AutoRecover

    def _fill_recent_menu(self):
        menu = self.recent_menu
        menu.delete(0, "end")
        items = self.settings.recent()
        if not items:
            menu.add_command(label="(Chưa có tệp nào)", state="disabled")
            return
        for i, path in enumerate(items, 1):
            menu.add_command(label=f"{i}  {os.path.basename(path)}", underline=0,
                             accelerator=os.path.dirname(path), command=lambda p=path: self.open_recent(p))
        menu.add_separator()
        menu.add_command(label="Xóa danh sách gần đây", command=self.settings.clear_recent)

    def open_recent(self, path):
        if not os.path.exists(path):
            messagebox.showerror(APP_NAME, f"Không tìm thấy tệp:\n{path}", parent=self.root)
            return
        if self.confirm_discard():
            self.open_path(path)

    def autosave_options(self):
        minutes = simpledialog.askinteger(
            "Lưu tự động", "Lưu bản khôi phục sau mỗi bao nhiêu phút?\n(0 = tắt; bản khôi phục giúp lấy lại dữ liệu\n"
                          "khi máy tắt đột ngột hoặc chương trình bị lỗi)",
            initialvalue=self.settings.autosave_minutes, minvalue=0, maxvalue=120, parent=self.root)
        if minutes is not None:
            self.settings.autosave_minutes = minutes

    def _autosave_tick(self):
        """Every minute: a heartbeat; with unsaved changes, an AutoRecover copy every few minutes."""
        try:
            minutes = self.settings.autosave_minutes
            self.recovery.beat(path=self.path)
            last = self.recovery.saved_at
            busy = self._autosave_thread is not None and self._autosave_thread.is_alive()
            if (self.dirty and minutes and not busy and not self.is_editing() and not self.drag
                    and (last is None or time.time() - last >= minutes * 60)):
                # copy the workbook now (fast), write it out on another thread
                self._autosave_thread = self.recovery.save_in_background(self.wb.snapshot(), self.sheet)
                self.root.after(500, self._autosave_done)
        finally:
            self.root.after(60_000, self._autosave_tick)

    def _autosave_done(self):
        if self._autosave_thread is not None and self._autosave_thread.is_alive():
            self.root.after(500, self._autosave_done)
            return
        self._autosave_thread = None
        if self.recovery.saved_at and self.mode_label.cget("text") == "Sẵn sàng":
            self.set_mode(f"Đã lưu bản khôi phục lúc {dt.datetime.now():%H:%M}")

    def _offer_recovery(self):
        """After a crash: offer the AutoRecover copies the previous sessions left behind."""
        entries = self.recovery.orphans()
        if not entries:
            return
        options = [(f"{dt.datetime.fromtimestamp(e['saved']):%d/%m/%Y %H:%M}  —  "
                    f"{os.path.basename(e['path']) if e['path'] else 'Sổ làm việc chưa lưu'}", str(i))
                   for i, e in enumerate(entries)]
        options.append(("Không khôi phục (xóa các bản này)", "none"))
        choice = dialogs.ChoiceDialog(
            self.root, "Khôi phục tài liệu", options, "0",
            prompt="Lần trước PyExcel đã đóng đột ngột khi còn thay đổi chưa lưu.\n"
                   "Chọn bản khôi phục muốn mở (Hủy: để lại, hỏi lại lần sau):").show()
        if choice is None:
            return
        if choice != "none":
            entry = entries[int(choice)]
            if not self.confirm_discard():
                return
            try:
                wb = fe.load_workbook(entry["file"])
            except Exception as err:  # a damaged copy is reported, not fatal
                messagebox.showerror(APP_NAME, f"Không mở được bản khôi phục:\n{err}", parent=self.root)
                return
            self._load(wb, entry["path"])
            self.dirty = True  # recovered work still needs saving
            self.update_title()
            self.set_mode("Đã khôi phục — hãy lưu lại (Ctrl+S)")
        for entry in entries:
            self.recovery.forget(entry)


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # crisp text on high-DPI screens
        except (AttributeError, OSError):
            pass
    root = tk.Tk()
    scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
    width = min(round(1360 * scale), int(root.winfo_screenwidth() * 0.94))
    height = min(round(800 * scale), int(root.winfo_screenheight() * 0.88))
    root.geometry(f"{width}x{height}")
    root.minsize(round(720 * scale), round(480 * scale))
    if sys.platform == "win32":
        root.state("zoomed")  # Excel opens maximised
    SpreadsheetApp(root, sys.argv[1] if len(sys.argv) > 1 else None)
    root.mainloop()


if __name__ == "__main__":
    main()
