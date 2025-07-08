import time
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as mbox
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import datetime
import openpyxl

# --- Biến driver toàn cục ---
browser = None
delete_buttons = []
dang_xoa_hs_trung = False
dong_test_hien_tai = None  # lưu dòng hiện tại khi test
dang_xoa_hs_7980 = False
dong_hien_tai_7980 = None  # Dòng hiện tại để duyệt danh sách 7980


# --- Hàm khởi động Chrome và điền thông tin tự động ---
def launchBrowser():
    browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    browser.get("https://gdbhyt.baohiemxahoi.gov.vn/")
    
    try:
        # Đợi và nhập vào trường "macskcb"
        WebDriverWait(browser, 15).until(
            EC.presence_of_element_located((By.ID, "macskcb"))
        ).send_keys("01820")

        # Nhập vào "username"
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.ID, "username"))
        ).send_keys("001194021054")

        # Nhập vào "password"
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.ID, "password"))
        ).send_keys("Donggiahuy@123")

        # ✅ Focus vào ô Captcha (dòng bạn cần)
        WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.ID, "Captcha_TB_I"))
        )
        browser.execute_script("document.getElementById('Captcha_TB_I').focus();")

    except TimeoutException:
        print("Không tìm thấy đủ các trường nhập sau 15 giây.")
    
    return browser


# --- Giao diện GUI ---
root = tk.Tk()
root.title("Tự động hóa cổng bảo hiểm")
root.geometry("800x650")

# Hàm ghi log
def ghi_log(message):
    text_box.config(state='normal')
    text_box.insert(tk.END, message + "\n")
    text_box.see(tk.END)  # Tự động cuộn xuống dòng cuối
    text_box.config(state='disabled')

# --- Hàm mở trình duyệt ---
def mo_chrome():
    global browser
    try:
        browser = launchBrowser()
        if browser:
            ghi_log("✅ Nhập Captcha và bấm 'Đăng nhập' thủ công.")
        else:
            ghi_log("❌ Không thể khởi động trình duyệt.")
    except WebDriverException as e:
        ghi_log(f"❌ Lỗi khi mở Chrome: {e}")



# --- Hàm chung để chọn menu và lấy danh sách hồ sơ ---
def lay_danh_sach_ho_so(menu_ids, combobox_ids, output_var_name, ten_ho_so):
    global browser
    if browser is None:
        status_label.config(text="⚠️ Bạn cần mở trình duyệt trước!", fg="orange")
        return

    try:
        # 1. Đóng popup phiên bản nếu có
        try:
            popup = browser.find_element(By.ID, "popupInfoVersion_PW-1")
            if popup.is_displayed():
                browser.find_element(By.ID, "popupInfoVersion_HCB-1").click()
                time.sleep(0.5)
        except:
            pass  # Không có popup thì bỏ qua

        # 2. Click lần lượt các menu theo ID
        for menu_id in menu_ids:
            WebDriverWait(browser, 10).until(
                EC.element_to_be_clickable((By.ID, menu_id))
            ).click()

        # 3. Chọn các giá trị trong combobox lọc
        for combo_id, item_id in combobox_ids:
            WebDriverWait(browser, 10).until(
                EC.element_to_be_clickable((By.ID, combo_id))
            ).click()
            WebDriverWait(browser, 5).until(
                EC.element_to_be_clickable((By.ID, item_id))
            ).click()

        # 4. Chọn tháng từ Combobox giao diện người dùng
        thang_chon = combo_thang.get()                     # Ví dụ: "Tháng 7"
        so_thang = int(thang_chon.split()[-1])             # Lấy số 7
        index_thang = so_thang - 1                         # Index = 6

        WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.ID, "cbx_thang_I"))
        ).click()
        WebDriverWait(browser, 5).until(
            EC.element_to_be_clickable((By.ID, f"cbx_thang_DDD_L_LBI{index_thang}T0"))
        ).click()

        # 5. Bấm nút Tìm kiếm
        WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.ID, "bt_TimKiem"))
        ).click()

        # 6. Lấy kết quả tổng số hồ sơ
        summary_element = WebDriverWait(browser, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "dxp-summary"))
        )
        summary_text = summary_element.text  # Ví dụ: "Page 1 of 18 (348 items)"

        import re
        match = re.search(r"\((\d+)\s+items\)", summary_text)
        count = int(match.group(1)) if match else 0

        # 7. Gán biến toàn cục theo tên output_var_name
        globals()[output_var_name] = count

        # 8. Ghi log kết quả
        ghi_log(f"✅ Đã lấy được {count} hồ sơ {ten_ho_so}")

    except Exception as e:
        status_label.config(
            text=f"❌ Lỗi khi xử lý: {e}",
            fg="red"
        )




# Load hồ sơ trùng
def load_ho_so_trung():
    menu_ids = [
        "HeaderMenu_DXI2_T",         # 🧭 Menu cấp 1: "Hồ sơ đề nghị thanh toán"
        "HeaderMenu_DXI2i0_T",       # 📄 Menu cấp 2: "Hồ sơ XML"
        "HeaderMenu_DXI2i0i2_T",     # 📋 Menu cấp 3: "Danh sách đề nghị thanh toán"
    ]
    combobox_ids = [
        ("cb_TrangThaiHS_I", "cb_TrangThaiHS_DDD_L_LBI2T0"),  # ✅ Trạng thái hồ sơ: "Hồ sơ trùng"
        ("cb_TrangThaiTT_I", "cb_TrangThaiTT_DDD_L_LBI0T0"),  # 💰 Trạng thái thanh toán: "Tất cả"
    ]
    lay_danh_sach_ho_so(menu_ids, combobox_ids, "so_ho_so_trung", "trùng")



# Load hồ sơ 7980
def load_ho_so_7980():
    menu_ids = [
        "HeaderMenu_DXI2_T",         # 🧭 Menu cấp 1: "Hồ sơ đề nghị thanh toán"
        "HeaderMenu_DXI2i1_T",       # 📁 Menu cấp 2: "Hồ sơ 7980"
        "HeaderMenu_DXI2i1i2_",     # 📋 Menu cấp 3: "Danh sách đề nghị thanh toán 7980a"
    ]
    combobox_ids = [
        ("cb_TrangThaiHS_I", "cb_TrangThaiHS_DDD_L_LBI0T0"),  # ✅ Trạng thái hồ sơ: "Tất cả"
        ("cb_TrangThaiTT_I", "cb_TrangThaiTT_DDD_L_LBI0T0"),  # 💰 Trạng thái thanh toán: "Tất cả"
    ]
    lay_danh_sach_ho_so(menu_ids, combobox_ids, "so_ho_so_7980", "7980")








# Xóa hồ sơ trùng
def xoa_danh_sach_ho_so_trung():
    global so_ho_so_trung, dang_xoa_hs_trung

    if not so_ho_so_trung or so_ho_so_trung <= 0:
        ghi_log("⚠️ Không có hồ sơ trùng để xóa. Vui lòng bấm 'OK' trước.")
        return

    if not dang_xoa_hs_trung:
        return  # Nếu không ở chế độ xóa, thoát

    def xoa_tung_ho_so(i):
        global dang_xoa_hs_trung

        if i >= so_ho_so_trung:
            mbox.showinfo("Hoàn thành", "✅ Đã xóa hết tất cả các hồ sơ.")
            ghi_log("✅ Đã xóa xong toàn bộ hồ sơ trùng.")
            btn_delete_hs_trung.config(text="Xóa HS Trùng")
            dang_xoa_hs_trung = False
            return

        if not dang_xoa_hs_trung:
            ghi_log("⏹️ Đã dừng thao tác xóa theo yêu cầu.")
            btn_delete_hs_trung.config(text="Xóa HS Trùng")
            return

        try:
            # Lấy dòng đầu tiên và icon xóa
            row = WebDriverWait(browser, 5).until(
                EC.presence_of_element_located((By.XPATH, "//tr[contains(@id, 'gvDanhSachHoSo_DXDataRow')]"))
            )
            cols = row.find_elements(By.TAG_NAME, "td")
            ho_ten = cols[8].text.strip() if len(cols) > 8 else "Không rõ tên"
            icon_xoa = cols[26].find_element(By.TAG_NAME, "img")

            # Click icon xóa
            browser.execute_script("arguments[0].click();", icon_xoa)
            ghi_log(f"🗑️ Đã xóa hồ sơ: {ho_ten}")

            # Đợi popup xác nhận xóa và click nút "Có"
            WebDriverWait(browser, 5).until(
                EC.visibility_of_element_located((By.ID, "PopupThongBaoXoa_PWH-1"))
            )
            btn_co = WebDriverWait(browser, 3).until(
                EC.element_to_be_clickable((By.ID, "btnCo"))
            )
            btn_co.click()

            # Đợi popup thông báo thành công và đóng lại
            WebDriverWait(browser, 5).until(
                EC.visibility_of_element_located((By.ID, "popup_message_PW-1"))
            )
            btn_dong_thong_bao = WebDriverWait(browser, 3).until(
                EC.element_to_be_clickable((By.ID, "popup_message_HCB-1"))
            )
            btn_dong_thong_bao.click()

        except Exception as e:
            ghi_log(f"❌ Lỗi khi xóa dòng {i+1}: {e}")

        # Tiếp tục sau một chút delay
        root.after(700, lambda: xoa_tung_ho_so(i + 1))

    xoa_tung_ho_so(0)




def toggle_xoa_ho_so_trung():
    global dang_xoa_hs_trung

    if not dang_xoa_hs_trung:
        # Bắt đầu xóa
        dang_xoa_hs_trung = True
        btn_delete_hs_trung.config(text="⏹️ Dừng")
        xoa_danh_sach_ho_so_trung()
    else:
        # Dừng thao tác xóa
        dang_xoa_hs_trung = False
        btn_delete_hs_trung.config(text="Xóa HS Trùng")








# Hàm để chọn file Excel
def chon_file_excel():
    global dong_test_hien_tai
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],
        title="Chọn file Excel"
    )
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)
        dong_test_hien_tai = None  # Reset dòng test
        ghi_log(f"📂 Đã chọn file: {file_path}")

# Test file excel
def test_in_thong_tin_excel():
    global dong_test_hien_tai

    ma_the_col = combo_mt.get().strip().upper()
    ho_ten_col = combo_ht.get().strip().upper()
    ngay_vao_col = combo_nv.get().strip().upper()
    ngay_ra_col = combo_nr.get().strip().upper()

    file_path = entry_file_path.get().strip()
    if not file_path:
        ghi_log("❌ Chưa chọn file Excel.")
        return

    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
    except Exception as e:
        ghi_log(f"❌ Lỗi khi mở file Excel: {e}")
        return

    try:
        start = int(entry_start.get().strip())
        end = int(entry_end.get().strip())
    except ValueError:
        ghi_log("⚠️ Vui lòng nhập số nguyên cho dòng bắt đầu và kết thúc.")
        return

    if dong_test_hien_tai is None:
        dong_test_hien_tai = start

    if dong_test_hien_tai > end:
        ghi_log("✅ Đã test đến dòng cuối cùng.")
        return

    try:
        ma_the_val = ws[f"{ma_the_col}{dong_test_hien_tai}"].value
        ho_ten_val = ws[f"{ho_ten_col}{dong_test_hien_tai}"].value
        ngay_vao_val = ws[f"{ngay_vao_col}{dong_test_hien_tai}"].value
        ngay_ra_val = ws[f"{ngay_ra_col}{dong_test_hien_tai}"].value

        # Gộp thông tin thành 1 dòng log
        log_line = f"{dong_test_hien_tai}: {ma_the_val} | {ho_ten_val} | {ngay_vao_val} | {ngay_ra_val}"
        ghi_log(log_line)

    except Exception as e:
        ghi_log(f"❌ Lỗi khi đọc dòng {dong_test_hien_tai}: {e}")

    dong_test_hien_tai += 1


# Duyệt qua danh sách excel, tìm kiếm, so sánh và xóa hồ sơ 7980
def xoa_ho_so_7980():
    global dong_hien_tai, dang_xoa_hs_7980

    if not dang_xoa_hs_7980:
        # Bắt đầu tiến trình xoá
        btn_delete_hs_7980.config(text="⏹️ Dừng xóa")
        ghi_log("🚀 Bắt đầu xoá hồ sơ 79/80...")
        dong_hien_tai = None
        dang_xoa_hs_7980 = True
        xoa_tiep_dong_7980()
    else:
        # Dừng tiến trình xoá
        dang_xoa_hs_7980 = False
        btn_delete_hs_7980.config(text="Xóa HS 79/80")
        ghi_log("⏹️ Đã dừng xoá theo yêu cầu.")

def xoa_tiep_dong_7980():
    global dong_hien_tai, dang_xoa_hs_7980

    if not dang_xoa_hs_7980:
        return

    # --- Đọc dữ liệu từ file Excel ---
    file_path = entry_file_path.get().strip()
    if not file_path:
        ghi_log("❌ Chưa chọn file Excel.")
        return

    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
    except Exception as e:
        ghi_log(f"❌ Lỗi khi mở file Excel: {e}")
        return

    try:
        start = int(entry_start.get().strip())
        end = int(entry_end.get().strip())
    except ValueError:
        ghi_log("⚠️ Vui lòng nhập số nguyên cho dòng bắt đầu và kết thúc.")
        return

    if dong_hien_tai is None:
        dong_hien_tai = start

    if dong_hien_tai > end:
        ghi_log("✅ Đã duyệt hết tất cả các dòng.")
        btn_delete_hs_7980.config(text="Xóa HS 79/80")
        dang_xoa_hs_7980 = False
        return

    try:
        # --- Lấy dữ liệu từ dòng hiện tại ---
        ma_the_col = combo_mt.get().strip().upper()
        ho_ten_col = combo_ht.get().strip().upper()
        ngay_vao_col = combo_nv.get().strip().upper()
        ngay_ra_col = combo_nr.get().strip().upper()

        ma_the_val = ws[f"{ma_the_col}{dong_hien_tai}"].value
        ho_ten_val = str(ws[f"{ho_ten_col}{dong_hien_tai}"].value).strip()
        ngay_vao_val = str(ws[f"{ngay_vao_col}{dong_hien_tai}"].value).strip()
        ngay_ra_val = str(ws[f"{ngay_ra_col}{dong_hien_tai}"].value).strip()

        if not ma_the_val:
            ghi_log(f"{dong_hien_tai}: ⚠️ Không có mã thẻ. Bỏ qua.")
            dong_hien_tai += 1
            root.after(500, xoa_tiep_dong_7980)
            return

        # --- Tìm lại input mỗi lần để tránh stale element ---
        input_box = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.ID, "gvDanhSachBHYT7980_DXFREditorcol3_I"))
        )

        # --- Xóa dữ liệu cũ trong input mã thẻ ---
        input_box.clear()

        # --- Chờ loading sau khi clear input ---
        try:
            WebDriverWait(browser, 5).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingDiv_EIS"))
            )
        except:
            pass

        # --- Tìm lại input box lần nữa trước khi nhập mã thẻ ---
        input_box = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.ID, "gvDanhSachBHYT7980_DXFREditorcol3_I"))
        )
        input_box.send_keys(str(ma_the_val))

        # --- Chờ loading sau khi nhập mã thẻ ---
        try:
            WebDriverWait(browser, 5).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingDiv_EIS"))
            )
        except:
            pass

        # --- Tìm danh sách kết quả ---
        rows = browser.find_elements(By.XPATH, "//tr[starts-with(@id, 'gvDanhSachBHYT7980_DXDataRow')]")
        found = False

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 19:
                ho_ten = cols[5].text.strip()
                ngay_vao = cols[8].text.strip().replace(" SA", "").replace(" CH", "")
                ngay_ra = cols[9].text.strip().replace(" SA", "").replace(" CH", "")

                if ho_ten == ho_ten_val and ngay_vao == ngay_vao_val and ngay_ra == ngay_ra_val:
                    try:
                        # 🔄 Tìm lại cột trước khi bấm xoá
                        cols = row.find_elements(By.TAG_NAME, "td")
                        delete_btn = cols[18].find_element(By.TAG_NAME, "input")

                        browser.execute_script("arguments[0].click();", delete_btn)

                        # ⏳ Chờ popup xác nhận xóa hiện ra
                        WebDriverWait(browser, 5).until(
                            EC.visibility_of_element_located((By.ID, "PopupThongBaoXoa_PWH-1"))
                        )

                        # ✅ Tự động bấm vào nút "Không"
                        btn_khong = WebDriverWait(browser, 3).until(
                            EC.element_to_be_clickable((By.ID, "btnKhong_CD"))
                        )
                        btn_khong.click()

                        ghi_log(f"{dong_hien_tai}: 🗑️ Đã xoá: {ho_ten}")
                        found = True
                        break
                    except Exception as e:
                        ghi_log(f"{dong_hien_tai}: ❌ Không thể xoá: {e}")

        if not found:
            ghi_log(f"{dong_hien_tai}: ❌ Không tìm thấy hồ sơ phù hợp để xoá.")

    except Exception as e:
        ghi_log(f"{dong_hien_tai}: ❌ Lỗi: {e}")

    dong_hien_tai += 1
    root.after(700, xoa_tiep_dong_7980)


























# --- Nút Đăng nhập căn giữa ---
btn_login = tk.Button(root, text="Đăng nhập cổng bảo hiểm", font=("Arial", 12), command=mo_chrome)
btn_login.pack(pady=10)

import datetime

# Lấy tháng hiện tại của hệ thống để truyền vào combobox
thang_hien_tai = datetime.datetime.now().month
index_mac_dinh = thang_hien_tai - 1  # Vì Combobox index bắt đầu từ 0

# --- Combobox chọn tháng ---
selected_thang = tk.StringVar()
combo_thang = ttk.Combobox(root, textvariable=selected_thang, font=("Arial", 11), width=10, state="readonly")
combo_thang['values'] = [f"Tháng {i}" for i in range(1, 13)]
combo_thang.current(index_mac_dinh)
combo_thang.pack(pady=5)


# --- Frame chứa 2 cột, KHÔNG khung viền ---
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# --- Cột trái: Hồ sơ trùng ---
left_column = tk.LabelFrame(button_frame, text="Hồ sơ trùng", font=("Arial", 10, "bold"), bd=2, relief="groove", padx=10, pady=10)
left_column.pack(side="left", padx=20)

btn_load_hs_trung = tk.Button(left_column, text="Load hồ sơ trùng", font=("Arial", 12), command=load_ho_so_trung)
btn_load_hs_trung.pack(pady=5)

btn_delete_hs_trung = tk.Button(left_column, text="Xóa HS Trùng", font=("Arial", 12), command=toggle_xoa_ho_so_trung)
btn_delete_hs_trung.pack(pady=5)

# --- Cột phải: Hồ sơ 79/80 ---
right_column = tk.LabelFrame(button_frame, text="Hồ sơ 79/80", font=("Arial", 10, "bold"), bd=2, relief="groove", padx=10, pady=10)
right_column.pack(side="left", padx=20)

btn_load_hs_7980 = tk.Button(right_column, text="Load hồ sơ 79/80", font=("Arial", 12), command=load_ho_so_7980)
btn_load_hs_7980.pack(pady=5)

# --- Dòng chọn file Excel ---
file_select_frame = tk.Frame(right_column)
file_select_frame.pack(pady=5)

entry_file_path = tk.Entry(file_select_frame, width=25, font=("Arial", 10))
entry_file_path.pack(side="left", padx=5)

btn_browse_file = tk.Button(file_select_frame, text="Chọn file", command=chon_file_excel)
btn_browse_file.pack(side="left")

# --- Dòng chọn cột từ Excel ---
column_select_frame = tk.Frame(right_column)
column_select_frame.pack(pady=5)

# Danh sách chữ cái A-Z
chu_cai_list = [chr(i) for i in range(65, 91)]  # Từ 'A' đến 'Z'

# Tạo từng nhãn và combobox
label_mt = tk.Label(column_select_frame, text="Mã thẻ")
label_mt.pack(side="left", padx=2)
combo_mt = ttk.Combobox(column_select_frame, values=chu_cai_list, width=3)
combo_mt.pack(side="left", padx=2)

label_ht = tk.Label(column_select_frame, text="Họ tên")
label_ht.pack(side="left", padx=2)
combo_ht = ttk.Combobox(column_select_frame, values=chu_cai_list, width=3)
combo_ht.pack(side="left", padx=2)

label_nv = tk.Label(column_select_frame, text="Ngày vào")
label_nv.pack(side="left", padx=2)
combo_nv = ttk.Combobox(column_select_frame, values=chu_cai_list, width=3)
combo_nv.pack(side="left", padx=2)

label_nr = tk.Label(column_select_frame, text="Ngày ra")
label_nr.pack(side="left", padx=2)
combo_nr = ttk.Combobox(column_select_frame, values=chu_cai_list, width=3)
combo_nr.pack(side="left", padx=2)

# --- Hàng nhập dòng bắt đầu / kết thúc ---
row_range_frame = tk.Frame(right_column)
row_range_frame.pack(pady=5)

# Nhãn và input: Dòng bắt đầu
label_start = tk.Label(row_range_frame, text="Dòng bắt đầu")
label_start.pack(side="left", padx=2)
entry_start = tk.Entry(row_range_frame, width=6)
entry_start.pack(side="left", padx=2)

# Nhãn và input: Dòng kết thúc
label_end = tk.Label(row_range_frame, text="Dòng kết thúc")
label_end.pack(side="left", padx=2)
entry_end = tk.Entry(row_range_frame, width=6)
entry_end.pack(side="left", padx=2)

# Nút Test
btn_test_range = tk.Button(row_range_frame, text="Test", command=test_in_thong_tin_excel)
btn_test_range.pack(side="left", padx=5)


btn_delete_hs_7980 = tk.Button(right_column, text="Xóa HS 79/80", font=("Arial", 12), command=xoa_ho_so_7980)
btn_delete_hs_7980.pack(pady=5)

# --- Status và Text Box ---
status_label = tk.Label(root, text="", font=("Arial", 10))
status_label.pack()

text_box = tk.Text(root, height=25, font=("Arial", 11))
text_box.pack(padx=10, pady=10, fill="both")
text_box.config(state='disabled')

root.mainloop()
