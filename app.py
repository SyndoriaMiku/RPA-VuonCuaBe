import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.scrolledtext as scrolledtext
import openpyxl
import time
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

def choose_file():
    file_path = filedialog.askopenfilename(
        title="Chọn File Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
    )
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)
        
def start_automation_thread():
    # Làm mờ nút chạy để ngăn người dùng bấm nhiều lần
    btn_run.config(state=tk.DISABLED)
    
    # Tạo một luồng phụ để chạy tool
    thread = threading.Thread(target=run_automation_worker)
    thread.daemon = True # Đảm bảo luồng phụ tự tắt khi bạn ấn X đóng phần mềm
    thread.start()

def run_automation_worker():
    # Gọi hàm chạy chính
    run_automation()
    # Sau khi hàm chính chạy xong (hoặc bị lỗi), bật lại nút bấm
    btn_run.config(state=tk.NORMAL)        

def run_automation():
    file_path = entry_file_path.get()
    sheet_name = entry_sheet.get()
    
    if not file_path:
        messagebox.showerror("Lỗi", "Vui lòng chọn file Excel trước khi chạy.")
        return
    
    # ---------------------------------------------------------
    # PHẦN 1: ĐỌC TOÀN BỘ DỮ LIỆU EXCEL TỪ Ô B4
    # ---------------------------------------------------------
    data_list = []
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        if sheet_name:
            if sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
            else:
                messagebox.showerror("Lỗi", f"Không tìm thấy Sheet '{sheet_name}'.")
                return
        else:
            sheet = wb.active
            
        current_row = 4
        while True:
            # Lấy giá trị cột B (Mã hàng)
            code = sheet[f"B{current_row}"].value
            
            # Điều kiện dừng: Nếu ô B trống (None) thì kết thúc vòng lặp
            if not code:
                break
                
            # Lấy giá trị cột D (Số lượng)
            qty_val = sheet[f"D{current_row}"].value
            qty = float(qty_val) if qty_val is not None else 0
            
            # Lưu vào danh sách
            data_list.append({
                "code": str(code).strip(),
                "qty": qty,
                "row": current_row
            })
            current_row += 1
            
    except Exception as e:
        messagebox.showerror("Lỗi Đọc Excel", f"Có lỗi xảy ra khi đọc file:\n{e}")
        return

    if not data_list:
        messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu nào bắt đầu từ ô B4.")
        return

    # ---------------------------------------------------------
    # PHẦN 2: CHẠY VÒNG LẶP TRÊN WEB (SELENIUM)
    # ---------------------------------------------------------
    log_messages = []
    
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Chrome(options=chrome_options)
        
        # --- BẢN VÁ MỚI: TỰ ĐỘNG CHUYỂN ĐÚNG TAB KIOTVIET ---
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if "kiotviet.vn" in driver.current_url.lower():
                break
        # ----------------------------------------------------

        wait = WebDriverWait(driver, 10)
        
        txt_log.insert(tk.END, f"\n----------------------------------------\n")
        txt_log.insert(tk.END, f"Bắt đầu chạy {len(data_list)} mã hàng...\n")
        
        # Vì chạy đa luồng, dùng tk.END thay thế root.update() chỗ này cho an toàn
        txt_log.see(tk.END) 

        # Vòng lặp duyệt qua từng dòng dữ liệu lấy từ Excel
        for item in data_list:
            code = item["code"]
            qty_excel = item["qty"]
            
            try:
                # Step 1: Tìm ô input và click vào nó trước
                search_input = wait.until(EC.element_to_be_clickable((By.ID, "productSearchInput")))
                
                # Giả lập thao tác click của người dùng
                search_input.click()
                time.sleep(0.5)
                
                search_input.send_keys(Keys.CONTROL + "a")
                search_input.send_keys(Keys.DELETE)
                search_input.send_keys(code)
                
                # Chờ cứng 2 giây cho web load kết quả mới
                time.sleep(2) 
                
                # Step 2: Click vào kết quả bằng Javascript (Né AdGuard)
                first_result = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".output-complete ul li")))
                driver.execute_script("arguments[0].click();", first_result)
                
                # Step 3: Đợi input số lượng hiện ra
                input_td = wait.until(EC.presence_of_element_located((
                    By.CSS_SELECTOR, "table[role='grid'] tbody tr:nth-child(1) td.cell-order-number input"
                )))
                
                # Step 4: Lấy số lượng trên web
                list_cells = driver.find_elements(By.CSS_SELECTOR, "table[role='grid'] tbody tr:nth-child(1) td.cell-quantity.txtR")
                txt_web_value = list_cells[0].text.strip()
                txt_web_value = txt_web_value.replace(",", "").replace(" ", "")
                qty_web = float(txt_web_value) if txt_web_value else 0
                
                # Step 5: Tính tổng và kiểm tra điều kiện
                total_value = qty_web + qty_excel
                
                if total_value < 0:
                    error_msg = f"- Mã hàng {code} số lượng cân chỉnh không hợp lệ (Tổng: {total_value})"
                    log_messages.append(error_msg)
                    txt_log.insert(tk.END, error_msg + "\n")
                    txt_log.see(tk.END)
                    continue 
                
                # Step 6: Điền số lượng bằng API AngularJS
                driver.execute_script("""
                    var elm = arguments[0];
                    var val = arguments[1];
                    var $el = angular.element(elm);
                    $el.val(val).triggerHandler('input');
                    $el.triggerHandler('change');
                    $el.triggerHandler('blur');
                """, input_td, total_value)
                
                
                time.sleep(1) 
                
            except Exception as e_row:
                
                chi_tiet_loi = str(e_row)
                if "Stacktrace" in chi_tiet_loi or "Timeout" in chi_tiet_loi:
                    error_msg = f"- BỎ QUA: Mã {code} - Không tìm thấy sản phẩm trên KiotViet."
                else:
                    error_msg = f"- LỖI Web: Mã {code}. Chi tiết: {chi_tiet_loi.split(chr(10))[0]}"
                
                log_messages.append(error_msg)
                txt_log.insert(tk.END, error_msg + "\n")
                txt_log.see(tk.END)
                continue

        # Sau khi chạy xong toàn bộ danh sách
        txt_log.insert(tk.END, "--- HOÀN TẤT QUÁ TRÌNH ---\n")
        txt_log.see(tk.END)
        messagebox.showinfo("Thành công", f"Đã chạy xong {len(data_list)} mã hàng!\nXem Log để biết chi tiết.")

    except Exception as e:
        messagebox.showerror("Lỗi Hệ Thống", f"Không thể kết nối với Web:\n{e}")

# --- THIẾT KẾ GIAO DIỆN (GUI) ---
root = tk.Tk()
root.title("Auto Robot")
root.geometry("850x650")
root.eval('tk::PlaceWindow . center')

tk.Label(root, text="File Excel:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_file_path = tk.Entry(root, width=45)
entry_file_path.grid(row=0, column=1, padx=5, pady=10)
btn_browse = tk.Button(root, text="Chọn File", command=choose_file)
btn_browse.grid(row=0, column=2, padx=5, pady=10)

tk.Label(root, text="Tên Sheet (bỏ trống=mặc định):").grid(row=1, column=0, padx=10, pady=10, sticky="e")
entry_sheet = tk.Entry(root, width=20)
entry_sheet.grid(row=1, column=1, padx=5, pady=10, sticky="w")

# ĐÃ SỬA: Thay run_automation thành start_automation_thread để kích hoạt đa luồng
btn_run = tk.Button(root, text="▶ CHẠY TỰ ĐỘNG", command=start_automation_thread, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"))
btn_run.grid(row=2, column=0, columnspan=3, pady=15, ipadx=20, ipady=5)

# Khung Log
tk.Label(root, text="Nhật ký chạy (Log):").grid(row=3, column=0, padx=10, sticky="nw")
txt_log = scrolledtext.ScrolledText(root, width=65, height=10, fg="red", bg="#f9f9f9")
txt_log.grid(row=4, column=0, columnspan=3, padx=10, pady=5)

root.mainloop()