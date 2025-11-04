import os
import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ============================================================
# 🧩 PHẦN 1: XỬ LÝ FILE EXCEL
# ============================================================

def read_login_data(file_path):
    """Đọc dữ liệu đăng nhập từ file Excel"""
    df = pd.read_excel(file_path, engine="openpyxl")
    return df

def write_result(file_path, df):
    """Ghi kết quả đăng nhập ra file mới"""
    df.to_excel(file_path, index=False, engine="openpyxl")
    print(f"✅ Kết quả đã được lưu vào: {file_path}")

def add_result_columns(df):
    """Thêm cột kết quả nếu chưa có"""
    for col in ["result", "message", "timestamp"]:
        if col not in df.columns:
            df[col] = ""
    return df

def update_row(df, index, result, message):
    """Cập nhật kết quả cho 1 dòng"""
    df.at[index, "result"] = result
    df.at[index, "message"] = message
    df.at[index, "timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ============================================================
# 🌐 PHẦN 2: TỰ ĐỘNG LOGIN VỚI SELENIUM
# ============================================================

URL = "https://the-internet.herokuapp.com/login"

def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--log-level=3")

    # 🚫 Tắt toàn bộ password manager + leak detection
    prefs = {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "profile.password_manager_leak_detection": False,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)

    # 🧩 Dùng profile tạm để Chrome “sạch”, không lưu mật khẩu
    options.add_argument("--user-data-dir=C:\\Temp\\SeleniumProfile")

    # ⚙️ Bỏ cảnh báo "Chrome is being controlled by automated test software"
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def attempt_login(driver, username, password):
    """Thực hiện đăng nhập và trả kết quả"""
    driver.get(URL)
    try:
        driver.find_element(By.ID, "username").send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

        wait = WebDriverWait(driver, 5)
        try:
            success = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".flash.success")))
            return "Success", success.text.strip()
        except:
            error = driver.find_element(By.CSS_SELECTOR, ".flash.error")
            return "Fail", error.text.strip()
    except Exception as e:
        return "Error", str(e)


# ============================================================
# 🚀 PHẦN 3: CHẠY CHÍNH
# ============================================================

INPUT_FILE = r"C:\Users\ADMIN\Documents\Automation Testing\selenium_excel_demo\data\login_data.xlsx"
OUTPUT_FILE = r"C:\Users\ADMIN\Documents\Automation Testing\selenium_excel_demo\data\login_result.xlsx"

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Không tìm thấy file input tại:\n{INPUT_FILE}")
        return

    df = read_login_data(INPUT_FILE)
    df = add_result_columns(df)

    driver = setup_driver()

    try:
        for i, row in df.iterrows():
            user = str(row["username"])
            pwd = str(row["password"])
            print(f"🔹 Thử đăng nhập: {user} / {pwd}")
            result, message = attempt_login(driver, user, pwd)
            update_row(df, i, result, message)
            print(f"   → {result}: {message[:60]}")
            time.sleep(1)
    finally:
        driver.quit()

    write_result(OUTPUT_FILE, df)
    print("\n✅ Hoàn tất kiểm thử batch login!")

if __name__ == "__main__":
    main()
