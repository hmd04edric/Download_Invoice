import os
import time
import traceback
import pandas as pd
import xml.etree.ElementTree as ET
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC

# ========== CẤU HÌNH ========== #
INPUT_FILE = r'week4\file_input\input.xlsx'
OUTPUT_FILE = r'week4\file_output\output.xlsx'
DOWNLOAD_FOLDER = 'week4/file_output'


def handle_input():
    return pd.read_excel(INPUT_FILE, dtype=str)


def open_browser(download_path):
    options = Options()
    prefs = {
        "download.default_directory": os.path.abspath(download_path),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)


# ========== CHECK LOAD ========== #
def check_load_success(driver, system):
    try:
        if system == "fpt":
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div[3]/div/div[2]/div[2]'))
            )
        elif system == "misa":
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CLASS_NAME, "invNo"))
            )
        elif system == "van":
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "frameViewInvoice"))
            )
        print(f"{system.upper()}: Tra cứu thành công")
        return True
    except:
        return False


def check_load_fail(driver, system): 
    try:
        if system == "fpt":
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/div/div'))
                )
                print("FPT: Không tìm thấy hóa đơn")
                return True
            except:
                pass
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div[3]/div/div[1]/div/div[2]/div[2]'))
                )
                print("FPT: MST không đúng")
                return True
            except:
                pass

        elif system == "misa":
            try:
                popup = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.ID, "showPopupInvoicNotExist"))
                )
                if popup.is_displayed():
                    print("MISA: Không tìm thấy hóa đơn")
                    return True
            except:
                pass

        elif system == "van":
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "Bkav_alert_dialog"))
                )
                print("VAN: Mã tra cứu không đúng.")
                return True
            except:
                pass

            time.sleep(2)
            frames = driver.find_elements(By.ID, "frameViewInvoice")
            if not frames:
                print("VAN: Không tìm thấy iframe -> lỗi tra cứu")
                return True

        return False
    except:
        return False


# ========== XML ========== #
def get_latest_xml(download_folder):
    files = [f for f in os.listdir(download_folder) if f.lower().endswith(".xml")]
    if not files:
        return None
    files = sorted(files, key=lambda x: os.path.getmtime(os.path.join(download_folder, x)), reverse=True)
    return os.path.join(download_folder, files[0])


def extract_invoice_data_from_xml(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        dlh_don = root.find(".//DLHDon")
        if dlh_don is None:
            print("Không tìm thấy phần tử DLHDon trong file.")
            return {}

        so_hd = dlh_don.findtext(".//SHDon", default="")
        nban = dlh_don.find(".//NBan")
        ten_nban = nban.findtext("Ten", default="")
        mst_nban = nban.findtext("MST", default="")
        dchi_nban = nban.findtext("DChi", default="")
        sdt_nban = nban.findtext("SDThoai", default="") or ""
        stk_nban = nban.findtext("STKNHang", default="")

        nmua = dlh_don.find(".//NMua")
        ten_nmua = nmua.findtext("Ten", default="")
        mst_nmua = nmua.findtext("MST", default="")
        dchi_nmua = nmua.findtext("DChi", default="")
        stk_nmua = nmua.findtext("STKNHang", default="")

        return {
            "Số hóa đơn": so_hd,
            "Đơn vị bán hàng": ten_nban,
            "Mã số thuế bán": mst_nban,
            "Địa chỉ bán": dchi_nban,
            "Điện thoại": sdt_nban,
            "Số tài khoản": stk_nban,
            "Họ tên người mua hàng": ten_nmua,
            "Mã số thuế mua": mst_nmua,
            "Địa chỉ mua": dchi_nmua,
            "Số tài khoản người mua": stk_nmua
        }

    except Exception as e:
        print("Lỗi xử lý XML:", e)
        traceback.print_exc()
        return {}


# ========== XỬ LÝ FPT ========== #
def process_fpt_invoice(driver, url, ma_so_thue, ma_tra_cuu):
    print("FPT: Nhập mã tra cứu...")
    driver.get(url)
    time.sleep(2)

    try:
        driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[3]/div/div[1]/div/div[2]/div/input').send_keys(ma_so_thue)
        driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[3]/div/div[1]/div/div[3]/div/input').send_keys(ma_tra_cuu)
        driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div[3]/div/div[1]/div/div[4]/div[2]/div/button').click()
        print("FPT: Đã nhấn tra cứu.")
    except Exception as e:
        print("FPT: Lỗi khi nhập hoặc nhấn tra cứu:", e)
        traceback.print_exc()
        return False

    time.sleep(3)
    if check_load_fail(driver, "fpt"):
        return False
    elif check_load_success(driver, "fpt"):
        print("FPT: Tra cứu thành công")
        try:
            print("FPT: Đang click nút tải XML...")
            btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div/div[3]/div/div[1]/div/div[4]/div[2]/div/button'))
            )
            btn.click()
            print("FPT: Đã tải XML thành công.")
            return True
        except Exception as e:
            print("FPT: Lỗi khi tải XML:", e)
            traceback.print_exc()
            return False

    return False


# ========== XỬ LÝ MISA ========== #
def process_misa_invoice(driver, url, ma_tra_cuu):
    print("MISA: Nhập mã tra cứu...")
    driver.get(url)
    time.sleep(2)

    try:
        driver.find_element(By.XPATH, '//*[@id="txtCode"]').send_keys(ma_tra_cuu)
        driver.find_element(By.XPATH, '//*[@id="btnSearchInvoice"]').click()
        print("MISA: Đã nhấn tra cứu.")
    except Exception as e:
        print("MISA: Lỗi khi nhập hoặc nhấn tra cứu:", e)
        traceback.print_exc()
        return False

    time.sleep(3)
    if check_load_fail(driver, "misa"):
        return False

    elif check_load_success(driver, "misa"):
        print("MISA: Tra cứu thành công")
        try:
            print("MISA: Đang click nút xem hóa đơn...")
            btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-content-container"]/div[1]/div[2]/div[12]/div'))
            )
            btn.click()

            print("MISA: Đang click nút tải XML...")
            xml_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-content-container"]/div[1]/div[2]/div[12]/div/div/div[2]'))
            )
            xml_element.click()
            print("MISA: Đã tải XML thành công.")
            return True
        except Exception as e:
            print("MISA: Lỗi khi tải XML:", e)
            traceback.print_exc()
            return False
    

    return False


# ========== XỬ LÝ VAN ========== #
def process_van_invoice(driver, url, ma_tra_cuu):
    print("VAN: Nhập mã tra cứu...")
    driver.get(url)
    time.sleep(2)

    try:
        driver.find_element(By.ID, "txtInvoiceCode").send_keys(ma_tra_cuu)
        driver.find_element(By.ID, "Button1").click()
        print("VAN: Đã nhấn tra cứu.")
    except Exception as e:
        print("VAN: Lỗi khi nhập hoặc nhấn tra cứu:", e)
        traceback.print_exc()
        return False

    time.sleep(2)
    if check_load_fail(driver, "van"):
        return False
    elif check_load_success(driver, "van"):
        print("VAN: Tra cứu thành công")

        try:
            print("VAN: Đang kiểm tra iframe...")
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameViewInvoice'))
            )
            print("VAN: Đã switch vào iframe.")
        except Exception as e:
            print("VAN: Lỗi khi switch vào iframe:", e)
            traceback.print_exc()
            return False

        try:
            print("VAN: Đang tìm và click LinkDownXML...")
            link_xml = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, 'LinkDownXML'))
            )
            driver.execute_script("arguments[0].click();", link_xml)
            print("VAN: Đã tải XML thành công.")
            driver.switch_to.default_content()
            return True
        except Exception as e:
            print("VAN: Lỗi khi tải XML:", e)
            traceback.print_exc()
            return False

    

    return False


# ========== CHẠY TOÀN BỘ ========== #
def process_invoice(df):
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

    if os.path.exists(OUTPUT_FILE):
        os.remove(OUTPUT_FILE)
    for f in os.listdir(DOWNLOAD_FOLDER):
        if f.lower().endswith(".xml"):
            try:
                os.remove(os.path.join(DOWNLOAD_FOLDER, f))
            except Exception as e:
                print(f"Lỗi xóa file XML cũ {f}: {e}")

    for index, row in df.iterrows():
        ma_so_thue = str(row['Mã số thuế']).strip()
        ma_tra_cuu = str(row['Mã tra cứu']).strip()
        url = str(row['URL']).strip()

        driver = open_browser(DOWNLOAD_FOLDER)
        result = False

        try:
            if "fpt" in url:
                result = process_fpt_invoice(driver, url, ma_so_thue, ma_tra_cuu)
            elif "meinvoice" in url:
                result = process_misa_invoice(driver, url, ma_tra_cuu)
            elif "van.ehoadon" in url:
                result = process_van_invoice(driver, url, ma_tra_cuu)

            df.at[index, "Status"] = "success" if result else "fail"

            if result:
                time.sleep(2)
                xml_file = get_latest_xml(DOWNLOAD_FOLDER)
                if xml_file:
                    info = extract_invoice_data_from_xml(xml_file)
                    for key, val in info.items():
                        df.at[index, key] = val
        except Exception as e:
            print("Lỗi xử lý:", e)
            traceback.print_exc()
            df.at[index, "Status"] = "fail"
        finally:
            time.sleep(2)
            driver.quit()

    df.to_excel(OUTPUT_FILE, index=False)


# ========== CHẠY CHÍNH ========== #
def main():
    df = handle_input()
    process_invoice(df)

if __name__ == "__main__":
    main()
