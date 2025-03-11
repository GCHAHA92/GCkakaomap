import logging
import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
import datetime
import holidays
import os
import sys
import pyautogui
import time
import pyperclip

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 평일/휴일 시간 설정 함수
def get_time_settings(date, weekday_start_hour, weekday_start_minute, weekday_end_hour, weekday_end_minute,
                       holiday_start_hour, holiday_start_minute, holiday_end_hour, holiday_end_minute):
    """날짜에 따라 평일/휴일 시간 설정을 반환합니다."""
    if date.weekday() < 5:  # 0-4: 월-금 (평일)
        start_hour = weekday_start_hour
        start_minute = weekday_start_minute
        end_hour = weekday_end_hour
        end_minute = weekday_end_minute
        logging.info(f"{date} (평일) 시간 설정")
    else:
        start_hour = holiday_start_hour
        start_minute = holiday_start_minute
        end_hour = holiday_end_hour
        end_minute = holiday_end_minute
        logging.info(f"{date} (휴일) 시간 설정")
    return start_hour, start_minute, end_hour, end_minute

# 날짜 반복 함수
def apply_overtime_for_dates(driver, start_date, end_date, work_details, weekday_start_hour, weekday_start_minute,
                              weekday_end_hour, weekday_end_minute, holiday_start_hour, holiday_start_minute,
                              holiday_end_hour, holiday_end_minute):
    """날짜 범위에 대해 초과 근무를 신청합니다."""
    current_date = start_date
    while current_date <= end_date:
        # 평일/휴일 판단 및 시간 설정
        start_hour, start_minute, end_hour, end_minute = get_time_settings(
            current_date, weekday_start_hour, weekday_start_minute, weekday_end_hour, weekday_end_minute,
            holiday_start_hour, holiday_start_minute, holiday_end_hour, holiday_end_minute
        )

        # 초과 근무 신청
        if not apply_overtime(driver, current_date, work_details, start_hour, start_minute, end_hour, end_minute):
            logging.error(f"{current_date} 초과 근무 신청 실패")
            messagebox.showerror("오류", f"{current_date} 초과 근무 신청 실패")
            return False  # 실패 시 종료

        # 다음 날짜로 이동
        current_date += datetime.timedelta(days=1)
    return True  # 성공

# 초과 근무 신청 함수 (기존 코드와 동일)
def apply_overtime(driver, date, work_details, start_hour, start_minute, end_hour, end_minute):
    try:
        # 초과근무 일자 입력 (pyautogui 사용)
        work_date_field = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'mainframe.WorkFrame.form.divWork_tat_atdc/TatAtdc01001M02.form.divForm.form.divDetail.form.divSubDetail02.form.cboWrkSeNm.comboedit:input'))
        )
        driver.execute_script("arguments[0].scrollIntoView();", work_date_field)
        work_date_field.click()
        time.sleep(0.5)
        pyautogui.press('tab')
        pyautogui.write(date.strftime("%Y-%m-%d"), interval=0.1)
        pyautogui.press('tab', presses=2)
        logging.info(f"{date} 초과근무 일자 입력 완료 (pyautogui)")

        time.sleep(1)
        pyautogui.write(start_hour, interval=0.3)  # 시작 시
        pyautogui.write(start_minute, interval=0.3)  # 시작 분
        pyautogui.write(end_hour, interval=0.3)  # 종료 시
        pyautogui.write(end_minute, interval=0.3)  # 종료 분
        logging.info("시간 입력 완료 (pyautogui)")

        # 근무 할 일 입력 (클립보드 방식)
        work_details_field = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'mainframe.WorkFrame.form.divWork_tat_atdc/TatAtdc01001M02.form.divForm.form.divDetail.form.divSubDetail02.form.edtWrkDtlsCn:input'))
        )
        driver.execute_script("arguments[0].scrollIntoView();", work_details_field)
        work_details_field.click()
        pyperclip.copy(work_details)
        work_details_field.send_keys(Keys.CONTROL + "v")
        logging.info("근무 할 일 입력 완료")

        # 저장 버튼 클릭
        save_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="mainframe.WorkFrame.form.divWork_tat_atdc/TatAtdc01001M02.form.divForm.form.divDetail.form.btnSave:icontext"]'))
        )
        save_button.click()
        logging.info("저장 클릭 완료")

        # 확인 팝업 처리
        confirm_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="mainframe.WorkFrame.msgConfirm.form.btnOk:icontext"]'))
        )
        confirm_button.click()
        logging.info("최종 확인 완료")

    except Exception as e:
        logging.error("오류 발생:", e)
        return False  # 실패

    return True  # 성공

class OvertimeAppGUI:
    def __init__(self, root):
        self.root = root
        root.title("초과 근무 신청")

        # 아이콘 설정 (ICO 만 지원)
        icon_path_ico = "icon.ico"
        if os.path.exists(icon_path_ico):
            self.set_icon_bitmap(root, "icon.ico")
        else:
            logging.warning("아이콘 파일 없음")

        # 시작 날짜
        ttk.Label(root, text="시작 날짜 (YYYYMMDD):").grid(row=0, column=0, padx=10, pady=5)
        self.start_date_entry = ttk.Entry(root)
        self.start_date_entry.grid(row=0, column=1, padx=10, pady=5)
        default_start_date = datetime.date.today() + datetime.timedelta(days=1)
        self.start_date_entry.insert(0, default_start_date.strftime("%Y%m%d"))

        # 종료 날짜
        ttk.Label(root, text="종료 날짜 (YYYYMMDD):").grid(row=1, column=0, padx=10, pady=5)
        self.end_date_entry = ttk.Entry(root)
        self.end_date_entry.grid(row=1, column=1, padx=10, pady=5)

        # 근무 내용
        ttk.Label(root, text="근무 내용:").grid(row=2, column=0, padx=10, pady=5)
        self.work_details_entry = ttk.Entry(root)
        self.work_details_entry.grid(row=2, column=1, padx=10, pady=5)

        # 시간 목록 생성
        hours = [str(i).zfill(2) for i in range(24)]
        minutes = [str(i).zfill(2) for i in range(0, 60, 10)]

        # 평일 시간
        ttk.Label(root, text="평일 시작 시간 (HH:MM):").grid(row=3, column=0, padx=10, pady=5)
        self.weekday_start_hour_combo = ttk.Combobox(root, width=3, values=hours, state="readonly")
        self.weekday_start_hour_combo.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        self.weekday_start_hour_combo.set("06")  # 기본 시작 시간 06시
        self.weekday_start_minute_combo = ttk.Combobox(root, width=3, values=minutes, state="readonly")
        self.weekday_start_minute_combo.grid(row=3, column=1, padx=50, pady=5, sticky=tk.W)
        self.weekday_start_minute_combo.set("00")  # 기본 시작 분 00분

        ttk.Label(root, text="평일 종료 시간 (HH:MM):").grid(row=4, column=0, padx=10, pady=5)
        self.weekday_end_hour_combo = ttk.Combobox(root, width=3, values=hours, state="readonly")
        self.weekday_end_hour_combo.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
        self.weekday_end_hour_combo.set("23")  # 기본 종료 시간 23시
        self.weekday_end_minute_combo = ttk.Combobox(root, width=3, values=minutes, state="readonly")
        self.weekday_end_minute_combo.grid(row=4, column=1, padx=50, pady=5, sticky=tk.W)
        self.weekday_end_minute_combo.set("00")  # 기본 종료 분 00분

        # 휴일 시간
        ttk.Label(root, text="휴일 시작 시간 (HH:MM):").grid(row=5, column=0, padx=10, pady=5)
        self.holiday_start_hour_combo = ttk.Combobox(root, width=3, values=hours, state="readonly")
        self.holiday_start_hour_combo.grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)
        self.holiday_start_hour_combo.set("06")  # 기본 시작 시간 06시
        self.holiday_start_minute_combo = ttk.Combobox(root, width=3, values=minutes, state="readonly")
        self.holiday_start_minute_combo.grid(row=5, column=1, padx=50, pady=5, sticky=tk.W)
        self.holiday_start_minute_combo.set("00")  # 기본 시작 분 00분

        ttk.Label(root, text="휴일 종료 시간 (HH:MM):").grid(row=6, column=0, padx=10, pady=5)
        self.holiday_end_hour_combo = ttk.Combobox(root, width=3, values=hours, state="readonly")
        self.holiday_end_hour_combo.grid(row=6, column=1, padx=5, pady=5, sticky=tk.W)
        self.holiday_end_hour_combo.set("21")  # 기본 종료 시간 21시
        self.holiday_end_minute_combo = ttk.Combobox(root, width=3, values=minutes, state="readonly")
        self.holiday_end_minute_combo.grid(row=6, column=1, padx=50, pady=5, sticky=tk.W)
        self.holiday_end_minute_combo.set("00")  # 기본 종료 분 00분

        # 실행 버튼
        ttk.Button(root, text="실행", command=self.run_script).grid(row=7, column=0, columnspan=2, padx=10, pady=10)

        # 상태 표시줄
        self.status_label = ttk.Label(root, text="대기 중...")
        self.status_label.grid(row=8, column=0, columnspan=2, sticky=tk.W + tk.E, padx=10, pady=5)

    def set_icon_bitmap(self, window, icon_filename):
        """아이콘 설정 함수"""
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.abspath(".")

            icon_path = os.path.join(base_path, icon_filename)
            window.iconbitmap(icon_path)
            logging.info(f"아이콘 설정 성공: {icon_path}")
        except Exception as e:
            logging.error(f"아이콘 설정 실패: {e}")

    def run_script(self):
        """실행 버튼 클릭 이벤트 핸들러"""
        self.update_status("초기화 중...")  # 상태 업데이트

        start_date_str = self.start_date_entry.get()
        end_date_str = self.end_date_entry.get()
        work_details = self.work_details_entry.get()

        weekday_start_hour = self.weekday_start_hour_combo.get()
        weekday_start_minute = self.weekday_start_minute_combo.get()
        weekday_end_hour = self.weekday_end_hour_combo.get()
        weekday_end_minute = self.weekday_end_minute_combo.get()

        holiday_start_hour = self.holiday_start_hour_combo.get()
        holiday_start_minute = self.holiday_start_minute_combo.get()
        holiday_end_hour = self.holiday_end_hour_combo.get()
        holiday_end_minute = self.holiday_end_minute_combo.get()

        try:
            start_date = datetime.datetime.strptime(start_date_str, "%Y%m%d").date()
            end_date = datetime.datetime.strptime(end_date_str, "%Y%m%d").date()
        except ValueError as e:
            logging.error(f"잘못된 날짜 형식: {e}")
            messagebox.showerror("오류", "잘못된 날짜 형식입니다. YYYYMMDD 형식으로 입력해주세요.")
            self.update_status("날짜 형식 오류")
            return

        self.update_status("엣지 드라이버 초기화 중...")

        try:
            service = Service(EdgeChromiumDriverManager().install())
            options = webdriver.EdgeOptions()
            driver = webdriver.Edge(service=service, options=options)
        except Exception as e:
            logging.error(f"엣지 드라이버 초기화 실패: {e}")
            messagebox.showerror("오류", f"엣지 드라이버 초기화 실패: {e}")
            self.update_status("드라이버 초기화 실패")
            return

        try:
            self.update_status("웹 사이트 접속 중...")
            website_url = "https://seoul.insarang.go.kr/nexin/index.html" # 설정 파일 대신 직접 URL 지정
            driver.get(website_url)
            driver.maximize_window()

            self.update_status("로그인 시도 중...")
            try:
                login_wait_time = 10 # 설정 파일 대신 직접 값 지정
                login_button_xpath = '//*[@id="mainframe.WorkFrame.form.divCenter.form.divLogin.form.btnLogin:icontext"]' # 설정 파일 대신 직접 XPath 지정
                login_btn = WebDriverWait(driver, login_wait_time).until(
                    EC.element_to_be_clickable((By.XPATH, login_button_xpath))
                )
                login_btn.click()
                logging.info("로그인 버튼 클릭 완료")
            except (NoSuchElementException, TimeoutException) as e:
                logging.error(f"로그인 버튼 클릭 실패: {e}")
                messagebox.showerror("오류", "로그인 버튼을 찾을 수 없습니다.")
                self.update_status("로그인 실패")
                return

            self.update_status("인증서 선택 대기 중... 수동으로 인증서 선택 후 확인 버튼을 클릭해주세요.")

            try:
                iframe_wait_time = 10 # 설정 파일 대신 직접 값 지정
                WebDriverWait(driver, iframe_wait_time).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, "dscert"))
                )
                logging.info("iframe(dscert) 진입 완료")

                # 수동 입력을 위한 대기
                logging.info("인증서 비밀번호를 수동으로 입력해주세요. 완료 후 확인 버튼을 눌러주세요.")
                manual_input_wait_time = 600 # 설정 파일 대신 직접 값 지정
                WebDriverWait(driver, manual_input_wait_time).until(  # 10분 동안 대기
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_confirm_iframe"]'))
                )
                logging.info("인증서 비밀번호 입력 완료 확인")

                confirm_button_wait_time = 10 # 설정 파일 대신 직접 값 지정

                # 가려지는 요소가 사라질 때까지 대기
                try:
                    overlay_wait_time = 5  # 최대 대기 시간 (초)
                    WebDriverWait(driver, overlay_wait_time).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, "blockUI blockOverlay"))
                    )
                    logging.info("가려지는 요소가 사라질 때까지 대기 완료")
                except TimeoutException:
                    logging.warning("가려지는 요소가 지정된 시간 내에 사라지지 않음")

                # 클릭 가능한 상태가 될 때까지 대기
                try:
                    ok_btn = WebDriverWait(driver, confirm_button_wait_time).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_confirm_iframe"]'))
                    )
                    ok_btn.click()
                    logging.info("확인 버튼 클릭 완료")
                except (NoSuchElementException, TimeoutException, ElementClickInterceptedException) as e:
                    logging.error(f"확인 버튼 클릭 실패: {e}")
                    messagebox.showerror("오류", "확인 버튼을 클릭할 수 없습니다. 다시 시도해주세요.")
                    self.update_status("확인 버튼 클릭 실패")
                    return

                driver.switch_to.default_content()
            except (NoSuchElementException, TimeoutException) as e:
                logging.error(f"인증서 처리 실패: {e}")
                messagebox.showerror("오류", "인증서 처리에 실패했습니다. 다시 시도해주세요.")
                self.update_status("인증서 처리 실패")
                return

            self.update_status("메뉴 클릭 중...")

            try:
                menu_wait_time = 10 # 설정 파일 대신 직접 값 지정
                WebDriverWait(driver, menu_wait_time).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="mainframe.WorkFrame.form.divTop.form.divTopMenu.form.btnMenuMN00002542:icontext"]'))
                ).click()
                logging.info("출장관리 클릭")

                WebDriverWait(driver, menu_wait_time).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="mainframe.WorkFrame.form.divLeft.form.divMenu.form.btnMenuMN00002551:icontext"]'))
                ).click()
                logging.info("초과근무 클릭")

                WebDriverWait(driver, menu_wait_time).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="mainframe.WorkFrame.form.divLeft.form.divMenu.form.divMenuMN00002551.form.btnMenuMN00002600:icontext"]'))
                ).click()
                logging.info("초과근무 내역 클릭")
            except (NoSuchElementException, TimeoutException) as e:
                logging.error(f"메뉴 클릭 실패: {e}")
                messagebox.showerror("오류", "메뉴 클릭에 실패했습니다.")
                self.update_status("메뉴 클릭 실패")
                return

            self.update_status("사전 신청 클릭 중...")

            try:
                prereg_wait_time = 20 # 설정 파일 대신 직접 값 지정
                pre_reg_btn = WebDriverWait(driver, prereg_wait_time).until(
                    EC.element_to_be_clickable((By.ID, "mainframe.WorkFrame.form.divWork_tat_MN00002600.form.divForm.form.divDetail.form.btnPreReg:icontext"))
                )
                pre_reg_btn.click()
                logging.info("사전신청 클릭")
            except (NoSuchElementException, TimeoutException) as e:
                logging.error(f"사전 신청 클릭 실패: {e}")
                messagebox.showerror("오류", "사전 신청 클릭에 실패했습니다.")
                self.update_status("사전 신청 클릭 실패")
                return

            self.update_status("날짜별 초과 근무 신청 중...")
            # 날짜 범위에 대해 초과 근무 신청
            if not apply_overtime_for_dates(driver, start_date, end_date, work_details, weekday_start_hour,
                                           weekday_start_minute, weekday_end_hour, weekday_end_minute,
                                           holiday_start_hour, holiday_start_minute, holiday_end_hour,
                                           holiday_end_minute):
                self.update_status("초과 근무 신청 실패 (날짜 중 하나)")
                return  # 실패 시 종료

            messagebox.showinfo("완료", "초과 근무 신청이 완료되었습니다.")
            self.update_status("완료")

        except Exception as e:
            logging.exception("예상치 못한 오류 발생")
            messagebox.showerror("오류", f"예상치 못한 오류 발생: {e}")
            self.update_status("예상치 못한 오류")

        finally:
            self.update_status("드라이버 종료 중...")
            try:
                driver.quit()
            except Exception as e:
                logging.error(f"드라이버 종료 중 오류 발생: {e}")
            self.update_status("완료")

    def update_status(self, message):
        """상태 표시줄 업데이트"""
        self.status_label.config(text=message)
        self.root.update_idletasks()  # GUI 업데이트

    def set_icon_bitmap(self, window, icon_filename):
        """아이콘 설정 함수"""
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.abspath(".")

            icon_path = os.path.join(base_path, icon_filename)
            window.iconbitmap(icon_path)
            logging.info(f"아이콘 설정 성공: {icon_path}")
        except Exception as e:
            logging.error(f"아이콘 설정 실패: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OvertimeAppGUI(root)
    root.mainloop()
