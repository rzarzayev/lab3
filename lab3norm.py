from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import os
from dotenv import load_dotenv

class AttendanceScraper:
    def __init__(self):
        load_dotenv()
        self.driver = webdriver.Edge()
        self.username = os.getenv("NAME")
        self.password = os.getenv("PASSWORD")

        if self.username is None or self.password is None:
            raise Exception(".env faylından istifadəçi adı və ya şifrə tapılmadı.")

    def login(self):
        url = "https://sso.aztu.edu.az/"
        self.driver.get(url)
        print("KOICA giriş səhifəsinə keçilir...")

        try:
            username_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/section/div/div[1]/div/div/form/div[1]/input"))
            )
            username_field.send_keys(self.username)

            password_field = self.driver.find_element(By.XPATH, "/html/body/section/div/div[1]/div/div/form/div[2]/input")
            password_field.send_keys(self.password)

            login_button = self.driver.find_element(By.XPATH, "/html/body/section/div/div[1]/div/div/form/div[3]/button")
            login_button.click()
            print("Giriş məlumatları daxil edildi.")
        except Exception as e:
            print(f"Məlumatların daxil edilməsi zamanı səhv: {e}")
        finally:
            time.sleep(5)

    def open_menu(self):
        try:
            menu_icon = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "i.fas.fa-bars"))
            )
            menu_icon.click()
        except Exception as e:
            print(f"Menyu açılarkən xəta baş verdi: {e}")

    def navigate_to_student_section(self):
        try:
            student_link = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Tələbə keçid"))
            )
            student_link.click()
        except Exception as e:
            print(f"Tələbə keçidinə yönləndirilərkən xəta baş verdi: {e}")

    def open_subjects(self):
        try:
            subjects_menu = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Fənlər')]//parent::a"))
            )
            subjects_menu.click()
        except Exception as e:
            print(f"Fənlər bölməsi tapılmadı: {e}")

    def select_python_course(self):
        try:
            python_course = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Python proqramlaşdırma dili"))
            )
            python_course.click()
        except Exception as e:
            print(f"Python kursu seçilərkən xəta baş verdi: {e}")

    def scrape_attendance(self):
        try:
            time.sleep(3)
            attendance_button = self.driver.find_element(By.LINK_TEXT, "Davamiyyət")
            attendance_button.click()

            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//th[contains(@class, 'list_text6')]/font"))
            )

            # Tarixləri və davamiyyət statuslarını almaq
            date_elements = self.driver.find_elements(By.XPATH, "//th[contains(@class, 'list_text6')]/font")
            dates = [date.text for date in date_elements if date.text.strip() != '']

            attendance_elements = self.driver.find_elements(By.XPATH, "//tr[contains(@class, 'attend-tr')]//span")
            attendance_status = [
                att.text.strip() for att in attendance_elements if att.text.strip() in ["i/e", "q/b"]
            ]

            # Məlumatları birləşdirmək və Excel faylına yazmaq
            attendance_data = list(zip(dates, attendance_status))
            attendance_df = pd.DataFrame(attendance_data, columns=["Tarix", "Davamiyyət"])
            attendance_df.to_excel('attendance_data.xlsx', index=False)

            print("Məlumatlar uğurla 'attendance_data.xlsx' adlı Excel faylına yazıldı.")
        except Exception as e:
            print(f"Davamiyyət məlumatları alınarkən xəta baş verdi: {e}")

    def close(self):
        self.driver.quit()

    def run(self):
        try:
            self.login()
            self.open_menu()
            self.navigate_to_student_section()
            self.open_subjects()
            self.select_python_course()
            self.scrape_attendance()
        finally:
            self.close()

if __name__ == "__main__":
    scraper = AttendanceScraper()
    scraper.run()