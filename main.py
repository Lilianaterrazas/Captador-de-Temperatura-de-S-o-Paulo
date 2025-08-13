import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from datetime import datetime
import os

CLIMA_SITE = "https://www.accuweather.com/pt/br/s%C3%A3o-paulo/45881/current-weather/45881"
ARQUIVO_EXCEL = "temperatura_sao_paulo.xlsx"

def capturar_data_clima():
    driver = None
    try:
        chrome_options = Options()

        driver = webdriver.Chrome(options=chrome_options)
        driver.get(CLIMA_SITE)

        temperature_selector = (By.CSS_SELECTOR, 'div.temp')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(temperature_selector))
        temperature_element = driver.find_element(*temperature_selector)
        temperature = temperature_element.text.strip()

        humidity_selector = (By.XPATH, "//div[contains(text(), 'Umidade')]/following-sibling::div")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(humidity_selector))
        humidity_element = driver.find_element(*humidity_selector)
        humidity = humidity_element.text.strip()

        if '%' not in humidity:
            humidity += '%'

        current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not os.path.exists(ARQUIVO_EXCEL):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Temperatura São Paulo"
            sheet.append(["Data/Hora", "Temperatura", "Umidade do Ar"])
        else:
            workbook = openpyxl.load_workbook(ARQUIVO_EXCEL)
            sheet = workbook.active

        sheet.append([current_datetime, temperature, humidity])
        workbook.save(ARQUIVO_EXCEL)
        messagebox.showinfo("Sucesso", f"Dados capturados e salvos em {ARQUIVO_EXCEL}\nTemperatura: {temperature}, Umidade: {humidity}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao capturar os dados: {e}\nPor favor, verifique os seletores de elementos do AccuWeather ou se o Chrome está instalado e o chromedriver.exe é compatível.")
    finally:
        if driver:
            driver.quit()

def criar_interface():
    root = tk.Tk()
    root.title("Captador de Temperatura de São Paulo - AccuWeather")
    root.geometry("400x200")

    label = tk.Label(root, text="Previsão do tempo de São Paulo (AccuWeather)", font=("Arial", 12))
    label.pack(pady=20)
 
    capture_button = tk.Button(root, text="Buscar previsão", command=capturar_data_clima, font=("Arial", 12))
    capture_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()