import os
import shutil
import time
import xml.etree.ElementTree as ET
import tkinter as tk
from threading import Thread, Event
import pdfplumber
import re

# Define a pasta de downloads
DOWNLOADS_PATH = "C:/Users/FISCAL01/Downloads"

# Extensões de arquivos a considerar
VALID_EXTENSIONS = {'.pdf', '.txt', '.xml', '.xls', '.xlsx'}

# Variável global para controlar o número de arquivos esperados
files_to_wait = 4

def extract_company_name(file_path, is_single_file=False):
    try:
        if is_single_file and file_path.lower().endswith('.pdf'):
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    texto_pagina = page.extract_text()
                    if texto_pagina and "Contribuinte" in texto_pagina:
                        match = re.search(r'\bContribuinte\s', texto_pagina, re.IGNORECASE)
                        if match:
                            start_idx = match.end()
                            end_idx = texto_pagina.find('\n', start_idx) if texto_pagina.find('\n', start_idx) != -1 else len(texto_pagina)
                            company_name = texto_pagina[start_idx:end_idx].strip()
                            invalid_chars = '<>:"/\\|?*'
                            for char in invalid_chars:
                                company_name = company_name.replace(char, '')
                            return company_name if company_name else "Sem Empresa"
            return "Sem Empresa"
        elif file_path.lower().endswith('.xml'):
            tree = ET.parse(file_path)
            root = tree.getroot()
            for elem in root.iter('NomeFantasia'):
                company_name = elem.text.strip()
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    company_name = company_name.replace(char, '')
                return company_name if company_name else "Sem Empresa"
            return "Sem Empresa"
        return "Sem Empresa"
    except Exception as e:
        return "Sem Empresa"

def organize_files(stop_event, status_label, log_text):
    global files_to_wait
    processed_files = []
    
    while not stop_event.is_set():
        files = os.listdir(DOWNLOADS_PATH)
        valid_files = [
            f for f in files 
            if os.path.join(DOWNLOADS_PATH, f) not in processed_files 
            and not os.path.isdir(os.path.join(DOWNLOADS_PATH, f))
            and any(f.lower().endswith(ext) for ext in VALID_EXTENSIONS)
        ]
        
        if len(valid_files) == files_to_wait:
            company_name = None
            is_single_file = (files_to_wait == 1)
            for file_name in valid_files:
                file_path = os.path.join(DOWNLOADS_PATH, file_name)
                company_name = extract_company_name(file_path, is_single_file)
                break
            
            if not company_name:
                company_name = "Sem Empresa"
            
            folder_path = os.path.join(DOWNLOADS_PATH, company_name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            
            for file_name in valid_files:
                file_path = os.path.join(DOWNLOADS_PATH, file_name)
                new_file_path = os.path.join(folder_path, file_name)
                try:
                    shutil.move(file_path, new_file_path)
                    processed_files.append(file_path)
                    log_text.insert(tk.END, f"Arquivo '{file_name}' movido para '{company_name}'\n")
                    log_text.see(tk.END)
                except Exception as e:
                    log_text.insert(tk.END, f"Erro ao mover '{file_name}': {e}\n")
                    log_text.see(tk.END)
        
        status_label.config(text=f"Monitorando... ({len(valid_files)}/{files_to_wait} arquivos detectados)")
        time.sleep(5)
    
    status_label.config(text="Parado")
    log_text.insert(tk.END, "Monitoramento encerrado.\n")
    log_text.see(tk.END)

def start_monitoring(stop_event, thread_ref, status_label, log_text):
    global files_to_wait
    if not thread_ref[0] or not thread_ref[0].is_alive():
        stop_event.clear()
        thread_ref[0] = Thread(target=organize_files, args=(stop_event, status_label, log_text))
        thread_ref[0].start()
        status_label.config(text=f"Monitorando... (0/{files_to_wait} arquivos detectados)")
        log_text.insert(tk.END, "Monitoramento iniciado.\n")
        log_text.see(tk.END)

def stop_monitoring(stop_event, status_label):
    stop_event.set()
    status_label.config(text="Parado")

def clear_log(log_text):
    log_text.delete(1.0, tk.END)
    log_text.insert(tk.END, "Log limpo.\n")
    log_text.see(tk.END)

def set_less_than_100(status_label, log_text):
    global files_to_wait
    files_to_wait = 4
    log_text.insert(tk.END, "Configurado para 'Com Movimento menor que 100' (4 arquivos).\n")
    log_text.see(tk.END)
    status_label.config(text=f"Monitorando... (0/{files_to_wait} arquivos detectados)" if thread_ref[0] and thread_ref[0].is_alive() else "Parado")

def set_more_than_100(status_label, log_text):
    global files_to_wait
    files_to_wait = 5
    log_text.insert(tk.END, "Configurado para 'Com Movimento mais que 100' (5 arquivos).\n")
    log_text.see(tk.END)
    status_label.config(text=f"Monitorando... (0/{files_to_wait} arquivos detectados)" if thread_ref[0] and thread_ref[0].is_alive() else "Parado")

def set_no_movement(status_label, log_text):
    global files_to_wait
    files_to_wait = 1
    log_text.insert(tk.END, "Configurado para 'Sem Movimento' (1 arquivo PDF).\n")
    log_text.see(tk.END)
    status_label.config(text=f"Monitorando... (0/{files_to_wait} arquivos detectados)" if thread_ref[0] and thread_ref[0].is_alive() else "Parado")

def main():
    global thread_ref
    # Configura a janela
    root = tk.Tk()
    root.title("Monitor de Pasta")
    root.geometry("550x450")  # Aumentei um pouco para caber os novos botões
    
    # Label para status
    status_label = tk.Label(root, text="Parado", font=("Arial", 12, "bold"))
    status_label.pack(pady=10)
    
    # Área de log
    log_text = tk.Text(root, height=10, width=50, font=("Arial", 10))
    log_text.pack(pady=10)
    
    # Botões
    stop_event = Event()
    thread_ref = [None]
    
    start_button = tk.Button(root, text="Iniciar Monitoramento", 
                           command=lambda: start_monitoring(stop_event, thread_ref, status_label, log_text))
    start_button.pack(pady=5)
    
    stop_button = tk.Button(root, text="Parar Monitoramento", 
                          command=lambda: stop_monitoring(stop_event, status_label))
    stop_button.pack(pady=5)
    
    clear_button = tk.Button(root, text="Limpar Log", 
                           command=lambda: clear_log(log_text))
    clear_button.pack(pady=5)
    
    less_than_100_button = tk.Button(root, text="Com Movimento menor que 100", 
                                   command=lambda: set_less_than_100(status_label, log_text))
    less_than_100_button.pack(pady=5)
    
    more_than_100_button = tk.Button(root, text="Com Movimento mais que 100", 
                                   command=lambda: set_more_than_100(status_label, log_text))
    more_than_100_button.pack(pady=5)
    
    no_movement_button = tk.Button(root, text="Sem Movimento", 
                                 command=lambda: set_no_movement(status_label, log_text))
    no_movement_button.pack(pady=5)
    
    # Inicia a interface
    root.mainloop()

if __name__ == "__main__":
    main()