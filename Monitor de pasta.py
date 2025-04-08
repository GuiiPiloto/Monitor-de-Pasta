import os
import shutil
import time
import xml.etree.ElementTree as ET
import tkinter as tk
from threading import Thread, Event
from tkinter import messagebox, filedialog
import pdfplumber
import re
import zipfile
import sys

# Detecta a pasta de downloads do usuário dinamicamente ou permite seleção manual
def get_downloads_path(root):
    try:
        # Tenta o caminho padrão primeiro
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        if os.path.exists(downloads_path):
            return downloads_path
        
        # Se não encontrar, pede pro usuário selecionar manualmente
        messagebox.showwarning("Pasta não encontrada", "Não foi possível encontrar a pasta de downloads automaticamente.\nPor favor, selecione a pasta de downloads manualmente.")
        selected_path = filedialog.askdirectory(
            parent=root,
            title="Selecione a pasta de downloads",
            mustexist=True
        )
        
        if selected_path and os.path.exists(selected_path):
            return selected_path
        else:
            raise FileNotFoundError("Nenhuma pasta de downloads foi selecionada ou a pasta não existe.")
    except Exception as e:
        raise FileNotFoundError(f"Erro ao determinar a pasta de downloads: {e}")

# Função original pra extrair o nome da empresa (usada pros modos antigos)
def extract_company_name(file_path, log_text):
    try:
        if file_path.lower().endswith('.xml'):
            try:
                tree = ET.parse(file_path)
                root = tree.getroot()
                # Pega o primeiro <xNome> que encontrar, independentemente de ser <emit> ou <dest>
                xnome_elem = root.find('.//xNome')
                if xnome_elem is not None:
                    company_name = xnome_elem.text.strip()
                else:
                    log_text.insert(tk.END, "Tag <xNome> não encontrada no XML.\n")
                    log_text.see(tk.END)
                    return "Sem Empresa"
            except ET.ParseError as e:
                log_text.insert(tk.END, f"Erro ao parsear XML: {e}\n")
                log_text.see(tk.END)
                return "Sem Empresa"
            
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                company_name = company_name.replace(char, '')
            return company_name if company_name else "Sem Empresa"
        elif file_path.lower().endswith('.pdf'):
            try:
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
            except Exception as e:
                log_text.insert(tk.END, f"Erro ao ler PDF: {e}\n")
                log_text.see(tk.END)
                return "Sem Empresa"
        return "Sem Empresa"
    except Exception as e:
        log_text.insert(tk.END, f"Erro ao ler arquivo: {e}\n")
        log_text.see(tk.END)
        return "Sem Empresa"

# Função pra extrair o nome da empresa no modo ZIP (só do <emit><xNome>)
def extract_company_name_for_zip(file_path, log_text):
    try:
        if file_path.lower().endswith('.xml'):
            try:
                namespaces = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
                tree = ET.parse(file_path)
                root = tree.getroot()
                
                # Tenta encontrar <emit> com namespace
                emit_elem = root.find('.//nfe:emit', namespaces)
                if emit_elem is not None:
                    xnome_elem = emit_elem.find('nfe:xNome', namespaces)
                    if xnome_elem is not None:
                        company_name = xnome_elem.text.strip()
                        log_text.insert(tk.END, f"Nome encontrado com namespace (modo ZIP): {company_name}\n")
                        log_text.see(tk.END)
                    else:
                        log_text.insert(tk.END, "Tag <xNome> não encontrada dentro de <emit> com namespace (modo ZIP).\n")
                        log_text.see(tk.END)
                else:
                    log_text.insert(tk.END, "Tag <emit> não encontrada com namespace, tentando sem namespace (modo ZIP)...\n")
                    log_text.see(tk.END)
                    # Tenta sem namespace
                    emit_elem = root.find('.//emit')
                    if emit_elem is not None:
                        xnome_elem = emit_elem.find('xNome')
                        if xnome_elem is not None:
                            company_name = xnome_elem.text.strip()
                            log_text.insert(tk.END, f"Nome encontrado sem namespace (modo ZIP): {company_name}\n")
                            log_text.see(tk.END)
                        else:
                            log_text.insert(tk.END, "Tag <xNome> não encontrada dentro de <emit> sem namespace (modo ZIP).\n")
                            log_text.see(tk.END)
                            return "Sem Empresa"
                    else:
                        log_text.insert(tk.END, "Tag <emit> não encontrada sem namespace (modo ZIP).\n")
                        log_text.see(tk.END)
                        return "Sem Empresa"
                
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    company_name = company_name.replace(char, '')
                return company_name if company_name else "Sem Empresa"
            except ET.ParseError as e:
                log_text.insert(tk.END, f"Erro ao parsear XML (modo ZIP): {e}\n")
                log_text.see(tk.END)
                return "Sem Empresa"
        return "Sem Empresa"
    except Exception as e:
        log_text.insert(tk.END, f"Erro ao ler XML (modo ZIP): {e}\n")
        log_text.see(tk.END)
        return "Sem Empresa"

def extract_zip(zip_path, extract_to):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        return True
    except Exception as e:
        return False

# Extensões de arquivos a considerar
VALID_EXTENSIONS = {'.pdf', '.txt', '.xml', '.xls', '.xlsx', '.zip'}

# Variáveis globais
files_to_wait = 4
zip_mode = False
zip_company_name = None
DOWNLOADS_PATH = None

def organize_files(stop_event, status_label, log_text):
    global files_to_wait, zip_mode, zip_company_name
    if DOWNLOADS_PATH is None:
        log_text.insert(tk.END, "Erro: Não foi possível determinar a pasta de downloads. Encerrando.\n")
        log_text.see(tk.END)
        status_label.config(text="Erro: Pasta de downloads não encontrada")
        return
    
    processed_files = []
    
    while not stop_event.is_set():
        try:
            files = os.listdir(DOWNLOADS_PATH)
            valid_files = [
                f for f in files 
                if os.path.join(DOWNLOADS_PATH, f) not in processed_files 
                and not os.path.isdir(os.path.join(DOWNLOADS_PATH, f))
                and any(f.lower().endswith(ext) for ext in VALID_EXTENSIONS)
            ]
            
            if zip_mode and valid_files:
                # Procura o primeiro arquivo que seja um ZIP
                zip_file = None
                for file_name in valid_files:
                    file_path = os.path.join(DOWNLOADS_PATH, file_name)
                    if file_path.lower().endswith('.zip'):
                        zip_file = (file_name, file_path)
                        break
                
                if zip_file is None:
                    # Se não encontrou um ZIP, ignora os outros arquivos e continua monitorando
                    log_text.insert(tk.END, f"Aguardando ZIP... ({len(valid_files)} arquivos ignorados)\n")
                    log_text.see(tk.END)
                    time.sleep(5)
                    continue
                
                file_name, file_path = zip_file
                temp_dir = os.path.join(DOWNLOADS_PATH, "temp_extract")
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir, ignore_errors=True)
                os.makedirs(temp_dir, exist_ok=True)
                
                if extract_zip(file_path, temp_dir):
                    xml_count = 0
                    if zip_company_name is None:  # Primeiro ZIP (Saída)
                        for extracted_file in os.listdir(temp_dir):
                            extracted_path = os.path.join(temp_dir, extracted_file)
                            if extracted_path.lower().endswith('.xml'):
                                zip_company_name = extract_company_name_for_zip(extracted_path, log_text)
                                break
                        subfolder = "SAÍDA"
                        status_label.config(text="Aguardando ZIP de Entrada")
                    else:  # Segundo ZIP (Entrada)
                        subfolder = "ENTRADA"
                        status_label.config(text="Aguardando ZIP de Saída")
                    
                    folder_path = os.path.join(DOWNLOADS_PATH, zip_company_name, subfolder)
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path, exist_ok=True)
                    
                    for extracted_file in os.listdir(temp_dir):
                        extracted_path = os.path.join(temp_dir, extracted_file)
                        new_file_path = os.path.join(folder_path, extracted_file)
                        try:
                            shutil.move(extracted_path, new_file_path)
                            if extracted_file.lower().endswith('.xml'):
                                xml_count += 1
                        except Exception as e:
                            log_text.insert(tk.END, f"Erro ao mover arquivo extraído '{extracted_file}': {e}\n")
                            log_text.see(tk.END)
                    
                    log_text.insert(tk.END, f"Processado ZIP de {subfolder}: {zip_company_name} - {xml_count} XMLs movidos\n")
                    log_text.see(tk.END)
                    shutil.rmtree(temp_dir, ignore_errors=True)
                    os.remove(file_path)
                    processed_files.append(file_path)
                    if subfolder == "ENTRADA":
                        zip_company_name = None  # Reseta pra próximo par
                else:
                    log_text.insert(tk.END, f"Erro ao descompactar '{file_name}'\n")
                    log_text.see(tk.END)
                    processed_files.append(file_path)
            
            elif not zip_mode and len(valid_files) == files_to_wait:
                company_name = None
                for file_name in valid_files:
                    file_path = os.path.join(DOWNLOADS_PATH, file_name)
                    company_name = extract_company_name(file_path, log_text)
                    break
                
                if not company_name:
                    company_name = "Sem Empresa"
                
                folder_path = os.path.join(DOWNLOADS_PATH, company_name)
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path, exist_ok=True)
                
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
            
            status_label.config(text=f"Monitorando... ({len(valid_files)}/{files_to_wait} arquivos detectados)" if not zip_mode else status_label.cget("text"))
            time.sleep(5)
        except Exception as e:
            log_text.insert(tk.END, f"Erro no loop de monitoramento: {e}\n")
            log_text.see(tk.END)
            time.sleep(5)
    
    status_label.config(text="Parado")
    log_text.insert(tk.END, "Monitoramento encerrado.\n")
    log_text.see(tk.END)

def start_monitoring(stop_event, thread_ref, status_label, log_text):
    global files_to_wait, zip_mode
    if not thread_ref[0] or not thread_ref[0].is_alive():
        stop_event.clear()
        thread_ref[0] = Thread(target=organize_files, args=(stop_event, status_label, log_text))
        thread_ref[0].start()
        status_label.config(text=f"Monitorando... (0/{files_to_wait} arquivos detectados)" if not zip_mode else "Aguardando ZIP de Saída")
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
    global files_to_wait, zip_mode, zip_company_name
    if zip_mode and zip_company_name is not None:
        if not messagebox.askyesno("Confirmação", "Tem Certeza? O processamento de ZIPs não está completo"):
            return
    files_to_wait = 4
    zip_mode = False
    zip_company_name = None
    log_text.insert(tk.END, "Configurado para 'Com Movimento menor que 100' (4 arquivos).\n")
    log_text.see(tk.END)
    status_label.config(text=f"Monitorando... (0/{files_to_wait} arquivos detectados)" if thread_ref[0] and thread_ref[0].is_alive() else "Parado")

def set_more_than_100(status_label, log_text):
    global files_to_wait, zip_mode, zip_company_name
    if zip_mode and zip_company_name is not None:
        if not messagebox.askyesno("Confirmação", "Tem Certeza? O processamento de ZIPs não está completo"):
            return
    files_to_wait = 5
    zip_mode = False
    zip_company_name = None
    log_text.insert(tk.END, "Configurado para 'Com Movimento maior que 100' (5 arquivos).\n")
    log_text.see(tk.END)
    status_label.config(text=f"Monitorando... (0/{files_to_wait} arquivos detectados)" if thread_ref[0] and thread_ref[0].is_alive() else "Parado")

def set_no_movement(status_label, log_text):
    global files_to_wait, zip_mode, zip_company_name
    if zip_mode and zip_company_name is not None:
        if not messagebox.askyesno("Confirmação", "Tem Certeza? O processamento de ZIPs não está completo"):
            return
    files_to_wait = 1
    zip_mode = False
    zip_company_name = None
    log_text.insert(tk.END, "Configurado para 'Sem Movimento' (1 arquivo PDF).\n")
    log_text.see(tk.END)
    status_label.config(text=f"Monitorando... (0/{files_to_wait} arquivos detectados)" if thread_ref[0] and thread_ref[0].is_alive() else "Parado")

def set_zip_mode(status_label, log_text):
    global zip_mode, zip_company_name
    zip_mode = True
    zip_company_name = None
    log_text.insert(tk.END, "Modo ZIP ativado.\n")
    log_text.see(tk.END)
    status_label.config(text="Aguardando ZIP de Saída" if thread_ref[0] and thread_ref[0].is_alive() else "Parado")

def reset_zip_mode(status_label, log_text):
    global zip_company_name
    zip_company_name = None
    log_text.insert(tk.END, "Modo ZIP resetado.\n")
    log_text.see(tk.END)
    status_label.config(text="Aguardando ZIP de Saída" if thread_ref[0] and thread_ref[0].is_alive() else "Parado")

def main():
    global thread_ref, DOWNLOADS_PATH
    root = tk.Tk()
    root.title("📂 Monitor de Pasta - JETTAX")
    root.geometry("700x680")

    frame_top = tk.Frame(root)
    frame_top.pack(pady=10)

    status_label = tk.Label(frame_top, text="⏹️ Parado", font=("Segoe UI", 13, "bold"))
    status_label.pack()

    frame_log = tk.Frame(root)
    frame_log.pack(pady=10)

    log_text = tk.Text(frame_log, height=18, width=80, font=("Consolas", 10), wrap=tk.WORD, relief=tk.SOLID, bd=1)
    log_text.pack(padx=10, pady=5)

    # Tenta determinar a pasta de downloads
    try:
        DOWNLOADS_PATH = get_downloads_path(root)
        log_text.insert(tk.END, f"📥 Pasta de downloads detectada: {DOWNLOADS_PATH}\n")
        log_text.insert(tk.END, "✅ Aplicativo iniciado com sucesso.\n")
        log_text.see(tk.END)
    except FileNotFoundError as e:
        log_text.insert(tk.END, f"❌ Erro: {e}\n")
        log_text.see(tk.END)
        status_label.config(text="🚫 Erro: Pasta de downloads não encontrada")
        root.update()
        root.destroy()
        return

    stop_event = Event()
    thread_ref = [None]

    frame_buttons = tk.Frame(root)
    frame_buttons.pack(pady=10)

    button_pad = {'padx': 5, 'pady': 3, 'fill': 'x'}

    tk.Button(frame_buttons, text="▶️ Iniciar Monitoramento", command=lambda: start_monitoring(stop_event, thread_ref, status_label, log_text)).pack(**button_pad)
    tk.Button(frame_buttons, text="⏹️ Parar Monitoramento", command=lambda: stop_monitoring(stop_event, status_label)).pack(**button_pad)
    tk.Button(frame_buttons, text="🧹 Limpar Log", command=lambda: clear_log(log_text)).pack(**button_pad)
    tk.Button(frame_buttons, text="📉 Movimento < 100", command=lambda: set_less_than_100(status_label, log_text)).pack(**button_pad)
    tk.Button(frame_buttons, text="📈 Movimento > 100", command=lambda: set_more_than_100(status_label, log_text)).pack(**button_pad)
    tk.Button(frame_buttons, text="📄 Sem Movimento", command=lambda: set_no_movement(status_label, log_text)).pack(**button_pad)
    tk.Button(frame_buttons, text="🗜️ Processar ZIPs", command=lambda: set_zip_mode(status_label, log_text)).pack(**button_pad)
    tk.Button(frame_buttons, text="🔁 Resetar ZIP", command=lambda: reset_zip_mode(status_label, log_text)).pack(**button_pad)

    root.mainloop()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        os.chdir(sys._MEIPASS)
    main()