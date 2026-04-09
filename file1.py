from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox
import os
from datetime import datetime

def finish():
    root.destroy()

def on_date_key(event):
    entry = event.widget
    if event.keysym in ("BackSpace", "Tab", "Left", "Right"): return
    if not event.char.isdigit(): return "break"
    text = entry.get()
    if len(text) >= 10: return "break"
    if len(text) == 2 or len(text) == 5: entry.insert(END, ".")

def setup_placeholder(entry, placeholder):
    entry.insert(0, placeholder)
    entry.config(foreground="grey")
    def on_focus_in(event):
        if entry.get() == placeholder:
            entry.delete(0, END)
            entry.config(foreground="black")
    def on_focus_out(event):
        if entry.get() == "":
            entry.insert(0, placeholder)
            entry.config(foreground="grey")
    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)

# Добавлена метка ошибки для файлов
def create_file_selector(parent, label_text, placeholder, is_folder=False):
    container = ttk.Frame(parent, padding=[5, 2])
    container.pack(anchor=NW, padx=10)
    lbl = ttk.Label(container, text=label_text, font=("Times New Roman", 11, "bold"), foreground="blue")
    lbl.pack(anchor=NW, padx=10)
    
    border_frame = Frame(container, highlightbackground="grey", highlightthickness=1, bg="white")
    border_frame.pack(fill='x', padx=10, pady=5)
    
    entry = ttk.Entry(border_frame, width=50)
    entry.pack(side=LEFT, padx=(5, 5), pady=5)
    setup_placeholder(entry, placeholder)

    btn_select = ttk.Button(border_frame, text="Обзор...", width=10, 
                            command=lambda: select_item(entry, is_folder, placeholder))
    btn_select.pack(side=LEFT, padx=2)
    def open_path():
        path = entry.get()
        if os.path.exists(path) and path != placeholder:
            os.startfile(path)
        else:
            messagebox.showwarning("Внимание", "Путь не выбран или не существует")

    btn_open = ttk.Button(border_frame, text="👁", width=3, command=open_path)
    btn_open.pack(side=LEFT, padx=2)
    
    # Метка для ошибок конкретно этого файла
    err_lbl = ttk.Label(container, text="", font=("Arial", 8), foreground="red")
    err_lbl.pack(anchor=NW, padx=15)
    
    return entry, err_lbl

def select_item(entry, is_folder, placeholder):
    path = filedialog.askdirectory() if is_folder else filedialog.askopenfilename(filetypes=(("Word files", "*.docx *.doc"), ("all files", "*.*")))
    if path:
        entry.delete(0, END)
        entry.insert(0, os.path.normpath(path))
        entry.config(foreground="black")

def create_date_selector(parent):
    container = ttk.Frame(parent, padding=[5, 2])
    container.pack(fill='x', padx=10)
    ttk.Label(container, text="Период практики (дд.мм.гггг):", font=("Times New Roman", 11, "bold"), foreground="blue").pack(anchor=NW, padx=10)
    df = ttk.Frame(container); df.pack(anchor=NW, padx=10, pady=5)
    
    s_date = ttk.Entry(df, width=15); s_date.pack(side=LEFT)
    setup_placeholder(s_date, "дд.мм.гггг"); s_date.bind("<KeyPress>", on_date_key)
    ttk.Label(df, text=" — ").pack(side=LEFT)
    
    e_date = ttk.Entry(df, width=15); e_date.pack(side=LEFT)
    setup_placeholder(e_date, "дд.мм.гггг"); e_date.bind("<KeyPress>", on_date_key)
    
    error_label = ttk.Label(container, text="", font=("Arial", 8), foreground="red")
    error_label.pack(anchor=NW, padx=10)
    return s_date, e_date, error_label

def create_type_selector(parent):
    container = ttk.Frame(parent, padding=[5, 2])
    container.pack(fill='x', padx=10)
    lbl = ttk.Label(container, text="Форма отчета для студентов:", font=("Times New Roman", 11, "bold"), foreground="blue")
    lbl.pack(anchor=NW, padx=10)
    
    vals = ["Бакалавриат", "Магистратура"]
    combobox = ttk.Combobox(container, values=vals, state="readonly", width=30)
    combobox.set(vals[0])
    combobox.pack(anchor=NW, padx=10, pady=5)
    
    vals_mag = ["Учебная практика — НИР", "Производственная практика — ТП", "Производственная практика — Педагогическая", "Производственная практика — НИР", "Производственная практика — Преддипломная"]
    vals_bak = ["Производственная практика — Технологическая", "Производственная практика — Преддипломная"]
    
    combobox_mag_bak = ttk.Combobox(container, state="readonly",foreground="grey", width=100, font=("Times New Roman", 10, "italic"))
    combobox_mag_bak.set("---Выберите форму отчета---")
    combobox_mag_bak.pack(anchor=NW, padx=10, pady=5)
    
    err_lbl = ttk.Label(container, text="", font=("Arial", 8), foreground="red")
    err_lbl.pack(anchor=NW, padx=10)

    def on_selected(event):
        err_lbl.config(text="") # Очистка при выборе
        selected = combobox.get()
        combobox_mag_bak.set("---Выберите форму отчета---")
        if selected == "Магистратура":
            combobox_mag_bak.configure(values=vals_mag)
        else:
            combobox_mag_bak.configure(values=vals_bak)
            
    combobox.bind("<<ComboboxSelected>>", on_selected)
    return combobox, combobox_mag_bak, err_lbl

def start_gen():
    # 0. Очистка всех ошибок
    for lbl in [err_ved, err_pr, err_fld, date_error_lbl, err_type]:
        lbl.config(text="")
    
    # 1. Проверка файлов
    file_checks = [
        (entry_ved, err_ved, "Укажите файл ведомости..."),
        (entry_pr, err_pr, "Укажите файл приказа..."),
        (entry_folder, err_fld, "Выберите папку...")
    ]
    
    for entry, label, placeholder in file_checks:
        path = entry.get()
        if path == placeholder or not path:
            label.config(text="⚠️ Это поле обязательно для заполнения")
            return
        if not os.path.exists(path):
            label.config(text="❌ Путь не существует")
            return

    # 2. Проверка дат
    try:
        d1, d2 = entry_start_date.get(), entry_end_date.get()
        if d1 == "дд.мм.гггг" or d2 == "дд.мм.гггг":
            date_error_lbl.config(text="⚠️ Введите обе даты периода")
            return
        
        date_start = datetime.strptime(d1, "%d.%m.%Y")
        date_end = datetime.strptime(d2, "%d.%m.%Y")
        if date_start > date_end:
            date_error_lbl.config(text="❌ Дата начала позже даты окончания")
            return
        if date_end > datetime.now():
            date_error_lbl.config(text="❌ Дата окончания не может быть в будущем")
            return
    except ValueError:
        date_error_lbl.config(text="❌ Ошибка в формате даты")
        return

    # 3. Проверка выбора формы отчета
    if combo_sub_type.get() == "---Выберите форму отчета---":
        err_type.config(text="⚠️ Выберите конкретную форму практики из списка")
        return

    messagebox.showinfo("RepGen", "Все проверки пройдены! Начинаю сборку отчетов...")

# --- Сборка ---
root = Tk()
root.title("RepGen")
root.iconbitmap(default = "web.ico")
root.geometry("850x650") 

ttk.Label(text="Параметры формирования отчета", font=("Times New Roman", 16)).pack(pady=10)
ttk.Separator(root, orient='horizontal').pack(fill='x', padx=10, pady=5)

entry_ved, err_ved = create_file_selector(root, "Ведомость", "Укажите файл ведомости...")
entry_pr, err_pr = create_file_selector(root, "Приказ", "Укажите файл приказа...")
entry_folder, err_fld = create_file_selector(root, "Папка с отчетами студентов", "Выберите папку...", True)

entry_start_date, entry_end_date, date_error_lbl = create_date_selector(root)
combo_main_type, combo_sub_type, err_type = create_type_selector(root)

btn_gen = ttk.Button(root, text="СГЕНЕРИРОВАТЬ ОТЧЕТЫ", command=start_gen)
btn_gen.pack(pady=20)

root.mainloop()