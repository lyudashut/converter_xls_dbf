import pandas as pd
import chardet
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import dbf

CONFIG_FILE = "config.json"

REPLACE_MAP = str.maketrans({
    "І": "И", "Ї": "И", "Є": "Е", "Ґ": "Г",
    "і": "и", "ї": "и", "є": "е", "ґ": "г",
    "’": "'", "—": "-", "–": "-", "«": '"', "»": '"',
    "№": "N", "…": "..."
})

def clean_text(value):
    if value is None:
        return ""
    value = str(value)
    value = value.translate(REPLACE_MAP)
    return value[:254]

def detect_encoding(file_path):
    with open(file_path, "rb") as f:
        result = chardet.detect(f.read())
    return result["encoding"]

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_config():
    cfg = {
        "delimiter": delimiter_var.get(),
        "encoding": encoding_var.get(),
        "output_dir": output_dir_var.get()
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def sanitize_columns(df):
    safe = []
    for i, col in enumerate(df.columns):
        col = str(col).strip()
        if not col or col.lower() == "nan":
            col = f"FIELD_{i}"
        if col[0].isdigit():
            col = "F_" + col
        col = col.replace(" ", "_")[:10]
        safe.append(col)
    df.columns = safe
    return df

def convert_file(input_file, delimiter, encoding_choice, output_dir, progress=None):
    print("\n=== Обробка файлу:", input_file)

    ext = os.path.splitext(input_file)[1].lower()

    if encoding_choice == "auto":
        encoding = detect_encoding(input_file)
        print("Автовизначене кодування:", encoding)
    else:
        encoding = encoding_choice

    if ext == ".csv":
        df = pd.read_csv(
            input_file,
            dtype=str,
            encoding=encoding,
            sep=None,
            engine="python"
        )
    else:
        df = pd.read_excel(input_file, dtype=str)

    if isinstance(df, pd.Series):
        df = df.to_frame()

    if df is None or df.empty:
        raise ValueError("Файл порожній або має неправильний формат.")

    # Приводимо ВСЕ до тексту
    df = df.astype(str)

    # Чистимо текст (pandas 3.x — без applymap)
    df = df.apply(lambda col: col.map(clean_text))

    # Прибираємо всі варіанти NaN/null
    df = df.replace(["nan", "NaN", "NAN", "None", "NONE", "null", "NULL"], "")

    # Назви колонок
    df = sanitize_columns(df)

    # Оптимальна ширина полів
    widths = {}
    for col in df.columns:
        max_len = df[col].astype(str).str.len().max()
        widths[col] = min(max(max_len, 1), 254)

    # Вихідний DBF
    base = os.path.splitext(os.path.basename(input_file))[0]
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, base + ".dbf")
    else:
        output_file = os.path.splitext(input_file)[0] + ".dbf"

    print("Створюємо DBF:", output_file)

    # Структура DBF
    structure = ";".join([f"{col} C({widths[col]})" for col in df.columns])

    # DBF у DOS-кодуванні (CP866)
    table = dbf.Table(output_file, structure, codepage="cp866")
    table.open(dbf.READ_WRITE)

    # Записуємо рядки — кодуємо в cp866, щоб уникнути проблем
    for _, row in df.iterrows():
        safe_row = []
        for val in row:
            val = str(val)
            try:
                val.encode("cp866")
            except:
                val = val.encode("cp866", errors="replace").decode("cp866")
            safe_row.append(val)
        table.append(tuple(safe_row))

    table.close()

    if progress:
        progress["value"] += 1

    print("Готово:", output_file)
    return True

def choose_single_file():
    file_path = filedialog.askopenfilename(
        title="Виберіть Excel або CSV",
        filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")]
    )
    if file_path:
        run_conversion([file_path])

def choose_folder():
    folder = filedialog.askdirectory(title="Виберіть папку з файлами")
    if folder:
        files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith((".xlsx", ".xls", ".csv"))
        ]
        run_conversion(files)

def choose_output_dir():
    folder = filedialog.askdirectory(title="Виберіть вихідну папку")
    if folder:
        output_dir_var.set(folder)
        save_config()

def run_conversion(files):
    if not files:
        messagebox.showerror("Помилка", "Немає файлів для конвертації")
        return

    delimiter = delimiter_var.get()
    encoding_choice = encoding_var.get()
    output_dir = output_dir_var.get()

    progress = ttk.Progressbar(root, length=350, mode="determinate", maximum=len(files))
    progress.pack(pady=10)

    success = 0

    for f in files:
        try:
            if convert_file(f, delimiter, encoding_choice, output_dir, progress):
                success += 1
        except Exception as e:
            print("\n=== ПОМИЛКА ===")
            print("Файл:", f)
            print("Причина:", e)

        root.update_idletasks()

    progress.destroy()
    save_config()

    messagebox.showinfo("Готово", f"Успішно конвертовано: {success} із {len(files)} файлів")

def open_output_dir():
    path = output_dir_var.get()
    if not path:
        messagebox.showerror("Помилка", "Вихідна папка не вибрана")
        return
    if not os.path.exists(path):
        messagebox.showerror("Помилка", "Папка не існує")
        return
    os.startfile(path)

# --- GUI ---
root = tk.Tk()
root.title("Конвертер Excel/CSV → DBF (Pro v3.9, CP866)")
root.geometry("520x380")

cfg = load_config()

label = tk.Label(root, text="Конвертер Excel/CSV → DBF (Pro v3.9, CP866)", font=("Arial", 14))
label.pack(pady=10)

delimiter_var = tk.StringVar(value=cfg.get("delimiter", ";"))
tk.Label(root, text="Розділювач CSV:").pack()
tk.OptionMenu(root, delimiter_var, ";", ",", "|", "\t").pack()

encoding_var = tk.StringVar(value=cfg.get("encoding", "auto"))
tk.Label(root, text="Кодування CSV:").pack()
tk.OptionMenu(root, encoding_var, "auto", "cp1251", "utf-8").pack()

output_dir_var = tk.StringVar(value=cfg.get("output_dir", ""))
tk.Label(root, text="Вихідна папка (для DBF):").pack(pady=(10, 0))
frame_out = tk.Frame(root)
frame_out.pack(pady=5)
entry_out = tk.Entry(frame_out, textvariable=output_dir_var, width=40)
entry_out.pack(side=tk.LEFT, padx=5)
btn_out = tk.Button(frame_out, text="Обрати...", command=choose_output_dir)
btn_out.pack(side=tk.LEFT)

tk.Button(root, text="Обрати файл", font=("Arial", 12),
          command=choose_single_file).pack(pady=8)
tk.Button(root, text="Обрати папку", font=("Arial", 12),
          command=choose_folder).pack(pady=8)
tk.Button(root, text="Відкрити вихідну папку", font=("Arial", 11),
          command=open_output_dir).pack(pady=8)

root.mainloop()