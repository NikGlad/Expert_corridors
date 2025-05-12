import tkinter as tk
from tkinter import messagebox
from docx import Document

# === Извлечение разделов по ключевым словам ===
def extract_sections(doc, keywords):
    result = []
    capture = False
    current_keyword = None

    for para in doc.paragraphs:
        text = para.text.strip()

        for keyword in keywords:
            if keyword in text and not text.endswith("конец"):
                capture = True
                current_keyword = keyword
                result.append((para, keyword))
                break
            elif keyword in text and text.endswith("конец") and capture and current_keyword == keyword:
                result.append((para, keyword))
                capture = False
                current_keyword = None
                break
        else:
            if capture:
                result.append((para, current_keyword))

    return result

# === Создание документа ===
def generate_doc():
    try:
        src_doc = Document("исходный.docx")
        dst_doc = Document()

        selected_keywords = []

        # ПРАВИТЬ. Обработка основной опции "Ф1.3 Ширина коридора в зависимости от длины" с подопциями
        if var_f13_width.get():
            if f13_width_subvar.get() == "до 40м":
                selected_keywords.append("Длина коридора до 40м")
            elif f13_width_subvar.get() == "более 40м":
                selected_keywords.append("Длина коридора более 40м")
            else:
                messagebox.showwarning("Внимание", "Вы выбрали 'Ф1.3 Ширина коридора в зависимости от длины', но не указали длину коридора.")
                return
        # ПРАВИТЬ.
        if var_corridor_width.get():
            selected_keywords.append("Ширина коридора. Для всех зданий")
        if var_f11_corridor_width.get():
            selected_keywords.append("Ф1.1 Ширина коридора.")

        if not selected_keywords:
            messagebox.showwarning("Внимание", "Выберите хотя бы один раздел.")
            return

        sections = extract_sections(src_doc, selected_keywords)

        for para, _ in sections:
            new_para = dst_doc.add_paragraph()
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                new_run.font.name = run.font.name
                new_run.font.size = run.font.size
                new_run.font.color.rgb = run.font.color.rgb if run.font.color else None

        dst_doc.save("готовый.docx")
        messagebox.showinfo("Готово", "Файл 'готовый.docx' успешно создан.")

    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

# === Интерфейс ===
root = tk.Tk()
root.title("Expert")
root.geometry("800x800")

# Прокрутка
top_frame = tk.Frame(root)
top_frame.pack(fill="both", expand=True)

canvas = tk.Canvas(top_frame)
scrollbar = tk.Scrollbar(top_frame, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# ПРАВИТЬ Переменные
var_f13_width = tk.BooleanVar()
var_corridor_width = tk.BooleanVar()
var_f11_corridor_width = tk.BooleanVar()

# ПРАВИТЬ Переменные для вложенных кнопок
f13_width_subvar = tk.StringVar(value="")

# ПРАВИТЬ, ЕСЛИ ЕСТЬ ВЛОЖЕННЫЕ КНОПКИ. Отображение вложенных кнопок
def toggle_vstroyka_suboptions():
    if var_f13_width.get():
        f13_width_subframe.pack(anchor="w", padx=20)
    else:
        f13_width_subframe.pack_forget()
        f13_width_subvar.set("")

# Интерфейс выбора разделов
tk.Label(scrollable_frame, text="Выберите необходимые параметры:").pack(anchor="w", pady=(10, 5))



# РАЗДЕЛЫ
tk.Checkbutton(scrollable_frame, text="Ширина коридора. Для всех зданий", variable=var_corridor_width).pack(anchor="w")
tk.Checkbutton(scrollable_frame, text="Ф1.1 Ширина коридора.", variable=var_f11_corridor_width).pack(anchor="w")

 # ПРАВИТЬ, ЕСЛИ ЕСТЬ ВЛОЖЕННЫЕ КНОПКИ. РАЗДЕЛ + вложенные радиокнопки
f13_width_frame = tk.Frame(scrollable_frame)
f13_width_frame.pack(anchor="w", fill="x")

tk.Checkbutton(f13_width_frame, text="Ф1.3 Ширина коридора в зависимости от длины", variable=var_f13_width,
               command=toggle_vstroyka_suboptions).pack(anchor="w")

f13_width_subframe = tk.Frame(f13_width_frame)
f13_width_subframe.pack(anchor="w", padx=20)
f13_width_subframe.pack_forget()

# ПРАВИТЬ, ЕСЛИ ЕСТЬ ВЛОЖЕННЫЕ КНОПКИ
tk.Label(f13_width_subframe, text="Выберите длину коридора:").pack(anchor="w")
tk.Radiobutton(f13_width_subframe, text="Длина коридора до 40м",
               variable=f13_width_subvar, value="до 40м").pack(anchor="w")
tk.Radiobutton(f13_width_subframe, text="Длина коридора более 40м",
               variable=f13_width_subvar, value="более 40м").pack(anchor="w")
# конец РАЗДЕЛ + вложенные радиокнопки

# Кнопка
bottom_frame = tk.Frame(root)
bottom_frame.pack(fill="x", pady=10)
tk.Button(bottom_frame, text="Создать готовый.docx", command=generate_doc,
          height=2, font=("Arial", 11, "bold")).pack()

# Запуск
root.mainloop()
