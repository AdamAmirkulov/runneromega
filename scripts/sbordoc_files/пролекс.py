import re
import os
import pdfplumber

# УКАЖИ ПУТЬ К ПАПКЕ С PDF
folder = r"C:\Users\User\Desktop\А Омега караганда"

# Казахские и русские буквы (верхний и нижний регистр)
KZ = r"А-ЯЁа-яёӘәҒғҚқҢңӨөҰұҮүҺһІі"

for filename in os.listdir(folder):
    if not filename.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(folder, filename)

    with pdfplumber.open(pdf_path) as pdf:
        text = " ".join(page.extract_text() or "" for page in pdf.pages)
        text = re.sub(r"\s+", " ", text)

    match = re.search(
        rf"от\s+([{KZ}][{KZ}\s]+?),\s*[\d\.]+г\.р\.,\s*ИИН:\s*(\d{{12}})",
        text
    )
    if not match:
        print(f"Не найдено: {filename}")
        continue

    fio = re.sub(r"\s+", " ", match.group(1).strip())
    iin = match.group(2)

    new_name = f"Постановление об отмене ИН, {fio}, {iin}.pdf"
    new_path = os.path.join(folder, new_name)

    if os.path.exists(new_path):
        print(f"Уже существует, пропускаем: {new_name}")
        continue

    os.rename(pdf_path, new_path)
    print(f"OK: {new_name}")