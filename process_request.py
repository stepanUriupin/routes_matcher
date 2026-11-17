"""
Все словари и настройки лежат в отдельных файлах .py:
    dict_layout.py      где в заявке лежат данные (пункты, ячейки, месяцы)
    dict_blocks.py      порты и города (тобишь блоки)
    dict_incoterms.py   термины Инкотермс 2020
    dict_types.py       колонки таблицы типов и какие поля заявки используем
"""
# тк я еще не получил окончательную форму заявки, крайне тяжело импортировать данные тк я до конца не понимаю какая будет заявка на самом деле и как на самом деле выглядит маршрут
# это бета версия на заявке, которая прикреплена в гитхабе в папке которая будет называться что то типо test_data или что то такое

# p.s тесты писала иишка, но я так поглядел они вроде норм


import json
import re
import sys
from pathlib import Path
import pandas as pd
from docx import Document
from openpyxl import load_workbook

from dict_blocks import BLOCKS, PORT_SUFFIX
from dict_incoterms import INCOTERMS, INCOTERMS_ALIASES, SEA_ONLY
from dict_layout import (DOCX_TABLE, DOCX_VALUE_COL, FIELDS, HEADER_PARAGRAPHS, MONTHS,
                         XLSX_FIRST_FIELD_ROW, XLSX_HEADER_CELLS, XLSX_LABEL_COL, XLSX_VALUE_COL)
from dict_types import (EMPTY_MARKS, FIELD_DESTINATION, FIELD_HUB, FIELD_INCOTERM, FIELD_ORIGIN,
                        REQUIRED_COLUMNS, SEA, TYPES_COLUMNS)



# ОБЩИЕ ФУНКЦИЯ


def norm(value):
    """Единый вид для сравнения: без лишних пробелов, без разницы в регистре."""
    return " ".join(str(value if value is not None else "").split()).casefold()


def as_text(value):
    """Поле заявки бывает строкой или списком строк - приводим к строке."""
    if isinstance(value, (list, tuple)):
        return " ".join(str(v) for v in value)
    return str(value or "")


def is_empty(value):
    return norm(value) in EMPTY_MARKS


def _lines(text):
    return [t.strip() for t in str(text or "").splitlines() if t.strip()]


# Справочник блоков в виде пар (что ищем в тексте, имя блока)
BLOCK_ENTRIES = [(norm(synonym), block) for block, synonyms in BLOCKS.items() for synonym in synonyms]



# ЧТЕНИЕ ЗАЯВКИ
# Каждая функция дает одно и тоже
# header_lines - строки шапки, labels[n] / values[n] - подпись и значение пункта n
# (значение - список строк), customer - заказчик


def read_docx(path):
    doc = Document(path)
    header_lines = [p.text.strip() for p in doc.paragraphs[:HEADER_PARAGRAPHS] if p.text.strip()]
    table = doc.tables[DOCX_TABLE]
    labels, values = {}, {}
    for n in FIELDS:
        if n - 1 >= len(table.rows):  # если пункт отсутствует в заявке
            labels[n], values[n] = "", []
            continue
        row = table.rows[n - 1]
        labels[n] = row.cells[0].text.strip()
        values[n] = _lines(row.cells[DOCX_VALUE_COL].text)
    # заказчик - абзац сразу после «Заказчик:»
    texts = [p.text.strip() for p in doc.paragraphs]
    customer = ""
    for i, t in enumerate(texts):
        if t.startswith("Заказчик"):
            customer = texts[i + 1] if i + 1 < len(texts) else ""
            break
    return header_lines, labels, values, customer


def read_xlsx(path):
    ws = load_workbook(path, data_only=True).active
    header_lines = [str(ws[c].value).strip() for c in XLSX_HEADER_CELLS if ws[c].value]
    labels, values = {}, {}
    for n in FIELDS:
        r = XLSX_FIRST_FIELD_ROW + n - 1
        labels[n] = str(ws[f"{XLSX_LABEL_COL}{r}"].value or "").strip()
        values[n] = _lines(ws[f"{XLSX_VALUE_COL}{r}"].value)
    # заказчик - ячейка справа от «Заказчик:»
    customer = ""
    for row in ws.iter_rows():
        if row[0].value and str(row[0].value).startswith("Заказчик"):
            customer = str(row[1].value or "").strip()
            break
    return header_lines, labels, values, customer


READERS = {".docx": read_docx, ".xlsx": read_xlsx}


def parse_header(lines):
    # Номер заявки, договор и дата оформления - из строк шапки.
    text = " ".join(lines)
    res = {}
    m = re.search(r"Заявка[^№]*№\s*(\d+)", text)
    if m:
        res["request_number"] = m.group(1)
    m = re.search(r"Договор\s*№\s*(?:No\.?\s*)?(\S+)\s+от\s+(\d{2}\.\d{2}\.\d{4})", text)
    if m:
        res["contract_number"], res["contract_date"] = m.group(1), m.group(2)
    m = re.search(r"(\d{1,2})\W{0,3}\s*([а-яё]+)\s*(\d{4})", text.lower())
    if m and m.group(2) in MONTHS:
        res["request_date"] = f"{int(m.group(1)):02d}.{MONTHS[m.group(2)]:02d}.{m.group(3)}"
    return res


def extract_request(path):
    # читает заявку -> словарь с шапкой, заказчиком и полями
    reader = READERS.get(Path(path).suffix.lower())
    if reader is None:
        raise ValueError(f"Формат заявки не поддерживается: {path}")
    header_lines, labels, values, customer = reader(path)

    # защита от сдвига формы: подпись пункта N должна начинаться с «N.»
    warnings = [f"пункт {n}: подпись «{labels[n][:30]}» не начинается с «{n}.»"
                for n in FIELDS if not labels[n].startswith(f"{n}.")]

    result = parse_header(header_lines)
    result["customer"] = customer
    result["fields"] = {FIELDS[n]: values[n] for n in FIELDS}
    result["warnings"] = warnings
    return result



# ТАБЛИЦА ТИПОВ


def load_types(path):
    df = pd.read_excel(path, header=None, dtype=str).fillna("")

    # строка заголовков - та, где есть колонка Базис_поставки (над ней может быть название таблицы)
    header_row = next((i for i in range(len(df))
                       if "Базис_поставки" in [str(v).strip() for v in df.iloc[i]]), None)
    if header_row is None:
        raise ValueError("В таблице типов не найдена колонка «Базис_поставки»")
    headers = {j: TYPES_COLUMNS[str(v).strip()] for j, v in enumerate(df.iloc[header_row])
               if str(v).strip() in TYPES_COLUMNS}

    types, warnings = [], []
    for i in range(header_row + 1, len(df)):
        row = {key: str(df.iat[i, j]).strip() for j, key in headers.items()}
        if all(is_empty(row.get(k)) for k in REQUIRED_COLUMNS):
            continue  # пустая заготовка
        missing = [k for k in REQUIRED_COLUMNS if is_empty(row.get(k))]
        if missing:
            warnings.append(f"тип ID {row.get('id') or '?'} пропущен: не заполнено {', '.join(missing)}")
            continue
        types.append(row)
    return types, warnings



# РАСПОЗНАВАНИЕ И ОПРЕДЕЛЕНИЕ ТИПА


def find_block(text, extra_names=()):
    # находит "лучшее" (тобишь самое длинное) совпадение, если нет то Ноне
    text = norm(text)
    candidates = BLOCK_ENTRIES + [(norm(n), n) for n in extra_names if not is_empty(n)]
    best = None
    for key, block in candidates:
        if key and re.search(rf"(?<!\w){re.escape(key)}(?!\w)", text):
            if best is None or len(key) > len(best[0]):
                best = (key, block)
    return best[1] if best else None


def find_incoterm(text):
    # Достаёт код базиса из строки (если он ваще там есть) upd требует пересмотра, как только уточнится формат заявки
    upper = as_text(text).upper()
    for alias, real in INCOTERMS_ALIASES.items():
        upper = re.sub(rf"\b{alias}\b", real, upper)
    for code in INCOTERMS:
        if re.search(rf"\b{code}\b", upper):
            return code
    return None


def detect_type(request, types_path):
    types, warnings = load_types(types_path)
    review = []  # посмотреть вручную / feedback

    def result(numbers, status, details):
        return {"numbers": numbers, "status": status, "details": details,
                "needs_review": review, "warnings": warnings}

    # 1. базис поставки
    incoterm = find_incoterm(request.get(FIELD_INCOTERM))
    if incoterm is None:
        review.append({"поле": "базис поставки", "значение": as_text(request.get(FIELD_INCOTERM)).strip()})

    # 2. определяет блок
    names = lambda col: [t[col] for t in types]
    origin_text, hub_text = as_text(request.get(FIELD_ORIGIN)), as_text(request.get(FIELD_HUB))
    origin = find_block(origin_text, names("origin"))
    hub = find_block(hub_text, names("hub"))
    destination = find_block(as_text(request.get(FIELD_DESTINATION)), names("destination"))
    details = {"базис": incoterm, "место погрузки": origin, "место выгрузки": hub,
               "место назначения": destination}

    # 3. море если оба порты
    both_ports = bool(origin and hub and origin.endswith(PORT_SUFFIX) and hub.endswith(PORT_SUFFIX))
    for label, block, text in (("место погрузки", origin, origin_text), ("место выгрузки", hub, hub_text)):
        if not (block and block.endswith(PORT_SUFFIX)):
            review.append({"поле": label, "значение": text.strip(),
                           "причина": "не найдено в блоках портов"})
    if not both_ports:
        status = "conflict" if incoterm in SEA_ONLY else "not_sea"
        return result([0], status, details)
    details["способ"] = SEA

    # 4. место назначения aka destination
    if destination is None:
        review.append({"поле": "место назначения", "значение": as_text(request.get(FIELD_DESTINATION)).strip(),
                       "причина": "не найдено в блоках и в таблице типов"})
        return result([], "needs_review", details)
    if incoterm is None:
        return result([], "needs_review", details)

    # 5. поиск нужных строк в таблице
    found = [int(float(t["id"])) for t in types
             if t["incoterm"].upper() == incoterm and norm(t["mode"]) == norm(SEA)
             and norm(t["origin"]) == norm(origin) and norm(t["hub"]) == norm(hub)
             and norm(t["destination"]) == norm(destination)]
    return result(found, "ok" if found else "not_in_list", details)



# ВСЁ ВМЕСТЕ


def process_request(request_path, types_path):
    """Заявка (Word/Excel) -> данные заявки и номер типа перевозки."""
    request = extract_request(request_path)
    found = detect_type(request["fields"], types_path)
    return {
        "request_number": request.get("request_number"),
        "request_date": request.get("request_date"),
        "customer": request.get("customer"),
        "numbers": found["numbers"],
        "status": found["status"],
        "details": found["details"],
        "needs_review": found["needs_review"],
        "warnings": request["warnings"] + found["warnings"],
    }


# прост запуск че бубнить то
if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    print(json.dumps(process_request(sys.argv[1], sys.argv[2]), ensure_ascii=False, indent=2))


