#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_catalog_from_xlsx.py
============================
Пересобирает HTML-каталог библиотеки из одного файла Excel.

ИСПОЛЬЗОВАНИЕ:
    python update_catalog_from_xlsx.py Каталог.xlsx catalog_template.html
    python update_catalog_from_xlsx.py Каталог.xlsx catalog_template.html index.html

Ключевые особенности:
    - Автоматически определяет столбцы по заголовкам.
    - Удаляет полные дубликаты (название + автор + инв. номер + год + издательство).
      Книги с одинаковым номером, но разными годами/издательствами СОХРАНЯЮТСЯ.
    - Обложки подставляются по инвентарному номеру из папки covers/.
"""

import sys
import os
import re
import json
import datetime

try:
    import pandas as pd
except ImportError:
    sys.exit(
        "ОШИБКА: не найден пакет pandas. Установите зависимости:\n"
        "    pip install pandas openpyxl"
    )

# ---------------------------------------------------------------------------
# Соответствие "смысл столбца" -> ключевые слова для поиска в заголовках xlsx
# ---------------------------------------------------------------------------
HEADER_HINTS = {
    "title": ["назв", "title", "книг"],
    "author": ["автор", "author"],
    "publisher": ["издат", "publisher"],
    "year": ["год", "year"],
    "copies": ["экземп", "кол-во", "количество", "шт.", "шт", "copies"],
    "section": ["раздел", "секц", "section", "категор"],
    "inventory": ["инвентар", "инв.", "inventory", "номер"],
    "date_added": ["добавл", "дата", "date"],
    "cover": ["обложк", "cover", "фото", "изображ"],
}
FIELD_ORDER = ["title", "author", "publisher", "year", "copies",
               "section", "inventory", "date_added", "cover"]

SECTION_FIX = {
    "ESPA?OL": "ESPAÑOL",
    "FRAN?AIS": "FRANÇAIS",
    "Italiana": "ITALIANA",
    "Энциклопедии": "Энциклопедия",
    "НА иностр. Язык": "на иностр. язык",
}

STOP_WORDS = {
    "КАК", "ЕЁ", "ЕГО", "ИЛИ", "ДЛЯ", "ПРИ", "ОТ", "ДО", "ПО", "ИЗ", "НА",
    "И", "В", "К", "С", "О", "У", "ЙОГИ", "КНИГА", "ТОМ",
}


def detect_columns(df):
    """Сопоставляет столбцы xlsx с нужными полями по заголовкам."""
    header_row = [str(c).strip().lower() for c in df.columns]
    mapping = {}
    for field, hints in HEADER_HINTS.items():
        for i, col_name in enumerate(header_row):
            if any(h in col_name for h in hints):
                mapping[field] = i
                break

    if len(mapping) >= 3:
        print(f"Обнаружены заголовки столбцов, сопоставлено полей: {len(mapping)}/9")
        return mapping, True

    print("Заголовки не распознаны — использую порядок столбцов A-I.")
    mapping = {field: i for i, field in enumerate(FIELD_ORDER) if i < len(df.columns)}
    return mapping, False


def check_files_exist(xlsx_path, template_path):
    """Проверяет, что входные файлы существуют."""
    problems = []

    if not os.path.isfile(xlsx_path):
        problems.append(f"Не найден файл каталога: {xlsx_path}")
    elif not xlsx_path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        problems.append(f"Файл '{xlsx_path}' не похож на Excel-файл.")

    if not os.path.isfile(template_path):
        problems.append(f"Не найден файл шаблона: {template_path}")

    if problems:
        print("ОШИБКА — скрипт остановлен:\n")
        for p in problems:
            print(" ✗ " + p)
        sys.exit(1)


def clean_year(raw):
    """Извлекает год из значения."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return 0
    s = str(raw).strip()
    if not s or s.lower() == "nan":
        return 0
    m = re.search(r"\d{3,4}", s)
    return int(m.group()) if m else 0


def clean_copies(raw):
    """Извлекает количество экземпляров."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return 1
    s = str(raw).strip()
    m = re.search(r"\d+", s)
    return int(m.group()) if m else 1


def clean_date(raw):
    """Приводит дату к формату ГГГГ-ММ-ДД."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    if isinstance(raw, (datetime.datetime, datetime.date)):
        return raw.strftime("%Y-%m-%d")
    s = str(raw).strip()
    if not s or s.lower() == "nan":
        return ""
    s = s.replace(".", "-").replace("/", "-").replace(" ", "-")
    parts = [p for p in s.split("-") if p]
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        y, m, d = parts
        if len(y) != 4:
            d, m, y = parts
        return f"{y}-{m.zfill(2)}-{d.zfill(2)}"
    return s


def clean_str(raw):
    """Приводит к строке, убирает пробелы."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    return str(raw).strip()


COVER_EXT_RE = re.compile(r"\.(jpe?g|png|webp|gif|bmp|svg)([?#].*)?$", re.IGNORECASE)


def clean_cover(raw):
    """Возвращает значение столбца 'Обложка', только если оно похоже на файл/ссылку."""
    s = clean_str(raw)
    if not s:
        return ""
    low = s.lower()
    if low.startswith("http://") or low.startswith("https://"):
        return s
    if COVER_EXT_RE.search(low):
        return s
    return ""


def infer_type(section, title):
    """Определяет тип книги по разделу и названию."""
    s = section.lower()
    t = title.lower()
    if "курс" in s:
        return "курс"
    if "священ" in s:
        return "священный текст"
    if "агиограф" in s:
        return "биография"
    if "йога" in s and "курс" not in s:
        return "практическое руководство"
    if "аюрвед" in s:
        return "Аюрведа"
    if "медицин" in s:
        return "медицина"
    if "иностр" in s or "язык" in s:
        return "самоучитель"
    if "книги гуру" in s:
        return "книга гуру"
    if "детск" in s:
        return "для детей"
    if "словар" in s:
        return "словарь"
    if "журнал" in t or "issue" in t or "vol." in t:
        return "журнал"
    return "книга"


def extract_tags(title, section):
    """Извлекает теги из названия и раздела."""
    tags = [section]
    words = re.findall(r"[А-ЯЁ][а-яёА-ЯЁ\-]{2,}", title)
    seen = set(tags)
    count = 0
    for w in words:
        if w.upper() in STOP_WORDS or w in seen:
            continue
        seen.add(w)
        tags.append(w)
        count += 1
        if count >= 3:
            break
    return tags


def main():
    if len(sys.argv) not in (3, 4):
        print(__doc__)
        sys.exit(1)

    xlsx_path, template_path = sys.argv[1], sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) == 4 else "index.html"

    check_files_exist(xlsx_path, template_path)

    print(f"Читаю {xlsx_path} ...")
    try:
        raw = pd.read_excel(xlsx_path, sheet_name=0, header=None, dtype=object)
        with_header = pd.read_excel(xlsx_path, sheet_name=0, dtype=object)
    except Exception as e:
        print(f"ОШИБКА при чтении xlsx-файла: {e}")
        sys.exit(1)

    if len(with_header) == 0 and len(raw) == 0:
        print(f"ОШИБКА: файл '{xlsx_path}' не содержит данных.")
        sys.exit(1)

    mapping, has_header = detect_columns(with_header)
    df = with_header if has_header else raw
    if has_header:
        df = df.reset_index(drop=True)

    records = []
    seen_keys = set()
    skipped_empty = 0
    dupe_count = 0
    ignored_covers = 0

    for _, row in df.iterrows():
        def get(field):
            idx = mapping.get(field)
            if idx is None or idx >= len(row):
                return None
            return row.iloc[idx]

        title = clean_str(get("title"))
        if not title or title.lower() in ("nan", "название", "title"):
            skipped_empty += 1
            continue

        author = clean_str(get("author"))
        publisher = clean_str(get("publisher"))
        year = clean_year(get("year"))
        copies = clean_copies(get("copies"))
        section = clean_str(get("section")) or "Без раздела"
        section = SECTION_FIX.get(section, section)
        inventory = clean_str(get("inventory"))
        date_added = clean_date(get("date_added"))
        cover_raw = clean_str(get("cover"))
        cover = clean_cover(get("cover"))

        if cover_raw and not cover:
            ignored_covers += 1

        # Ключ для проверки дубликатов:
        # учитывает название, автора, номер, год и издательство —
        # поэтому разные издания одной книги НЕ удаляются
        key = (
            title.lower(),
            author.lower(),
            inventory.lower(),
            str(year),
            publisher.lower(),
        )

        if key in seen_keys:
            dupe_count += 1
            continue

        seen_keys.add(key)

        records.append({
            "title": title,
            "author": author,
            "publisher": publisher,
            "year": year,
            "copies": copies,
            "section": section,
            "inventory": inventory,
            "date_added": date_added,
            "language": "ru",
            "type": infer_type(section, title),
            "tags": extract_tags(title, section),
            "cover": cover,
        })

    records.sort(key=lambda r: r["title"].lower())

    print(f"Готово: {len(records)} книг. Пропущено пустых строк: {skipped_empty}. "
          f"Дубликатов убрано: {dupe_count}.")

    if ignored_covers:
        print(f"⚠ В столбце 'Обложка' у {ignored_covers} книг(и) найдено значение, "
              f"не похожее на имя файла или ссылку — оно проигнорировано.")

    if len(records) == 0:
        print("ОШИБКА: после обработки не осталось ни одной книги.")
        sys.exit(1)

    # Превью
    print("\nПервые записи для проверки:")
    for r in records[:3]:
        print(f"  • «{r['title']}» — {r['author'] or '(автор не указан)'}, "
              f"{r['year'] or 'год не указан'}, инв. № {r['inventory'] or '—'}")

    # Защита: не перезаписываем шаблон
    if os.path.abspath(output_path) == os.path.abspath(template_path):
        print("ОШИБКА: имя выходного файла совпадает с шаблоном.")
        sys.exit(1)

    try:
        with open(template_path, encoding="utf-8") as f:
            html = f.read()
    except Exception as e:
        print(f"ОШИБКА при чтении шаблона: {e}")
        sys.exit(1)

    if "__CATALOG_DATA__" not in html:
        print("ОШИБКА: в шаблоне не найден маркер __CATALOG_DATA__.")
        sys.exit(1)

    data_json = json.dumps(records, ensure_ascii=False)
    html = html.replace("__CATALOG_DATA__", data_json)

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
    except Exception as e:
        print(f"ОШИБКА при сохранении файла: {e}")
        sys.exit(1)

    print(f"\nСохранено: {output_path}")


if __name__ == "__main__":
    main()