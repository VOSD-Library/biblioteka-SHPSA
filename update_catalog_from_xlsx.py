#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_catalog_from_xlsx.py
============================
Пересобирает HTML-каталог библиотеки из одного файла Excel.

ИСПОЛЬЗОВАНИЕ:
    python update_catalog_from_xlsx.py Каталог.xlsx catalog_template.html
    python update_catalog_from_xlsx.py Каталог.xlsx catalog_template.html имя_файла.html

Аргументы:
    1) путь к xlsx-файлу с каталогом
    2) путь к HTML-шаблону (catalog_template.html — идёт вместе со скриптом,
       НЕ редактируйте его руками, в нём весь дизайн и логика сайта)
    3) (необязательно) имя итогового файла. Если не указать — скрипт сам
       назовёт файл по текущей дате и времени, например:
           catalog 2026 09 06 17 44.html
       и положит его рядом с xlsx-файлом.

ОЖИДАЕМАЯ СТРУКТУРА xlsx:
    Один лист, одна строка = одна книга. Столбцы (порядок как в исходном каталоге):
        A: Название
        B: Автор
        C: Издательство
        D: Год
        E: Экземпляров
        F: Раздел
        G: Инвентарный номер
        H: Дата добавления (ГГГГ-ММ-ДД, ГГГГ.ММ.ДД или ГГГГ ММ ДД)
        I: Обложка (необязательно) — имя файла или ссылка на картинку обложки

    Если в первой строке файла есть заголовки (например: "Название", "Автор",
    "Раздел" и т.д.) — скрипт сам их найдёт и сопоставит столбцы по смыслу,
    независимо от их порядка. Если заголовков нет — берёт данные по порядку
    столбцов A–H (I — необязательно), как описано выше.

    Лишние столбцы игнорируются.

ОБЛОЖКИ КНИГ:
    Есть два способа показать реальную обложку вместо цветной подложки —
    можно использовать оба сразу:

    1) Автоматически по инвентарному номеру (самый простой способ):
       Создайте рядом с готовым HTML-файлом папку с именем "covers" и
       положите туда фото обложки в формате .jpg с именем ТОЧНО как
       инвентарный номер книги, например:
           covers/ШПСА 00545.jpg
       Сайт сам найдёт нужный файл. Ничего в xlsx менять не нужно.

    2) Через столбец "Обложка" в xlsx — впишите туда имя файла или прямую
       ссылку на картинку в интернете. Это удобно, если обложки не в
       формате .jpg или названы иначе.

    Если обложки нет ни там, ни там — книга покажется с цветной подложкой
    и первой буквой названия, как раньше. Ничего не ломается.

УСТАНОВКА ЗАВИСИМОСТЕЙ (один раз):
    pip install pandas openpyxl --break-system-packages
    (на Windows/Mac обычно достаточно: pip install pandas openpyxl)

Что делает скрипт:
    - проверяет, что все указанные файлы существуют и открываются
    - читает xlsx и проверяет, что столбцы сопоставлены правильно
      (например, что в столбце "Год" на самом деле годы, а не текст)
    - убирает пустые строки и очевидные строки-заголовки, случайно попавшие в данные
    - убирает дубликаты (по названию + автору + инвентарному номеру)
    - определяет "тип" книги и извлекает несколько тегов из названия для фильтров
    - подставляет всё это в HTML-шаблон и сохраняет готовый файл с именем,
      содержащим дату и время создания
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
        "    pip install pandas openpyxl --break-system-packages\n"
        "(или без флага --break-system-packages, если это ваш личный компьютер)"
    )

# ---------------------------------------------------------------------------
# Соответствие "смысл столбца" -> ключевые слова для поиска в заголовках xlsx
# ---------------------------------------------------------------------------
HEADER_HINTS = {
    "title": ["назв", "title", "книг"],
    "author": ["автор", "author"],
    "publisher": ["издат", "publisher"],
    "year": ["год", "year"],
    "copies": ["экземп", "кол-во", "количество", "copies"],
    "section": ["раздел", "секц", "section", "категор"],
    "inventory": ["инвентар", "инв.", "inventory", "номер"],
    "date_added": ["добавл", "дата", "date"],
    "cover": ["обложк", "cover", "фото", "изображ"],
}
FIELD_ORDER = ["title", "author", "publisher", "year", "copies", "section", "inventory", "date_added", "cover"]

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
    """Пытается сопоставить столбцы xlsx с нужными полями по заголовкам.
    Если заголовки не распознаны как минимум для 3 полей — считает, что
    заголовков нет, и использует порядок столбцов A-H."""
    header_row = [str(c).strip().lower() for c in df.columns]
    mapping = {}
    for field, hints in HEADER_HINTS.items():
        for i, col_name in enumerate(header_row):
            if any(h in col_name for h in hints):
                mapping[field] = i
                break

    if len(mapping) >= 3:
        print(f"Обнаружены заголовки столбцов, сопоставлено полей: {len(mapping)}/8")
        return mapping, True

    print("Заголовки не распознаны — использую порядок столбцов A-H (без строки заголовка).")
    mapping = {field: i for i, field in enumerate(FIELD_ORDER) if i < len(df.columns)}
    return mapping, False


def check_files_exist(xlsx_path, template_path):
    """Проверяет, что входные файлы существуют и имеют правильное расширение.
    При проблеме — печатает понятное сообщение и завершает работу."""
    problems = []

    if not os.path.isfile(xlsx_path):
        problems.append(f"Не найден файл каталога: {xlsx_path}\n"
                         f"   Проверьте, что имя написано без опечаток и файл лежит в этой же папке.")
    elif not xlsx_path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        problems.append(f"Файл '{xlsx_path}' не похож на Excel-файл (.xlsx/.xls). "
                         f"Проверьте, тот ли файл вы указали.")

    if not os.path.isfile(template_path):
        problems.append(f"Не найден файл шаблона: {template_path}\n"
                         f"   Скачайте catalog_template.html и положите его в эту же папку.")

    if problems:
        print("ОШИБКА — скрипт остановлен, ничего не изменено:\n")
        for p in problems:
            print(" ✗ " + p)
        sys.exit(1)


def validate_mapping(df, mapping, has_header):
    """Проверяет здравый смысл сопоставленных столбцов и возвращает список
    предупреждений. Не останавливает работу сама — только предупреждает,
    кроме случая, когда столбец 'Название' совсем не удаётся определить."""
    warnings = []
    n = len(df)
    sample_n = min(n, 300)
    sample = df.iloc[:sample_n]

    def col_values(field):
        idx = mapping.get(field)
        if idx is None or idx >= len(df.columns):
            return None
        return sample.iloc[:, idx]

    # --- Название: обязательное поле ---
    title_vals = col_values("title")
    if title_vals is None:
        print("ОШИБКА — скрипт остановлен: не удалось определить столбец 'Название'.\n"
              "Проверьте структуру xlsx-файла (см. справку: python update_catalog_from_xlsx.py).")
        sys.exit(1)
    non_empty_titles = title_vals.dropna().astype(str).str.strip()
    non_empty_titles = non_empty_titles[non_empty_titles != ""]
    if len(non_empty_titles) == 0:
        print("ОШИБКА — скрипт остановлен: столбец 'Название' пуст.\n"
              "Проверьте, что в файле действительно есть названия книг в нужном столбце.")
        sys.exit(1)
    numeric_titles = non_empty_titles.str.match(r"^\d+([.,]\d+)?$")
    if numeric_titles.mean() > 0.5:
        warnings.append(
            "Столбец 'Название' в основном состоит из одних чисел — похоже, столбцы "
            "перепутаны местами. Проверьте результат перед публикацией."
        )

    # --- Год: должен содержать похожие на годы числа ---
    year_vals = col_values("year")
    if year_vals is not None:
        non_empty = year_vals.dropna().astype(str).str.strip()
        non_empty = non_empty[(non_empty != "") & (non_empty.str.lower() != "nan")]
        if len(non_empty) >= 5:
            looks_like_year = non_empty.apply(
                lambda s: bool(re.fullmatch(r"(19|20)\d{2}(\.0)?", s.strip()))
            )
            if looks_like_year.mean() < 0.5:
                warnings.append(
                    "Столбец 'Год' — меньше половины значений похожи на годы (например, 2020). "
                    "Проверьте, не перепутан ли этот столбец с другим."
                )

    # --- Экземпляров: должны быть небольшие числа ---
    copies_vals = col_values("copies")
    if copies_vals is not None:
        non_empty = copies_vals.dropna().astype(str).str.strip()
        non_empty = non_empty[(non_empty != "") & (non_empty.str.lower() != "nan")]
        if len(non_empty) >= 5:
            numeric = non_empty.apply(lambda s: bool(re.fullmatch(r"\d+(\.0)?", s.strip())))
            if numeric.mean() < 0.5:
                warnings.append(
                    "Столбец 'Экземпляров' — многие значения не похожи на количество "
                    "(целое число). Проверьте, не перепутан ли этот столбец с другим."
                )

    # --- Раздел: не должен быть уникальным для каждой строки (иначе это не раздел) ---
    section_vals = col_values("section")
    if section_vals is not None:
        non_empty = section_vals.dropna().astype(str).str.strip()
        non_empty = non_empty[non_empty != ""]
        if len(non_empty) >= 20:
            uniqueness = non_empty.nunique() / len(non_empty)
            if uniqueness > 0.9:
                warnings.append(
                    "Столбец 'Раздел' — почти все значения уникальны (как будто это не "
                    "название раздела, а что-то вроде названия книги или инвентарного номера). "
                    "Проверьте, не перепутан ли этот столбец с другим."
                )

    return warnings


def print_preview(records, limit=3):
    """Печатает несколько первых распознанных книг, чтобы можно было
    быстро глазами проверить, что данные легли в правильные поля."""
    print("\nПроверьте на глаз, что поля не перепутаны (первые записи):")
    for r in records[:limit]:
        print(f"  • «{r['title']}» — {r['author'] or '(автор не указан)'}, "
              f"{r['year'] or 'год не указан'}, раздел: {r['section']}, "
              f"инв. № {r['inventory'] or '—'}")
    print()


def clean_year(raw):
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return 0
    s = str(raw).strip()
    if not s or s.lower() == "nan":
        return 0
    m = re.search(r"\d{3,4}", s)
    return int(m.group()) if m else 0


def clean_copies(raw):
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return 1
    s = str(raw).strip()
    m = re.search(r"\d+", s)
    return int(m.group()) if m else 1


def clean_date(raw):
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
        if len(y) != 4:  # на случай формата ДД-ММ-ГГГГ
            d, m, y = parts
        return f"{y}-{m.zfill(2)}-{d.zfill(2)}"
    return s


def clean_str(raw):
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    return str(raw).strip()


def infer_type(section, title):
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


def make_auto_filename(xlsx_path):
    """Генерирует имя выходного файла на основе текущей даты и времени,
    например: catalog 2026 09 06 17 44.html — и кладёт его рядом с xlsx."""
    now = datetime.datetime.now()
    stamp = now.strftime("%Y %m %d %H %M")
    folder = os.path.dirname(os.path.abspath(xlsx_path))
    return os.path.join(folder, f"catalog {stamp}.html")


def main():
    if len(sys.argv) not in (3, 4):
        print(__doc__)
        sys.exit(1)

    xlsx_path, template_path = sys.argv[1], sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) == 4 else make_auto_filename(xlsx_path)

    # --- Защита №1: файлы вообще существуют? ---
    check_files_exist(xlsx_path, template_path)

    print(f"Читаю {xlsx_path} ...")
    try:
        raw = pd.read_excel(xlsx_path, sheet_name=0, header=None, dtype=object)
        with_header = pd.read_excel(xlsx_path, sheet_name=0, dtype=object)
    except Exception as e:
        print(f"ОШИБКА при чтении xlsx-файла: {e}\n"
              f"Проверьте, что файл не открыт в Excel в этот момент, и что это "
              f"настоящий .xlsx-файл, а не переименованный .csv.")
        sys.exit(1)

    if len(with_header) == 0 and len(raw) == 0:
        print(f"ОШИБКА: файл '{xlsx_path}' не содержит данных на первом листе.")
        sys.exit(1)

    mapping, has_header = detect_columns(with_header)
    df = with_header if has_header else raw
    if has_header:
        df = df.reset_index(drop=True)
    else:
        df = raw

    # --- Защита №2: похожи ли данные в столбцах на то, что там должно быть? ---
    warnings = validate_mapping(df, mapping, has_header)
    if warnings:
        print("\nВНИМАНИЕ — возможны проблемы с сопоставлением столбцов:")
        for w in warnings:
            print(" ⚠ " + w)
        print("Скрипт всё равно продолжит и создаст файл — но обязательно "
              "откройте его и проверьте карточки книг.\n")

    records = []
    seen_keys = set()
    skipped_empty = 0
    dupe_count = 0

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
        cover = clean_str(get("cover"))

        key = (title.lower(), author.lower(), inventory.lower())
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

    print(f"Готово: {len(records)} книг. Пропущено пустых строк: {skipped_empty}. Дубликатов убрано: {dupe_count}.")

    if len(records) == 0:
        print("ОШИБКА: после обработки не осталось ни одной книги. "
              "Файл не создан — проверьте исходный xlsx.")
        sys.exit(1)

    print_preview(records)

    # --- Защита №3: не даём случайно испортить сам шаблон ---
    if os.path.abspath(output_path) == os.path.abspath(template_path):
        print("ОШИБКА: имя итогового файла совпадает с именем шаблона "
              "(catalog_template.html). Это перезаписало бы сам шаблон — остановлено.\n"
              "Укажите другое имя для результата или не указывайте третий "
              "аргумент, чтобы имя подобралось автоматически.")
        sys.exit(1)

    try:
        with open(template_path, encoding="utf-8") as f:
            html = f.read()
    except Exception as e:
        print(f"ОШИБКА при чтении шаблона: {e}")
        sys.exit(1)

    if "__CATALOG_DATA__" not in html:
        print(
            "ОШИБКА: в файле шаблона не найден маркер __CATALOG_DATA__.\n"
            "Убедитесь, что вы указали правильный catalog_template.html "
            "и не редактировали его вручную."
        )
        sys.exit(1)

    data_json = json.dumps(records, ensure_ascii=False)
    html = html.replace("__CATALOG_DATA__", data_json)

    # Актуализируем число книг в мета-тегах (description / og / twitter),
    # чтобы превью ссылки в Telegram/WhatsApp/VK всегда показывало верную цифру.
    book_count = len(records)
    html = re.sub(r"\d+ книг[а]?,", f"{book_count} книг,", html)
    html = re.sub(r"\d+ книг[а]? с обложками", f"{book_count} книг с обложками", html)

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
    except Exception as e:
        print(f"ОШИБКА при сохранении файла: {e}")
        sys.exit(1)

    print(f"Сохранено: {output_path}")


if __name__ == "__main__":
    main()
