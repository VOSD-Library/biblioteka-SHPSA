#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
compress_covers.py
==================
Сжимает обложки в папке covers/ до MAX_WIDTH пикселей по ширине
и заданного качества JPEG.

ВАЖНО: скрипт пропускает файлы, которые уже сжаты (ширина ≤ MAX_WIDTH).
Это значит, что повторный запуск НЕ испортит уже сжатые обложки
и НЕ будет обрабатывать их заново.

ИСПОЛЬЗОВАНИЕ:
    python compress_covers.py
"""

import os
from PIL import Image

# ===== Настройки =====
COVERS_DIR = r"E:\BibCatalog\covers"
MAX_WIDTH = 800     # Максимальная ширина в пикселях
QUALITY = 75        # Качество JPEG (75-85 — хороший баланс)


def compress_images(directory):
    """Сжимает JPG-изображения в папке. Пропускает уже сжатые (ширина ≤ MAX_WIDTH)."""
    processed = 0
    skipped = 0
    errors = 0

    # Собираем все файлы рекурсивно
    files_to_process = []
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith(('.jpg', '.jpeg')):
                files_to_process.append(os.path.join(root, filename))

    total = len(files_to_process)
    print(f"Найдено JPG-файлов: {total}")
    print(f"Максимальная ширина: {MAX_WIDTH}px, качество: {QUALITY}")
    print("-" * 60)

    for i, filepath in enumerate(files_to_process, 1):
        filename = os.path.basename(filepath)
        try:
            with Image.open(filepath) as img:
                original_width = img.width
                original_size = os.path.getsize(filepath)

                # ===== ГЛАВНАЯ ПРОВЕРКА: файл уже сжат? =====
                if img.width <= MAX_WIDTH:
                    skipped += 1
                    print(f"[{i}/{total}] ⏭  Пропущено (уже сжато): {filename} "
                          f"({original_width}px, {original_size // 1024} КБ)")
                    continue

                # Конвертируем в RGB, если нужно
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')

                # Изменяем размер
                ratio = MAX_WIDTH / img.width
                new_height = int(img.height * ratio)
                img = img.resize((MAX_WIDTH, new_height), Image.LANCZOS)

                # Сохраняем, перезаписывая оригинал
                img.save(filepath, 'JPEG', quality=QUALITY, optimize=True)

                new_size = os.path.getsize(filepath)
                processed += 1
                print(f"[{i}/{total}] ✓ Сжато: {filename}  "
                      f"({original_width}px → {MAX_WIDTH}px, "
                      f"{original_size // 1024} КБ → {new_size // 1024} КБ)")
        except Exception as e:
            errors += 1
            print(f"[{i}/{total}] ✗ Ошибка при обработке {filename}: {e}")

    # Итоги
    print("-" * 60)
    print(f"Готово! Сжато: {processed}, пропущено (уже сжато): {skipped}, ошибок: {errors}")


if __name__ == "__main__":
    compress_images(COVERS_DIR)