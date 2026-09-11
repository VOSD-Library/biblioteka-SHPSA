#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
publish.py
==========
Единый скрипт публикации каталога библиотеки ВОСД на GitHub Pages.

Что делает:
    1. Очищает старый publish_log.txt (чтобы лог не разрастался).
    2. Пересобирает index.html из Каталог.xlsx + catalog_template.html.
    3. Проверяет git status — есть ли изменения.
    4. Если есть — делает: git add . / git commit / git push.
    5. Пишет лог в publish_log.txt.

ИСПОЛЬЗОВАНИЕ:
    Двойной клик по publish.bat
    или вручную в CMD:  python publish.py
"""

import os
import sys
import subprocess
import datetime

# ===== Настройки =====
WORK_DIR = r"E:\BibCatalog"
XLSX_FILE = "Каталог.xlsx"
TEMPLATE_FILE = "catalog_template.html"
OUTPUT_FILE = "index.html"
GENERATOR = "update_catalog_from_xlsx.py"
LOG_FILE = "publish_log.txt"


def log(msg, also_print=True):
    """Пишет строку в лог и в консоль."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    if also_print:
        print(line)
    try:
        with open(os.path.join(WORK_DIR, LOG_FILE), "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        print(f"⚠ Не удалось записать в лог: {e}")


def run(cmd, capture=False):
    """Запускает команду, возвращает (код, stdout, stderr)."""
    try:
        result = subprocess.run(
            cmd,
            cwd=WORK_DIR,
            shell=True,
            capture_output=capture,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        return result.returncode, result.stdout or "", result.stderr or ""
    except Exception as e:
        return -1, "", str(e)


def main():
    # ===== Шаг 0: очистить старый лог =====
    log_path = os.path.join(WORK_DIR, LOG_FILE)
    try:
        if os.path.exists(log_path):
            os.remove(log_path)
    except Exception:
        pass

    log("=" * 60)
    log(f"Запуск публикации. Папка: {WORK_DIR}")

    if not os.path.isdir(WORK_DIR):
        log(f"✗ Папка не найдена: {WORK_DIR}")
        sys.exit(1)

    # ===== Шаг 1: пересборка index.html =====
    log("Шаг 1/3: пересборка index.html из Каталог.xlsx ...")
    code, out, err = run(
        f'python "{GENERATOR}" "{XLSX_FILE}" "{TEMPLATE_FILE}" "{OUTPUT_FILE}"',
        capture=True,
    )
    if out:
        for line in out.strip().splitlines():
            log("    " + line)
    if err:
        for line in err.strip().splitlines():
            log("    [stderr] " + line)
    if code != 0:
        log(f"✗ Ошибка генератора (код {code}). Публикация остановлена.")
        sys.exit(1)
    log("✓ index.html успешно пересобран.")

    # ===== Шаг 2: проверка изменений в git =====
    log("Шаг 2/3: проверяю изменения в git ...")
    code, out, err = run("git status --porcelain", capture=True)
    if code != 0:
        log("✗ Похоже, папка не является git-репозиторием.")
        log(f"   Вывод git: {err.strip()}")
        sys.exit(1)

    changes = out.strip()
    if not changes:
        log("ℹ Изменений нет — коммитить нечего. Публикация не требуется.")
        log("=" * 60)
        return

    log("Обнаружены изменения:")
    for line in changes.splitlines():
        log("    " + line)

    # ===== Шаг 3: git add / commit / push =====
    log("Шаг 3/3: git add / commit / push ...")

    run("git add .", capture=True)
    log("✓ git add выполнен.")

    commit_msg = f"Обновление каталога от {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
    code, out, err = run(f'git commit -m "{commit_msg}"', capture=True)
    if code != 0:
        log(f"✗ git commit не удался (код {code}): {err.strip()}")
        sys.exit(1)
    log(f"✓ git commit выполнен: {commit_msg}")

    code, out, err = run("git push", capture=True)
    if code != 0:
        log(f"✗ git push не удался (код {code}):")
        for line in (out + err).strip().splitlines():
            log("    " + line)
        log("   Проверьте: 1) есть ли интернет, 2) настроен ли remote origin, 3) есть ли права на push.")
        sys.exit(1)
    log("✓ git push выполнен успешно.")
   log("✓ Сайт обновится через 1–2 минуты: https://vosd-library.github.io/biblioteka-SHPSA/")
    log("=" * 60)


if __name__ == "__main__":
    main()