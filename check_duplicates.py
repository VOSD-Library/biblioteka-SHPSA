#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Находит полные дубликаты (название + автор + инв. номер) в Каталог.xlsx."""
import pandas as pd

df = pd.read_excel('Каталог.xlsx', dtype=object)

# Ищем полные дубликаты
key_cols = []
for col in df.columns:
    c = str(col).lower()
    if 'назв' in c or 'автор' in c or 'инв' in c:
        key_cols.append(col)

print("Проверяю дубликаты по столбцам:", key_cols)
print()

dupes = df[df.duplicated(subset=key_cols, keep=False)]

if len(dupes) == 0:
    print("✓ Полных дубликатов нет")
else:
    print(f"⚠ Найдено {len(dupes)} строк с полными дубликатами:")
    for _, row in dupes.sort_values(key_cols).iterrows():
        print(f"  • {row.get('Название', '?')} / {row.get('Автор', '?')} / {row.get('Инв. номер', '?')}")